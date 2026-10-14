"""Generic consolidated-1099 text parser.

Works on ``pdftotext -layout`` output with the common layout:

* section headers naming the term and 8949 box ("Short-term ... Box A", "Unknown term"),
  repeated on every page;
* a security header line carrying a CUSIP (validated by its check digit), whose
  description may wrap onto following lines;
* one row per lot, columns in :attr:`GenericParser.columns` order, with any trailing text
  (e.g. "Original basis: 353.43", W/D codes) kept as notes;
* a summary page with one line per box ("Box A ... proceeds cost mkt wash gain") and
  1099-DIV/INT/OID/MISC box lines ("1a Total ordinary dividends ... 1,234.56").

Broker-specific layouts subclass this and change the class attributes / hooks. Any row
that cannot be mapped raises :class:`ParseError` rather than being guessed.
"""

from __future__ import annotations

import re
from decimal import Decimal
from typing import Optional

from ..io import parse_date
from ..models import (
    BOXES,
    BoxTotal,
    FormTotal,
    Lot,
    Package,
    Payer,
    to_decimal,
)
from .base import AMOUNT, EIN_RE, BrokerParser, ParseError, find_cusip, register

DATE_TOKEN = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
QTY_TOKEN = re.compile(r"^[\d,]*\d(\.\d+)?$")
AMOUNT_TOKEN = re.compile(rf"^{AMOUNT}$")
BLANK_TOKENS = {"...", "--", "-", "—"}
NA_TOKENS = {"N/A", "NA"}
CODE_SUFFIX = re.compile(r"^(.*\d)([A-Z]{1,2})$")  # "1,234.56W" -> amount + code

FORM_HEADERS = {
    "DIV": re.compile(r"1099-DIV", re.I),
    "INT": re.compile(r"1099-INT", re.I),
    "OID": re.compile(r"1099-OID", re.I),
    "MISC": re.compile(r"1099-MISC", re.I),
    "B": re.compile(r"1099-B\b", re.I),
    "DA": re.compile(r"1099-DA\b", re.I),
}


def clean_description(text: str) -> str:
    text = re.sub(r"Symbol\s*:?\s*[A-Z.]+", "", text, flags=re.I)
    text = re.sub(r"CUSIP\s*:?", "", text, flags=re.I)
    text = re.sub(r"\s*/\s*(?=/|$)", "", text)
    return re.sub(r"\s+", " ", text).strip(" /,|-")


@register
class GenericParser(BrokerParser):
    name = "generic"
    detect_keywords = ("1099-b", "1099-div", "1099-int", "1099-da")

    #: lot row column order; types: qty, date, amount, amount_or_na
    columns: tuple[tuple[str, str], ...] = (
        ("quantity", "qty"),
        ("date_acquired", "date"),
        ("date_sold", "date"),
        ("proceeds", "amount"),
        ("cost", "amount_or_na"),
        ("market_discount", "amount"),
        ("wash_sale", "amount"),
        ("broker_gain", "amount_or_na"),
    )
    #: summary line amount order, keyed by how many amounts the line has
    box_total_columns = {
        5: ("proceeds", "cost", "market_discount", "wash_sale", "gain"),
        4: ("proceeds", "cost", "wash_sale", "gain"),
    }
    payer_name: Optional[str] = None
    skip_line = re.compile(r"\b(sub)?totals?\b|^\s*(security|description|quantity|date)\b",
                           re.I)

    # -- section detection -----------------------------------------------------
    def section_box(self, line: str) -> Optional[str]:
        low = line.lower()
        if "unknown term" in low or "term unknown" in low:
            return "X"
        if not re.search(r"(short|long)[\s-]*term", low):
            return None
        m = re.search(r"\bbox\s+([a-hjkx])\b", line, re.I) or \
            re.search(r"\(([ABDEGHJK])\)", line)
        if m:
            return m.group(1).upper()
        covered = "noncovered" not in low and "not reported" not in low
        short = "short" in low
        digital = "digital" in low or "1099-da" in low
        if digital:
            return ("G" if covered else "H") if short else ("J" if covered else "K")
        return ("A" if covered else "B") if short else ("D" if covered else "E")

    # -- token helpers ----------------------------------------------------------
    @staticmethod
    def _split_code(token: str) -> tuple[str, str]:
        m = CODE_SUFFIX.match(token)
        if m and AMOUNT_TOKEN.match(m.group(1)):
            return m.group(1), m.group(2)
        return token, ""

    def _token_ok(self, token: str, kind: str) -> bool:
        if kind == "date":
            return bool(DATE_TOKEN.match(token)) or token.upper() in {"VARIOUS"} | NA_TOKENS
        if kind == "qty":
            return bool(QTY_TOKEN.match(token))
        if kind in ("amount", "amount_or_na"):
            if token in BLANK_TOKENS or AMOUNT_TOKEN.match(token):
                return True
            return kind == "amount_or_na" and token.upper() in NA_TOKENS
        return False

    def looks_like_lot(self, line: str) -> bool:
        tokens = line.split()
        if len(tokens) < len(self.columns):
            return False
        first_kind = self.columns[0][1]
        return self._token_ok(tokens[0], first_kind) and sum(
            1 for t in tokens if AMOUNT_TOKEN.match(self._split_code(t)[0])) >= 2

    def parse_lot(self, line: str, box: str, desc: str, cusip: str, page: int) -> Lot:
        tokens = line.split()
        values, codes = {}, []
        for i, (field, kind) in enumerate(self.columns):
            if i >= len(tokens):
                raise ParseError(f"p.{page}: too few columns in lot row: {line.strip()!r}")
            token, code = self._split_code(tokens[i])
            if code:
                codes.append(code)
            if not self._token_ok(token, kind):
                raise ParseError(f"p.{page}: column {field!r} got {token!r} in lot row "
                                 f"{line.strip()!r}; adjust the parser's column layout")
            values[field] = token
        notes = " ".join(codes + tokens[len(self.columns):])
        cost = to_decimal(values.get("cost"))
        return Lot(
            box=box,
            description=desc,
            cusip=cusip,
            quantity=to_decimal(values.get("quantity")),
            date_acquired=parse_date(values.get("date_acquired")),
            date_sold=parse_date(values.get("date_sold")),
            proceeds=to_decimal(values["proceeds"]) or Decimal("0"),
            cost=cost,
            market_discount=to_decimal(values.get("market_discount")) or Decimal("0"),
            wash_sale=to_decimal(values.get("wash_sale")) or Decimal("0"),
            broker_gain=to_decimal(values.get("broker_gain")),
            notes=notes,
            source_page=page,
        )

    # -- header -----------------------------------------------------------------
    def parse_header(self, pages: list[str], pkg: Package) -> None:
        head = "\n".join(pages[:3])
        m = re.search(r"(?:tax\s+year|for\s+(?:calendar\s+)?year)\s*:?\s*(20\d{2})", head, re.I) \
            or re.search(r"\b(20\d{2})\s+(?:form\s+)?1099", head, re.I)
        if m:
            pkg.tax_year = int(m.group(1))
        if re.search(r"\bCORRECTED\b(?!\s*\(if checked\))", head):
            pkg.corrected = True
            m = re.search(r"(no changes?[^\n]*)", head, re.I)
            if m:
                pkg.corrected_message = m.group(1).strip()
        m = re.search(r"RECIPIENT'?S?\s+(?:TIN|identification number)[^\n\d]*([\dX*-]{4,})",
                      head, re.I)
        if m:
            from ..models import mask_last4
            pkg.recipient_tin_last4 = mask_last4(m.group(1)) or ""
        account = ""
        m = re.search(r"Account\s*(?:No\.?|Number|#)?\s*:?\s*([X*\d-]*\d{4})\b", head, re.I)
        if m:
            account = m.group(1)
        m = re.search(r"(?:PAYER'?S?\s+(?:TIN|federal\s+id\w*(?:\s+number)?)|Federal\s+ID"
                      r"(?:\s+Number)?)[^\n\d]{0,30}(\d{2}-\d{7})", head, re.I)
        ein = m.group(1) if m else ""
        if not ein:
            m = EIN_RE.search(head)
            ein = m.group(1) if m else ""
        forms = [f"1099-{k}" for k, rx in FORM_HEADERS.items()
                 if rx.search("\n".join(pages))]
        pkg.forms_present = forms
        if ein:
            pkg.payers.append(Payer(name=self.payer_name or pkg.broker, ein=ein,
                                    account_last4=account, forms=forms))

    # -- summary totals -------------------------------------------------------------
    def parse_summary(self, pages: list[str], pkg: Package) -> None:
        amt = re.compile(AMOUNT + r"|\.\.\.|--")
        current_form = None
        seen_boxes, seen_income = set(), set()
        for pno, text in enumerate(pages, 1):
            for line in text.splitlines():
                for form, rx in FORM_HEADERS.items():
                    if rx.search(line) and form in ("DIV", "INT", "OID", "MISC"):
                        current_form = form
                    elif rx.search(line):
                        current_form = None
                # box total lines
                m = re.search(r"\bBox\s+([ABDEGHJKX])\b", line)
                if m and m.group(1) not in seen_boxes:
                    tail = line[m.end():]
                    amounts = amt.findall(tail)
                    layout = self.box_total_columns.get(len(amounts))
                    if layout:
                        vals = dict(zip(layout, (to_decimal(a) for a in amounts)))
                        pkg.box_totals.append(BoxTotal(
                            box=m.group(1), proceeds=vals["proceeds"] or Decimal("0"),
                            cost=vals.get("cost"),
                            market_discount=vals.get("market_discount") or Decimal("0"),
                            wash_sale=vals.get("wash_sale") or Decimal("0"),
                            gain=vals.get("gain"), source_page=pno))
                        seen_boxes.add(m.group(1))
                        continue
                # income box lines
                if current_form:
                    m = re.match(rf"^\s*(\d{{1,2}}[a-h]?)\s+[A-Za-z][^\n]*?\s({AMOUNT})\s*$",
                                 line)
                    if m and (current_form, m.group(1)) not in seen_income:
                        value = to_decimal(m.group(2))
                        seen_income.add((current_form, m.group(1)))
                        pkg.form_totals.append(FormTotal(current_form, m.group(1), value, pno))

    # -- lots ---------------------------------------------------------------------
    def parse_lots(self, pages: list[str], pkg: Package) -> None:
        box = None
        desc, cusip, collecting_desc = "", "", False
        for pno, text in enumerate(pages, 1):
            for line in text.splitlines():
                if not line.strip():
                    continue
                new_box = self.section_box(line)
                if new_box:
                    box, collecting_desc = new_box, False
                    continue
                if box is None:
                    continue
                if self.looks_like_lot(line):
                    if not cusip and not desc:
                        raise ParseError(f"p.{pno}: lot row before any security header: "
                                         f"{line.strip()!r}")
                    pkg.lots.append(self.parse_lot(line, box, desc, cusip, pno))
                    collecting_desc = False
                    continue
                if self.skip_line.search(line):
                    continue
                found = find_cusip(line)
                if found:
                    cusip = found
                    before = line[: line.find(found)]
                    desc = clean_description(before)
                    collecting_desc = True
                    continue
                if collecting_desc and not re.search(r"\d", line):
                    desc = clean_description(f"{desc} {line}")

    def parse(self, pages: list[str], source_file: str = "") -> Package:
        pkg = Package(client="", tax_year=0, broker=self.payer_name or "Unknown broker",
                      source_file=source_file)
        self.parse_header(pages, pkg)
        self.parse_summary(pages, pkg)
        self.parse_lots(pages, pkg)
        unknown = [b for b in {l.box for l in pkg.lots} if b not in BOXES]
        if unknown:
            raise ParseError(f"Unknown boxes {unknown}")
        return pkg
