"""Form 8949 entry lines, wash-sale notes and preparer flags."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Optional

from .models import (
    BOXES,
    DIGITAL_BOXES,
    NONCOVERED_BOXES,
    ZERO,
    IncomeLine,
    Lot,
    Package,
)

# Where each income box goes on the return. Software-neutral on purpose.
INCOME_BOXES: dict[tuple[str, str], tuple[str, str]] = {
    ("DIV", "1a"): ("Total ordinary dividends", "Sch B Part II / 1040 line 3b"),
    ("DIV", "1b"): ("Qualified dividends", "1040 line 3a"),
    ("DIV", "2a"): ("Total capital gain distributions", "Sch D line 13"),
    ("DIV", "2b"): ("Unrecaptured Sec. 1250 gain", "Unrecaptured 1250 gain worksheet"),
    ("DIV", "2c"): ("Section 1202 gain", "Form 8949 / Sch D (1202 exclusion)"),
    ("DIV", "2d"): ("Collectibles (28%) gain", "28% rate gain worksheet"),
    ("DIV", "2e"): ("Section 897 ordinary dividends", "Informational (foreign persons)"),
    ("DIV", "2f"): ("Section 897 capital gain", "Informational (foreign persons)"),
    ("DIV", "3"): ("Nondividend distributions", "Reduce basis; gain if basis exhausted"),
    ("DIV", "4"): ("Federal income tax withheld", "1040 line 25b"),
    ("DIV", "5"): ("Section 199A dividends", "Form 8995 / 8995-A (QBI REIT dividends)"),
    ("DIV", "6"): ("Investment expenses", "Not deductible (already in 1a)"),
    ("DIV", "7"): ("Foreign tax paid", "Sch 3 line 1 or Form 1116"),
    ("DIV", "9"): ("Cash liquidation distributions", "Basis recovery, then capital gain"),
    ("DIV", "10"): ("Noncash liquidation distributions", "Basis recovery, then capital gain"),
    ("DIV", "12"): ("Exempt-interest dividends", "1040 line 2a"),
    ("DIV", "13"): ("Specified private activity bond interest dividends",
                    "Form 6251 line 2g (AMT)"),
    ("DIV", "16"): ("State tax withheld", "State return"),
    ("INT", "1"): ("Interest income", "Sch B Part I / 1040 line 2b"),
    ("INT", "2"): ("Early withdrawal penalty", "Sch 1 line 18"),
    ("INT", "3"): ("Interest on U.S. savings bonds & Treasury obligations",
                   "Sch B / 1040 line 2b; state subtraction"),
    ("INT", "4"): ("Federal income tax withheld", "1040 line 25b"),
    ("INT", "5"): ("Investment expenses", "Not deductible"),
    ("INT", "6"): ("Foreign tax paid", "Sch 3 line 1 or Form 1116"),
    ("INT", "8"): ("Tax-exempt interest", "1040 line 2a"),
    ("INT", "9"): ("Specified private activity bond interest", "Form 6251 line 2g (AMT)"),
    ("INT", "10"): ("Market discount", "Sch B (if election to include currently)"),
    ("INT", "11"): ("Bond premium", "Sch B reduction (ABP adjustment)"),
    ("INT", "12"): ("Bond premium on Treasury obligations", "Reduces box 3 on Sch B"),
    ("INT", "13"): ("Bond premium on tax-exempt bond", "Reduces box 8 / line 2a"),
    ("INT", "17"): ("State tax withheld", "State return"),
    ("OID", "1"): ("Original issue discount", "Sch B (net of acquisition premium)"),
    ("OID", "2"): ("Other periodic interest", "Sch B"),
    ("OID", "3"): ("Early withdrawal penalty", "Sch 1 line 18"),
    ("OID", "4"): ("Federal income tax withheld", "1040 line 25b"),
    ("OID", "5"): ("Market discount", "Sch B (if election)"),
    ("OID", "6"): ("Acquisition premium", "Reduces OID on Sch B"),
    ("OID", "8"): ("OID on U.S. Treasury obligations", "Sch B; state subtraction"),
    ("OID", "9"): ("Investment expenses", "Not deductible"),
    ("OID", "10"): ("Bond premium", "Sch B reduction"),
    ("OID", "11"): ("Tax-exempt OID", "1040 line 2a"),
    ("MISC", "2"): ("Royalties", "Sch E"),
    ("MISC", "3"): ("Other income", "Sch 1 line 8z"),
    ("MISC", "4"): ("Federal income tax withheld", "1040 line 25b"),
    ("MISC", "8"): ("Substitute payments in lieu of dividends or interest",
                    "Sch 1 line 8z (ordinary, NOT qualified)"),
}

OTHER_CATEGORIES: dict[str, tuple[str, str]] = {
    "section_1256": ("Section 1256 contracts", "Form 6781"),
    "margin_interest": ("Margin interest paid", "Form 4952 (investment interest)"),
    "accrued_interest_paid": ("Accrued interest paid on purchases", "Sch B reduction"),
    "adr_fee": ("ADR fees", "Not deductible"),
    "investment_fee": ("Other investment fees", "Not deductible"),
    "partnership_distribution": ("Partnership (PTP) distributions", "Check for K-1"),
    "state_withholding": ("State tax withheld", "State return"),
    "other": ("Other", "Review"),
}


def income_label(form: str, box: str) -> tuple[str, str]:
    return INCOME_BOXES.get((form, box), (f"1099-{form} box {box}", "Review"))


@dataclass
class EntryLine:
    """One Form 8949 summary line (basis reported to IRS)."""

    box: str
    ordinary: bool
    proceeds: Decimal
    cost: Decimal
    wash: Decimal
    market_discount: Decimal
    lot_count: int

    @property
    def term_label(self) -> str:
        t = {"ST": "Short-term", "LT": "Long-term"}[BOXES[self.box].term]
        if BOXES[self.box].digital_asset:
            t = "Digital asset " + t.lower()
        return t + (" - Ordinary" if self.ordinary else "")

    @property
    def code(self) -> str:
        return "W" if self.wash else ""

    @property
    def gain(self) -> Decimal:
        return self.proceeds - self.cost + self.wash


@dataclass
class Flag:
    category: str
    issue: str
    amount: Optional[Decimal] = None
    status: str = "Open"


@dataclass
class Analysis:
    entry_lines: list[EntryLine]
    noncovered: dict[str, list[Lot]]
    wash_notes: dict[int, str]  # id(lot) -> note
    flags: list[Flag] = field(default_factory=list)
    qualified_div_at_risk: list[dict] = field(default_factory=list)
    digital_asset_question_yes: bool = False


def ensure_income_detail(pkg: Package) -> None:
    """Give every keyed box amount a detail line, so detail = box ties out by construction
    only when the broker printed no payment-level detail (labelled as such)."""
    have = {(l.form, l.box) for l in pkg.income_lines}
    for t in pkg.form_totals:
        if (t.form, t.box) not in have and t.amount:
            pkg.income_lines.append(IncomeLine(
                form=t.form, box=t.box, amount=t.amount,
                description="Per 1099 summary page (no payment-level detail)",
                source_page=t.source_page))


def _entry_lines(pkg: Package) -> list[EntryLine]:
    boxes = ["A", "D"]
    if pkg.digital_lots or any(t.box in DIGITAL_BOXES for t in pkg.box_totals):
        boxes += ["G", "J"]
    lines = []
    for box in boxes:
        for ordinary in (False, True):
            lots = [l for l in pkg.lots if l.box == box and l.ordinary == ordinary]
            if ordinary and not lots:
                continue
            lines.append(EntryLine(
                box=box, ordinary=ordinary,
                proceeds=sum((l.proceeds for l in lots), ZERO),
                cost=sum((l.cost or ZERO for l in lots), ZERO),
                wash=sum((l.wash_sale for l in lots), ZERO),
                market_discount=sum((l.market_discount for l in lots), ZERO),
                lot_count=len(lots)))
    return lines


def _wash_notes(pkg: Package) -> dict[int, str]:
    notes = {}
    sales_by_cusip = defaultdict(list)
    for lot in pkg.lots:
        sales_by_cusip[(lot.account, lot.cusip or lot.description)].append(lot)
    for lot in pkg.lots:
        if not lot.wash_sale:
            continue
        parts = []
        pre = lot.pre_wash_gain
        if pre is not None and (pre >= 0 or lot.wash_sale > -pre + Decimal("0.005")):
            parts.append("Disallowed amount exceeds this line's pre-adjustment loss: broker "
                         "aggregated several lots; reported as the broker computed.")
        if isinstance(lot.date_sold, date):
            same = sales_by_cusip[(lot.account, lot.cusip or lot.description)]
            later = [l for l in same if isinstance(l.date_sold, date)
                     and l.date_sold > lot.date_sold]
            if later:
                last_date = max(l.date_sold for l in later)
                last = [l for l in later if l.date_sold == last_date]
                if all(not l.wash_sale for l in last):
                    parts.append(f"Later sales of this security through {last_date:%m/%d/%Y} "
                                 "with no wash on the final sale: position likely closed and "
                                 "deferred loss realized this year (confirm no shares at 12/31).")
                else:
                    parts.append("Final sale in the year also carries a wash adjustment: "
                                 "replacement shares may be held at 12/31 (loss carries into "
                                 "next year's basis).")
            else:
                parts.append("No later sale of this security on the 1099: replacement shares "
                             "may be held at 12/31 (deferred loss carries into next year's "
                             "basis). Cannot be determined from the 1099; confirm with client.")
        if lot.is_digital:
            parts.append("Digital asset: non-security digital assets are not subject to wash "
                         "sale rules (confirm no law change for the year); spot crypto "
                         "ETFs/trusts are securities and are covered.")
        notes[id(lot)] = " ".join(parts)
    return notes


def _qualified_dividend_screen(pkg: Package) -> list[dict]:
    at_risk = []
    for div in pkg.income_lines:
        if div.form != "DIV" or div.box != "1b" or not div.cusip:
            continue
        if not isinstance(div.pay_date, date):
            continue
        for lot in pkg.lots:
            if lot.cusip != div.cusip:
                continue
            if not (isinstance(lot.date_acquired, date) and isinstance(lot.date_sold, date)):
                continue
            held = (lot.date_sold - lot.date_acquired).days
            if lot.date_acquired <= div.pay_date <= lot.date_sold and held <= 60:
                at_risk.append({
                    "security": div.description or lot.description,
                    "cusip": div.cusip,
                    "pay_date": div.pay_date,
                    "dividend": div.amount,
                    "lot_acquired": lot.date_acquired,
                    "lot_sold": lot.date_sold,
                    "days_held": held,
                })
    return at_risk


def analyze(pkg: Package) -> Analysis:
    flags: list[Flag] = []
    add = lambda cat, issue, amt=None: flags.append(Flag(cat, issue, amt))  # noqa: E731

    noncovered = {b: [l for l in pkg.lots if l.box == b] for b in NONCOVERED_BOXES}
    noncovered = {b: lots for b, lots in noncovered.items() if lots}

    # Package-level
    if pkg.corrected:
        msg = f' Broker message: "{pkg.corrected_message}".' if pkg.corrected_message else ""
        add("Corrected 1099", "Package is CORRECTED. If the return was already filed, "
            "determine whether an amendment is needed." + msg)
    if not pkg.payers:
        add("Payer", "No payer name / EIN recorded; key from the 1099 header.")
    for box in BOXES:
        if pkg.lots_in([box]) and pkg.box_total(box) is None:
            add("Tie-out", f"Box {box}: lots present but the broker's summary total was not "
                "keyed, so the box tie-out is incomplete.")

    # Wash sales
    wash_notes = _wash_notes(pkg)
    wash_total = sum((l.wash_sale for l in pkg.lots), ZERO)
    if wash_total:
        add("Wash sales", "Broker tracks only same-account, identical-CUSIP wash sales. Ask "
            "about purchases in IRAs (loss permanently disallowed), spouse's accounts, other "
            "brokers and options on the same securities.", wash_total)
        agg = [l for l in pkg.lots if l.wash_sale and "aggregated" in wash_notes.get(id(l), "")]
        if agg:
            add("Wash sales", f"{len(agg)} wash line(s) have a disallowed amount larger than "
                "the line's own loss (broker aggregation). Reported as the broker computed.",
                sum((l.wash_sale for l in agg), ZERO))

    # Not reported / unknown term
    unknown = noncovered.get("X", [])
    if unknown:
        add("Unknown term", f"{len(unknown)} unknown-term lot(s): holding period AND basis "
            "must come from the client before entry.", sum((l.proceeds for l in unknown), ZERO))
    for box in ("B", "E", "H", "K"):
        lots = noncovered.get(box, [])
        missing = [l for l in lots if l.cost is None]
        if missing:
            add("Noncovered basis", f"Box {box}: {len(missing)} lot(s) with no basis on the "
                "1099; client must supply basis and date acquired.",
                sum((l.proceeds for l in missing), ZERO))
        if lots:
            add("Noncovered", f"Box {box}: {len(lots)} lot(s) basis not reported to IRS: enter "
                "each lot individually and verify basis with the client.",
                sum((l.proceeds for l in lots), ZERO))
    if noncovered.get("H") or noncovered.get("K"):
        add("Digital assets", "Noncovered 1099-DA lots (Box H/K): basis comes from broker "
            "records only; request support.")

    # Digital assets
    da_yes = bool(pkg.digital_lots) or "DA" in pkg.forms_present or "1099-DA" in pkg.forms_present
    if da_yes:
        add("Digital assets", "Digital asset transactions present: answer YES to the Form "
            "1040 digital asset question.")

    # CVRs / zero basis
    for lot in pkg.lots:
        text = f"{lot.description} {lot.notes}".upper()
        if ("CVR" in text or "CONTINGENT VALUE" in text) and not lot.cost:
            add("Zero basis", f"{lot.description}: CVR with zero/no basis; confirm basis "
                "allocation with the client.", lot.proceeds)

    # Market discount
    md = sum((l.market_discount for l in pkg.lots), ZERO)
    if md:
        add("Market discount", "Accrued market discount on sales is ordinary income; confirm "
            "it is reported as ordinary, not capital gain.", md)

    # Income items
    totals = defaultdict(lambda: ZERO)
    for line in pkg.income_lines:
        totals[(line.form, line.box)] += line.amount
    liq = totals[("DIV", "9")] + totals[("DIV", "10")]
    if liq:
        add("Liquidation distributions", "1099-DIV box 9/10 are not dividends: they recover "
            "basis first, then become capital gain. Get basis for the liquidated security.", liq)
    if totals[("DIV", "3")]:
        add("Nondividend distributions", "1099-DIV box 3 reduces basis; confirm basis is not "
            "exhausted (excess is capital gain).", totals[("DIV", "3")])
    if totals[("MISC", "8")]:
        add("Substitute payments", "1099-MISC box 8 substitute payments are ordinary income, "
            "not qualified dividends.", totals[("MISC", "8")])
    if totals[("MISC", "3")]:
        add("Other income", "1099-MISC box 3 goes to Schedule 1 other income.",
            totals[("MISC", "3")])
    if totals[("OID", "6")] or totals[("OID", "10")] or totals[("INT", "11")]:
        add("OID / premium", "Take the taxable OID/interest after acquisition premium and "
            "bond premium adjustments.")
    if totals[("OID", "11")]:
        add("Tax-exempt OID", "Tax-exempt OID goes to 1040 line 2a.", totals[("OID", "11")])

    qd_risk = _qualified_dividend_screen(pkg)
    if qd_risk:
        secs = sorted({r["security"] for r in qd_risk})
        add("Qualified dividends", "Holding-period screen: lots held 60 days or less straddle "
            f"qualified dividend payments ({', '.join(secs)}). Ex-date is not on the 1099; "
            "ask whether other shares were held through year-end.",
            sum((r["dividend"] for r in qd_risk), ZERO))
    elif any(l.form == "DIV" and l.box == "1b" and not isinstance(l.pay_date, date)
             for l in pkg.income_lines):
        add("Qualified dividends", "Qualified dividends have no payment dates in the input, so "
            "the 60-day holding-period screen could not be run.")

    # Other items
    other = defaultdict(lambda: ZERO)
    for item in pkg.other_items:
        other[item.category] += item.amount
    if other["section_1256"]:
        add("Section 1256", "Section 1256 aggregate profit/(loss) goes to Form 6781 "
            "(60/40), not Form 8949.", other["section_1256"])
    if other["margin_interest"]:
        add("Margin interest", "Margin interest goes to Form 4952: check net investment "
            "income and ask whether the client elects to treat qualified dividends/gains as "
            "investment income.", other["margin_interest"])
    fees = other["adr_fee"] + other["investment_fee"]
    if fees:
        add("Fees", "ADR and other investment fees are not deductible.", fees)
    if other["partnership_distribution"]:
        add("Partnerships", "Publicly traded partnership distributions: request the K-1(s).",
            other["partnership_distribution"])

    state_amt = state_treasury_amount(pkg)
    if state_amt:
        add("State", "U.S. Treasury / direct federal obligation interest is a state "
            "subtraction in most states (including Colorado).", state_amt)

    add("Completeness", "Confirm there are no separate joint, retirement or other-broker "
        "1099s for this client.")

    return Analysis(
        entry_lines=_entry_lines(pkg),
        noncovered=noncovered,
        wash_notes=wash_notes,
        flags=flags,
        qualified_div_at_risk=qd_risk,
        digital_asset_question_yes=da_yes,
    )


def fund_ordinary_dividends(pkg: Package, cusip: str) -> Decimal:
    return sum((l.amount for l in pkg.income_lines
                if l.form == "DIV" and l.box == "1a" and l.cusip == cusip), ZERO)


def state_treasury_amount(pkg: Package) -> Decimal:
    total = ZERO
    for ob in pkg.state_obligations:
        base = ob.dividend_amount if ob.dividend_amount is not None else \
            fund_ordinary_dividends(pkg, ob.cusip)
        total += (base * ob.percent).quantize(Decimal("0.01"))
    total += sum((l.amount for l in pkg.income_lines if l.form == "INT" and l.box == "3"), ZERO)
    total += sum((l.amount for l in pkg.income_lines if l.form == "OID" and l.box == "8"), ZERO)
    return total
