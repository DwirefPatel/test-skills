"""Canonical data model for a consolidated brokerage 1099 package.

Every parser (PDF, CSV, JSON, or an upstream system in a larger build) produces a
:class:`Package`. Everything downstream (validation, analysis, workbook) consumes only
this model, so new brokers plug in by writing a parser and nothing else.

Money is held as ``Decimal`` so tie-outs are exact to the cent.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Optional, Union

DateLike = Union[date, str, None]  # str is used for "VARIOUS" / "N/A"

ZERO = Decimal("0")


@dataclass(frozen=True)
class BoxInfo:
    code: str
    term: str  # "ST", "LT" or "UNKNOWN"
    basis_reported: bool
    digital_asset: bool
    label: str


# Form 8949 boxes. "X" is used for unknown-term lots (Box 2 "Unknown" on the 1099-B).
BOXES: dict[str, BoxInfo] = {
    "A": BoxInfo("A", "ST", True, False, "Short-term, basis reported to IRS"),
    "B": BoxInfo("B", "ST", False, False, "Short-term, basis NOT reported"),
    "D": BoxInfo("D", "LT", True, False, "Long-term, basis reported to IRS"),
    "E": BoxInfo("E", "LT", False, False, "Long-term, basis NOT reported"),
    "X": BoxInfo("X", "UNKNOWN", False, False, "Unknown term (term and basis from client)"),
    "G": BoxInfo("G", "ST", True, True, "Digital asset short-term, basis reported"),
    "H": BoxInfo("H", "ST", False, True, "Digital asset short-term, basis NOT reported"),
    "J": BoxInfo("J", "LT", True, True, "Digital asset long-term, basis reported"),
    "K": BoxInfo("K", "LT", False, True, "Digital asset long-term, basis NOT reported"),
}
SECURITY_BOXES = ["A", "B", "D", "E", "X"]
DIGITAL_BOXES = ["G", "H", "J", "K"]
REPORTED_BOXES = ["A", "D", "G", "J"]
NONCOVERED_BOXES = ["B", "E", "X", "H", "K"]

INCOME_FORMS = ["DIV", "INT", "OID", "MISC"]


def mask_last4(value: Optional[str]) -> Optional[str]:
    """Reduce a TIN or account number to its last 4 digits/characters."""
    if value is None:
        return None
    chars = re.sub(r"[^0-9A-Za-z]", "", str(value))
    if not chars:
        return None
    return "****" + chars[-4:]


def to_decimal(value) -> Optional[Decimal]:
    """Parse numbers as printed on 1099s: commas, $, (parentheses) negatives, '...' blanks."""
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(Decimal("0.01"))
    s = str(value).strip()
    if s.upper() in {"", "N/A", "NA"}:
        return None
    if s in {"...", "--", "-", "—"}:
        return ZERO
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").strip()
    if s.endswith("-"):  # some brokers print trailing minus
        neg, s = True, s[:-1]
    d = Decimal(s)
    return -d if neg else d


@dataclass
class Payer:
    name: str
    ein: str  # shown in full per workpaper standard
    account_last4: str
    forms: list[str] = field(default_factory=list)

    def __post_init__(self):
        self.account_last4 = mask_last4(self.account_last4) or ""


@dataclass
class Lot:
    box: str
    description: str
    cusip: str
    quantity: Optional[Decimal]
    date_acquired: DateLike
    date_sold: DateLike
    proceeds: Decimal
    cost: Optional[Decimal]  # None = N/A (not reported / unknown)
    wash_sale: Decimal = ZERO
    market_discount: Decimal = ZERO
    broker_gain: Optional[Decimal] = None  # None = N/A on the 1099
    ordinary: bool = False  # e.g. "Box A - Ordinary" sections
    account: str = ""
    notes: str = ""
    source_page: Optional[int] = None

    def __post_init__(self):
        self.box = self.box.upper()
        if self.box not in BOXES:
            raise ValueError(f"Unknown 8949 box {self.box!r} for {self.description}")

    @property
    def info(self) -> BoxInfo:
        return BOXES[self.box]

    @property
    def calc_gain(self) -> Optional[Decimal]:
        if self.cost is None:
            return None
        return self.proceeds - self.cost + self.wash_sale

    @property
    def pre_wash_gain(self) -> Optional[Decimal]:
        if self.cost is None:
            return None
        return self.proceeds - self.cost

    @property
    def is_digital(self) -> bool:
        return self.info.digital_asset


@dataclass
class BoxTotal:
    """Broker-reported totals for one 8949 box, keyed from the 1099-B summary page."""

    box: str
    proceeds: Decimal
    cost: Optional[Decimal]
    wash_sale: Decimal = ZERO
    market_discount: Decimal = ZERO
    gain: Optional[Decimal] = None
    source_page: Optional[int] = None


@dataclass
class IncomeLine:
    """A payment-level 1099-DIV/INT/OID/MISC line (or a summary-only line)."""

    form: str
    box: str
    amount: Decimal
    description: str = ""
    cusip: str = ""
    pay_date: DateLike = None
    source_page: Optional[int] = None

    def __post_init__(self):
        self.form = self.form.upper().replace("1099-", "")
        self.box = str(self.box).lower()


@dataclass
class FormTotal:
    """Box amount keyed from the 1099-DIV/INT/OID/MISC summary page."""

    form: str
    box: str
    amount: Decimal
    source_page: Optional[int] = None

    def __post_init__(self):
        self.form = self.form.upper().replace("1099-", "")
        self.box = str(self.box).lower()


@dataclass
class OtherItem:
    """Supplemental / non-reportable items and Section 1256 figures.

    category: margin_interest, accrued_interest_paid, adr_fee, investment_fee,
    partnership_distribution, section_1256, state_withholding, other.
    """

    category: str
    description: str
    amount: Decimal
    source_page: Optional[int] = None


@dataclass
class StateObligation:
    """U.S. Treasury / direct federal obligation % of a fund's dividends."""

    description: str
    cusip: str
    percent: Decimal  # stored as a fraction: 0.4321 = 43.21%
    dividend_amount: Optional[Decimal] = None  # None = use the fund's 1099-DIV 1a lines
    source_page: Optional[int] = None


@dataclass
class Package:
    client: str
    tax_year: int
    broker: str
    statement_date: DateLike = None
    corrected: bool = False
    corrected_message: str = ""
    prior_changes: list[str] = field(default_factory=list)
    recipient_tin_last4: str = ""
    payers: list[Payer] = field(default_factory=list)
    forms_present: list[str] = field(default_factory=list)
    lots: list[Lot] = field(default_factory=list)
    box_totals: list[BoxTotal] = field(default_factory=list)
    income_lines: list[IncomeLine] = field(default_factory=list)
    form_totals: list[FormTotal] = field(default_factory=list)
    other_items: list[OtherItem] = field(default_factory=list)
    state_obligations: list[StateObligation] = field(default_factory=list)
    source_file: str = ""

    def __post_init__(self):
        self.recipient_tin_last4 = mask_last4(self.recipient_tin_last4) or ""

    # convenience lookups -------------------------------------------------
    def box_total(self, box: str) -> Optional[BoxTotal]:
        return next((t for t in self.box_totals if t.box == box), None)

    def form_total(self, form: str, box: str) -> Optional[FormTotal]:
        return next(
            (t for t in self.form_totals if t.form == form and t.box == box.lower()), None
        )

    def lots_in(self, boxes) -> list[Lot]:
        return [lot for lot in self.lots if lot.box in boxes]

    @property
    def security_lots(self) -> list[Lot]:
        return self.lots_in(SECURITY_BOXES)

    @property
    def digital_lots(self) -> list[Lot]:
        return self.lots_in(DIGITAL_BOXES)

    @property
    def primary_payer_ein(self) -> str:
        return self.payers[0].ein if self.payers else ""

    def payer_ein_for_box(self, box: str) -> str:
        """EIN of the payer whose form covers this box (1099-DA for digital boxes)."""
        form = "1099-DA" if BOXES[box].digital_asset else "1099-B"
        for payer in self.payers:
            if any(f.upper().replace("FORM ", "") == form for f in payer.forms):
                return payer.ein
        return self.primary_payer_ein
