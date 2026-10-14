from decimal import Decimal

import pytest

from brokerage_1099.parsers import ParseError, cusip_valid, find_cusip, parse_text
from brokerage_1099.validation import validate

PAGE1 = """
EXAMPLE BROKERAGE LLC                              2025 FORM 1099 CONSOLIDATED
PAYER'S TIN: 12-3456789                            RECIPIENT'S TIN: ***-**-6789
Account Number: XXXX-5678
                    Summary of Proceeds, Gains & Losses
Short-term transactions Box A     15,150.00    16,400.00     0.00    500.00     (750.00)
Long-term transactions Box D      20,100.00     5,200.00     0.00      0.00    14,900.00

1099-DIV   Dividends and Distributions
1a Total ordinary dividends ....................          62.40
1b Qualified dividends .........................          62.40
1099-INT   Interest Income
1 Interest income ..............................          88.12
"""

PAGE2 = """
1099-B  Proceeds From Broker and Barter Exchange Transactions
SHORT TERM TRANSACTIONS FOR COVERED TAX LOTS - Box A
Quantity  Acquired  Sold  Proceeds  Cost  Mkt disc  Wash  Gain/loss
APPLE INC / CUSIP: 037833100 / Symbol: AAPL
  50.000   01/10/25   03/03/25   11,250.00   12,000.00   ...   ...       (750.00)
  20.000   03/10/25   04/15/25    3,900.00    4,400.00   ...   500.00W     0.00
  Security total:                15,150.00   16,400.00         500.00    (750.00)
LONG TERM TRANSACTIONS FOR COVERED TAX LOTS - Box D
MICROSOFT CORP / CUSIP: 594918104
  / Symbol: MSFT
  40.000   06/03/19   07/14/25   20,100.00    5,200.00   ...   ...      14,900.00   Original basis: 5,200.00
"""


def test_cusip_check_digit():
    assert cusip_valid("037833100")
    assert not cusip_valid("037833101")
    assert find_cusip("APPLE INC 037833100 AAPL") == "037833100"
    assert find_cusip("Account 123456789") is None or cusip_valid("123456789")


def test_generic_parser_end_to_end():
    pkg = parse_text([PAGE1, PAGE2], source_file="x.pdf")
    assert pkg.tax_year == 2025
    assert pkg.recipient_tin_last4 == "****6789"
    assert pkg.payers[0].ein == "12-3456789"
    assert pkg.payers[0].account_last4 == "****5678"
    assert len(pkg.lots) == 3
    a1, a2, d1 = pkg.lots
    assert (a1.box, a1.cusip, a1.description) == ("A", "037833100", "APPLE INC")
    assert a1.broker_gain == Decimal("-750.00")
    assert a2.wash_sale == Decimal("500.00") and "W" in a2.notes
    assert d1.box == "D" and d1.description == "MICROSOFT CORP"
    assert "Original basis" in d1.notes
    assert d1.source_page == 2
    assert {t.box for t in pkg.box_totals} == {"A", "D"}
    assert pkg.form_total("DIV", "1a").amount == Decimal("62.40")
    assert pkg.form_total("INT", "1").amount == Decimal("88.12")
    report = validate(pkg)
    assert not [c for c in report.failures if c.name.startswith(("Lot", "Box"))], \
        report.summary()


def test_misaligned_row_raises_instead_of_guessing():
    bad = PAGE2.replace("3,900.00    4,400.00", "3,900.00    04/01/25")
    with pytest.raises(ParseError):
        parse_text([PAGE1, bad])
