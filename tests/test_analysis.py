from datetime import date
from decimal import Decimal

from brokerage_1099 import analyze
from brokerage_1099.models import IncomeLine, Lot, Package


def test_entry_lines(sample_pkg):
    an = analyze(sample_pkg)
    lines = {(l.box, l.ordinary): l for l in an.entry_lines}
    st = lines[("A", False)]
    assert (st.proceeds, st.cost, st.wash) == (Decimal("31200.00"), Decimal("32500.00"),
                                                Decimal("850.00"))
    assert st.gain == Decimal("-450.00")
    assert st.code == "W"
    assert lines[("A", True)].gain == Decimal("210.00")
    assert lines[("D", False)].gain == Decimal("13850.00")
    assert lines[("J", False)].lot_count == 0  # shown as "None reported"


def test_noncovered_grouping(sample_pkg):
    an = analyze(sample_pkg)
    assert set(an.noncovered) == {"B", "E", "X", "H"}


def test_flags_cover_skill_requirements(sample_pkg):
    an = analyze(sample_pkg)
    cats = {f.category for f in an.flags}
    for expected in ("Wash sales", "Unknown term", "Noncovered basis", "Digital assets",
                     "Zero basis", "Substitute payments", "Section 1256", "Margin interest",
                     "Fees", "State", "Completeness"):
        assert expected in cats, expected
    assert an.digital_asset_question_yes
    assert all(f.status == "Open" for f in an.flags)


def test_aggregated_wash_line_detected(sample_pkg):
    an = analyze(sample_pkg)
    tesla = next(l for l in sample_pkg.lots if l.description == "TESLA INC" and l.wash_sale)
    assert "aggregated" in an.wash_notes[id(tesla)]


def _pkg(lots, income=()):
    return Package(client="T", tax_year=2025, broker="B", lots=list(lots),
                   income_lines=list(income))


def test_wash_replacement_held_note():
    lot = Lot("A", "XYZ", "988498101", Decimal(1), date(2025, 1, 1), date(2025, 12, 1),
              Decimal(90), Decimal(100), wash_sale=Decimal(10), broker_gain=Decimal(0))
    an = analyze(_pkg([lot]))
    assert "replacement shares may be held at 12/31" in an.wash_notes[id(lot)]


def test_qualified_dividend_holding_period_screen():
    lot = Lot("A", "XYZ", "988498101", Decimal(10), date(2025, 3, 1), date(2025, 3, 20),
              Decimal(1000), Decimal(990), broker_gain=Decimal(10))
    div = IncomeLine("DIV", "1b", Decimal("12.00"), "XYZ", "988498101", date(2025, 3, 10))
    an = analyze(_pkg([lot], [div]))
    assert an.qualified_div_at_risk and an.qualified_div_at_risk[0]["days_held"] == 19
    assert any(f.category == "Qualified dividends" for f in an.flags)


def test_corrected_package_flags_amendment(sample_pkg):
    sample_pkg.corrected = True
    sample_pkg.corrected_message = "No changes from prior statement"
    an = analyze(sample_pkg)
    flag = next(f for f in an.flags if f.category == "Corrected 1099")
    assert "amendment" in flag.issue and "No changes" in flag.issue
