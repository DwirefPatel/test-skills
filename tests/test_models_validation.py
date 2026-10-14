from decimal import Decimal

from brokerage_1099 import package_from_dict, package_to_dict, validate
from brokerage_1099.analysis import ensure_income_detail
from brokerage_1099.io import merge_overrides
from brokerage_1099.models import mask_last4, to_decimal


def test_to_decimal_handles_1099_number_styles():
    assert to_decimal("(1,234.56)") == Decimal("-1234.56")
    assert to_decimal("$1,000.00") == Decimal("1000.00")
    assert to_decimal("12.50-") == Decimal("-12.50")
    assert to_decimal("...") == Decimal("0")
    assert to_decimal("N/A") is None
    assert to_decimal(None) is None


def test_mask_last4():
    assert mask_last4("123-45-6789") == "****6789"
    assert mask_last4("ACCT-00005678") == "****5678"
    assert mask_last4("") is None


def test_sample_package_masks_tin_and_accounts(sample_pkg):
    assert sample_pkg.recipient_tin_last4 == "****6789"
    assert all(p.account_last4.startswith("****") for p in sample_pkg.payers)
    assert sample_pkg.payers[0].ein == "12-3456789"  # payer EIN shown in full


def test_json_round_trip(sample_pkg):
    again = package_from_dict(package_to_dict(sample_pkg))
    assert len(again.lots) == len(sample_pkg.lots)
    assert again.lots[0].proceeds == sample_pkg.lots[0].proceeds


def test_sample_validates(sample_pkg):
    ensure_income_detail(sample_pkg)
    report = validate(sample_pkg)
    assert report.ok, report.summary()
    assert len(report.checks) > 20


def test_lot_mismatch_is_reported_not_plugged(sample_dict):
    sample_dict["lots"][0]["broker_gain"] = "-700.00"  # should be -750.00
    pkg = package_from_dict(sample_dict)
    report = validate(pkg)
    names = {c.name for c in report.failures}
    assert "Lot 1 gain" in names
    assert "Box A gain/(loss)" in names
    assert pkg.lots[0].broker_gain == Decimal("-700.00")  # input untouched


def test_box_total_mismatch(sample_dict):
    for t in sample_dict["box_totals"]:
        if t["box"] == "D":
            t["proceeds"] = "27400.01"
    pkg = package_from_dict(sample_dict)
    ensure_income_detail(pkg)
    report = validate(pkg)
    assert [c.name for c in report.failures] == [
        "Box D proceeds", "Box D proceeds - cost + wash = reported gain"]


def test_missing_box_total_is_a_warning(sample_dict):
    sample_dict["box_totals"] = [t for t in sample_dict["box_totals"] if t["box"] != "E"]
    report = validate(package_from_dict(sample_dict))
    assert any("Box E" in w for w in report.warnings)


def test_merge_overrides_replaces_box_totals_and_header(sample_pkg):
    merged = merge_overrides(sample_pkg, {
        "client": "Other Client",
        "box_totals": [{"box": "A", "proceeds": "1", "cost": "1", "gain": "0"}],
    })
    assert merged.client == "Other Client"
    assert merged.box_total("A").proceeds == Decimal("1")
    assert merged.box_total("D") is not None
