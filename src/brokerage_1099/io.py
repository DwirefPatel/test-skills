"""Load and save :class:`Package` objects as JSON, and import lots from CSV.

The JSON format is the integration contract for the larger build: any upstream system
(OCR, broker API, another tool) can hand over a package in this shape.
"""

from __future__ import annotations

import csv
import json
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from .models import (
    BoxTotal,
    DateLike,
    FormTotal,
    IncomeLine,
    Lot,
    OtherItem,
    Package,
    Payer,
    StateObligation,
    to_decimal,
)

_DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y")


def parse_date(value) -> DateLike:
    if value is None or isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return s.upper()  # "VARIOUS", "N/A", "INHERITED"...


def _bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "x"}


def _int(value):
    if value in (None, ""):
        return None
    return int(value)


def _dec(value, default=None):
    d = to_decimal(value)
    return default if d is None else d


def lot_from_dict(d: dict) -> Lot:
    return Lot(
        box=d["box"],
        description=d.get("description", ""),
        cusip=d.get("cusip", "") or "",
        quantity=_dec(d.get("quantity")),
        date_acquired=parse_date(d.get("date_acquired")),
        date_sold=parse_date(d.get("date_sold")),
        proceeds=_dec(d.get("proceeds"), Decimal("0")),
        cost=_dec(d.get("cost")),
        wash_sale=_dec(d.get("wash_sale"), Decimal("0")),
        market_discount=_dec(d.get("market_discount"), Decimal("0")),
        broker_gain=_dec(d.get("broker_gain")),
        ordinary=_bool(d.get("ordinary", False)),
        account=d.get("account", "") or "",
        notes=d.get("notes", "") or "",
        source_page=_int(d.get("source_page")),
    )


def package_from_dict(d: dict) -> Package:
    return Package(
        client=d["client"],
        tax_year=int(d["tax_year"]),
        broker=d["broker"],
        statement_date=parse_date(d.get("statement_date")),
        corrected=_bool(d.get("corrected", False)),
        corrected_message=d.get("corrected_message", ""),
        prior_changes=list(d.get("prior_changes", [])),
        recipient_tin_last4=d.get("recipient_tin_last4", d.get("recipient_tin", "")),
        payers=[
            Payer(
                name=p["name"],
                ein=p["ein"],
                account_last4=p.get("account_last4", p.get("account", "")),
                forms=list(p.get("forms", [])),
            )
            for p in d.get("payers", [])
        ],
        forms_present=list(d.get("forms_present", [])),
        lots=[lot_from_dict(x) for x in d.get("lots", [])],
        box_totals=[
            BoxTotal(
                box=x["box"].upper(),
                proceeds=_dec(x.get("proceeds"), Decimal("0")),
                cost=_dec(x.get("cost")),
                wash_sale=_dec(x.get("wash_sale"), Decimal("0")),
                market_discount=_dec(x.get("market_discount"), Decimal("0")),
                gain=_dec(x.get("gain")),
                source_page=_int(x.get("source_page")),
            )
            for x in d.get("box_totals", [])
        ],
        income_lines=[
            IncomeLine(
                form=x["form"],
                box=x["box"],
                amount=_dec(x["amount"]),
                description=x.get("description", ""),
                cusip=x.get("cusip", "") or "",
                pay_date=parse_date(x.get("pay_date")),
                source_page=_int(x.get("source_page")),
            )
            for x in d.get("income_lines", [])
        ],
        form_totals=[
            FormTotal(
                form=x["form"],
                box=x["box"],
                amount=_dec(x["amount"]),
                source_page=_int(x.get("source_page")),
            )
            for x in d.get("form_totals", [])
        ],
        other_items=[
            OtherItem(
                category=x["category"],
                description=x.get("description", ""),
                amount=_dec(x["amount"]),
                source_page=_int(x.get("source_page")),
            )
            for x in d.get("other_items", [])
        ],
        state_obligations=[
            StateObligation(
                description=x.get("description", ""),
                cusip=x.get("cusip", "") or "",
                percent=_dec(x["percent"]),
                dividend_amount=_dec(x.get("dividend_amount")),
                source_page=_int(x.get("source_page")),
            )
            for x in d.get("state_obligations", [])
        ],
        source_file=d.get("source_file", ""),
    )


def _jsonable(obj: Any):
    if isinstance(obj, Decimal):
        return str(obj)
    if isinstance(obj, date):
        return obj.isoformat()
    if is_dataclass(obj):
        return {f.name: _jsonable(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _jsonable(v) for k, v in obj.items()}
    return obj


def package_to_dict(pkg: Package) -> dict:
    return _jsonable(pkg)


def load_package(path) -> Package:
    with open(path, encoding="utf-8") as fh:
        return package_from_dict(json.load(fh))


def save_package(pkg: Package, path) -> Path:
    path = Path(path)
    path.write_text(json.dumps(package_to_dict(pkg), indent=2), encoding="utf-8")
    return path


def merge_overrides(pkg: Package, overrides: dict) -> Package:
    """Overlay preparer-keyed items (summary-page totals, payers, income) onto a parsed package.

    Scalar header fields replace; ``box_totals`` / ``form_totals`` replace by key; list
    fields such as ``income_lines`` and ``other_items`` are appended.
    """
    base = package_to_dict(pkg)
    for key in ("client", "tax_year", "broker", "statement_date", "corrected",
                "corrected_message", "recipient_tin_last4", "forms_present", "prior_changes"):
        if key in overrides:
            base[key] = overrides[key]
    if "payers" in overrides:
        base["payers"] = overrides["payers"]
    for key, keyfn in (
        ("box_totals", lambda x: x["box"].upper()),
        ("form_totals", lambda x: (x["form"].upper().replace("1099-", ""), str(x["box"]).lower())),
    ):
        if key in overrides:
            merged = {keyfn(x): x for x in base.get(key, [])}
            merged.update({keyfn(x): x for x in overrides[key]})
            base[key] = list(merged.values())
    for key in ("lots", "income_lines", "other_items", "state_obligations"):
        if key in overrides:
            base[key] = base.get(key, []) + overrides[key]
    return package_from_dict(base)


CSV_LOT_COLUMNS = [
    "box", "description", "cusip", "quantity", "date_acquired", "date_sold", "proceeds",
    "cost", "market_discount", "wash_sale", "broker_gain", "ordinary", "account", "notes",
    "source_page",
]


def load_lots_csv(path) -> list[Lot]:
    """Read lots from a CSV with the canonical column names (see ``CSV_LOT_COLUMNS``)."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        missing = {"box", "proceeds"} - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"{path}: CSV is missing required columns {sorted(missing)}")
        return [lot_from_dict({k: (v if v != "" else None) for k, v in row.items()})
                for row in reader]


__all__ = [
    "load_package", "save_package", "package_from_dict", "package_to_dict",
    "merge_overrides", "load_lots_csv", "parse_date", "CSV_LOT_COLUMNS",
]
