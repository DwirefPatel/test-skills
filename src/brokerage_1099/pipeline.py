"""High-level entry points: the API the larger build calls.

    from brokerage_1099 import process_package, process_pdf
    result = process_pdf("client.pdf", out_dir="out", overrides={...})
    result.workbook_path, result.ok, result.headline, result.to_dict()
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Optional

from .analysis import Analysis, analyze, ensure_income_detail, state_treasury_amount
from .io import merge_overrides
from .models import ZERO, Package
from .recalc import VerifyResult, recalc_and_verify
from .validation import ValidationReport, validate
from .workbook import BuildResult, build_workbook, default_filename


@dataclass
class RunResult:
    package: Package
    validation: ValidationReport
    analysis: Analysis
    build: BuildResult
    verify: Optional[VerifyResult] = None
    messages: list[str] = field(default_factory=list)

    @property
    def workbook_path(self) -> Path:
        return self.build.path

    @property
    def ok(self) -> bool:
        """True when the Python tie-outs pass and (if recalculated) every check cell is 0."""
        if not self.validation.ok:
            return False
        return self.verify.ok if self.verify is not None else True

    @property
    def headline(self) -> str:
        return headline(self)

    def to_dict(self) -> dict:
        p, an = self.package, self.analysis
        d = {
            "workbook": str(self.workbook_path),
            "ok": self.ok,
            "client": p.client,
            "tax_year": p.tax_year,
            "broker": p.broker,
            "corrected": p.corrected,
            "payers": [{"name": x.name, "ein": x.ein, "account": x.account_last4}
                       for x in p.payers],
            "entry_lines": [{
                "term": l.term_label, "box": l.box, "proceeds": str(l.proceeds),
                "cost": str(l.cost), "code": l.code, "wash": str(l.wash),
                "gain": str(l.gain), "lots": l.lot_count} for l in an.entry_lines],
            "noncovered": {b: {"lots": len(v),
                               "proceeds": str(sum((x.proceeds for x in v), ZERO))}
                           for b, v in an.noncovered.items()},
            "income": _income_totals(p),
            "state_treasury_subtraction": str(state_treasury_amount(p)),
            "flags": [{"category": f.category, "issue": f.issue,
                       "amount": None if f.amount is None else str(f.amount),
                       "status": f.status} for f in an.flags],
            "validation": {
                "checks": len(self.validation.checks),
                "failures": [{"name": c.name, "expected": str(c.expected),
                              "actual": str(c.actual), "diff": str(c.diff),
                              "location": c.location} for c in self.validation.failures],
                "warnings": self.validation.warnings,
            },
            "headline": self.headline,
        }
        if self.verify is not None:
            d["workbook_checks"] = {
                "recalculated": self.verify.recalc.ok,
                "recalc_error": self.verify.recalc.error,
                "formula_errors": self.verify.recalc.formula_errors,
                "check_cells": len(self.verify.checks),
                "failures": [{"label": c.label, "location": c.location, "value": c.value}
                             for c in self.verify.failures],
                "status": self.verify.status,
            }
        return d


def _income_totals(p: Package) -> dict:
    out: dict[str, Decimal] = {}
    for line in p.income_lines:
        key = f"{line.form} {line.box}"
        out[key] = out.get(key, ZERO) + line.amount
    return {k: str(v) for k, v in out.items() if v}


def _money(v: Decimal) -> str:
    return f"({abs(v):,.2f})" if v < 0 else f"{v:,.2f}"


def headline(result: RunResult) -> str:
    p, an = result.package, result.analysis
    parts = [f"{p.client} {p.tax_year} {p.broker}"
             + (" (CORRECTED)" if p.corrected else "")
             + (f", payer EIN {p.primary_payer_ein}" if p.payers else "") + "."]
    for l in an.entry_lines:
        if not l.lot_count:
            parts.append(f"{l.term_label} (Box {l.box}): none reported.")
            continue
        parts.append(f"{l.term_label} (Box {l.box}): proceeds {_money(l.proceeds)}, cost "
                     f"{_money(l.cost)}, wash {_money(l.wash)}"
                     f"{' (code W)' if l.wash else ''}, gain/(loss) {_money(l.gain)}.")
    if an.noncovered:
        nc = "; ".join(f"Box {b}: {len(v)} lot(s), proceeds "
                       f"{_money(sum((x.proceeds for x in v), ZERO))}"
                       for b, v in an.noncovered.items())
        parts.append(f"Enter individually: {nc}.")
    else:
        parts.append("No noncovered or unknown-term lots.")
    inc = _income_totals(p)
    if inc:
        parts.append("Income: " + ", ".join(f"{k} {_money(Decimal(v))}"
                                              for k, v in inc.items()) + ".")
    st = state_treasury_amount(p)
    if st:
        parts.append(f"State Treasury subtraction {_money(st)}.")
    v = result.verify
    if v is None:
        parts.append("Workbook not recalculated; Python tie-outs "
                     + ("all pass." if result.validation.ok else
                        f"FAILED ({len(result.validation.failures)})."))
    elif v.ok and result.validation.ok:
        parts.append(f"Tie-out: all {len(v.checks)} check cells = 0 (ALL TIED).")
    else:
        bad = [f"{c.location} = {c.value}" for c in v.failures]
        if v.recalc.error:
            bad.append(v.recalc.error)
        bad += v.recalc.formula_errors[:10]
        bad += [f"{c.name}: diff {c.diff}" for c in result.validation.failures]
        parts.append("Tie-out: CHECK - " + "; ".join(bad[:15]) + ".")
    open_flags = sum(1 for f in an.flags if f.status == "Open")
    parts.append(f"{open_flags} open flag(s).")
    return " ".join(parts)


def process_package(pkg: Package, out_dir="out", *, filename: str | None = None,
                    recalc: bool = True, timeout: int = 180) -> RunResult:
    """Validate, analyse, build the workbook and (optionally) recalc + verify it."""
    ensure_income_detail(pkg)
    report = validate(pkg)
    an = analyze(pkg)
    path = Path(out_dir) / (filename or default_filename(pkg))
    build = build_workbook(pkg, an, path)
    verify = recalc_and_verify(build, timeout) if recalc else None
    return RunResult(pkg, report, an, build, verify)


def process_pdf(pdf_path, out_dir="out", *, broker: str | None = None,
                overrides: dict | None = None, **kwargs) -> RunResult:
    """Parse a 1099 PDF, merge preparer-keyed overrides, then :func:`process_package`."""
    from .parsers import parse_pdf

    pkg = parse_pdf(pdf_path, broker=broker)
    if overrides:
        pkg = merge_overrides(pkg, overrides)
    return process_package(pkg, out_dir, **kwargs)
