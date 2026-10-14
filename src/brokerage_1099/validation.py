"""Pre-build tie-outs, computed in Python from the model (independent of Excel).

These are the "validate before building" checks: if any fails, the parser or keyed
input is wrong. Nothing here ever adjusts an input to force a tie.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional

from .models import BOXES, ZERO, Package

TOLERANCE = Decimal("0.005")


@dataclass
class CheckResult:
    name: str
    expected: Optional[Decimal]
    actual: Optional[Decimal]
    location: str = ""

    @property
    def diff(self) -> Decimal:
        return (self.actual or ZERO) - (self.expected or ZERO)

    @property
    def ok(self) -> bool:
        return abs(self.diff) < TOLERANCE


@dataclass
class ValidationReport:
    checks: list[CheckResult] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def failures(self) -> list[CheckResult]:
        return [c for c in self.checks if not c.ok]

    @property
    def ok(self) -> bool:
        return not self.failures

    def summary(self) -> str:
        lines = [f"{len(self.checks)} checks, {len(self.failures)} failed, "
                 f"{len(self.warnings)} warnings"]
        for c in self.failures:
            lines.append(f"  FAIL {c.name}: expected {c.expected} got {c.actual} "
                         f"(diff {c.diff}) {c.location}")
        for w in self.warnings:
            lines.append(f"  WARN {w}")
        return "\n".join(lines)


def _sum(values) -> Decimal:
    return sum((v for v in values if v is not None), ZERO)


def validate(pkg: Package) -> ValidationReport:
    report = ValidationReport()

    # 1. Every lot: proceeds - cost + wash = broker gain (N/A lots excluded)
    for i, lot in enumerate(pkg.lots, 1):
        if lot.cost is None or lot.broker_gain is None:
            continue
        report.checks.append(CheckResult(
            name=f"Lot {i} gain",
            expected=lot.broker_gain,
            actual=lot.calc_gain,
            location=f"{lot.description} {lot.cusip} sold {lot.date_sold} p.{lot.source_page}",
        ))

    # 2. Lot sums per box = broker box totals
    by_box = defaultdict(list)
    for lot in pkg.lots:
        by_box[lot.box].append(lot)
    for box in BOXES:
        lots = by_box.get(box, [])
        total = pkg.box_total(box)
        if total is None:
            if lots:
                report.warnings.append(
                    f"Box {box}: {len(lots)} lots but no broker summary total keyed; "
                    "box tie-out cannot be performed")
            continue
        page = f"summary p.{total.source_page}"
        report.checks.append(CheckResult(f"Box {box} proceeds", total.proceeds,
                                         _sum(l.proceeds for l in lots), page))
        if total.cost is not None:
            report.checks.append(CheckResult(f"Box {box} cost", total.cost,
                                             _sum(l.cost for l in lots), page))
        report.checks.append(CheckResult(f"Box {box} wash sale", total.wash_sale,
                                         _sum(l.wash_sale for l in lots), page))
        report.checks.append(CheckResult(f"Box {box} market discount", total.market_discount,
                                         _sum(l.market_discount for l in lots), page))
        if total.gain is not None:
            report.checks.append(CheckResult(f"Box {box} gain/(loss)", total.gain,
                                             _sum(l.broker_gain for l in lots), page))
            # Per-term: proceeds - cost + wash = broker-reported gain
            if total.cost is not None and BOXES[box].basis_reported:
                report.checks.append(CheckResult(
                    f"Box {box} proceeds - cost + wash = reported gain", total.gain,
                    total.proceeds - total.cost + total.wash_sale, page))

    # 3. Payment-level income = 1099 box amounts
    detail = defaultdict(lambda: ZERO)
    for line in pkg.income_lines:
        detail[(line.form, line.box)] += line.amount
    keyed = {(t.form, t.box): t for t in pkg.form_totals}
    for key in sorted(set(detail) | set(keyed)):
        form, box = key
        if key not in keyed:
            report.warnings.append(f"1099-{form} box {box}: detail lines present but no "
                                   "summary amount keyed")
            continue
        report.checks.append(CheckResult(
            f"1099-{form} box {box}", keyed[key].amount, detail.get(key, ZERO),
            f"summary p.{keyed[key].source_page}"))

    if not pkg.payers:
        report.warnings.append("No payer / EIN recorded")
    return report
