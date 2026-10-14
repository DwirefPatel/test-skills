"""Command line interface.

    brokerage-1099 build package.json -o out/            # JSON package -> workbook
    brokerage-1099 parse statement.pdf -o package.json   # PDF -> JSON package (review it)
    brokerage-1099 run statement.pdf --overrides keyed.json -o out/
    brokerage-1099 validate package.json                 # Python tie-outs only

Exit codes: 0 = all tied, 2 = built but a tie-out failed, 1 = error.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .io import load_lots_csv, load_package, merge_overrides, save_package
from .parsers import ParseError, parse_pdf, registered
from .pipeline import RunResult, process_package
from .validation import validate


def _print_result(result: RunResult, report_path: str | None) -> int:
    print(f"Workbook: {result.workbook_path}")
    print(result.headline)
    if not result.validation.ok or result.validation.warnings:
        print(result.validation.summary())
    if report_path:
        Path(report_path).write_text(json.dumps(result.to_dict(), indent=2, default=str))
        print(f"Report: {report_path}")
    return 0 if result.ok else 2


def _load_with_extras(args):
    pkg = load_package(args.package)
    if getattr(args, "lots_csv", None):
        pkg.lots.extend(load_lots_csv(args.lots_csv))
    return pkg


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="brokerage-1099",
                                 description="Brokerage consolidated 1099 -> Excel workpaper")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_build_opts(p):
        p.add_argument("-o", "--out-dir", default="out")
        p.add_argument("--filename", help="override <Client>_<Year>_<Broker>_1099_Summary.xlsx")
        p.add_argument("--no-recalc", action="store_true",
                       help="skip LibreOffice recalculation/read-back")
        p.add_argument("--timeout", type=int, default=180)
        p.add_argument("--report", help="write a JSON run report here")

    b = sub.add_parser("build", help="build the workbook from a JSON package")
    b.add_argument("package")
    b.add_argument("--lots-csv", help="append lots from a CSV with canonical columns")
    add_build_opts(b)

    v = sub.add_parser("validate", help="run the Python tie-outs only")
    v.add_argument("package")
    v.add_argument("--lots-csv")

    p = sub.add_parser("parse", help="parse a 1099 PDF into a JSON package for review")
    p.add_argument("pdf")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--broker", help="parser name (auto-detect if omitted)")
    p.add_argument("--overrides", help="JSON of preparer-keyed fields to merge")

    r = sub.add_parser("run", help="parse a PDF and build the workbook in one step")
    r.add_argument("pdf")
    r.add_argument("--broker")
    r.add_argument("--overrides")
    r.add_argument("--save-package", help="also write the parsed JSON package")
    add_build_opts(r)

    sub.add_parser("parsers", help="list registered broker parsers")

    args = ap.parse_args(argv)
    try:
        if args.cmd == "parsers":
            for cls in registered():
                print(f"{cls.name}: detects {', '.join(cls.detect_keywords)}")
            return 0
        if args.cmd == "validate":
            report = validate(_load_with_extras(args))
            print(report.summary())
            return 0 if report.ok else 2
        if args.cmd == "build":
            pkg = _load_with_extras(args)
        else:
            pkg = parse_pdf(args.pdf, broker=args.broker)
            if args.overrides:
                pkg = merge_overrides(pkg, json.loads(Path(args.overrides).read_text()))
            if args.cmd == "parse":
                save_package(pkg, args.output)
                print(f"Parsed {len(pkg.lots)} lots, {len(pkg.box_totals)} box totals, "
                      f"{len(pkg.form_totals)} income boxes -> {args.output}")
                print(validate(pkg).summary())
                return 0
            if args.save_package:
                save_package(pkg, args.save_package)
        if not pkg.client:
            print("Package has no client name; set it in the JSON or --overrides.",
                  file=sys.stderr)
            return 1
        result = process_package(pkg, args.out_dir, filename=args.filename,
                                 recalc=not args.no_recalc, timeout=args.timeout)
        return _print_result(result, args.report)
    except (ParseError, ValueError, KeyError, FileNotFoundError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())


