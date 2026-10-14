# brokerage-1099

Python implementation of the **Brokerage 1099 Summary** workpaper skill. It turns a
client's consolidated brokerage 1099 (1099-B, 1099-DA, DIV, INT, OID, MISC, Section
1256, supplemental pages) into a tied-out Excel preparer workpaper, plus a one-paragraph
headline and a machine-readable JSON report.

It is built as a self-contained package with a stable API so it can be one module of a
larger tax-prep build.

> Preparer workpaper, not tax advice. Recipient TIN and account numbers are kept as last 4
> only; payer EINs are shown in full.

## What it produces

`<Client>_<Year>_<Broker>_1099_Summary.xlsx` with these tabs:

| Tab | Contents |
|---|---|
| **Summary** | Header and payer/EIN block; Form 8949 one-line entries (A/D, plus G/J for digital assets, plus "Ordinary" lines when non-zero); per-term tie-out (proceeds - cost + wash = broker gain, Diff = 0); wash-sale detail by security; orange **enter individually** block for B/E/X/H/K; income headline items with destinations; all check cells with an overall `ALL TIED` / `CHECK` status; color legend |
| **1099-B Lots** | One row per lot, keyed values in blue with source page, calc gain and Diff per lot, totals by box (SUMIFS), keyed broker summary, Diff block, 8949 entry lines |
| **By Security** | Per-CUSIP lot count, proceeds, basis, wash and net, with a check against the Lots total |
| **Wash Sales** | Every lot with a non-zero 1g, keyed off the amount rather than the W code. Includes per-line notes (aggregation; position closed vs. replacement possibly held at 12/31), by-security subtotals, and per-box detail vs. broker checks |
| **Noncovered Lots** | Orange tab, only when B/E/X/H/K exist. Has yellow client-supplied date/term/basis columns and calculated gain once they are filled in |
| **Crypto 1099-DA** | Only when digital assets exist. Same lot layout, plus by-asset subtotals |
| **Income Detail** | Payment-level DIV/INT/OID/MISC lines, box totals vs. 1099 box amounts (Diff = 0), supplemental items, and the state Treasury-obligation subtraction |
| **Flags & Questions** | Numbered open issues with amounts (wash-sale questions, unknown term, missing basis, qualified-dividend 60-day screen, liquidation distributions, CVRs, 1256, margin interest/4952, fees, OID, state, CORRECTED/amendment, completeness) |

Formatting follows the skill: Arial throughout. Blue = keyed, black = formula, green = cross-tab link, green fill = check, orange = enter individually, yellow = client must supply. Numbers are `#,##0.00` with negatives in parentheses. Every check is wrapped in `ROUND(...,2)`. Gridlines are off, header rows frozen, and lot tabs have autofilter.

## Tie-outs (never plugged)

1. **Before building**, `validation.py` checks the model in Python:
   - every lot: proceeds - cost + wash = broker gain/(loss)
   - each box's lots sum to the broker summary
   - per term: proceeds - cost + wash = reported gain
   - income detail = 1099 box amounts
2. **After building**, the workbook is recalculated with headless LibreOffice. Every registered check cell is read back with `data_only=True` and must be within 0.005 of 0, with zero formula errors.

If either step fails, the run exits with code 2 and names the failing cell or check. Inputs are never adjusted to force a tie.

## Install

```bash
pip install -e .            # requires Python 3.10+, openpyxl
pip install -e '.[dev]'     # + pytest
```

Optional system tools: `pdftotext` (poppler-utils) for PDF input, and LibreOffice
(`soffice`) for recalculation and verification. Without LibreOffice, use `--no-recalc`; the
Python tie-outs still run, and Excel calculates the formulas when the file is opened.

## CLI

```bash
# JSON package -> workbook (+ JSON run report)
brokerage-1099 build examples/sample_package.json -o out --report out/report.json

# PDF -> JSON package for review (then fix/key anything missing and run build)
brokerage-1099 parse statement.pdf -o package.json --overrides keyed.json

# PDF -> workbook in one step
brokerage-1099 run statement.pdf --overrides keyed.json -o out --save-package pkg.json

brokerage-1099 validate package.json      # Python tie-outs only
brokerage-1099 parsers                    # list registered broker parsers
```

Exit codes: `0` = all tied, `2` = built but a tie-out failed (details printed), `1` = error.

`--overrides` is a small JSON of preparer-keyed items: client name, payers/EINs, and
summary-page box totals. Header fields replace, `box_totals`/`form_totals` replace by key, and lists append.

## Python API (for the larger build)

```python
from brokerage_1099 import load_package, package_from_dict, process_package, process_pdf

result = process_package(package_from_dict(data), out_dir="out")   # or load_package(path)
# result = process_pdf("statement.pdf", out_dir="out", overrides={"client": "Jane Doe"})

result.ok               # True only if every tie-out passes
result.workbook_path    # Path to the .xlsx
result.headline         # chat/summary paragraph
result.to_dict()        # JSON-safe: entry lines, noncovered, income, flags, check results
```

The lower-level pieces can also be used on their own: `validate(pkg)`, `analyze(pkg)`,
`workbook.build_workbook(pkg, analysis, path)` and `recalc.recalc_and_verify(build)`.

### Input contract

`examples/sample_package.json` is the reference. Money values are strings or numbers,
dates are ISO strings or `"VARIOUS"`, and a `null` cost or gain means N/A on the 1099.

| Key | Meaning |
|---|---|
| `client`, `tax_year`, `broker`, `statement_date`, `corrected`, `corrected_message`, `prior_changes`, `recipient_tin_last4` | Header (TIN is masked automatically) |
| `payers[]` | `name`, `ein`, `account_last4` (masked), `forms` |
| `lots[]` | `box` (A B D E X G H J K, where X = unknown term), `description`, `cusip`, `quantity`, `date_acquired`, `date_sold`, `proceeds`, `cost`, `wash_sale`, `market_discount`, `broker_gain`, `ordinary`, `account`, `notes`, `source_page` |
| `box_totals[]` | Broker 1099-B summary per box: `proceeds`, `cost`, `wash_sale`, `market_discount`, `gain`, `source_page` |
| `income_lines[]` | Payment-level `form` (DIV/INT/OID/MISC), `box`, `amount`, `description`, `cusip`, `pay_date` |
| `form_totals[]` | 1099 box amounts from the summary page |
| `other_items[]` | `category`: `margin_interest`, `accrued_interest_paid`, `adr_fee`, `investment_fee`, `partnership_distribution`, `section_1256`, `state_withholding`, `other` |
| `state_obligations[]` | Fund Treasury %: `cusip`, `percent` (fraction), optional `dividend_amount` |

Lots can also be loaded from CSV with the same column names (`--lots-csv`, `load_lots_csv`).

## Adding a broker parser

The generic parser (`parsers/generic.py`) reads `pdftotext -layout` output. It works as follows:
- It tracks the current box from section headers that repeat on every page.
- It groups lots by CUSIP, detecting CUSIPs by their check digit, and handles descriptions that wrap.
- It keeps trailing text such as "Original basis" or W/D codes as notes.
- It reads summary-page box totals and DIV/INT/OID/MISC box lines.
- When a row doesn't fit the column layout, it raises `ParseError` instead of guessing.

Real broker layouts differ (column order, unknown-term sections with N/A columns, and so on). Each broker gets a small subclass:

```python
from brokerage_1099.parsers import register
from brokerage_1099.parsers.generic import GenericParser

@register
class SchwabParser(GenericParser):
    name = "schwab"
    detect_keywords = ("charles schwab",)
    payer_name = "Charles Schwab & Co., Inc."
    columns = (("quantity", "qty"), ("date_acquired", "date"), ("date_sold", "date"),
               ("proceeds", "amount"), ("cost", "amount_or_na"), ("wash_sale", "amount"),
               ("broker_gain", "amount_or_na"))
```

Import the module in `parsers/__init__.py`. Then confirm against a real statement that `validate()` passes: lot sums per box must equal the broker's totals to the cent.

## Layout

```
src/brokerage_1099/
  models.py       canonical data model (Decimal money, masking, 8949 box metadata)
  io.py           JSON/CSV load-save, overrides merge
  validation.py   Python tie-outs (pre-build)
  analysis.py     8949 entry lines, wash notes, flags, income destinations
  workbook.py     openpyxl workbook builder (all tabs, links, check registry)
  recalc.py       LibreOffice recalculation + check-cell read-back
  pipeline.py     process_package / process_pdf, headline, JSON report
  cli.py          command line
  parsers/        pdftotext extraction, registry, generic layout parser
examples/         sample package and its generated workbook/report
tests/            pytest suite (LibreOffice tests auto-skip if soffice is missing)
```

## Not included yet

- Broker-specific parsers (Wells Fargo, Robinhood, Schwab, Fidelity, Webull, E*Trade,
  Coinbase). Each one needs real sample statements to tune its layout.
- Payment-level income detail from PDFs. The generic parser reads summary box amounts only, and payment detail can be supplied in JSON/CSV.
- Tax-software screen names. Destinations are given as form/line references.
