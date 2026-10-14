"""Form 8949 / 1099-B definition (Drake screen 8949).

One row = one Drake 8949 instance:
  * basis-reported boxes (A, D, G, J): ONE summary line per box, with the wash-sale
    adjustment (code W) in column (g) - the line keyed from the 1099 workpaper;
  * basis NOT reported / unknown term (B, E, H, K and unknown): one line per lot.

Rows are normally written by `python run.py import-1099`, which only runs after the
brokerage workpaper ties out. Lots whose term or basis is missing on the 1099 come in
with those cells blank, so `validate` stops until the client supplies them.
"""
from engine.fields import Field
from engine.util import money

from .base import FormDef
from .payer_boxes import TSJ, reference_fields

BOXES = list("ABCDEFGHIJKL")
SHORT = set("ABCGHI")
REPORTED = set("ADGJ")              # basis reported to the IRS
ADJ_CODES = set("BTNHDQXRWLESCMOPYZ")


class F8949Form(FormDef):
    code = "8949"
    title = "Form 8949 / 1099-B - Sales and Dispositions of Capital Assets"
    instance_label = "8949 line"
    source_label = "1099-B workpaper"
    sheet = "8949"
    id_fields = ["lot_ref", "description"]
    check_field = "description"

    fields = [
        Field("ts", "TSJ - Taxpayer / Spouse / Joint", "upper", readback=False, choices=TSJ,
              help="T, S or J (blank = T)"),
        Field("box_8949", "Form 8949 box (A-L)", "code", choices=BOXES,
              help="A/D = basis reported; B/E = not reported; G-L = digital assets. Blank = unknown "
                   "term: enter B or E (H or K) once the client confirms the holding period."),
        Field("description", "1a - Description of property", "text", required=True, maxlen=60),
        Field("date_acquired", "1b - Date acquired", "date", help="MM/DD/YYYY, VARIOUS or INHERIT"),
        Field("date_sold", "1c - Date sold or disposed", "date", required=True),
        Field("proceeds", "1d - Proceeds", "money", required=True),
        Field("cost", "1e - Cost or other basis", "money",
              help="Blank on a noncovered lot = not on the 1099; get it from the client"),
        Field("adj_code", "(f) Adjustment code(s)", "upper", help="e.g. W (wash sale), D, B"),
        Field("adj_amount", "(g) Amount of adjustment", "money"),
        Field("fed_wh", "4 - Federal income tax withheld", "money"),
        # reference only - not typed (not in field_order)
        Field("entry_type", "Entry type (summary line or individual lot)", "text"),
        Field("lot_ref", "Line reference (unique per line)", "text", required=True),
        Field("broker_gain", "Gain/(loss) per the 1099 (tie-out)", "money"),
        Field("payer_name", "Payer's name", "text"),
        Field("payer_ein", "Payer's TIN (EIN)", "ein"),
    ] + reference_fields()

    def rules(self, row, settings):
        e, w = [], []
        box = row.get("box_8949", "").strip().upper()
        cost_raw = row.get("cost", "").strip()
        etype = row.get("entry_type", "")
        if not box:
            e.append("8949 box is blank (unknown term): get the holding period from the client, then "
                     "enter B or E (H or K for digital assets)")
        if not cost_raw:
            if box and box not in REPORTED:
                e.append("Cost basis is blank: basis is not on the 1099, so get it from the client "
                         "(enter 0 only if it really is zero)")
            elif box:
                e.append("Cost basis is required on a basis-reported line")
        if not row.get("date_acquired", "").strip():
            e.append("Date acquired is blank: get it from the client (or VARIOUS / INHERIT)")
        codes = row.get("adj_code", "").strip().upper()
        bad = sorted(set(codes) - ADJ_CODES)
        if bad:
            e.append(f"Adjustment code(s) {''.join(bad)} not valid for Form 8949")
        adj = money(row.get("adj_amount", ""))
        if adj and not codes:
            e.append("Adjustment amount has no adjustment code")
        if codes and not adj:
            w.append("Adjustment code with no adjustment amount")
        if "W" in codes and adj < 0:
            w.append("Wash sale (W) adjustment is negative; it should increase the gain")
        broker = row.get("broker_gain", "").strip()
        if broker and cost_raw and codes in ("", "W"):
            calc = money(row["proceeds"]) - money(cost_raw) + adj
            if abs(calc - money(broker)) > 0.005:
                e.append(f"Proceeds - cost + adjustment = {calc:,.2f} but the 1099 gain/(loss) is "
                         f"{money(broker):,.2f}")
        if etype.startswith("Summary") and codes:
            w.append("Summary line with adjustments: attach the broker statement to the e-file "
                     "(PDF attachment or Form 8453)")
        if "ordinary" in etype.lower():
            w.append("Ordinary-income portion of the box (e.g. discount): confirm the treatment in "
                     "Drake before filing")
        if box and box not in REPORTED and etype.startswith("Summary"):
            e.append("Basis-not-reported boxes must be entered lot by lot, not as a summary")
        return e, w

    def totals(self, rows, settings):
        out = []
        boxes = sorted({r.get("box_8949", "").strip().upper() for r in rows} - {""})
        for box in boxes:
            sel = [r for r in rows if r.get("box_8949", "").strip().upper() == box]
            proceeds = sum(money(r.get("proceeds", "")) for r in sel)
            cost = sum(money(r.get("cost", "")) for r in sel)
            adj = sum(money(r.get("adj_amount", "")) for r in sel)
            part = "Part I (short-term)" if box in SHORT else "Part II (long-term)"
            tie = f"Form 8949 {part}, box {box}"
            out += [(f"Box {box} proceeds", proceeds, tie),
                    (f"Box {box} cost basis", cost, tie),
                    (f"Box {box} adjustments", adj, tie),
                    (f"Box {box} gain/(loss)", proceeds - cost + adj, f"{tie} -> Schedule D")]
        return out
