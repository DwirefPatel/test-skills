"""Form W-2 definition."""
from engine.fields import Field
from engine.util import money

from .base import FormDef

BOX12_CODES = ["A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q",
               "R", "S", "T", "V", "W", "Y", "Z", "AA", "BB", "DD", "EE", "FF", "GG", "HH", "II"]
SS_RATE, MEDICARE_RATE, ADDL_MEDICARE_THRESHOLD = 0.062, 0.0145, 200000


class W2Form(FormDef):
    code = "W2"
    title = "Form W-2 - Wage and Tax Statement"
    instance_label = "W-2"
    source_label = "W-2"
    sheet = "W2"
    id_fields = ["employer_ein", "employer_name"]
    check_field = "employer_ein"

    fields = [
        Field("ts", "TS - Taxpayer / Spouse", "ts", readback=False, help="T or S (blank = T)"),
        Field("state_code", "State code (top of screen)", "upper", help="Only if your Drake W2 screen uses it"),
        Field("employer_ein", "b - Employer identification number (EIN)", "ein", required=True),
        Field("employer_name", "c - Employer name", "text", required=True),
        Field("employer_street", "c - Employer street address", "text"),
        Field("employer_city", "c - Employer city", "text"),
        Field("employer_state", "c - Employer state", "state"),
        Field("employer_zip", "c - Employer ZIP", "zip"),
        Field("box1_wages", "1 - Wages, tips, other compensation", "money"),
        Field("box2_fed_wh", "2 - Federal income tax withheld", "money"),
        Field("box3_ss_wages", "3 - Social security wages", "money"),
        Field("box4_ss_tax", "4 - Social security tax withheld", "money"),
        Field("box5_medicare_wages", "5 - Medicare wages and tips", "money"),
        Field("box6_medicare_tax", "6 - Medicare tax withheld", "money"),
        Field("box7_ss_tips", "7 - Social security tips", "money"),
        Field("box8_alloc_tips", "8 - Allocated tips", "money"),
        Field("box10_dep_care", "10 - Dependent care benefits", "money"),
        Field("box11_nonqual", "11 - Nonqualified plans", "money"),
    ]
    for L in "abcd":
        fields += [Field(f"box12{L}_code", f"12{L} - Code", "code", choices=BOX12_CODES),
                   Field(f"box12{L}_amt", f"12{L} - Amount", "money")]
    del L
    fields += [
        Field("box13_statutory", "13 - Statutory employee", "check"),
        Field("box13_retirement", "13 - Retirement plan", "check"),
        Field("box13_sick_pay", "13 - Third-party sick pay", "check"),
        Field("box14_desc", "14 - Other (description)", "text"),
        Field("box14_amt", "14 - Other (amount)", "money"),
        Field("box15_state", "15 - State", "state"),
        Field("box15_state_id", "15 - Employer's state ID number", "text"),
        Field("box16_state_wages", "16 - State wages, tips, etc.", "money"),
        Field("box17_state_tax", "17 - State income tax", "money"),
        Field("box18_local_wages", "18 - Local wages, tips, etc.", "money"),
        Field("box19_local_tax", "19 - Local income tax", "money"),
        Field("box20_locality", "20 - Locality name", "text"),
    ]

    def rules(self, row, settings):
        e, w = [], []
        m = lambda k: money(row.get(k, ""))
        b1, b2, b3, b4 = m("box1_wages"), m("box2_fed_wh"), m("box3_ss_wages"), m("box4_ss_tax")
        b5, b6, b7 = m("box5_medicare_wages"), m("box6_medicare_tax"), m("box7_ss_tips")
        year = str(settings.get("tax_year"))
        base = settings.get("ss_wage_base", {}).get(year)

        if b2 > b1 > 0:
            w.append("Box 2 withholding is more than Box 1 wages")
        if b3 + b7 and abs(b4 - (b3 + b7) * SS_RATE) > 1.0:
            w.append(f"Box 4 ({b4:,.2f}) is not 6.2% of Boxes 3+7 ({(b3 + b7) * SS_RATE:,.2f})")
        if base and b3 + b7 > base:
            w.append(f"Boxes 3+7 ({b3 + b7:,.2f}) exceed the {year} SS wage base ({base:,.0f})")
        if b5:
            expected = b5 * MEDICARE_RATE + max(0, b5 - ADDL_MEDICARE_THRESHOLD) * 0.009
            if abs(b6 - expected) > 1.0:
                w.append(f"Box 6 ({b6:,.2f}) differs from expected Medicare tax ({expected:,.2f})")
        if b1 and not b5 and not b3:
            w.append("Box 1 has wages but Boxes 3 and 5 are blank")
        for L in "abcd":
            code, amt = row.get(f"box12{L}_code", "").strip(), row.get(f"box12{L}_amt", "").strip()
            if bool(code) != bool(amt):
                e.append(f"Box 12{L} needs both a code and an amount")
        if row.get("box14_amt", "").strip() and not row.get("box14_desc", "").strip():
            w.append("Box 14 has an amount but no description")
        if m("box16_state_wages") and not row.get("box15_state", "").strip():
            e.append("Box 16 has state wages but Box 15 state is blank")
        return e, w

    def totals(self, rows, settings):
        tot = lambda k, who=None: sum(money(r.get(k, "")) for r in rows
                                      if who is None or (r.get("ts", "").strip().upper() or "T") == who)
        out = [("Total Box 1 wages", tot("box1_wages"), "Form 1040, line 1a"),
               ("Total Box 2 federal withholding", tot("box2_fed_wh"), "Form 1040, line 25a")]
        if any((r.get("ts", "").strip().upper() == "S") for r in rows):
            out += [("  Taxpayer Box 1", tot("box1_wages", "T"), "W-2 summary"),
                    ("  Spouse Box 1", tot("box1_wages", "S"), "W-2 summary")]
        states = sorted({r.get("box15_state", "").strip().upper() for r in rows} - {""})
        for st in states:
            amt = sum(money(r.get("box17_state_tax", "")) for r in rows
                      if r.get("box15_state", "").strip().upper() == st)
            out.append((f"Total Box 17 {st} withholding", amt, f"{st} return withholding"))
        return out
