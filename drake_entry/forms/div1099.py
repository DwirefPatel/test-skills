"""Form 1099-DIV definition (Drake screen DIV)."""
from engine.fields import Field

from .payer_boxes import PayerBoxesForm, payer_fields, reference_fields


class DIVForm(PayerBoxesForm):
    code = "DIV"
    title = "Form 1099-DIV - Dividends and Distributions"
    instance_label = "1099-DIV"
    source_label = "1099-DIV"
    sheet = "DIV"

    fields = payer_fields() + [
        Field("box1a", "1a - Total ordinary dividends", "money"),
        Field("box1b", "1b - Qualified dividends", "money"),
        Field("box2a", "2a - Total capital gain distributions", "money"),
        Field("box2b", "2b - Unrecaptured Sec. 1250 gain", "money"),
        Field("box2c", "2c - Section 1202 gain", "money"),
        Field("box2d", "2d - Collectibles (28%) gain", "money"),
        Field("box2e", "2e - Section 897 ordinary dividends", "money"),
        Field("box2f", "2f - Section 897 capital gain", "money"),
        Field("box3", "3 - Nondividend distributions", "money"),
        Field("box4", "4 - Federal income tax withheld", "money"),
        Field("box5", "5 - Section 199A dividends", "money"),
        Field("box6", "6 - Investment expenses", "money"),
        Field("box7", "7 - Foreign tax paid", "money"),
        Field("box8", "8 - Foreign country or U.S. possession", "text"),
        Field("box9", "9 - Cash liquidation distributions", "money"),
        Field("box10", "10 - Noncash liquidation distributions", "money"),
        Field("box12", "12 - Exempt-interest dividends", "money"),
        Field("box13", "13 - Specified private activity bond interest dividends", "money"),
        Field("box14_state", "14 - State", "state"),
        Field("box15_state_id", "15 - State identification no.", "text"),
        Field("box16_state_wh", "16 - State tax withheld", "money"),
    ] + reference_fields()

    tieout = [
        ("box1a", "Total 1a ordinary dividends", "Form 1040, line 3b / Sch B Part II"),
        ("box1b", "Total 1b qualified dividends", "Form 1040, line 3a"),
        ("box2a", "Total 2a capital gain distributions", "Schedule D, line 13"),
        ("box4", "Total box 4 federal withholding", "Form 1040, line 25b"),
        ("box5", "Total box 5 Sec. 199A dividends", "Form 8995"),
        ("box7", "Total box 7 foreign tax", "Sch 3 line 1 / Form 1116"),
        ("box12", "Total box 12 exempt-interest dividends", "Form 1040, line 2a"),
    ]

    def rules(self, row, settings):
        e, w = [], []
        b1a, b1b, b2a = self.m(row, "box1a"), self.m(row, "box1b"), self.m(row, "box2a")
        if b1b > b1a:
            e.append("Box 1b qualified dividends are more than Box 1a ordinary dividends")
        if self.m(row, "box5") > b1a:
            w.append("Box 5 Section 199A dividends are more than Box 1a")
        parts = sum(self.m(row, b) for b in ("box2b", "box2c", "box2d", "box2f"))
        if parts > b2a + 0.005:
            w.append("Boxes 2b-2f add up to more than Box 2a")
        if self.m(row, "box7") and not row.get("box8", "").strip():
            w.append("Box 7 foreign tax has no Box 8 country (RIC: enter 'RIC' or the country)")
        if self.m(row, "box9") or self.m(row, "box10"):
            w.append("Box 9/10 liquidation distributions are not dividends: basis recovery first, "
                     "then capital gain")
        if self.m(row, "box16_state_wh") and not row.get("box14_state", "").strip():
            e.append("Box 16 has state withholding but Box 14 state is blank")
        return e, w
