"""Form 1099-INT definition (Drake screen INT)."""
from engine.fields import Field

from .payer_boxes import PayerBoxesForm, payer_fields, reference_fields


class INTForm(PayerBoxesForm):
    code = "INT"
    title = "Form 1099-INT - Interest Income"
    instance_label = "1099-INT"
    source_label = "1099-INT"
    sheet = "INT"

    fields = payer_fields() + [
        Field("box1", "1 - Interest income", "money"),
        Field("box2", "2 - Early withdrawal penalty", "money"),
        Field("box3", "3 - Interest on U.S. Savings Bonds and Treasury obligations", "money"),
        Field("box4", "4 - Federal income tax withheld", "money"),
        Field("box5", "5 - Investment expenses", "money"),
        Field("box6", "6 - Foreign tax paid", "money"),
        Field("box7", "7 - Foreign country or U.S. possession", "text"),
        Field("box8", "8 - Tax-exempt interest", "money"),
        Field("box9", "9 - Specified private activity bond interest", "money"),
        Field("box10", "10 - Market discount", "money"),
        Field("box11", "11 - Bond premium", "money"),
        Field("box12", "12 - Bond premium on Treasury obligations", "money"),
        Field("box13", "13 - Bond premium on tax-exempt bond", "money"),
        Field("box15_state", "15 - State", "state"),
        Field("box16_state_id", "16 - State identification no.", "text"),
        Field("box17_state_wh", "17 - State tax withheld", "money"),
    ] + reference_fields()

    tieout = [
        ("box1", "Total box 1 interest", "Sch B Part I / Form 1040, line 2b"),
        ("box3", "Total box 3 Treasury interest", "Sch B; state subtraction"),
        ("box4", "Total box 4 federal withholding", "Form 1040, line 25b"),
        ("box8", "Total box 8 tax-exempt interest", "Form 1040, line 2a"),
    ]

    def rules(self, row, settings):
        e, w = [], []
        if self.m(row, "box12") > self.m(row, "box3") > 0:
            w.append("Box 12 Treasury bond premium is more than Box 3")
        if self.m(row, "box13") > self.m(row, "box8") > 0:
            w.append("Box 13 tax-exempt bond premium is more than Box 8")
        if self.m(row, "box6") and not row.get("box7", "").strip():
            w.append("Box 6 foreign tax has no Box 7 country")
        if self.m(row, "box17_state_wh") and not row.get("box15_state", "").strip():
            e.append("Box 17 has state withholding but Box 15 state is blank")
        return e, w
