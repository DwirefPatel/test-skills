"""Form 1099-MISC definition (Drake screen 99M) - the boxes a brokerage 1099 uses."""
from engine.fields import Field

from .payer_boxes import PayerBoxesForm, payer_fields, reference_fields


class MISCForm(PayerBoxesForm):
    code = "99M"
    title = "Form 1099-MISC - Miscellaneous Information"
    instance_label = "1099-MISC"
    source_label = "1099-MISC"
    sheet = "99M"

    fields = payer_fields() + [
        Field("box2", "2 - Royalties", "money"),
        Field("box3", "3 - Other income", "money"),
        Field("box4", "4 - Federal income tax withheld", "money"),
        Field("box8", "8 - Substitute payments in lieu of dividends or interest", "money"),
        Field("box15_state_wh", "15 - State tax withheld", "money"),
        Field("box16_state", "16 - State", "state"),
    ] + reference_fields()

    tieout = [
        ("box2", "Total box 2 royalties", "Schedule E"),
        ("box3", "Total box 3 other income", "Sch 1, line 8z"),
        ("box8", "Total box 8 substitute payments", "Sch 1, line 8z (ordinary, not qualified)"),
        ("box4", "Total box 4 federal withholding", "Form 1040, line 25b"),
    ]

    def rules(self, row, settings):
        e, w = [], []
        if self.m(row, "box3") or self.m(row, "box8"):
            w.append("Box 3/8 on Drake's 99M screen: confirm the 'For' link sends it to Schedule 1 "
                     "other income (box 8 is ordinary income, not a qualified dividend)")
        if self.m(row, "box15_state_wh") and not row.get("box16_state", "").strip():
            e.append("Box 15 has state withholding but Box 16 state is blank")
        return e, w
