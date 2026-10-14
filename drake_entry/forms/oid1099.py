"""Form 1099-OID definition (Drake screen OID)."""
from engine.fields import Field

from .payer_boxes import PayerBoxesForm, payer_fields, reference_fields


class OIDForm(PayerBoxesForm):
    code = "OID"
    title = "Form 1099-OID - Original Issue Discount"
    instance_label = "1099-OID"
    source_label = "1099-OID"
    sheet = "OID"

    fields = payer_fields() + [
        Field("box1", "1 - Original issue discount", "money"),
        Field("box2", "2 - Other periodic interest", "money"),
        Field("box3", "3 - Early withdrawal penalty", "money"),
        Field("box4", "4 - Federal income tax withheld", "money"),
        Field("box5", "5 - Market discount", "money"),
        Field("box6", "6 - Acquisition premium", "money"),
        Field("box8", "8 - OID on U.S. Treasury obligations", "money"),
        Field("box9", "9 - Investment expenses", "money"),
        Field("box10", "10 - Bond premium", "money"),
        Field("box11", "11 - Tax-exempt OID", "money"),
        Field("box12_state", "12 - State", "state"),
        Field("box14_state_wh", "14 - State tax withheld", "money"),
    ] + reference_fields()

    tieout = [
        ("box1", "Total box 1 OID", "Sch B (net of acquisition premium)"),
        ("box2", "Total box 2 other periodic interest", "Sch B"),
        ("box8", "Total box 8 Treasury OID", "Sch B; state subtraction"),
        ("box4", "Total box 4 federal withholding", "Form 1040, line 25b"),
        ("box11", "Total box 11 tax-exempt OID", "Form 1040, line 2a"),
    ]

    def rules(self, row, settings):
        e, w = [], []
        if self.m(row, "box6") > self.m(row, "box1") > 0:
            w.append("Box 6 acquisition premium is more than Box 1 OID")
        if self.m(row, "box6") or self.m(row, "box10"):
            w.append("Acquisition/bond premium present: confirm Drake reports the taxable OID net "
                     "of the premium")
        if self.m(row, "box14_state_wh") and not row.get("box12_state", "").strip():
            e.append("Box 14 has state withholding but Box 12 state is blank")
        return e, w
