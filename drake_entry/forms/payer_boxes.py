"""Shared base for 1099 screens that are one payer + a set of money boxes (DIV, INT, OID, 99M)."""
from engine.fields import Field
from engine.util import money

from .base import FormDef

TSJ = ["T", "S", "J"]


def payer_fields():
    return [
        Field("ts", "TSJ - Taxpayer / Spouse / Joint", "upper", readback=False, choices=TSJ,
              help="T, S or J (blank = T)"),
        Field("payer_ein", "Payer's TIN (EIN)", "ein", required=True),
        Field("payer_name", "Payer's name", "text", required=True),
    ]


def reference_fields():
    """Kept in the input workbook for traceability; not typed into Drake (not in field_order)."""
    return [
        Field("source_doc", "Source (set by import-1099; identifies the brokerage statement)", "text"),
        Field("source_page", "Source page(s) on the 1099", "text"),
    ]


class PayerBoxesForm(FormDef):
    id_fields = ["payer_ein", "payer_name"]
    check_field = "payer_ein"
    #: (field name, label) of the boxes to total in the tie-out, with where each ties to
    tieout = []

    def m(self, row, name):
        return money(row.get(name, ""))

    def totals(self, rows, settings):
        out = []
        for name, label, tie in self.tieout:
            amt = sum(self.m(r, name) for r in rows)
            if amt:
                out.append((label, amt, tie))
        return out
