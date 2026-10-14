"""
Form registry. To add a form:
  1. copy forms/w2.py -> forms/int1099.py, define its fields/rules/totals
  2. create forms/99I.json (screen code + tab order)
  3. add it to REGISTRY below and to "forms_enabled" in settings.json
  4. python run.py make-input --force   (adds its tab to the input workbook)
  5. calibrate + map-fields on a test return, then test with verify before real use
"""
from pathlib import Path

from .div1099 import DIVForm
from .f8949 import F8949Form
from .int1099 import INTForm
from .misc1099 import MISCForm
from .oid1099 import OIDForm
from .w2 import W2Form

REGISTRY = {
    "W2": (W2Form, "W2.json"),
    "8949": (F8949Form, "8949.json"),
    "DIV": (DIVForm, "DIV.json"),
    "INT": (INTForm, "INT.json"),
    "OID": (OIDForm, "OID.json"),
    "99M": (MISCForm, "99M.json"),
}


def load_forms(codes, base_dir):
    out = {}
    for code in codes:
        if code not in REGISTRY:
            raise SystemExit(f"Unknown form '{code}'. Available: {', '.join(REGISTRY)}")
        cls, cfg = REGISTRY[code]
        out[code] = cls(Path(base_dir) / "forms" / cfg)
    return out
