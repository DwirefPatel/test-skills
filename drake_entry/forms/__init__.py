"""
Form registry. To add a form:
  1. copy forms/w2.py -> forms/int1099.py, define its fields/rules/totals
  2. create forms/99I.json (screen code + tab order)
  3. add it to REGISTRY below and to "forms_enabled" in settings.json
  4. python run.py make-input --force   (adds its tab to the input workbook)
  5. calibrate + map-fields on a test return, then test with verify before real use
"""
from pathlib import Path

from .w2 import W2Form

REGISTRY = {
    "W2": (W2Form, "W2.json"),
}


def load_forms(codes, base_dir):
    out = {}
    for code in codes:
        if code not in REGISTRY:
            raise SystemExit(f"Unknown form '{code}'. Available: {', '.join(REGISTRY)}")
        cls, cfg = REGISTRY[code]
        out[code] = cls(Path(base_dir) / "forms" / cfg)
    return out
