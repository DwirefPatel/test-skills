"""
Base class every form definition extends.

A form = a Python file (fields + validation rules + tie-out totals)
       + a JSON file (Drake screen code, tab order, label overrides, field map).
The engine handles everything else: typing, verification, correction,
audit log and tracker.
"""
import json
from pathlib import Path

from engine.util import sha256_file, sha256_text


class FormDef:
    code = ""              # short code, e.g. "W2"
    title = ""             # section title in the tracker
    instance_label = ""    # "W-2" -> "W-2 #1"
    source_label = ""      # source document name
    sheet = ""             # tab name in Client_Input.xlsx
    id_fields = []         # fields that identify one document (e.g. employer EIN)
    check_field = None     # read first in verify/correct to confirm the right instance
    fields = []            # list of engine.fields.Field

    def __init__(self, cfg_path):
        self.cfg_path = Path(cfg_path)
        self.cfg = json.loads(self.cfg_path.read_text(encoding="utf-8"))
        self.by_name = {f.name: f for f in self.fields}
        for name, label in self.cfg.get("label_overrides", {}).items():
            if name in self.by_name:
                self.by_name[name].label = label
        unknown = [n for n in self.field_order if n != "__SKIP__" and n not in self.by_name]
        if unknown:
            raise ValueError(f"{self.cfg_path.name}: unknown fields in field_order: {unknown}")

    # ---- screen configuration (JSON)
    @property
    def screen_code(self):
        return self.cfg["screen_code"]

    @property
    def field_order(self):
        return self.cfg["field_order"]

    @property
    def field_ids(self):
        return self.cfg.get("field_ids", {})

    @property
    def focus_check(self):
        return bool(self.cfg.get("focus_check")) and bool(self.field_ids)

    def keys(self, settings_keys):
        return {**settings_keys, **self.cfg.get("keys", {})}

    def save_cfg(self):
        self.cfg_path.write_text(json.dumps(self.cfg, indent=2), encoding="utf-8")

    def version(self):
        """Fingerprint of this form's code + config, recorded in the audit log."""
        import inspect
        src = Path(inspect.getfile(type(self)))
        return sha256_text(sha256_file(src) + sha256_file(self.cfg_path))[:16]

    # ---- documents
    def label(self, name):
        return self.by_name[name].label if name in self.by_name else name

    def fingerprint(self, client_key, row):
        parts = [self.code, client_key] + [self.by_name[f].norm(row.get(f, "")) for f in self.id_fields]
        return sha256_text("|".join(parts))[:16]

    def describe(self, row):
        return " | ".join(str(row.get(f, "")).strip() for f in self.id_fields if str(row.get(f, "")).strip())

    # ---- validation
    def validate(self, row, settings):
        errors, warnings = [], []
        for f in self.fields:
            problem = f.check(row.get(f.name, ""))
            if problem:
                errors.append(f"{f.label} {problem}")
        if not errors:
            e, w = self.rules(row, settings)
            errors += e
            warnings += w
        return errors, warnings

    def rules(self, row, settings):
        """Form-specific checks. Return (errors, warnings)."""
        return [], []

    def totals(self, rows, settings):
        """Totals to tie out to the return: list of (label, amount, where_to_tie)."""
        return []
