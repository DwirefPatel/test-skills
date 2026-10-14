"""Field types: how a source value is validated, typed into Drake, and compared."""
import re
from datetime import datetime

from .util import digits, money

TRUE_WORDS = ("X", "Y", "YES", "TRUE", "1")
FALSE_WORDS = ("N", "NO", "FALSE", "0")
DATE_WORDS = ("VARIOUS", "INHERIT")
DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y", "%m%d%Y")


def parse_date(raw):
    """'1/5/25', '2025-01-05', '01052025' -> datetime; None if not a date."""
    raw = str(raw or "").strip()
    if " " in raw and raw.endswith("00:00:00"):   # Excel date cell read as text
        raw = raw.split(" ")[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


class Field:
    """
    kind: text | upper | state | code | ts | money | ein | zip | check | date
          (date is typed as MMDDYYYY; VARIOUS / INHERIT are typed as words)
    readback: whether the value can be read back from Drake (checkbox-style
              fields usually can't, so they're shown as 'Not read back').
    """

    def __init__(self, name, label, kind="text", required=False, readback=None,
                 choices=None, maxlen=None, help=""):
        self.name, self.label, self.kind = name, label, kind
        self.required = required
        self.readback = (kind != "check") if readback is None else readback
        self.choices = {c.upper() for c in choices} if choices else None
        self.maxlen = maxlen
        self.help = help

    def keys(self, raw):
        """Value as typed into Drake."""
        raw = str(raw or "").strip()
        if raw == "":
            return ""
        k = self.kind
        if k == "money":
            n = money(raw)
            return str(int(n)) if n == int(n) else f"{n:.2f}"
        if k in ("ein", "zip"):
            return digits(raw)[:9]
        if k == "check":
            return "X" if raw.upper() in TRUE_WORDS else ""
        if k == "date":
            if raw.upper() in DATE_WORDS:
                return raw.upper()
            d = parse_date(raw)
            return d.strftime("%m%d%Y") if d else raw
        if k in ("upper", "state", "code", "ts"):
            return raw.upper()
        return raw

    def norm(self, v):
        """Canonical form for comparing the source value with what Drake shows."""
        v = str(v or "").strip()
        if self.kind == "money":
            try:
                return f"{money(v):.2f}"
            except ValueError:
                return v
        if self.kind in ("ein", "zip"):
            return digits(v)
        if self.kind == "check":
            return "X" if v.upper() in TRUE_WORDS else ""
        if self.kind == "date":
            d = parse_date(v)
            return d.strftime("%m%d%Y") if d else re.sub(r"[^A-Z0-9]", "", v.upper())
        return re.sub(r"[^A-Z0-9]", "", v.upper())

    def check(self, raw):
        """Format problem with the source value, or None if it's fine."""
        raw = str(raw or "").strip()
        if raw == "":
            return "is required" if self.required else None
        k = self.kind
        if k == "money":
            try:
                money(raw)
            except ValueError:
                return f"is not a number ({raw!r})"
        if k == "ein" and len(digits(raw)) != 9:
            return "must be 9 digits"
        if k == "zip" and len(digits(raw)) not in (5, 9):
            return "must be 5 or 9 digits"
        if k == "state" and not re.fullmatch(r"[A-Za-z]{2}", raw):
            return "must be a 2-letter state code"
        if k == "ts" and raw.upper() not in ("T", "S"):
            return "must be T or S"
        if k == "date" and raw.upper() not in DATE_WORDS and not parse_date(raw):
            return "must be a date (MM/DD/YYYY) or VARIOUS / INHERIT"
        if k == "check" and raw.upper() not in TRUE_WORDS + FALSE_WORDS:
            return "must be X or blank"
        if self.choices and raw.upper() not in self.choices:
            return f"'{raw}' is not a valid code"
        if self.maxlen and len(raw) > self.maxlen:
            return f"is longer than {self.maxlen} characters"
        return None
