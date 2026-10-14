"""Small shared helpers."""
import getpass
import hashlib
import platform
import re

VERSION = "2.0.0"


def digits(s):
    return re.sub(r"\D", "", str(s or ""))


def mask(ssn):
    d = digits(ssn)
    return f"***-**-{d[-4:]}" if len(d) >= 4 else "****"


def money(s):
    """'$1,234.50' -> 1234.5 ; '(100)' -> -100 ; '' -> 0.0 ; raises ValueError on junk."""
    t = str(s or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    if t == "":
        return 0.0
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    n = float(t)
    return -n if neg else n


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_text(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def who():
    try:
        user = getpass.getuser()
    except Exception:
        user = "unknown"
    return user, platform.node()
