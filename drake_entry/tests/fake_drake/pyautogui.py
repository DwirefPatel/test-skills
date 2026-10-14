"""
FAKE Drake for testing - stands in for the real pyautogui module.
Simulates Drake's screen selector, multi-instance screens (PgDn), tabbing,
field selection/overwrite, clipboard copy, Esc to exit, and optional faults:

  FAKE_DRAKE_STATE     path to a JSON file holding the fake returns (persists across runs)
  FAKE_CORRUPT=value   the first time this exact value is typed, store '99999' instead (keying error)
  FAKE_AUTOFILL=i:j=TEXT  after leaving field index i with a value, pre-fill blank field j with TEXT
                          (mimics Drake's EIN database filling in the employer name)
"""
import json
import os


class FailSafeException(Exception):
    pass


FAILSAFE, PAUSE = True, 0
_P = os.environ.get("FAKE_DRAKE_STATE", "fake_drake_state.json")


def _load():
    if os.path.exists(_P):
        return json.load(open(_P))
    return {"return": os.environ.get("FAKE_RETURN", "default"), "returns": {}, "screen": None, "inst": 0,
            "field": 0, "sel": False, "buf": "", "corrupted": False}


def _save(s):
    json.dump(s, open(_P, "w"), indent=1)


def _vals(s):
    ret = s["returns"].setdefault(os.environ.get("FAKE_RETURN", "default"), {})
    scr = ret.setdefault(s["screen"], [])
    while len(scr) <= s["inst"]:
        scr.append({})
    return scr[s["inst"]]


class _Win:
    title, left, top, width, height = "Drake Tax 2025 (simulated)", 0, 0, 800, 600


def getActiveWindow():
    return _Win()


def write(text, interval=0):
    s = _load()
    if s["screen"] is None:
        s["buf"] += text
    else:
        v = _vals(s)
        key = str(s["field"])
        if os.environ.get("FAKE_CORRUPT") == text and not s["corrupted"]:
            text, s["corrupted"] = "99999", True
        v[key] = text if s["sel"] or key not in v else v[key] + text
        s["sel"] = False
    _save(s)


def press(key):
    s = _load()
    if key == "enter" and s["screen"] is None:
        s["screen"], s["inst"], s["field"], s["buf"] = s["buf"].upper(), 0, 0, ""
        _vals(s)
    elif key == "tab" and s["screen"]:
        af = os.environ.get("FAKE_AUTOFILL")
        if af:
            src, rest = af.split(":")
            dst, txt = rest.split("=", 1)
            v = _vals(s)
            if str(s["field"]) == src and v.get(src) and not v.get(dst):
                v[dst] = txt
        s["field"] += 1
        s["sel"] = False
    elif key == "pagedown" and s["screen"]:
        s["inst"] += 1
        s["field"], s["sel"] = 0, False
        _vals(s)
    elif key == "esc":
        if s["screen"]:
            scr = s["returns"][os.environ.get("FAKE_RETURN", "default")][s["screen"]]
            while scr and not any(scr[-1].values()):   # blank instances aren't saved
                scr.pop()
        s["screen"], s["buf"], s["sel"] = None, "", False
    elif key in ("home", "end"):
        s["sel"] = False
    elif key == "delete" and s["screen"] and s["sel"]:
        _vals(s)[str(s["field"])] = ""
        s["sel"] = False
    _save(s)


def hotkey(*keys):
    s = _load()
    if keys == ("shift", "end"):
        s["sel"] = True
    elif keys == ("ctrl", "c") and s["screen"]:
        import pyperclip
        pyperclip.copy(_vals(s).get(str(s["field"]), ""))
    _save(s)


def screenshot(path, region=None):
    open(path, "wb").write(b"fake screenshot")
