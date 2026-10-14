"""
Checks that values land where intended:
  field identity (before typing) - which control has the cursor, via Windows UI Automation
  readback (after typing)        - the field's contents, via UI Automation or clipboard
  screenshots                    - evidence of each finished screen
"""
import re
import time
from pathlib import Path

from .util import sha256_file

POS_TOLERANCE = 6  # pixels


class Verifier:
    def __init__(self, settings, kb, screenshot_dir):
        self.v = settings.get("verify", {})
        self.kb = kb
        self.shot_dir = Path(screenshot_dir)
        self._auto = None
        try:
            import uiautomation as auto
            self._auto = auto
        except ImportError:
            pass

    # ---- UI Automation
    def focused(self):
        if not self._auto:
            return None
        try:
            c = self._auto.GetFocusedControl()
            if c is None:
                return None
            r = c.BoundingRectangle
            top = c.GetTopLevelControl()
            w = top.BoundingRectangle if top else r
            try:
                val = c.GetValuePattern().Value
            except Exception:
                val = None
            return {"auto_id": c.AutomationId or "", "name": c.Name or "",
                    "class": c.ClassName or "", "rel_x": r.left - w.left, "rel_y": r.top - w.top,
                    "w": r.right - r.left, "h": r.bottom - r.top, "value": val}
        except Exception:
            return None

    @staticmethod
    def identity(info):
        if not info:
            return None
        return {k: info[k] for k in ("auto_id", "name", "class", "rel_x", "rel_y", "w", "h")}

    @staticmethod
    def same_field(expected, actual):
        if not expected or not actual:
            return None
        if expected.get("auto_id"):
            return expected["auto_id"] == actual.get("auto_id")
        return (abs(expected["rel_x"] - actual["rel_x"]) <= POS_TOLERANCE and
                abs(expected["rel_y"] - actual["rel_y"]) <= POS_TOLERANCE and
                abs(expected["w"] - actual["w"]) <= POS_TOLERANCE)

    def check_focus(self, form, field):
        """True = right field, False = wrong field, None = can't tell / not set up."""
        if self.kb.dry or not form.focus_check:
            return None
        return self.same_field(form.field_ids.get(field), self.identity(self.focused()))

    # ---- readback
    def enabled(self):
        return bool(self.v.get("readback")) and not self.kb.dry

    def read_value(self):
        """Current field's contents, or None if it can't be read. Never changes the field."""
        if not self.enabled():
            return None
        method = self.v.get("readback_method", "auto")
        if method in ("auto", "uia"):
            info = self.focused()
            if info and info.get("value") is not None:
                return info["value"]
            if method == "uia":
                return None
        import pyperclip
        sentinel = "__DRAKE_ENTRY_EMPTY__"
        pyperclip.copy(sentinel)
        self.kb.select_field()
        self.kb.hotkey(self.kb.k["copy"])
        time.sleep(self.kb.d.get("after_copy", 0.12))
        self.kb.press("end")
        got = pyperclip.paste()
        return "" if got == sentinel else got

    # ---- evidence
    def screenshot(self, run_id, label):
        if self.kb.dry or not self.v.get("screenshot"):
            return None, None
        folder = self.shot_dir / run_id
        folder.mkdir(parents=True, exist_ok=True)
        safe = re.sub(r"[^A-Za-z0-9_-]", "_", label)[:80]
        path = folder / f"{time.strftime('%H%M%S')}_{safe}.png"
        win = self.kb.pg.getActiveWindow()
        region = (win.left, win.top, win.width, win.height) if win else None
        self.kb.pg.screenshot(str(path), region=region)
        return path, sha256_file(path)
