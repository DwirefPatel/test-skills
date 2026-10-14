"""Keyboard driver for Drake data-entry screens (all keystrokes go through here)."""
import sys
import time


class Keyboard:
    def __init__(self, settings, dry_run=False):
        self.settings = settings
        self.k = dict(settings["keys"])
        self.d = settings["delays"]
        self.title = settings["window_title_must_contain"]
        self.dry = dry_run
        self.pg = None
        if not dry_run:
            import pyautogui
            pyautogui.FAILSAFE = True   # mouse into a screen corner = instant abort
            pyautogui.PAUSE = 0.02
            self.pg = pyautogui

    def use_form_keys(self, form):
        self.k = form.keys(self.settings["keys"])

    # ---- primitives
    def check_window(self):
        if self.dry:
            return
        win = self.pg.getActiveWindow()
        title = win.title if win else ""
        if self.title.lower() not in title.lower():
            raise WrongWindow(f"Active window is '{title}', not Drake. Nothing more was typed.")

    def press(self, key, wait=0.0):
        if self.dry:
            print(f"    [{key.upper()}]")
            return
        self.pg.press(key)
        if wait:
            time.sleep(wait)

    def hotkey(self, keys):
        if self.dry:
            print(f"    [{'+'.join(k.upper() for k in keys)}]")
            return
        self.pg.hotkey(*keys)

    def write(self, text):
        if self.dry:
            print(f"    type {text!r}")
        elif text:
            self.pg.write(text, interval=self.d["between_keys"])

    def countdown(self):
        if self.dry:
            return
        for s in range(int(self.d["countdown_seconds"]), 0, -1):
            print(f"  Starting in {s}...  (click into Drake now; mouse to a corner aborts)", end="\r")
            time.sleep(1)
        print(" " * 80, end="\r")

    # ---- screen navigation
    def open_screen(self, form):
        """From the screen selector: open this form's screen at instance #1, field #1."""
        self.use_form_keys(form)
        self.check_window()
        self.write(form.screen_code)
        self.press(self.k["open_screen"], self.d["after_open_screen"])
        for key in self.k.get("open_existing_extra", []):
            self.press(key, self.d["after_open_screen"])

    def next_instance(self, times=1):
        for _ in range(times):
            self.check_window()
            self.press(self.k["new_instance"], self.d["after_new_instance"])

    def next_field(self, times=1):
        for _ in range(times):
            self.press(self.k["next_field"], self.d["after_field"])

    def exit_screen(self):
        self.press(self.k["exit_screen"], self.d["after_exit"])

    # ---- field editing
    def select_field(self):
        self.press(self.k["field_start"])
        self.hotkey(self.k["select_to_end"])

    def set_field(self, value):
        """Replace whatever is in the current field with value (blank = clear it)."""
        self.check_window()
        self.select_field()
        if value:
            self.write(value)
        else:
            self.press(self.k["clear"])


class WrongWindow(Exception):
    pass
