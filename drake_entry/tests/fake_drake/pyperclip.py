"""FAKE clipboard for tests."""
import json
import os

_P = os.environ.get("FAKE_DRAKE_STATE", "fake_drake_state.json") + ".clip"


def copy(t):
    json.dump(t, open(_P, "w"))


def paste():
    return json.load(open(_P)) if os.path.exists(_P) else ""
