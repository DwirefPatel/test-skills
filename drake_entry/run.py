"""
Drake data-entry automation - command line.

  python run.py make-input            create input/Client_Input.xlsx (one tab per enabled form)
  python run.py validate              check the input workbook; no typing
  python run.py enter                 enter new documents into Drake (verified field by field)
  python run.py enter --dry-run       show what would be typed
  python run.py verify                read-only: compare Drake with the source, log differences
  python run.py correct               verify, then fix the differences you approve
  python run.py calibrate --form W2   confirm the tab order on a TEST return
  python run.py map-fields --form W2  record field identities (wrong-field check)
  python run.py build-tracker         rebuild tracking/Entry_Tracker.xlsx
  python run.py audit-verify          check the audit log hasn't been altered

Common options:  --client 6789  (SSN last 4 or client_id)   --forms W2   --initials DK
Safety: move the mouse into any screen corner to abort instantly.
"""
import argparse
import shutil
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

from engine.inputs import make_template  # noqa: E402
from engine.runner import Session  # noqa: E402
from engine.tracker import build_tracker  # noqa: E402

EXAMPLES = {"W2": {"client_ssn": "123-45-6789", "client_name": "Jane Sample", "ts": "T",
                   "employer_ein": "84-1234567", "employer_name": "Front Range Hotels LLC",
                   "employer_street": "1600 Broadway", "employer_city": "Denver", "employer_state": "CO",
                   "employer_zip": "80202", "box1_wages": "52000.00", "box2_fed_wh": "4800.00",
                   "box3_ss_wages": "55000.00", "box4_ss_tax": "3410.00", "box5_medicare_wages": "55000.00",
                   "box6_medicare_tax": "797.50", "box12a_code": "D", "box12a_amt": "3000.00",
                   "box13_retirement": "X", "box15_state": "CO", "box15_state_id": "12345678",
                   "box16_state_wages": "52000.00", "box17_state_tax": "2288.00"}}


def main():
    ap = argparse.ArgumentParser(description="Drake data-entry automation", formatter_class=argparse.RawTextHelpFormatter,
                                 epilog=__doc__)
    ap.add_argument("command", choices=["make-input", "validate", "enter", "verify", "correct", "calibrate",
                                        "map-fields", "build-tracker", "audit-verify"])
    ap.add_argument("--client", help="only this client: SSN last 4 or client_id")
    ap.add_argument("--forms", help="comma-separated form codes (default: forms_enabled in settings.json)")
    ap.add_argument("--form", help="form code for calibrate / map-fields")
    ap.add_argument("--input", help="input workbook (default from settings.json)")
    ap.add_argument("--initials", help="preparer initials for the audit trail")
    ap.add_argument("--dry-run", action="store_true", help="enter: print keystrokes only")
    ap.add_argument("--force", action="store_true", help="make-input: replace an existing workbook (backs it up)")
    ap.add_argument("--assume-yes", action="store_true", help=argparse.SUPPRESS)  # automated tests only
    a = ap.parse_args()

    s = Session(BASE, a)
    cmd = a.command
    if cmd == "make-input":
        path = s.paths["input"]
        if path.exists() and not a.force:
            sys.exit(f"{path} already exists. Use --force to replace it (a backup copy is kept).")
        if path.exists():
            backup = path.with_name(f"{path.stem}_backup_{time.strftime('%Y%m%d_%H%M%S')}{path.suffix}")
            shutil.copy2(path, backup)
            print(f"Backed up existing workbook to {backup.name}")
        path.parent.mkdir(parents=True, exist_ok=True)
        make_template(path, s.forms, EXAMPLES)
        print(f"Created {path}")
    elif cmd == "validate":
        sys.exit(0 if s.mode_validate() else 1)
    elif cmd == "enter":
        s.mode_enter()
    elif cmd == "verify":
        s.mode_check(fix=False)
    elif cmd == "correct":
        s.mode_check(fix=True)
    elif cmd == "calibrate":
        s.mode_calibrate()
    elif cmd == "map-fields":
        s.mode_map_fields()
    elif cmd == "build-tracker":
        p = build_tracker(s)
        print(f"Tracker: {p}" if p else "Nothing logged yet.")
    elif cmd == "audit-verify":
        ok, n, problem = s.audit.verify()
        print(f"Audit log: {n} entries. " + ("PASSED - intact." if ok else f"FAILED - {problem}"))
        sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()
