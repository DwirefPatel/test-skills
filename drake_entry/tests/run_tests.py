"""
End-to-end tests against a SIMULATED Drake (tests/fake_drake). No real Drake needed.

  python tests/run_tests.py

Copies the tool into a temp folder, then exercises: validation, entry, a keying
error that stops entry, re-run protection against duplicates, verify, correct,
Drake pre-fill handling, an edited EIN in the input, audit-log tamper detection,
and the tracker workbook.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PKG = HERE.parent
FAKE = HERE / "fake_drake"
PASSED = []


def check(name, cond, detail=""):
    print(f"  [{'PASS' if cond else 'FAIL'}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    if not cond:
        raise SystemExit(f"\nTEST FAILED: {name}")
    PASSED.append(name)


def run(work, *args, ret="default", env=None):
    e = {**os.environ, "PYTHONPATH": str(FAKE), "FAKE_DRAKE_STATE": str(work / "fake_state.json"),
         "FAKE_RETURN": ret, **(env or {})}
    p = subprocess.run([sys.executable, str(work / "run.py"), *args, "--assume-yes", "--initials", "QA"],
                       cwd=work, env=e, capture_output=True, text=True)
    return p.returncode, p.stdout + p.stderr


def press_esc_in_drake(work, ret):
    """What the user is told to do after a stop: press Esc in Drake."""
    e = {**os.environ, "PYTHONPATH": str(FAKE), "FAKE_DRAKE_STATE": str(work / "fake_state.json"), "FAKE_RETURN": ret}
    subprocess.run([sys.executable, "-c", "import pyautogui; pyautogui.press('esc')"], env=e, cwd=work)


def fake_return(work, ret):
    s = json.load(open(work / "fake_state.json"))
    return s["returns"].get(ret, {}).get("W2", [])


def write_input(work, rows):
    from openpyxl import load_workbook
    path = work / "input" / "Client_Input.xlsx"
    wb = load_workbook(path)
    ws = wb["W2"]
    names = [c.value for c in ws[2]]
    for i in range(4, 60):
        for j in range(1, len(names) + 1):
            ws.cell(i, j).value = None
    for i, r in enumerate(rows, 4):
        for j, n in enumerate(names, 1):
            if n in r:
                ws.cell(i, j).value = r[n]
    wb.save(path)


SAM_T = {"client_ssn": "123-45-6789", "client_name": "Sam Rivera", "ts": "T", "employer_ein": "84-1234567",
         "employer_name": "Front Range Hotels LLC", "employer_street": "1600 Broadway", "employer_city": "Denver",
         "employer_state": "CO", "employer_zip": "80202", "box1_wages": "52,000.00", "box2_fed_wh": "4800",
         "box3_ss_wages": "55000", "box4_ss_tax": "3400", "box5_medicare_wages": "55000",
         "box6_medicare_tax": "797.50", "box12a_code": "D", "box12a_amt": "3000", "box13_retirement": "X",
         "box15_state": "CO", "box15_state_id": "12345678", "box16_state_wages": "52000", "box17_state_tax": "2288"}
SAM_S = {"client_ssn": "123-45-6789", "client_name": "Sam Rivera", "ts": "S", "employer_ein": "84-7654321",
         "employer_name": "Aurora School District", "employer_city": "Aurora", "employer_state": "CO",
         "employer_zip": "80010", "box1_wages": "38500", "box2_fed_wh": "2900", "box3_ss_wages": "38500",
         "box4_ss_tax": "2387", "box5_medicare_wages": "38500", "box6_medicare_tax": "558.25",
         "box15_state": "CO", "box16_state_wages": "38500", "box17_state_tax": "1694"}
LEE = {"client_ssn": "987-65-4321", "client_name": "Lee Chen", "ts": "T", "employer_ein": "20-1112223",
       "employer_name": "Peak Logistics Inc", "employer_city": "Denver", "employer_state": "CO",
       "employer_zip": "80216", "box1_wages": "61000", "box2_fed_wh": "6100", "box3_ss_wages": "61000",
       "box4_ss_tax": "3782", "box5_medicare_wages": "61000", "box6_medicare_tax": "884.50",
       "box15_state": "CO", "box16_state_wages": "61000", "box17_state_tax": "2684"}
KIM = {"client_ssn": "555-11-2222", "client_name": "Kim Patel", "ts": "T", "employer_ein": "33-4445556",
       "employer_name": "Summit Dental PC", "employer_city": "Boulder", "employer_state": "CO",
       "employer_zip": "80302", "box1_wages": "47000", "box2_fed_wh": "3900", "box3_ss_wages": "47000",
       "box4_ss_tax": "2914", "box5_medicare_wages": "47000", "box6_medicare_tax": "681.50"}


def main():
    work = Path(tempfile.mkdtemp(prefix="drake_entry_test_"))
    shutil.copytree(PKG, work, dirs_exist_ok=True, ignore=shutil.ignore_patterns(
        "logs", "tracking", "evidence", "__pycache__", "*.pyc", "fake_state*"))
    st = json.load(open(work / "settings.json"))
    st["delays"] = {k: 0 for k in st["delays"]}
    json.dump(st, open(work / "settings.json", "w"), indent=2)
    run(work, "make-input", "--force")
    print(f"Test folder: {work}\n")

    print("1. Validation")
    bad = dict(LEE, employer_ein="20-111", box12a_code="ZZ", box12a_amt="100")
    write_input(work, [SAM_T, SAM_S, bad])
    code, out = run(work, "validate")
    check("bad EIN and invalid Box 12 code are errors", code == 1 and "must be 9 digits" in out
          and "not a valid code" in out, out)
    write_input(work, [SAM_T, SAM_S, LEE])
    code, out = run(work, "validate")
    check("clean input validates", code == 0, out)
    check("SS-tax warning raised for Sam's T W-2 (3400 vs 3410 expected)", "WARNING" in out and "Box 4" in out, out)

    print("\n2. Enter")
    code, out = run(work, "enter", "--client", "6789", ret="6789")
    inst = fake_return(work, "6789")
    check("two W-2s entered for Sam", code == 0 and len(inst) == 2, out)
    check("values landed in the right fields", inst[0].get("8") == "52000" and inst[0].get("2") == "841234567"
          and inst[1].get("0") == "S", str(inst[0]))
    check("tracker created", (work / "tracking" / "Entry_Tracker.xlsx").exists())

    print("\n3. Keying error stops entry")
    code, out = run(work, "enter", "--client", "4321", ret="4321", env={"FAKE_CORRUPT": "6100"})
    check("entry stopped on mismatch", code == 1 and "STOPPED" in out and "99999" in out, out)
    press_esc_in_drake(work, "4321")

    print("\n4. Re-running enter doesn't duplicate")
    code, out = run(work, "enter", "--client", "4321", ret="4321")
    check("incomplete W-2 skipped, not re-entered", "INCOMPLETE" in out and len(fake_return(work, "4321")) == 1, out)
    code, out = run(work, "enter", "--client", "6789", ret="6789")
    check("completed W-2s skipped", "already in Drake" in out and len(fake_return(work, "6789")) == 2, out)

    print("\n5. Verify then correct")
    code, out = run(work, "verify", "--client", "4321", ret="4321")
    check("verify finds the differences", "difference(s) logged" in out, out)
    code, out = run(work, "correct", "--client", "4321", ret="4321")
    check("correct fixes them", "correction(s) made and verified" in out, out)
    code, out = run(work, "verify", "--client", "4321", ret="4321")
    check("verify is clean afterwards", "Everything matches the source" in out, out)
    check("Drake now has the right Box 2", fake_return(work, "4321")[0].get("9") == "6100")

    print("\n6. Drake pre-fill is replaced and noted")
    write_input(work, [SAM_T, SAM_S, LEE, KIM])
    code, out = run(work, "enter", "--client", "2222", ret="2222", env={"FAKE_AUTOFILL": "2:3=SUMMIT DENTAL"})
    check("pre-filled employer name replaced with source", code == 0 and
          fake_return(work, "2222")[0].get("3") == "Summit Dental PC", out)
    log = (work / "logs" / "audit_log.jsonl").read_text()
    check("pre-fill noted in audit log", "pre-filled" in log)

    print("\n7. EIN corrected in the input after entry")
    lee2 = dict(LEE, employer_ein="20-1112299")
    write_input(work, [SAM_T, SAM_S, lee2, KIM])
    code, out = run(work, "enter", "--client", "4321", ret="4321")
    check("enter refuses (would duplicate)", "NOT entering" in out and len(fake_return(work, "4321")) == 1, out)
    code, out = run(work, "correct", "--client", "4321", ret="4321")
    check("correct matches it and fixes the EIN", "identity fields changed" in out and
          fake_return(work, "4321")[0].get("2") == "201112299", out)
    code, out = run(work, "verify", "--client", "4321", ret="4321")
    check("verify clean with the new EIN", "Everything matches" in out, out)

    print("\n8. Audit log integrity")
    code, out = run(work, "audit-verify")
    check("audit chain intact", code == 0 and "PASSED" in out, out)
    logp = work / "logs" / "audit_log.jsonl"
    lines = logp.read_text().splitlines()
    backup = list(lines)
    lines[5] = lines[5].replace('"QA"', '"XX"', 1) if '"QA"' in lines[5] else lines[5].replace("OK", "OX", 1)
    logp.write_text("\n".join(lines) + "\n")
    code, out = run(work, "audit-verify")
    check("tampering detected", code == 2 and "FAILED" in out, out)
    logp.write_text("\n".join(backup) + "\n")

    print("\n9. Tracker workbook")
    run(work, "build-tracker")
    from openpyxl import load_workbook
    wb = load_workbook(work / "tracking" / "Entry_Tracker.xlsx")
    check("all sheets present", all(s in wb.sheetnames for s in
                                    ["Summary", "Exceptions", "Corrections", "Tie-Out", "Run Log", "Audit Integrity"]),
          str(wb.sheetnames))
    check("one tab per client", sum(1 for s in wb.sheetnames if s[:4].isdigit()) == 3, str(wb.sheetnames))
    cor = [r for r in wb["Corrections"].iter_rows(min_row=5, values_only=True) if r[0]]
    check("corrections recorded with before/after", len(cor) >= 2 and any(r[5] == "99999" for r in cor), str(cor[:2]))

    print(f"\nAll {len(PASSED)} checks passed.  Test output kept in {work}")
    return work


if __name__ == "__main__":
    main()
