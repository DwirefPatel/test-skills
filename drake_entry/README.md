# Drake Data-Entry Automation

Keyboard automation that types source documents into Drake Tax, checks every field, fixes what's wrong with your approval, and keeps an audit trail a reviewer can rely on. No AI and no API calls, so it uses no credits.

W-2 is the first form. Other forms plug into the same engine (see *Adding a form*).

## Folder layout

| Path | What it is |
|---|---|
| `run.py` | The only thing you run. `python run.py --help` lists every command. |
| `settings.json` | Global settings: keys, delays, tax year, verification options. |
| `forms/w2.py` | W-2 definition: fields, validation rules, tie-out totals. |
| `forms/W2.json` | W-2 screen setup: screen code, tab order, field map. |
| `engine/` | Shared engine: typing, verification, correction, audit log, tracker. |
| `input/Client_Input.xlsx` | Where you enter source-document data, one tab per form. |
| `logs/audit_log.jsonl` | Tamper-evident audit log. This is the permanent record, and it's created on first run. |
| `evidence/screenshots/` | A screenshot of every finished screen, grouped by run. |
| `tracking/Entry_Tracker.xlsx` | Review workbook, rebuilt from the audit log after every run. |
| `tests/` | Simulated Drake plus 23 end-to-end checks (`python tests/run_tests.py`). |

## One-time setup (on the Drake PC)

1. Install Python 3.10+ from python.org, then run:
   `pip install pandas openpyxl pyautogui pygetwindow uiautomation pyperclip`
2. Run `python run.py make-input` to create the input workbook.
3. Put your initials in `settings.json` under `preparer_initials`, or pass `--initials` each time.
4. Open a test return and run `python run.py calibrate --form W2`. It types 1, 2, 3… into each field. Compare the numbers on screen with the printed list and fix `field_order` in `forms/W2.json` until every number lands in the right place. Delete the test W-2 afterwards.
5. On the same test return, run `python run.py map-fields --form W2`. This records each field's identity so the script can detect when the cursor is in the wrong field. If Drake doesn't expose its fields to Windows, the command says so and turns that check off. The other checks still apply.
6. Run `enter` on the test return with a made-up W-2, then run `verify` and `correct`. Confirm all three behave on your Drake version before using real returns.

Re-run steps 4 and 5 after each annual Drake update.

## Daily workflow

1. Fill in `input/Client_Input.xlsx`.
2. Run `python run.py validate`. Fix every error. Warnings, such as Box 4 not being 6.2% of Box 3, are shown and must be accepted at entry, and they are logged.
3. Run `python run.py enter`. For each client, the script asks you to open the return and click the screen selector box. It then types each W-2 and checks every field as it goes:
   - **Wrong-field check:** before typing, it confirms the cursor is in the intended field.
   - **Pre-fill handling:** if Drake pre-filled a field (for example from its EIN database), the source value replaces it and the change is noted.
   - **Readback:** after typing, it reads the field back from Drake. On any mismatch it stops immediately.
   - **Evidence:** it takes a screenshot of each finished W-2.
4. **Tie-out.** After each client, the script prints total Box 1, Box 2 and Box 17 by state. Compare them with the return in Drake's View mode, then answer y, n, or s (check later). Your answer is logged.
5. Open `tracking/Entry_Tracker.xlsx` and work through it:
   - **Exceptions tab** first: resolve everything marked High, and review Medium and Info items.
   - **Client tabs:** the reviewer initials fields or documents.
   - **Summary tab:** fill in Reviewed by and Review date.

Useful options: `--client 6789` for one client (SSN last 4 or client ID), `--dry-run` to see keystrokes without typing, and `--forms W2`.

## When something goes wrong

**If entry stops:**
1. Press Esc in Drake. What was typed so far is saved.
2. Run `python run.py correct --client 6789`. It reads every field back, lists the differences, fixes the ones you approve, and records the value before and after each fix plus your reason. It then offers the tie-out.

**If you changed a return by hand, or want a second check:** run `python run.py verify --client 6789`. This mode is read-only. It compares Drake with the source and logs the result without changing anything.

**If a source document was wrong:** correct the input workbook, then run `correct`. This works even if the EIN or employer name changed. The script matches the document by position, tells you it did so, and fixes Drake.

**If you run `enter` twice:** documents already in Drake are skipped, so there are no duplicates. If identity fields were edited in the input, `enter` refuses to add anything and tells you to run `correct` first.

**To abort instantly:** move the mouse into any screen corner. The script also refuses to type if the active window isn't Drake.

## What makes it audit-ready

- **Complete record.** Every run, field typed, readback, correction, screenshot, tie-out answer, accepted warning and stop is logged with a timestamp, preparer initials, Windows user and machine.
- **Tamper-evident.** Each log entry contains the SHA-256 of the entry before it. `python run.py audit-verify` detects any edit, deletion or reordering. The tracker's Audit Integrity tab shows the latest check.
- **Proof of source.** Each run records the SHA-256 of the input workbook, the settings and each form definition. You can later prove exactly which data and configuration produced an entry.
- **Evidence.** Each screenshot's SHA-256 is logged, and the tracker links to it.
- **Separation of duties.** The tracker has preparer, reviewer and date fields, plus an Exceptions tab with Resolved By and Resolution columns.
- **Masked identifiers.** Logs and the tracker show only the last 4 digits of the SSN. The full SSN exists only in the input workbook.
- **Rebuildable.** The tracker is rebuilt entirely from the log, so it can't drift from what actually happened. Run `python run.py build-tracker` at any time.

## Adding a form (for example, 1099-INT)

1. Copy `forms/w2.py` to `forms/int1099.py`. Set `code`, `title`, `sheet`, `id_fields` and `check_field`, and list the fields with their types. Add rules and totals as needed.
2. Create `forms/99I.json` with the screen code and tab order.
3. Register the form in `forms/__init__.py` and add it to `forms_enabled` in `settings.json`.
4. Run `python run.py make-input --force`. This adds the form's tab and backs up your current workbook.
5. Run `calibrate`, `map-fields`, then test `enter`/`verify`/`correct` on a test return.

Nothing in `engine/` changes. The new form automatically gets its own section in each client tab (starting on a new page), its own rows on the Summary tab, and all the same checks. For each client, all enabled forms are entered in one pass.

## Security and retention

The input workbook, screenshots, tracker and log all contain taxpayer information. Keep this folder on the secured PC, within your WISP and FTC Safeguards Rule controls. Keep `logs/`, `evidence/` and `tracking/` with the client file for your retention period. Never edit `logs/audit_log.jsonl` by hand, because any edit will show as tampering.

## Limits to know

- This tool has been tested against a simulated Drake, not the real program. The key assumptions are that Page Down moves to the next or a new instance, Esc saves and leaves the screen, and Home followed by Shift+End selects a field. Calibration and a test return confirm these on your version, and each key can be changed in `settings.json`.
- Checkbox-type fields (TS, Box 13) usually can't be read back. They're marked "Not read back" and should be checked against the screenshot.
- The script matches documents in Drake by the order it entered them. If someone deletes or reorders W-2s in Drake by hand, `verify` detects the wrong document and stops without changing anything.
- You remain responsible for the return. Review the tracker and the tie-out before filing.
