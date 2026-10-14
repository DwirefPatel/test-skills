"""
Run modes.

  validate    check the input workbook (no typing)
  enter       type new documents into Drake, verifying every field
  verify      read-only: read every field back from Drake and compare with the source
  correct     verify, then fix the differences you approve (logged before/after)
  calibrate   type field numbers into a TEST return to confirm the tab order
  map-fields  record each field's identity for the wrong-field check
"""
import json
import sys
import time
from collections import OrderedDict
from pathlib import Path

from forms import load_forms

from .audit import AuditLog
from .inputs import group_clients, load_rows
from .keyboard import Keyboard, WrongWindow
from .state import State
from .util import VERSION, digits, sha256_file, who
from .verify import Verifier

AFTER_STOP = ("In Drake, press Esc to leave the screen (what was typed is saved).\n"
              "  Then run:  python run.py correct --client {client}\n"
              "  It reads every field back, shows what differs, and fixes what you approve.")


class Halt(Exception):
    def __init__(self, message, guidance=""):
        super().__init__(message)
        self.guidance = guidance


class Session:
    def __init__(self, base, args):
        self.base = Path(base)
        self.args = args
        self.settings_path = self.base / "settings.json"
        self.settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        self.paths = {k: self.base / v for k, v in self.settings["paths"].items()}
        if getattr(args, "input", None):
            self.paths["input"] = Path(args.input)
        self.audit = AuditLog(self.paths["audit_log"])
        codes = ([c.strip().upper() for c in args.forms.split(",")] if getattr(args, "forms", None)
                 else self.settings["forms_enabled"])
        self.forms = load_forms(codes, self.base)
        self.run_id = time.strftime("%Y%m%d-%H%M%S")
        self.kb = self.vf = None
        self.dry = False
        self.initials = ""
        self.counts = OrderedDict(entered=0, verified=0, corrected=0, mismatches=0)

    # ------------------------------------------------------------------ plumbing
    def log(self, event, **data):
        return self.audit.append(event, self.run_id, **data)

    def rec(self, event, **data):
        """Log a data-entry event (skipped in dry runs)."""
        if not self.dry:
            self.log(event, **data)

    def ask(self, prompt, default=""):
        if getattr(self.args, "assume_yes", False):
            print(f"{prompt}{default}  [auto]")
            return default
        return input(prompt).strip()

    def start(self, mode, with_input=True):
        if with_input and not self.paths["input"].exists():
            raise SystemExit(f"Input workbook not found: {self.paths['input']}\nRun: python run.py make-input")
        user, machine = who()
        self.initials = (getattr(self.args, "initials", None) or self.settings.get("preparer_initials") or "").upper()
        while not self.initials:
            self.initials = self.ask("Your initials for the audit trail: ", "TEST").upper()
        data = dict(mode=mode, user=user, machine=machine, initials=self.initials, tool_version=VERSION,
                    settings_sha256=sha256_file(self.settings_path),
                    forms={c: f.version() for c, f in self.forms.items()})
        if with_input:
            data.update(input_file=str(self.paths["input"]), input_sha256=sha256_file(self.paths["input"]))
        self.log("run_start", **data)
        print(f"Run {self.run_id}  mode: {mode}  preparer: {self.initials}")

    def finish(self, status, tracker=True):
        summary = ", ".join(f"{v} {k}" for k, v in self.counts.items() if v)
        self.log("run_end", status=status, summary=summary)
        print(f"\nRun {status}. {summary}")
        if tracker:
            from .tracker import build_tracker
            path = build_tracker(self)
            if path:
                print(f"Tracker: {path}")

    def stop(self, exc):
        guidance = getattr(exc, "guidance", "")
        print(f"\nSTOPPED - {exc}")
        if guidance:
            print(f"  {guidance}")
        self.log("halt", message=str(exc))
        self.finish("stopped")
        sys.exit(1)

    def guarded(self, fn):
        try:
            fn()
        except (Halt, WrongWindow) as exc:
            self.stop(exc)
        except KeyboardInterrupt:
            self.stop(Halt("Cancelled with Ctrl+C."))
        except Exception as exc:
            if type(exc).__name__ == "FailSafeException":
                self.stop(Halt("Aborted: mouse moved to a screen corner."))
            self.log("halt", message=f"Unexpected error: {exc!r}")
            self.finish("error")
            raise

    def drive(self, dry=False):
        self.dry = dry
        self.kb = Keyboard(self.settings, dry_run=dry)
        self.vf = Verifier(self.settings, self.kb, self.paths["screenshots"])

    # ------------------------------------------------------------------ input
    def load_input(self):
        path = self.paths["input"]
        rows_by_form = OrderedDict((c, load_rows(path, f)) for c, f in self.forms.items())
        clients, errors = group_clients(rows_by_form)
        want = getattr(self.args, "client", None)
        if want:
            d = digits(want)
            clients = OrderedDict((s, c) for s, c in clients.items()
                                  if c["key"] == want or (len(d) >= 4 and s.endswith(d)))
        warnings = []
        for ssn, c in clients.items():
            for code, rows in c["forms"].items():
                form = self.forms[code]
                seen = {}
                for r in rows:
                    tag = f"{form.sheet} row {r['_row']} ({c['key']}, {form.describe(r)[:30]})"
                    e, w = form.validate(r, self.settings)
                    errors += [f"{tag}: {m}" for m in e]
                    warnings += [f"{tag}: {m}" for m in w]
                    fp = form.fingerprint(c["key"], r)
                    if fp in seen:
                        errors.append(f"{tag}: duplicate of row {seen[fp]} (same {', '.join(form.id_fields)})")
                    seen[fp] = r["_row"]
        return clients, errors, warnings

    def validated_clients(self, confirm=True, log_it=True):
        clients, errors, warnings = self.load_input()
        for w in warnings:
            print("  WARNING:", w)
        for e in errors:
            print("  ERROR:  ", e)
        if log_it:
            for w in warnings:
                self.log("validation", level="warning", message=w)
            for e in errors:
                self.log("validation", level="error", message=e)
        if errors:
            raise Halt(f"{len(errors)} error(s) in the input workbook.", "Fix them and run again.")
        if not clients:
            raise Halt("No matching rows in the input workbook.")
        if warnings and confirm:
            a = self.ask(f"\n{len(warnings)} warning(s) above. Accept and continue? (y/n): ", "y").lower()
            if a != "y":
                raise Halt("Stopped at validation warnings.")
            self.log("validation_accepted", count=len(warnings), by=self.initials)
        return clients

    # ------------------------------------------------------------------ field helpers
    def check_value(self, f, want):
        if not f.readback:
            return "Not read back", want
        actual = self.vf.read_value()
        if actual is None:
            return "Not read back", want
        return ("OK" if f.norm(actual) == f.norm(want) else "MISMATCH"), actual

    def assert_focus(self, form, name, base, raw, action):
        if self.vf.check_focus(form, name) is False:
            self.rec("field", **base, field=name, label=form.label(name), source=raw, typed="", drake="",
                     result="MISMATCH", note="Cursor was in a different field; nothing typed.", action=action)
            self.rec("instance_end", **base, complete=False)
            raise Halt(f"{base['client']} {form.instance_label} #{base['pos']}: the cursor is not in "
                       f"'{form.label(name)}'. The screen's tab order has shifted (pop-up or extra field).",
                       AFTER_STOP.format(client=base["client"]))

    def evidence(self, form, base):
        path, sha = self.vf.screenshot(self.run_id, f"{base['client'][-4:]}_{form.code}_{base['pos']}")
        if path:
            try:
                path = path.relative_to(self.base)
            except ValueError:
                pass
            self.rec("screenshot", **base, path=str(path), sha256=sha)

    # ------------------------------------------------------------------ validate
    def mode_validate(self):
        clients, errors, warnings = self.load_input()
        for w in warnings:
            print("  WARNING:", w)
        for e in errors:
            print("  ERROR:  ", e)
        print()
        for ssn, c in clients.items():
            docs = ", ".join(f"{len(rows)} {self.forms[k].instance_label}" for k, rows in c["forms"].items())
            print(f"  {c['key']:14s} {c['name'][:28]:28s} {docs}")
        print(f"\n{len(clients)} client(s), {len(errors)} error(s), {len(warnings)} warning(s).")
        print("Ready to enter." if not errors else "Fix the errors before entering.")
        return not errors

    # ------------------------------------------------------------------ enter
    def mode_enter(self):
        dry = getattr(self.args, "dry_run", False)
        self.start("enter (dry run)" if dry else "enter")
        self.guarded(lambda: self._enter(dry))

    def _enter(self, dry):
        clients = self.validated_clients()
        self.drive(dry)
        state = State(self.audit.events())
        for ssn, c in clients.items():
            plan = []
            for code, rows in c["forms"].items():
                form = self.forms[code]
                existing = state.instances_for(c["key"], code)
                known = {i["fingerprint"]: i for i in existing}
                fps = {form.fingerprint(c["key"], r) for r in rows}
                orphans = [i for i in existing if i["fingerprint"] not in fps]
                new = []
                for r in rows:
                    fp = form.fingerprint(c["key"], r)
                    if fp in known:
                        inst = known[fp]
                        why = ("already in Drake" if inst["complete"]
                               else "already in Drake but INCOMPLETE/MISMATCHED - use 'correct'")
                        print(f"  {c['key']} {form.instance_label} #{inst['pos']} {form.describe(r)}: {why} - skipped")
                        self.rec("skipped", client=c["key"], form=code, pos=inst["pos"], fingerprint=fp, reason=why)
                    else:
                        new.append((fp, r))
                if new and orphans:
                    print(f"  {c['key']} {form.instance_label}: {len(orphans)} document(s) in Drake no longer match "
                          f"the input (identity fields edited?). NOT entering new ones to avoid duplicates.\n"
                          f"    Run 'correct --client {c['key']}' first.")
                    self.rec("skipped", client=c["key"], form=code, pos=None, fingerprint="",
                             reason="Input identity fields changed; run correct first")
                    continue
                if new:
                    plan.append((form, new, (existing[-1]["pos"] + 1) if existing else 1, len(existing)))
            if not plan:
                continue

            print(f"\n=== {c['key']}  {c['name']} ===")
            for form, new, start, _ in plan:
                for n, (_, r) in enumerate(new):
                    print(f"    {form.instance_label} #{start + n}: {form.describe(r)}")
            if not dry:
                a = self.ask("Open this client's return in Drake, click the screen selector box,\n"
                             "then press Enter here (s = skip, q = quit): ", "").lower()
                if a == "q":
                    break
                if a == "s":
                    self.rec("skipped", client=c["key"], form="", pos=None, fingerprint="", reason="Skipped by user")
                    continue
                self.kb.countdown()
            for form, new, start, n_existing in plan:
                self.kb.open_screen(form)
                self.kb.next_instance(n_existing)
                for n, (fp, r) in enumerate(new):
                    if n:
                        self.kb.next_instance()
                    self.enter_instance(form, c, start + n, fp, r)
                self.kb.exit_screen()
            if not dry:
                self.tieout(c)

        self.finish("completed" if not dry else "dry run completed", tracker=not dry)

    def enter_instance(self, form, c, pos, fp, r):
        base = dict(client=c["key"], form=form.code, pos=pos)
        self.rec("instance_start", **base, client_name=c["name"], fingerprint=fp, desc=form.describe(r),
                 mode="enter", source_row=r["_row"])
        if self.dry:
            print(f"\n  -- {form.instance_label} #{pos}: {form.describe(r)}")
        clear_first = self.settings.get("entry", {}).get("clear_before_type", True)
        for idx, name in enumerate(form.field_order):
            if idx:
                self.kb.next_field()
            if name == "__SKIP__":
                continue
            f = form.by_name[name]
            raw = r.get(name, "")
            want = f.keys(raw)
            if self.dry:
                print(f"    {f.label[:44]:44s} -> {want!r}")
                continue
            self.assert_focus(form, name, base, raw, "enter")

            note, result, drake = "", None, None
            if clear_first:
                prior = self.vf.read_value() if f.readback else None
                if prior is not None and prior.strip() == want:
                    result, drake = "OK", prior                    # already correct (blank or pre-filled)
                    if want:
                        note = "Pre-filled by Drake; matched the source."
                else:
                    if prior and f.norm(prior) not in ("", "0.00"):
                        note = f"Drake had pre-filled {prior!r}; replaced with the source value."
                    self.kb.set_field(want)
            elif want:
                self.kb.write(want)
            if result is None:
                result, drake = self.check_value(f, want)
            if result == "Not read back" and not note:
                note = "Can't be read back; check the screenshot."

            self.rec("field", **base, field=name, label=f.label, source=raw, typed=want, drake=drake,
                     result=result, note=note, action="enter")
            if result == "MISMATCH":
                self.counts["mismatches"] += 1
                self.rec("instance_end", **base, complete=False)
                raise Halt(f"{c['key']} {form.instance_label} #{pos} '{f.label}': intended {want!r} "
                           f"but Drake shows {drake!r}.", AFTER_STOP.format(client=c["key"]))
        if not self.dry:
            self.evidence(form, base)
            self.rec("instance_end", **base, complete=True)
            self.counts["entered"] += 1

    def tieout(self, c):
        for code, rows in c["forms"].items():
            form = self.forms[code]
            totals = form.totals(rows, self.settings)
            if not totals:
                continue
            print(f"\n  Tie-out for {c['key']} {form.instance_label}s (all {len(rows)} in the input):")
            for label, amt, tie in totals:
                print(f"    {label:34s} {amt:>12,.2f}  -> {tie}")
            confirmed = "Not checked"
            if self.settings.get("entry", {}).get("prompt_tieout", True):
                a = self.ask("  Do these agree with the return in Drake (View mode)? y / n / s = check later: ",
                             "s").lower()
                confirmed = {"y": "Yes", "n": "No"}.get(a, "Not checked")
            for label, amt, tie in totals:
                self.log("tieout", client=c["key"], client_name=c["name"], form=code, label=label.strip(),
                         amount=round(amt, 2), tie_to=tie, confirmed=confirmed)
            if confirmed == "No":
                print(f"  Logged as an exception. Run 'verify --client {c['key']}' to find the difference.")

    # ------------------------------------------------------------------ verify / correct
    def mode_check(self, fix):
        self.start("correct" if fix else "verify")
        self.guarded(lambda: self._check(fix))

    def _check(self, fix):
        action = "correct" if fix else "verify"
        clients = self.validated_clients(confirm=False)
        self.drive()
        state = State(self.audit.events())
        for ssn, c in clients.items():
            targets = []
            for code, rows in c["forms"].items():
                form = self.forms[code]
                by_fp = OrderedDict((form.fingerprint(c["key"], r), r) for r in rows)
                insts = state.instances_for(c["key"], code)
                inst_fps = {i["fingerprint"] for i in insts}
                lost = [i for i in insts if i["fingerprint"] not in by_fp]
                unentered = [r for fp, r in by_fp.items() if fp not in inst_fps]
                for inst in insts:
                    if inst["fingerprint"] in by_fp:
                        targets.append((form, inst, by_fp[inst["fingerprint"]], False))
                if len(lost) == 1 and len(unentered) == 1:
                    print(f"  {c['key']} {form.instance_label} #{lost[0]['pos']}: identity fields changed in the "
                          f"input; comparing it with row {unentered[0]['_row']} ({form.describe(unentered[0])}).")
                    targets.append((form, lost[0], unentered[0], True))
                else:
                    for i in lost:
                        print(f"  {c['key']} {form.instance_label} #{i['pos']} ({i['desc']}) is in Drake but not in "
                              f"the input workbook - not checked.")
                        self.log("skipped", client=c["key"], form=code, pos=i["pos"], fingerprint=i["fingerprint"],
                                 reason="In Drake per audit log, not in input workbook")
                    for r in unentered:
                        print(f"  {c['key']} {form.describe(r)}: not entered yet - use 'enter'.")
            if not targets:
                continue

            print(f"\n=== {c['key']}  {c['name']} : {action} {len(targets)} document(s) ===")
            a = self.ask("Open this client's return in Drake, click the screen selector box,\n"
                         "then press Enter here (s = skip, q = quit): ", "").lower()
            if a == "q":
                break
            if a == "s":
                continue
            self.kb.countdown()
            diffs = self.read_pass(c, targets, action)
            if not diffs:
                print("  Everything matches the source.")
                self.tieout(c)
            elif fix:
                self.fix_pass(c, diffs, targets)
                self.tieout(c)
            else:
                print(f"  {len(diffs)} difference(s) logged. Run 'correct --client {c['key']}' to fix them.")
        self.finish("completed")

    def read_pass(self, c, targets, action):
        diffs = []
        by_form = OrderedDict()
        for form, inst, r, paired in targets:
            by_form.setdefault(form.code, []).append((inst, r, paired))
        for code, items in by_form.items():
            form = self.forms[code]
            self.kb.open_screen(form)
            cur = 1
            for inst, r, paired in sorted(items, key=lambda t: t[0]["pos"]):
                self.kb.next_instance(inst["pos"] - cur)
                cur = inst["pos"]
                base = dict(client=c["key"], form=code, pos=inst["pos"])
                self.log("instance_start", **base, client_name=c["name"], fingerprint=inst["fingerprint"],
                         desc=inst["desc"], mode=action)
                inst_diffs = []
                for idx, name in enumerate(form.field_order):
                    if idx:
                        self.kb.next_field()
                    if name == "__SKIP__":
                        continue
                    f = form.by_name[name]
                    raw = r.get(name, "")
                    want = f.keys(raw)
                    self.assert_focus(form, name, base, raw, action)
                    actual = self.vf.read_value() if f.readback else None
                    if actual is None:
                        self.log("field", **base, field=name, label=f.label, source=raw, typed=want, drake=want,
                                 result="Not read back", note="Can't be read back; check the screenshot.",
                                 action=action)
                        continue
                    ok = f.norm(actual) == f.norm(want)
                    if name == form.check_field and not ok and not paired:
                        self.log("field", **base, field=name, label=f.label, source=raw, typed=want, drake=actual,
                                 result="MISMATCH", note="WRONG DOCUMENT: this Drake instance belongs to another "
                                                         "document. Instance order differs from the log.",
                                 action=action)
                        self.log("instance_end", **base, complete=False)
                        self.kb.exit_screen()
                        raise Halt(f"{c['key']} {form.instance_label} #{inst['pos']} in Drake is not "
                                   f"{inst['desc']} (Drake shows {actual!r}).",
                                   "Someone may have deleted or reordered documents in Drake. Nothing was "
                                   "changed. Check the return manually, then run verify again.")
                    self.log("field", **base, field=name, label=f.label, source=raw, typed=want, drake=actual,
                             result="OK" if ok else "MISMATCH", note="" if ok else "Differs from the source.",
                             action=action)
                    if not ok:
                        inst_diffs.append(dict(form=form, inst=inst, r=r, paired=paired, idx=idx, name=name,
                                               f=f, raw=raw, want=want, before=actual))
                self.log("instance_end", **base, complete=not inst_diffs)
                self.counts["verified"] += 1
                self.counts["mismatches"] += len(inst_diffs)
                diffs += inst_diffs
            self.kb.exit_screen()
        return diffs

    def fix_pass(self, c, diffs, targets):
        print(f"\n  Differences for {c['key']}:")
        print(f"    {'#':>3}  {'Document':22s} {'Field':38s} {'Source':>16s}  {'Drake now':>16s}")
        for n, d in enumerate(diffs, 1):
            doc = f"{d['form'].instance_label} #{d['inst']['pos']}"
            print(f"    {n:>3}  {doc:22s} {d['f'].label[:38]:38s} {d['want']!s:>16s}  {d['before']!s:>16s}")
        a = self.ask("  Fix which? a = all, n = none, or numbers (e.g. 1,3): ", "a").lower()
        if a in ("n", ""):
            chosen = []
        elif a == "a":
            chosen = diffs
        else:
            picks = {int(x) for x in a.replace(" ", "").split(",") if x.isdigit()}
            chosen = [d for n, d in enumerate(diffs, 1) if n in picks]
        if not chosen:
            print("  No corrections made.")
            return
        reason = self.ask("  Reason (e.g. keying error / source document corrected): ", "Keying error") \
            or "Keying error"
        self.ask("  Click the screen selector box in this client's return, then press Enter here: ", "")
        self.kb.countdown()

        by_form = OrderedDict()
        for d in chosen:
            by_form.setdefault(d["form"].code, OrderedDict()).setdefault(d["inst"]["pos"], []).append(d)
        unchosen = {(d["form"].code, d["inst"]["pos"]) for d in diffs if d not in chosen}
        for code, by_pos in by_form.items():
            form = self.forms[code]
            self.kb.open_screen(form)
            cur_inst = 1
            for pos, items in sorted(by_pos.items()):
                self.kb.next_instance(pos - cur_inst)
                cur_inst, cur_idx = pos, 0
                base = dict(client=c["key"], form=code, pos=pos)
                all_ok = True
                for d in sorted(items, key=lambda x: x["idx"]):
                    self.kb.next_field(d["idx"] - cur_idx)
                    cur_idx = d["idx"]
                    self.assert_focus(form, d["name"], base, d["raw"], "correct")
                    self.kb.set_field(d["want"])
                    result, drake = self.check_value(d["f"], d["want"])
                    outcome = {"OK": "Fixed", "MISMATCH": "FAILED"}.get(result, "Typed - not read back")
                    self.log("correction", **base, field=d["name"], label=d["f"].label, before=d["before"],
                             after=drake, source=d["raw"], reason=reason, result=outcome)
                    self.log("field", **base, field=d["name"], label=d["f"].label, source=d["raw"], typed=d["want"],
                             drake=drake, result=result, note=f"Corrected from {d['before']!r} ({reason}).",
                             action="correct")
                    self.counts["corrected"] += 1
                    if result == "MISMATCH":
                        all_ok = False
                        self.log("instance_end", **base, complete=False)
                        raise Halt(f"Correction of '{d['f'].label}' did not take: Drake shows {drake!r}.",
                                   "Fix this field by hand in Drake, then run verify.")
                if items[0]["paired"] and all_ok:
                    r = items[0]["r"]
                    self.log("instance_rekey", **base, fingerprint=form.fingerprint(c["key"], r),
                             desc=form.describe(r))
                self.log("instance_end", **base, complete=all_ok and (code, pos) not in unchosen)
                self.evidence(form, base)
            self.kb.exit_screen()
        print(f"  {len(chosen)} correction(s) made and verified.")

    # ------------------------------------------------------------------ setup tools
    def pick_form(self):
        code = (getattr(self.args, "form", None) or next(iter(self.forms))).upper()
        if code not in self.forms:
            self.forms.update(load_forms([code], self.base))
        return self.forms[code]

    def mode_calibrate(self):
        form = self.pick_form()
        self.start(f"calibrate {form.code}", with_input=False)
        print(f"\nCALIBRATE {form.code}: open a TEST return and click the screen selector box.")
        print("The script opens the screen and types 1, 2, 3... into each field in tab order.\n")
        for i, name in enumerate(form.field_order, 1):
            print(f"  {i:3d}  {name if name == '__SKIP__' else form.label(name)}")
        self.ask("\nPress Enter here to start...", "")

        def go():
            self.drive()
            self.kb.countdown()
            self.kb.open_screen(form)
            print(f"\n  {'#':>3}  {'field':44s} read back")
            bad = 0
            for i, name in enumerate(form.field_order, 1):
                if i > 1:
                    self.kb.next_field()
                self.kb.set_field(str(i))
                got = self.vf.read_value()
                mark = "" if got is None else ("ok" if str(got).strip() == str(i) else "<-- CHECK")
                bad += mark == "<-- CHECK"
                print(f"  {i:3d}  {(name if name == '__SKIP__' else form.label(name))[:44]:44s} "
                      f"{'' if got is None else repr(got)} {mark}")
            print("\nNow compare the numbers on the Drake screen with this list and fix the tab order in "
                  f"forms/{form.cfg_path.name}. Then DELETE this test {form.instance_label} in Drake.")
            self.log("calibration", form=form.code, fields=len(form.field_order), readback_problems=bad)
            self.finish("completed", tracker=False)
        self.guarded(go)

    def mode_map_fields(self):
        form = self.pick_form()
        self.start(f"map-fields {form.code}", with_input=False)
        print(f"\nMAP FIELDS {form.code}: open a TEST return and click the screen selector box.")
        print("The script opens the screen and tabs through it WITHOUT typing.")
        self.ask("Press Enter here to start...", "")

        def go():
            self.drive()
            self.kb.countdown()
            self.kb.open_screen(form)
            ids, seen = {}, []
            for i, name in enumerate(form.field_order):
                if i:
                    self.kb.next_field()
                info = self.vf.focused()
                ident = self.vf.identity(info)
                if name != "__SKIP__":
                    ids[name] = ident
                seen.append(ident)
                desc = "(nothing detected)" if not info else (
                    f"id={info['auto_id'] or '-'} pos=({info['rel_x']},{info['rel_y']}) size={info['w']}x{info['h']}")
                print(f"  {name:22s} {desc}")
            self.kb.exit_screen()
            distinct = {json.dumps(x, sort_keys=True) for x in seen if x}
            usable = len(distinct) >= len(seen) * 0.9
            form.cfg["field_ids"] = ids
            form.cfg["focus_check"] = usable
            form.save_cfg()
            print(f"\n{'Mapped ' + str(len(distinct)) + ' distinct fields. Wrong-field check is ON.' if usable else 'Drake does not expose distinct fields on this PC; wrong-field check is OFF. Readback, screenshots and tie-outs still apply.'}")
            self.log("field_map_saved", form=form.code, focus_check=usable, distinct_fields=len(distinct),
                     config_sha256=sha256_file(form.cfg_path))
            self.finish("completed", tracker=False)
        self.guarded(go)
