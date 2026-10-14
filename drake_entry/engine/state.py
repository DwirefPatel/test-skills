"""
Rebuilds the current picture from the audit log (the log is the single source of truth):
which documents are in Drake, each field's latest status, corrections, tie-outs, runs.
"""
from collections import OrderedDict


class State:
    def __init__(self, events):
        self.runs = OrderedDict()
        self.instances = OrderedDict()   # (client, form, pos) -> instance dict
        self.corrections, self.tieouts, self.validations, self.halts, self.skipped = [], [], [], [], []
        for ev in events:
            self._apply(ev)

    def _by(self, run_id):
        return self.runs.get(run_id, {}).get("initials", "")

    def _apply(self, ev):
        e, d, rid, ts = ev["event"], ev["data"], ev["run_id"], ev["ts"]
        if e == "run_start":
            self.runs[rid] = {**d, "run_id": rid, "started": ts, "ended": "", "status": "running"}
        elif e == "run_end":
            if rid in self.runs:
                self.runs[rid].update(ended=ts, status=d.get("status", ""), summary=d.get("summary", ""))
        elif e == "instance_start":
            key = (d["client"], d["form"], d["pos"])
            inst = self.instances.get(key)
            if d.get("mode") == "enter" or inst is None:
                inst = {"client": d["client"], "client_name": d.get("client_name", ""), "form": d["form"],
                        "pos": d["pos"], "fingerprint": d["fingerprint"], "desc": d.get("desc", ""),
                        "entered_ts": ts, "entered_by": self._by(rid), "source_row": d.get("source_row"),
                        "fields": OrderedDict(), "complete": False, "screenshot": "", "screenshot_sha": "",
                        "last_checked": ""}
                self.instances[key] = inst
            if d.get("client_name"):
                inst["client_name"] = d["client_name"]
        elif e == "field":
            inst = self.instances.get((d["client"], d["form"], d["pos"]))
            if inst is not None:
                inst["fields"][d["field"]] = {**d, "ts": ts, "by": self._by(rid)}
                if d.get("action") in ("verify", "correct"):
                    inst["last_checked"] = f"{ts} {self._by(rid)}".strip()
        elif e == "instance_end":
            inst = self.instances.get((d["client"], d["form"], d["pos"]))
            if inst is not None:
                inst["complete"] = bool(d.get("complete"))
        elif e == "instance_rekey":
            inst = self.instances.get((d["client"], d["form"], d["pos"]))
            if inst is not None:
                inst["fingerprint"], inst["desc"] = d["fingerprint"], d["desc"]
        elif e == "screenshot":
            inst = self.instances.get((d["client"], d["form"], d["pos"]))
            if inst is not None:
                inst["screenshot"], inst["screenshot_sha"] = d["path"], d["sha256"]
        elif e == "correction":
            self.corrections.append({**d, "ts": ts, "by": self._by(rid), "run_id": rid})
        elif e == "tieout":
            self.tieouts.append({**d, "ts": ts, "by": self._by(rid), "run_id": rid})
        elif e == "validation":
            self.validations.append({**d, "ts": ts, "run_id": rid})
        elif e == "halt":
            self.halts.append({**d, "ts": ts, "run_id": rid})
        elif e == "skipped":
            self.skipped.append({**d, "ts": ts, "run_id": rid})

    def instances_for(self, client, form):
        return sorted((i for (c, f, _), i in self.instances.items() if c == client and f == form),
                      key=lambda i: i["pos"])

    @staticmethod
    def status(inst):
        results = [f["result"] for f in inst["fields"].values()]
        if "MISMATCH" in results:
            return "Open mismatch"
        if not inst["complete"]:
            return "Incomplete"
        if "Not read back" in results:
            return "Complete - some fields not read back"
        return "Complete - verified"
