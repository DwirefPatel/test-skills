"""
Append-only, hash-chained audit log (logs/audit_log.jsonl).

Every action (run start, each field typed / read back / corrected, screenshots,
tie-outs, stops) is one JSON line. Each line stores the SHA-256 of the previous
line, so any edit, deletion, or reordering breaks the chain and is detected by
`python run.py audit-verify`. Nothing in this file is ever rewritten.
"""
import json
import os
import time
from pathlib import Path

from .util import sha256_text

GENESIS = "GENESIS"


def _digest(rec):
    body = {k: v for k, v in rec.items() if k != "hash"}
    return sha256_text(json.dumps(body, sort_keys=True, ensure_ascii=False))


class AuditLog:
    def __init__(self, path):
        self.path = Path(path)
        self.seq, self.last = 0, GENESIS
        if self.path.exists():
            with open(self.path, encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        rec = json.loads(line)
                        self.seq, self.last = rec["seq"], rec["hash"]

    def append(self, event, run_id, **data):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        rec = {"seq": self.seq + 1, "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
               "run_id": run_id, "event": event, "data": data, "prev": self.last}
        rec["hash"] = _digest(rec)
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            f.flush()
            os.fsync(f.fileno())
        self.seq, self.last = rec["seq"], rec["hash"]
        return rec

    def events(self):
        if not self.path.exists():
            return []
        with open(self.path, encoding="utf-8") as f:
            return [json.loads(l) for l in f if l.strip()]

    def verify(self):
        """Returns (ok, events_checked, problem_or_None)."""
        prev, n = GENESIS, 0
        for n, rec in enumerate(self.events(), 1):
            if rec.get("seq") != n:
                return False, n, f"line {n}: sequence number is {rec.get('seq')} (a line was removed or reordered)"
            if rec.get("prev") != prev:
                return False, n, f"line {n}: link to previous line is broken"
            if _digest(rec) != rec.get("hash"):
                return False, n, f"line {n}: contents were altered after being written"
            prev = rec["hash"]
        return True, n, None
