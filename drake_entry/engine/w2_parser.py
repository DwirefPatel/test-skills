"""
W-2 text parser (pdftotext -layout output, or a .txt of it).

Reads the standard IRS box labels. A value is taken from the same line, after the label,
or from the next lines by column position (the IRS layout prints labels on one line and
amounts under them). Multiple copies of the same W-2 (Copy B / C / 2) are collapsed;
different employers in one file become separate W-2s.

Anything it cannot read is left blank and listed in `missing`. Checkbox boxes (13) and
box 12 codes are best-effort, so every parsed W-2 is flagged for review against the PDF.
"""
import re

from .util import digits

AMT = r"\(?-?\$?[\d,]*\d(?:\.\d{2})?\)?"
AMT_RE = re.compile(rf"(?<![\w/-]){AMT}(?![\w/-])")
EIN_RE = re.compile(r"\b(\d{2}-\d{7})\b")
SSN_RE = re.compile(r"\b(\d{3}-\d{2}-\d{4}|X{3}-X{2}-\d{4}|\*{3}-\*{2}-\d{4})\b", re.I)

# field -> label regex (box number + words)
LABELS = [
    ("box1_wages", r"\b1\s+Wages,?\s+tips"),
    ("box2_fed_wh", r"\b2\s+Federal\s+income\s+tax\s+withheld"),
    ("box3_ss_wages", r"\b3\s+Social\s+security\s+wages"),
    ("box4_ss_tax", r"\b4\s+Social\s+security\s+tax\s+withheld"),
    ("box5_medicare_wages", r"\b5\s+Medicare\s+wages"),
    ("box6_medicare_tax", r"\b6\s+Medicare\s+tax\s+withheld"),
    ("box7_ss_tips", r"\b7\s+Social\s+security\s+tips"),
    ("box8_alloc_tips", r"\b8\s+Allocated\s+tips"),
    ("box10_dep_care", r"\b10\s+Dependent\s+care"),
    ("box11_nonqual", r"\b11\s+Nonqualified\s+plans"),
    ("box16_state_wages", r"\b16\s+State\s+wages"),
    ("box17_state_tax", r"\b17\s+State\s+income\s+tax"),
    ("box18_local_wages", r"\b18\s+Local\s+wages"),
    ("box19_local_tax", r"\b19\s+Local\s+income\s+tax"),
]
ALL_LABELS = re.compile("|".join(f"(?:{p})" for _, p in LABELS) +
                        r"|\b(?:12[a-d]|13|14|15|20)\s+[A-Z]|\b[a-f]\s+(?:Employer|Employee|Control)",
                        re.I)


def _amount(tok):
    return tok.replace("$", "").replace(",", "").strip("()")


def _label_spans(line):
    return [(m.start(), m.end()) for m in ALL_LABELS.finditer(line)]


def _value_for(lines, i, start, end):
    """Amount for the label at lines[i][start:end]."""
    line = lines[i]
    spans = _label_spans(line)
    nxt = min([s for s, _ in spans if s > start] + [len(line)])
    m = AMT_RE.search(line[end:nxt])
    if m:
        return _amount(m.group(0))
    # value under the label: same column band on the following lines
    left = start
    right = nxt
    for j in range(i + 1, min(i + 4, len(lines))):
        if ALL_LABELS.search(lines[j]):
            break
        for m in AMT_RE.finditer(lines[j]):
            center = (m.start() + m.end()) / 2
            if left - 2 <= center < right + 2:
                return _amount(m.group(0))
    return ""


BOX12_CODES = {"A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S",
               "T", "V", "W", "Y", "Z", "AA", "BB", "DD", "EE", "FF", "GG", "HH", "II"}
CODE_AMT = re.compile(rf"\b([A-Z]{{1,2}})\s+({AMT})(?![\w/-])")


def _box12(lines, letter):
    """Code and amount printed after (or under) the '12a'..'12d' label."""
    for i, line in enumerate(lines):
        m = re.search(rf"\b12{letter}\b", line)
        if not m:
            continue
        for j, seg in [(i, line[m.end():])] + [(k, lines[k][max(0, m.start() - 2):m.start() + 30])
                                               for k in range(i + 1, min(i + 3, len(lines)))]:
            for cm in CODE_AMT.finditer(seg):
                if cm.group(1) in BOX12_CODES:
                    return cm.group(1), _amount(cm.group(2))
        return "", ""
    return "", ""


def _box13(lines):
    """Box 13 checkboxes: an X printed under the 'Statutory' / 'Retirement' / 'Third-party' words."""
    out = {}
    for i, line in enumerate(lines):
        if not re.search(r"\b13\s+Statutory", line, re.I):
            continue
        cols = {name: m.start() for name, word in (("box13_statutory", "Statutory"),
                                                   ("box13_retirement", "Retirement"),
                                                   ("box13_sick_pay", "Third-party"))
                for m in [re.search(word, line, re.I)] if m}
        for nxt in lines[i + 1:i + 4]:
            for m in re.finditer(r"(?<!\w)X(?!\w)", nxt):
                for name, col in cols.items():
                    if col - 2 <= m.start() <= col + 10:
                        out[name] = "X"
        break
    return out


def _employer_block(lines):
    out = {}
    for i, line in enumerate(lines):
        if re.search(r"\bc\s+Employer'?s\s+name", line, re.I):
            col = re.search(r"\bc\s+Employer", line, re.I).start()
            block = []
            for nxt in lines[i + 1:i + 6]:
                seg = nxt[max(0, col - 2):].strip()
                seg = re.split(r"\s{3,}", seg)[0].strip() if seg else ""
                if not seg or ALL_LABELS.search(seg) or EIN_RE.search(seg):
                    if block:
                        break
                    continue
                block.append(seg)
            if block:
                out["employer_name"] = block[0]
            if len(block) >= 2:
                m = re.match(r"^(.*?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", block[-1])
                if m:
                    out.update(employer_city=m.group(1).strip(" ,"), employer_state=m.group(2),
                               employer_zip=m.group(3))
                    if len(block) >= 3:
                        out["employer_street"] = block[1]
                else:
                    out["employer_street"] = block[1]
            break
    return out


def parse_w2_page(text):
    lines = text.splitlines()
    row, missing = {}, []
    joined = "\n".join(lines)
    ein = EIN_RE.search(joined)
    if ein:
        row["employer_ein"] = ein.group(1)
    ssn = SSN_RE.search(joined)
    row["_employee_ssn"] = ssn.group(1) if ssn else ""
    row.update(_employer_block(lines))
    for field, pat in LABELS:
        for i, line in enumerate(lines):
            m = re.search(pat, line, re.I)
            if m:
                v = _value_for(lines, i, m.start(), m.end())
                if v:
                    row[field] = v
                break
    for letter in "abcd":
        code, amt = _box12(lines, letter)
        if code:
            row[f"box12{letter}_code"], row[f"box12{letter}_amt"] = code, amt
    m = re.search(r"\b15\s+State\b[^\n]*\n(?:[^\n]*\n)?\s*([A-Z]{2})\s+([\w-]{4,})\s", joined)
    if m:
        row["box15_state"], row["box15_state_id"] = m.group(1), m.group(2)
    row.update(_box13(lines))
    for need in ("employer_ein", "employer_name", "box1_wages"):
        if not row.get(need):
            missing.append(need)
    return row, missing


def is_w2(text):
    t = text.upper()
    return "WAGE AND TAX STATEMENT" in t and bool(re.search(r"\bW-?2\b", t))


def parse_w2(pages):
    """-> list of (row, missing). Copies of the same W-2 are collapsed."""
    out, seen = [], set()
    for text in pages:
        if not is_w2(text):
            continue
        row, missing = parse_w2_page(text)
        key = (digits(row.get("employer_ein")), row.get("box1_wages"), row.get("box2_fed_wh"))
        if key in seen:
            continue
        seen.add(key)
        out.append((row, missing))
    return out
