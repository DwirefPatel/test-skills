"""Create a Dropbox file request for each tax client and save an Outlook draft with the link.

For every row in the client list (Excel):

1. Create (or reuse) a Dropbox file request titled
   "<Tejas & Anita Girase> - <2025> Tax Documents".
2. Build the request email from the firm's standard wording. New clients are also asked for
   last year's tax return and a Driver's License / State ID (both spouses when MFJ).
   Returning clients get the standard email only.
3. Save the email as a draft in Outlook, addressed to the client, with the
   "Document checklist & Questionnaire.xlsx" attached. Nothing is sent.

Results go to a log workbook. Rows already logged as done are skipped on later runs, so the
script is safe to re-run after fixing a bad row.

Run ``python send_tax_requests.py --help`` for options, and see README.md for setup.
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_ATTACHMENT = HERE / "templates" / "Document checklist & Questionnaire.xlsx"
DEFAULT_TOKEN_FILE = HERE / ".dropbox_token.json"

DEFAULT_SENDER_NAME = "Sandhya Arora"
DEFAULT_FIRM_NAME = "Mile High CPA & Advisors"
DEFAULT_BOOKING_URL = (
    "https://outlook.office365.com/owa/calendar/MileHighCPAAdvisors@mh-cpas.com/bookings/"
)

# Header aliases for the client list. Matching ignores case, spaces, dashes and underscores,
# so "First Name", "first_name" and "FIRSTNAME" are the same column.
COLUMN_ALIASES: dict[str, list[str]] = {
    "email": ["email", "email address", "e-mail", "client email", "taxpayer email", "to"],
    "first_name": ["first name", "taxpayer first name", "tp first name", "first"],
    "last_name": ["last name", "taxpayer last name", "tp last name", "last", "surname"],
    "spouse_first_name": ["spouse first name", "spouse", "spouse name", "sp first name"],
    "spouse_last_name": ["spouse last name", "sp last name"],
    "spouse_email": ["spouse email", "spouse email address", "cc"],
    "filing_status": ["filing status", "fs"],
    "client_type": ["client type", "new or returning", "new/returning", "client status", "type",
                    "new client"],
    "display_name": ["display name", "request name", "client name", "name"],
}
REQUIRED = ("email", "first_name")

MFJ_VALUES = {"mfj", "married filing jointly", "married filing joint", "married jointly",
              "joint", "married-joint", "2"}
NEW_VALUES = {"new", "new client", "n", "yes", "y", "true", "1", "prospect"}
RETURNING_VALUES = {"returning", "existing", "returning client", "existing client", "old",
                    "repeat", "no", "false", "0", "r", "e"}


# --------------------------------------------------------------------------------------
# Client list
# --------------------------------------------------------------------------------------

@dataclass
class Client:
    row: int
    email: str
    first_name: str
    last_name: str = ""
    spouse_first_name: str = ""
    spouse_last_name: str = ""
    spouse_email: str = ""
    filing_status: str = ""
    is_new: bool = False
    display_name_override: str = ""
    warnings: list[str] = field(default_factory=list)

    @property
    def is_mfj(self) -> bool:
        fs = self.filing_status.strip().lower()
        if fs:
            return fs in MFJ_VALUES
        # No filing status given: a spouse name or spouse email in the list means MFJ.
        return bool(self.spouse_first_name or self.spouse_email)

    @property
    def display_name(self) -> str:
        """'Tejas & Anita Girase', or 'Tejas Girase & Anita Patel' when last names differ."""
        if self.display_name_override:
            return self.display_name_override
        first, last = self.first_name.strip(), self.last_name.strip()
        sp_first = self.spouse_first_name.strip()
        sp_last = self.spouse_last_name.strip() or last
        if self.is_mfj and sp_first:
            if sp_last and sp_last.lower() != last.lower():
                return f"{first} {last} & {sp_first} {sp_last}".strip()
            return f"{first} & {sp_first} {last}".strip()
        return f"{first} {last}".strip()


def _norm(text: object) -> str:
    return re.sub(r"[\s_\-/]+", " ", str(text or "")).strip().lower()


def _cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def map_columns(headers: list[object], overrides: dict[str, str] | None = None) -> dict[str, int]:
    """Return {field: column index} for the header row."""
    normalized = [_norm(h) for h in headers]
    mapping: dict[str, int] = {}
    for fld, aliases in COLUMN_ALIASES.items():
        wanted = [_norm(overrides[fld])] if overrides and fld in overrides else \
            [_norm(a) for a in aliases]
        for alias in wanted:
            if alias in normalized:
                mapping[fld] = normalized.index(alias)
                break
    missing = [f for f in REQUIRED if f not in mapping]
    # A single full-name column ("Tejas & Anita Girase") can stand in for first name.
    if "first_name" in missing and "display_name" in mapping:
        missing.remove("first_name")
    if missing:
        raise ValueError(
            f"Client list is missing required column(s): {', '.join(missing)}. "
            f"Headers found: {[h for h in headers if h]}. Use --column FIELD=HEADER to map them."
        )
    return mapping


def parse_client_type(value: str, default_new: bool) -> tuple[bool, str | None]:
    v = _norm(value)
    if not v:
        return default_new, None
    if v in NEW_VALUES:
        return True, None
    if v in RETURNING_VALUES:
        return False, None
    return default_new, f"unrecognised client type {value!r}; treated as " \
                        f"{'new' if default_new else 'returning'}"


def split_full_name(name: str) -> tuple[str, str, str, str]:
    """'Tejas & Anita Girase' -> ('Tejas', 'Girase', 'Anita', 'Girase').

    Also handles 'Tejas Girase & Anita Patel', 'Tejas and Anita Girase' and 'Girase, Tejas'.
    """
    name = re.sub(r"\s+", " ", name).strip()
    if "," in name and "&" not in name and " and " not in name.lower():
        last, _, first = name.partition(",")
        return first.strip().split(" ")[0], last.strip(), "", ""
    halves = re.split(r"\s*(?:&|\band\b)\s*", name, maxsplit=1, flags=re.I)
    tp = halves[0].split(" ")
    sp = halves[1].split(" ") if len(halves) > 1 else []
    sp_first, sp_last = (sp[0], " ".join(sp[1:])) if sp else ("", "")
    first = tp[0]
    last = " ".join(tp[1:]) or sp_last
    return first, last, sp_first, sp_last or (last if sp_first else "")


EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def load_clients(path: Path, sheet: str | None = None, header_row: int | None = None,
                 overrides: dict[str, str] | None = None,
                 default_new: bool = False) -> list[Client]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Find the header row: the first row that maps the required columns.
    candidates = [header_row - 1] if header_row else range(min(len(rows), 15))
    mapping, start, last_err = None, 0, None
    for i in candidates:
        try:
            mapping = map_columns(list(rows[i]), overrides)
            start = i + 1
            break
        except (ValueError, IndexError) as exc:
            last_err = exc
    if mapping is None:
        raise last_err or ValueError("Client list is empty")

    clients: list[Client] = []
    for offset, values in enumerate(rows[start:], start=start + 1):
        get = lambda f: _cell(values[mapping[f]]) if f in mapping and mapping[f] < len(values) \
            else ""  # noqa: E731
        email, first, full = get("email"), get("first_name"), get("display_name")
        if not email and not first and not full:
            continue  # blank row
        last, sp_first, sp_last = get("last_name"), get("spouse_first_name"), \
            get("spouse_last_name")
        if not first and full:
            first, last2, sp_first2, sp_last2 = split_full_name(full)
            last, sp_first, sp_last = last or last2, sp_first or sp_first2, sp_last or sp_last2
        is_new, warn = parse_client_type(get("client_type"), default_new)
        c = Client(
            row=offset, email=email, first_name=first, last_name=last,
            spouse_first_name=sp_first, spouse_last_name=sp_last,
            spouse_email=get("spouse_email"), filing_status=get("filing_status"),
            is_new=is_new, display_name_override=full if "first_name" in mapping else "",
        )
        if warn:
            c.warnings.append(warn)
        if "client_type" not in mapping:
            c.warnings.append("no client type column; treated as "
                              f"{'new' if default_new else 'returning'}")
        clients.append(c)
    return clients


def validate_client(c: Client) -> str | None:
    """Return an error message if the row cannot be processed."""
    emails = [e.strip() for e in re.split(r"[;,]", c.email) if e.strip()]
    if not emails:
        return "missing email"
    bad = [e for e in emails if not EMAIL_RE.match(e)]
    if bad:
        return f"invalid email: {', '.join(bad)}"
    if not c.first_name:
        return "missing first name"
    if c.is_mfj and c.is_new and not c.spouse_first_name:
        c.warnings.append("MFJ but no spouse name; spouse's Driver's License asked for generically")
    return None


# --------------------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------------------

@dataclass
class EmailSettings:
    tax_year: int
    sender_name: str = DEFAULT_SENDER_NAME
    firm_name: str = DEFAULT_FIRM_NAME
    booking_url: str = DEFAULT_BOOKING_URL
    attach_checklist: bool = True


DISCLAIMER = (
    "This electronic mail message contains confidential and legally privileged information "
    "intended only for the use of the recipient. If the reader of this message is not the "
    "intended recipient, the reader is hereby notified that any dissemination, distribution, "
    "copying or other use of this message is strictly prohibited and is hereby instructed to "
    "notify the sender immediately by return email and destroy this copy of this message."
)


def request_title(c: Client, tax_year: int) -> str:
    return f"{c.display_name} - {tax_year} Tax Documents"


def extra_items(c: Client, tax_year: int) -> list[str]:
    """Documents asked for on top of the checklist, for new clients only.

    New clients: last year's return plus Driver's License / State ID, for both spouses when
    MFJ and for the client alone otherwise. Returning clients (MFJ or not): nothing extra.
    """
    if not c.is_new:
        return []
    items = [f"Copy of your {tax_year - 1} federal and state income tax returns"]
    if c.is_mfj:
        spouse = c.spouse_first_name.strip() or "your spouse"
        items.append(f"Driver's License / State ID (front and back) for both "
                     f"{c.first_name.strip()} and {spouse}")
    else:
        items.append("Driver's License / State ID (front and back)")
    return items


def build_email(c: Client, link: str, s: EmailSettings) -> tuple[str, str]:
    """Return (subject, HTML body)."""
    title = request_title(c, s.tax_year)
    e = html.escape
    p = '<p style="margin:0 0 12px 0">{}</p>'.format

    parts = [p(f"Hi {e(c.first_name.strip())},")]
    if s.attach_checklist:
        parts.append(p("Please find the attached checklist for documents upload &amp; the "
                       "questionnaire."))
    extras = extra_items(c, s.tax_year)
    if extras:
        intro = "Since this is our first year working together, please also upload:"
        items = "".join(f"<li>{e(i)}</li>" for i in extras)
        parts.append(p(intro) + f'<ul style="margin:0 0 12px 0">{items}</ul>')
    parts.append(p("Please use below link to upload the documents"))
    parts.append(p(f'<a href="{e(link, quote=True)}">{e(title)}</a>'))
    parts.append(p(f"Kind Regards,<br><br>{e(s.sender_name)}<br>"
                   f"On Behalf of {e(s.firm_name)}."))
    if s.booking_url:
        parts.append(p("You can now book and manage appointments using our booking page.<br>"
                       f'<a href="{e(s.booking_url, quote=True)}">{e(s.booking_url)}</a>'))
    parts.append(p("<b>P.S.: The Highest Compliment We Can Receive Is The Referral Of Your "
                   "Friends, Family &amp; Business Associates. Thank you for your Trust!</b>"))
    parts.append(f'<p style="margin:0;font-size:8pt;color:#666666">{e(DISCLAIMER)}</p>')

    body = ('<html><body><div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
            + "\n".join(parts) + "</div></body></html>")
    return title, body


# --------------------------------------------------------------------------------------
# Dropbox
# --------------------------------------------------------------------------------------

DROPBOX_SCOPES = ["account_info.read", "file_requests.read", "file_requests.write",
                  "files.metadata.read", "files.content.write"]


def dropbox_authorize(app_key: str, token_file: Path) -> None:
    """One-time: get a long-lived refresh token (PKCE, no app secret needed)."""
    from dropbox import DropboxOAuth2FlowNoRedirect

    flow = DropboxOAuth2FlowNoRedirect(app_key, use_pkce=True, token_access_type="offline",
                                       scope=DROPBOX_SCOPES)
    print("1. Open this URL and click Allow:\n  ", flow.start())
    code = input("2. Paste the authorization code here: ").strip()
    result = flow.finish(code)
    token_file.write_text(json.dumps({"app_key": app_key,
                                      "refresh_token": result.refresh_token}))
    print(f"Saved refresh token to {token_file}. Keep this file private.")


class DropboxRequests:
    def __init__(self, app_key: str, refresh_token: str, team_space: bool = False):
        import dropbox

        self._dropbox = dropbox
        dbx = dropbox.Dropbox(app_key=app_key, oauth2_refresh_token=refresh_token)
        if team_space:
            # Business accounts: create folders/requests in the team space, not the
            # member's personal folder.
            root = dbx.users_get_current_account().root_info.root_namespace_id
            dbx = dbx.with_path_root(dropbox.common.PathRoot.root(root))
        self.dbx = dbx
        self._existing: dict[tuple[str, str], object] | None = None

    def _load_existing(self) -> dict[tuple[str, str], object]:
        if self._existing is None:
            res = self.dbx.file_requests_list_v2()
            reqs = list(res.file_requests)
            while res.has_more:
                res = self.dbx.file_requests_list_continue(res.cursor)
                reqs.extend(res.file_requests)
            self._existing = {(r.title.strip().lower(), (r.destination or "").lower()): r
                              for r in reqs}
        return self._existing

    def get_or_create(self, title: str, destination: str,
                      deadline: datetime | None = None) -> tuple[str, bool]:
        """Return (url, created). Reuses an open request with the same title and folder."""
        fr = self._dropbox.file_requests
        existing = self._load_existing().get((title.strip().lower(), destination.lower()))
        if existing is not None and existing.is_open:
            return existing.url, False
        self._ensure_folder(destination)
        req = self.dbx.file_requests_create(
            title=title, destination=destination, open=True,
            deadline=fr.FileRequestDeadline(deadline=deadline) if deadline else None,
        )
        self._existing[(title.strip().lower(), destination.lower())] = req
        return req.url, True

    def _ensure_folder(self, path: str) -> None:
        try:
            self.dbx.files_create_folder_v2(path)
        except self._dropbox.exceptions.ApiError as exc:
            err = exc.error
            if not (err.is_path() and err.get_path().is_conflict()):
                raise


def safe_folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "-", name).strip(" .")


# --------------------------------------------------------------------------------------
# Outlook
# --------------------------------------------------------------------------------------

def _recipients(value: str) -> list[str]:
    return [e.strip() for e in re.split(r"[;,]", value) if e.strip()]


class OutlookDesktop:
    """Classic Outlook for Windows via COM (pip install pywin32). Saves to Drafts."""

    def __init__(self, send_as: str | None = None):
        import win32com.client  # type: ignore

        self.app = win32com.client.Dispatch("Outlook.Application")
        self.send_as = send_as

    def save_draft(self, to: list[str], cc: list[str], subject: str, html_body: str,
                   attachments: list[Path]) -> str:
        mail = self.app.CreateItem(0)  # olMailItem
        mail.To = "; ".join(to)
        if cc:
            mail.CC = "; ".join(cc)
        if self.send_as:
            mail.SentOnBehalfOfName = self.send_as
        mail.Subject = subject
        mail.HTMLBody = html_body
        for a in attachments:
            mail.Attachments.Add(str(a))
        mail.Save()
        return mail.EntryID


class OutlookGraph:
    """Outlook / Microsoft 365 via Microsoft Graph (works with new Outlook, Mac and web).

    Needs an Azure app registration (public client, delegated Mail.ReadWrite).
    """

    GRAPH = "https://graph.microsoft.com/v1.0"

    def __init__(self, client_id: str, tenant: str = "organizations",
                 cache_file: Path = HERE / ".msal_cache.json", mailbox: str | None = None):
        import msal
        import requests

        self.requests = requests
        cache = msal.SerializableTokenCache()
        if cache_file.exists():
            cache.deserialize(cache_file.read_text())
        app = msal.PublicClientApplication(
            client_id, authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=cache)
        scopes = ["Mail.ReadWrite"] + (["Mail.ReadWrite.Shared"] if mailbox else [])
        accounts = app.get_accounts()
        token = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
        if not token:
            flow = app.initiate_device_flow(scopes=scopes)
            print(flow["message"])
            token = app.acquire_token_by_device_flow(flow)
        if "access_token" not in token:
            raise RuntimeError(f"Microsoft sign-in failed: {token.get('error_description')}")
        if cache.has_state_changed:
            cache_file.write_text(cache.serialize())
        self.headers = {"Authorization": f"Bearer {token['access_token']}"}
        self.base = f"{self.GRAPH}/users/{mailbox}" if mailbox else f"{self.GRAPH}/me"

    def save_draft(self, to: list[str], cc: list[str], subject: str, html_body: str,
                   attachments: list[Path]) -> str:
        def addr(x):
            return [{"emailAddress": {"address": a}} for a in x]

        msg = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": addr(to),
            "ccRecipients": addr(cc),
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": a.name,
                "contentBytes": base64.b64encode(a.read_bytes()).decode(),
            } for a in attachments],
        }
        r = self.requests.post(f"{self.base}/messages", headers=self.headers, json=msg,
                               timeout=60)
        if r.status_code >= 300:
            raise RuntimeError(f"Graph error {r.status_code}: {r.text[:300]}")
        return r.json()["id"]


class PreviewOnly:
    """--dry-run: write each draft as an .html file instead of touching Outlook."""

    def __init__(self, out_dir: Path):
        self.out_dir = out_dir
        out_dir.mkdir(parents=True, exist_ok=True)

    def save_draft(self, to, cc, subject, html_body, attachments) -> str:
        path = self.out_dir / f"{safe_folder_name(subject)}.html"
        header = (f"<p><b>To:</b> {html.escape('; '.join(to))}<br><b>CC:</b> "
                  f"{html.escape('; '.join(cc))}<br><b>Subject:</b> {html.escape(subject)}<br>"
                  f"<b>Attachments:</b> {html.escape(', '.join(a.name for a in attachments))}"
                  "</p><hr>")
        path.write_text(html_body.replace("<body>", "<body>" + header, 1), encoding="utf-8")
        return str(path)


# --------------------------------------------------------------------------------------
# Results log
# --------------------------------------------------------------------------------------

LOG_HEADERS = ["Row", "Client", "Email", "Client Type", "MFJ", "Request Title",
               "Dropbox Link", "Request", "Draft", "Status", "Notes", "Timestamp"]


def load_done(log_path: Path) -> set[str]:
    """Emails already logged with status 'Done'."""
    if not log_path.exists():
        return set()
    import openpyxl

    wb = openpyxl.load_workbook(log_path, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = next(rows, ())
    try:
        e_i, s_i = header.index("Email"), header.index("Status")
    except ValueError:
        return set()
    done = {str(r[e_i]).strip().lower() for r in rows if r[s_i] == "Done" and r[e_i]}
    wb.close()
    return done


def append_log(log_path: Path, records: list[list]) -> None:
    import openpyxl
    from openpyxl.styles import Font

    if log_path.exists():
        wb = openpyxl.load_workbook(log_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Requests"
        ws.append(LOG_HEADERS)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        for col, width in zip("ABCDEFGHIJKL", (6, 28, 32, 11, 6, 44, 52, 10, 10, 9, 50, 20)):
            ws.column_dimensions[col].width = width
    for rec in records:
        ws.append(rec)
    wb.save(log_path)


# --------------------------------------------------------------------------------------
# Main loop
# --------------------------------------------------------------------------------------

def process(clients: list[Client], dropbox_client, outlook, settings: EmailSettings,
            dropbox_root: str, attachments: list[Path], done: set[str],
            deadline: datetime | None = None, cc_spouse: bool = True,
            log=print) -> list[list]:
    records = []
    for c in clients:
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        title = request_title(c, settings.tax_year) if c.first_name else ""
        base = [c.row, c.display_name, c.email, "New" if c.is_new else "Returning",
                "Yes" if c.is_mfj else "No", title]

        if c.email.strip().lower() in done:
            log(f"  row {c.row}: {c.display_name} - already done, skipped")
            continue
        err = validate_client(c)
        if err:
            log(f"  row {c.row}: SKIPPED - {err}")
            records.append(base + ["", "", "", "Error", err, stamp])
            continue

        link, req_state, draft_state = "", "", ""
        try:
            folder = f"{dropbox_root.rstrip('/')}/{safe_folder_name(c.display_name)}"
            link, created = dropbox_client.get_or_create(title, folder, deadline)
            req_state = "Created" if created else "Reused"
            subject, body = build_email(c, link, settings)
            cc = _recipients(c.spouse_email) if cc_spouse else []
            outlook.save_draft(_recipients(c.email), cc, subject, body, attachments)
            draft_state = "Saved"
            notes = "; ".join(c.warnings)
            records.append(base + [link, req_state, draft_state, "Done", notes, stamp])
            log(f"  row {c.row}: {c.display_name} <{c.email}> - request {req_state.lower()}, "
                f"draft saved{' (' + notes + ')' if notes else ''}")
        except Exception as exc:  # keep going; the log shows what to fix
            msg = f"{type(exc).__name__}: {exc}"
            records.append(base + [link, req_state, draft_state, "Error", msg, stamp])
            log(f"  row {c.row}: {c.display_name} - ERROR {msg}")
    return records


class FakeDropbox:
    """--dry-run stand-in: no API calls, placeholder link."""

    def get_or_create(self, title, destination, deadline=None):
        return "https://www.dropbox.com/request/DRY-RUN-PLACEHOLDER", True


def parse_args(argv=None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Create Dropbox file requests and Outlook drafts for tax clients.")
    sub = ap.add_subparsers(dest="cmd")

    auth = sub.add_parser("dropbox-auth", help="one-time Dropbox sign-in")
    auth.add_argument("--app-key", default=os.environ.get("DROPBOX_APP_KEY"))

    run = sub.add_parser("run", help="process the client list")
    run.add_argument("clients", type=Path, help="client list .xlsx")
    run.add_argument("--tax-year", type=int, default=datetime.now().year - 1,
                     help="tax year (default: last year)")
    run.add_argument("--sheet", help="worksheet name (default: first sheet)")
    run.add_argument("--header-row", type=int, help="header row number (default: auto-detect)")
    run.add_argument("--column", action="append", default=[], metavar="FIELD=HEADER",
                     help=f"map a column, e.g. --column email='Primary Email'. "
                          f"Fields: {', '.join(COLUMN_ALIASES)}")
    run.add_argument("--default-new", action="store_true",
                     help="treat rows with a blank client type as NEW (default: returning)")
    run.add_argument("--only", action="append", default=[], metavar="EMAIL",
                     help="process only this email (repeatable)")
    run.add_argument("--limit", type=int, help="process at most N rows (for a test batch)")

    run.add_argument("--dropbox-root", default=None,
                     help="Dropbox parent folder; each client gets <root>/<client name> "
                          "(default: '/File requests')")
    run.add_argument("--team-space", action="store_true",
                     help="Dropbox Business: create in the team space")
    run.add_argument("--deadline", help="optional upload deadline, YYYY-MM-DD (paid plans)")

    run.add_argument("--outlook", choices=["desktop", "graph"], default="desktop",
                     help="desktop = classic Outlook for Windows (default); graph = Microsoft "
                          "365 API (new Outlook / Mac)")
    run.add_argument("--send-as", help="desktop: 'Send on behalf of' mailbox address")
    run.add_argument("--graph-client-id", default=os.environ.get("GRAPH_CLIENT_ID"))
    run.add_argument("--graph-tenant", default=os.environ.get("GRAPH_TENANT", "organizations"))
    run.add_argument("--graph-mailbox", help="graph: shared mailbox to save drafts in")

    run.add_argument("--attachment", type=Path, action="append",
                     help=f"file to attach (default: {DEFAULT_ATTACHMENT.name})")
    run.add_argument("--no-attachment", action="store_true", help="attach nothing")
    run.add_argument("--no-cc-spouse", action="store_true",
                     help="don't CC the spouse email column")
    run.add_argument("--sender-name", default=DEFAULT_SENDER_NAME)
    run.add_argument("--firm-name", default=DEFAULT_FIRM_NAME)
    run.add_argument("--booking-url", default=DEFAULT_BOOKING_URL,
                     help="booking page link ('' to leave out)")

    run.add_argument("--log", type=Path, help="results workbook (default: "
                     "<clients>_requests_log.xlsx)")
    run.add_argument("--force", action="store_true",
                     help="redo rows already logged as Done")
    run.add_argument("--dry-run", action="store_true",
                     help="no Dropbox or Outlook calls; write HTML previews instead")
    run.add_argument("--yes", action="store_true", help="don't ask for confirmation")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    token_file = Path(os.environ.get("DROPBOX_TOKEN_FILE", DEFAULT_TOKEN_FILE))

    if args.cmd == "dropbox-auth":
        if not args.app_key:
            print("Pass --app-key or set DROPBOX_APP_KEY.")
            return 1
        dropbox_authorize(args.app_key, token_file)
        return 0
    if args.cmd != "run":
        parse_args(["--help"])
        return 1

    overrides = {}
    for item in args.column:
        fld, _, header = item.partition("=")
        if fld not in COLUMN_ALIASES or not header:
            print(f"Bad --column {item!r}. Fields: {', '.join(COLUMN_ALIASES)}")
            return 1
        overrides[fld] = header

    clients = load_clients(args.clients, args.sheet, args.header_row, overrides,
                           args.default_new)
    if args.only:
        wanted = {e.lower() for e in args.only}
        clients = [c for c in clients if c.email.strip().lower() in wanted]
    if args.limit:
        clients = clients[: args.limit]

    log_path = args.log or args.clients.with_name(f"{args.clients.stem}_requests_log.xlsx")
    done = set() if args.force else load_done(log_path)
    todo = [c for c in clients if c.email.strip().lower() not in done]

    attachments = [] if args.no_attachment else (args.attachment or [DEFAULT_ATTACHMENT])
    for a in attachments:
        if not a.exists():
            print(f"Attachment not found: {a}")
            return 1
    settings = EmailSettings(tax_year=args.tax_year, sender_name=args.sender_name,
                             firm_name=args.firm_name, booking_url=args.booking_url,
                             attach_checklist=bool(attachments))
    root = args.dropbox_root or "/File requests"
    deadline = datetime.strptime(args.deadline, "%Y-%m-%d") if args.deadline else None

    n_new = sum(c.is_new for c in todo)
    print(f"{len(clients)} client(s) in list, {len(clients) - len(todo)} already done, "
          f"{len(todo)} to process ({n_new} new, {len(todo) - n_new} returning).")
    print(f"Tax year {args.tax_year}; Dropbox folder {root}; Outlook: "
          f"{'PREVIEW ONLY (dry run)' if args.dry_run else args.outlook}.")
    if not todo:
        return 0
    if not (args.yes or args.dry_run):
        if input("Create file requests and save drafts? [y/N] ").strip().lower() != "y":
            return 1

    if args.dry_run:
        dbx = FakeDropbox()
        outlook = PreviewOnly(log_path.with_name(f"{args.clients.stem}_previews"))
    else:
        app_key = os.environ.get("DROPBOX_APP_KEY")
        refresh = os.environ.get("DROPBOX_REFRESH_TOKEN")
        if token_file.exists() and not refresh:
            saved = json.loads(token_file.read_text())
            app_key, refresh = app_key or saved["app_key"], saved["refresh_token"]
        if not (app_key and refresh):
            print("No Dropbox credentials. Run: python send_tax_requests.py dropbox-auth "
                  "--app-key YOUR_APP_KEY")
            return 1
        dbx = DropboxRequests(app_key, refresh, team_space=args.team_space)
        if args.outlook == "graph":
            if not args.graph_client_id:
                print("--outlook graph needs --graph-client-id (or GRAPH_CLIENT_ID).")
                return 1
            outlook = OutlookGraph(args.graph_client_id, args.graph_tenant,
                                   mailbox=args.graph_mailbox)
        else:
            outlook = OutlookDesktop(send_as=args.send_as)

    records = process(todo, dbx, outlook, settings, root, attachments, done, deadline,
                      cc_spouse=not args.no_cc_spouse)
    if not args.dry_run:
        append_log(log_path, records)
        print(f"Log: {log_path}")
    else:
        print(f"Previews: {outlook.out_dir}")
    errors = sum(r[9] == "Error" for r in records)
    print(f"Done: {len(records) - errors} draft(s) saved, {errors} error(s).")
    return 2 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
