import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tax_doc_requests"))
import send_tax_requests as t  # noqa: E402

SAMPLE = Path(__file__).resolve().parents[1] / "tax_doc_requests" / "sample_clients.xlsx"
SETTINGS = t.EmailSettings(tax_year=2025)
LINK = "https://www.dropbox.com/request/abc123"


def make_xlsx(tmp_path, rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    p = tmp_path / "clients.xlsx"
    wb.save(p)
    return p


def test_load_sample_and_names():
    clients = t.load_clients(SAMPLE)
    assert [c.display_name for c in clients] == [
        "Jane & John Sample", "Alex Example", "Priya Demo & Ravi Test"]
    assert [c.is_new for c in clients] == [False, True, True]
    assert [c.is_mfj for c in clients] == [True, False, True]


@pytest.mark.parametrize("full,expected", [
    ("Tejas & Anita Girase", ("Tejas", "Girase", "Anita", "Girase")),
    ("Tejas and Anita Girase", ("Tejas", "Girase", "Anita", "Girase")),
    ("Tejas Girase & Anita Patel", ("Tejas", "Girase", "Anita", "Patel")),
    ("Girase, Tejas", ("Tejas", "Girase", "", "")),
    ("Tejas Girase", ("Tejas", "Girase", "", "")),
])
def test_split_full_name(full, expected):
    assert t.split_full_name(full) == expected


def test_single_name_column_and_header_autodetect(tmp_path):
    p = make_xlsx(tmp_path, [
        ["2025 client list"], [],
        ["Name", "E-mail", "Client Type"],
        ["Tejas & Anita Girase", "tejasgirasr@gmail.com", "Returning"],
        [None, None, None],
    ])
    [c] = t.load_clients(p)
    assert c.first_name == "Tejas" and c.is_mfj and not c.is_new
    assert t.request_title(c, 2025) == "Tejas & Anita Girase - 2025 Tax Documents"


def test_missing_columns_message(tmp_path):
    p = make_xlsx(tmp_path, [["Foo", "Bar"], [1, 2]])
    with pytest.raises(ValueError, match="email"):
        t.load_clients(p)


def test_column_override(tmp_path):
    p = make_xlsx(tmp_path, [["Primary Email", "Given"], ["a@b.com", "Ann"]])
    [c] = t.load_clients(p, overrides={"email": "Primary Email", "first_name": "Given"})
    assert (c.email, c.first_name) == ("a@b.com", "Ann")


def test_returning_mfj_has_no_extra_asks():
    c = t.Client(row=2, email="x@y.com", first_name="Tejas", last_name="Girase",
                 spouse_first_name="Anita", filing_status="MFJ", is_new=False)
    subject, body = t.build_email(c, LINK, SETTINGS)
    assert subject == "Tejas & Anita Girase - 2025 Tax Documents"
    assert "Hi Tejas," in body and LINK in body
    assert "2024" not in body and "Driver" not in body and "also upload" not in body


def test_returning_single_has_no_extra_asks():
    c = t.Client(row=2, email="x@y.com", first_name="Sam", last_name="Solo",
                 filing_status="Single", is_new=False)
    _, body = t.build_email(c, LINK, SETTINGS)
    assert "2024" not in body and "Driver" not in body and "also upload" not in body


def test_new_mfj_client_asks_prior_return_and_both_licenses():
    c = t.Client(row=2, email="x@y.com", first_name="Tejas", last_name="Girase",
                 spouse_first_name="Anita", filing_status="Married Filing Jointly", is_new=True)
    _, body = t.build_email(c, LINK, SETTINGS)
    assert "2024 federal and state income tax returns" in body
    assert "for both Tejas and Anita" in body


def test_new_single_client_asks_one_license():
    c = t.Client(row=2, email="x@y.com", first_name="Alex", last_name="Example",
                 filing_status="Single", is_new=True)
    _, body = t.build_email(c, LINK, SETTINGS)
    assert "2024 federal" in body and "Driver&#x27;s License / State ID (front and back)<" in body


def test_validate_client():
    assert t.validate_client(t.Client(row=1, email="bad", first_name="A")) == "invalid email: bad"
    assert t.validate_client(t.Client(row=1, email="", first_name="A")) == "missing email"
    assert t.validate_client(t.Client(row=1, email="a@b.com; c@d.com", first_name="A")) is None


class RecordingOutlook:
    def __init__(self):
        self.drafts = []

    def save_draft(self, to, cc, subject, html_body, attachments):
        self.drafts.append((to, cc, subject))
        return "id"


class RecordingDropbox:
    def __init__(self):
        self.calls = []

    def get_or_create(self, title, destination, deadline=None):
        self.calls.append((title, destination))
        return LINK, True


def test_process_skips_done_and_logs_errors(tmp_path):
    clients = t.load_clients(SAMPLE)
    clients.append(t.Client(row=9, email="not-an-email", first_name="Bad"))
    dbx, ol = RecordingDropbox(), RecordingOutlook()
    recs = t.process(clients, dbx, ol, SETTINGS, "/File requests", [],
                     done={"alex@example.com"}, log=lambda *_: None)
    assert [d[2] for d in ol.drafts] == ["Jane & John Sample - 2025 Tax Documents",
                                         "Priya Demo & Ravi Test - 2025 Tax Documents"]
    assert ol.drafts[0][1] == ["john.sample@example.com"]  # spouse CC'd
    assert dbx.calls[0] == ("Jane & John Sample - 2025 Tax Documents",
                            "/File requests/Jane & John Sample")
    assert [r[9] for r in recs] == ["Done", "Done", "Error"]

    log = tmp_path / "log.xlsx"
    t.append_log(log, recs)
    assert t.load_done(log) == {"jane.sample@example.com", "priya.demo@example.com"}


def test_dry_run_end_to_end(tmp_path):
    clients = tmp_path / "c.xlsx"
    clients.write_bytes(SAMPLE.read_bytes())
    assert t.main(["run", str(clients), "--tax-year", "2025", "--dry-run"]) == 0
    previews = sorted((tmp_path / "c_previews").glob("*.html"))
    assert len(previews) == 3
    assert not (tmp_path / "c_requests_log.xlsx").exists()


def test_mfj_from_spouse_columns_without_filing_status(tmp_path):
    p = make_xlsx(tmp_path, [
        ["First Name", "Last Name", "Email", "Spouse First Name", "Spouse Email"],
        ["Tejas", "Girase", "tejasgirasr@gmail.com", "Anita", "anita@example.com"],
        ["Sam", "Solo", "sam@example.com", None, None],
    ])
    tejas, sam = t.load_clients(p)
    assert tejas.is_mfj and not tejas.is_new  # no client type column -> returning
    assert tejas.display_name == "Tejas & Anita Girase"
    assert not sam.is_mfj
