"""Unit test skrip sinkronisasi dengan Jira palsu (tanpa jaringan)."""
import hashlib
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import sync  # noqa: E402

HEAD = ["No", "Modul", "Feature", "Test Case", "Test Case ID", "Test Type", "Priority", "Requirement ID",
        "Pre-Condition", "Test Step", "Test Data", "Expected Result", "Current Result", "Status", "Severity",
        "Bug ID", "Bug Status", "Notes", "Capture"]


class FakeJira:
    base = "https://contoh.atlassian.net"

    def __init__(self):
        self.issues = {}
        self.created = []
        self.attached = []
        self.links = []

    def link_to_stories(self, key, requirements):
        self.links.append((key, requirements))
        return [f"STORY-{r}" for r in requirements]

    def browse_url(self, key):
        return f"{self.base}/browse/{key}"

    def find_open_bug(self, tc_id):
        for issue, labels in self.issues.values():
            if tc_id in labels and not issue.done:
                return issue
        return None

    def create_bug(self, tc, workbook_name):
        key = f"PB-{len(self.issues) + 1}"
        issue = sync.Issue(key=key, summary=f"[{tc.id}] {tc.get('Test Case')}", status="To Do")
        self.issues[key] = (issue, {tc.id})
        self.created.append((tc.id, tc.severity, tc.get("Current Result")))
        return issue

    def attach(self, key, path):
        self.attached.append((key, path.name))

    def get_issue(self, key):
        return self.issues[key][0] if key in self.issues else None

    def set_status(self, key, status, done):
        self.issues[key][0].status, self.issues[key][0].done = status, done


def make_workbook(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Requirements"
    ws.append(["REQ ID", "Requirement", "Modul", "Prioritas"])
    ws.append(["REQ-06", "Transfer memindahkan dana dengan benar", "Transfer", "High"])
    ws.append(["REQ-07", "Nominal harus positif", "Transfer", "High"])
    ws.append(["REQ-99", "Requirement tanpa test case", "-", "Low"])
    ws = wb.create_sheet("Transfer")
    ws["A1"] = "TEST CASE - ParaBank · Transfer"
    ws.append([])
    ws.append(HEAD)
    for i, r in enumerate(rows, 1):
        row = dict.fromkeys(HEAD)
        row.update({"No": i, "Modul": "Transfer", "Feature": "Transfer Funds", "Test Type": "Negative"}, **r)
        ws.append([row[h] for h in HEAD])
    wb.create_sheet("RTM")
    wb.create_sheet("Bug Report")
    wb.save(path)


def cell(path, tc_id, col):
    ws = openpyxl.load_workbook(path)["Transfer"]
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[HEAD.index("Test Case ID")] == tc_id:
            return r[HEAD.index(col)]


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


@pytest.fixture
def wb_path(tmp_path):
    p = tmp_path / "QA.xlsx"
    (tmp_path / "evidence").mkdir()
    (tmp_path / "evidence" / "PB-TRF-06.png").write_bytes(b"png")
    make_workbook(p, [
        {"Test Case": "Transfer nominal valid", "Test Case ID": "PB-TRF-01", "Test Type": "Positive",
         "Requirement ID": "REQ-06", "Status": "Passed"},
        {"Test Case": "Transfer nominal negatif", "Test Case ID": "PB-TRF-06", "Requirement ID": "REQ-07",
         "Status": "Failed", "Severity": "High", "Current Result": "Transfer -50 diterima",
         "Capture": "evidence/PB-TRF-06.png"},
        {"Test Case": "Transfer nominal 0", "Test Case ID": "PB-TRF-07", "Requirement ID": "REQ-07, REQ-06",
         "Status": None},
    ])
    return p


def test_baris_failed_dibuatkan_bug_dan_bug_id_ditulis(wb_path):
    jira = FakeJira()
    sync.run(wb_path, None, jira, root=wb_path.parent)
    assert jira.created == [("PB-TRF-06", "High", "Transfer -50 diterima")]
    assert jira.attached == [("PB-1", "PB-TRF-06.png")]
    assert jira.links == [("PB-1", ["REQ-07"])]
    assert cell(wb_path, "PB-TRF-06", "Bug ID") == "PB-1"
    assert cell(wb_path, "PB-TRF-06", "Bug Status") == "To Do"


def test_dijalankan_ulang_tidak_membuat_tiket_ganda_dan_tidak_menyimpan_ulang(wb_path):
    jira = FakeJira()
    sync.run(wb_path, None, jira, root=wb_path.parent)
    before = digest(wb_path)
    data = sync.run(wb_path, None, jira, root=wb_path.parent)
    assert len(jira.created) == 1
    assert not data["_result"].changed
    assert digest(wb_path) == before


def test_tiket_terbuka_yang_sudah_ada_dipakai_ulang(wb_path):
    jira = FakeJira()
    existing = sync.Issue(key="PB-50", summary="lama", status="In Progress")
    jira.issues["PB-50"] = (existing, {"PB-TRF-06"})
    data = sync.run(wb_path, None, jira, root=wb_path.parent)
    assert jira.created == []
    assert data["_result"].reused == [("PB-TRF-06", "PB-50")]
    assert cell(wb_path, "PB-TRF-06", "Bug ID") == "PB-50"


def test_status_jira_ditulis_balik_dan_memicu_retest(wb_path):
    jira = FakeJira()
    sync.run(wb_path, None, jira, root=wb_path.parent)
    jira.set_status("PB-1", "Done", True)
    data = sync.run(wb_path, None, jira, root=wb_path.parent)
    assert cell(wb_path, "PB-TRF-06", "Bug Status") == "Done"
    actions = {c["id"]: c["action"] for c in data["cases"]}
    assert actions["PB-TRF-06"] == "Perlu retest"
    assert data["bugs"][0]["state"] == "Closed"


def test_lulus_tetapi_tiket_masih_terbuka(wb_path):
    jira = FakeJira()
    sync.run(wb_path, None, jira, root=wb_path.parent)
    wb = openpyxl.load_workbook(wb_path)
    wb["Transfer"].cell(row=5, column=HEAD.index("Status") + 1, value="Passed")
    wb.save(wb_path)
    data = sync.run(wb_path, None, jira, root=wb_path.parent)
    assert {c["id"]: c["action"] for c in data["cases"]}["PB-TRF-06"] == "Lulus, tiket masih terbuka"


def test_rtm_dan_status_kosong_dihitung_not_run(wb_path):
    data = sync.run(wb_path, None, FakeJira(), root=wb_path.parent)
    rtm = {r["id"]: r for r in data["rtm"]}
    assert rtm["REQ-07"]["cases"] == "PB-TRF-06, PB-TRF-07"
    assert rtm["REQ-07"]["state"] == "Ada yang gagal"
    assert rtm["REQ-06"]["notrun"] == 1
    assert rtm["REQ-99"]["state"] == "Tidak ada TC"
    assert data["totals"] == {"Passed": 1, "Failed": 1, "Blocked": 0, "Not Run": 1}
    ws = openpyxl.load_workbook(wb_path)["RTM"]
    assert ws.cell(row=3, column=1).value == "REQ-06"


def test_mode_lokal_tidak_membuat_tiket(wb_path, tmp_path):
    data = sync.run(wb_path, tmp_path / "site", None, root=wb_path.parent)
    assert data["pending"] == ["PB-TRF-06"]
    assert cell(wb_path, "PB-TRF-06", "Bug ID") is None
    page = (tmp_path / "site" / "index.html").read_text(encoding="utf-8")
    assert "Jira belum terhubung" in page and "PB-TRF-06" in page
    assert (tmp_path / "site" / "QA.xlsx").exists()


def test_status_tidak_dikenal_menjadi_peringatan(wb_path):
    wb = openpyxl.load_workbook(wb_path)
    wb["Transfer"].cell(row=4, column=HEAD.index("Status") + 1, value="Lulus")
    wb.save(wb_path)
    data = sync.run(wb_path, None, None, root=wb_path.parent)
    assert any("PB-TRF-01" in w and "Lulus" in w for w in data["warnings"])


def test_capture_di_luar_repo_tidak_diunggah(wb_path, tmp_path):
    outside = tmp_path.parent / "rahasia.png"
    outside.write_bytes(b"x")
    wb = openpyxl.load_workbook(wb_path)
    wb["Transfer"].cell(row=5, column=HEAD.index("Capture") + 1, value="../rahasia.png")
    wb.save(wb_path)
    jira = FakeJira()
    sync.run(wb_path, None, jira, root=wb_path.parent)
    assert jira.attached == []
