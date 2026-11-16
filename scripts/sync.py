"""Sinkronisasi workbook QA manual ParaBank dengan Jira, lalu membuat dashboard.

Yang dilakukan setiap kali dijalankan:
  1. Membaca semua test case dari sheet modul (sheet yang punya kolom "Test Case ID").
  2. Baris Failed yang belum punya Bug ID dibuatkan tiket Bug di Jira (atau memakai tiket
     terbuka yang sudah ada untuk test case yang sama), lalu Bug ID ditulis ke Excel.
  3. Status setiap tiket ditarik dari Jira dan ditulis ke kolom Bug Status.
  4. Sheet RTM dan Bug Report dibangun ulang dari data terbaru.
  5. Workbook hanya disimpan jika ada nilai yang berubah.
  6. Dashboard HTML dan data.json ditulis ke folder site.

Tanpa variabel JIRA_* skrip berjalan dalam mode lokal: tidak ada tiket yang dibuat, tetapi RTM,
Bug Report, dan dashboard tetap diperbarui.
"""
from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

STATUSES = ("Passed", "Failed", "Blocked", "Not Run")
SEVERITIES = ("Critical", "High", "Medium", "Low")
PRIORITY_BY_SEVERITY = {"Critical": "Highest", "High": "High", "Medium": "Medium", "Low": "Low"}
GENERATED_SHEETS = ("Summary", "Test Plan", "Requirements", "RTM", "Bug Report", "Test Data", "How to use")
LABEL = "parabank-manual"
WIB = timezone(timedelta(hours=7))


# ---------------------------------------------------------------- data

@dataclass
class TestCase:
    sheet: str
    row: int
    cols: dict
    values: dict

    def get(self, name: str) -> str:
        v = self.values.get(name)
        return "" if v is None else str(v).strip()

    @property
    def id(self) -> str:
        return self.get("Test Case ID")

    @property
    def status(self) -> str:
        raw = self.get("Status")
        for s in STATUSES:
            if raw.lower() == s.lower():
                return s
        return "Not Run" if raw == "" else f"INVALID:{raw}"

    @property
    def requirements(self) -> list[str]:
        return [r.strip() for r in self.get("Requirement ID").replace(";", ",").split(",") if r.strip()]

    @property
    def severity(self) -> str:
        sev = self.get("Severity").capitalize()
        return sev if sev in SEVERITIES else "Medium"


@dataclass
class Issue:
    key: str
    summary: str = ""
    status: str = ""
    done: bool = False


@dataclass
class Result:
    changed: bool = False
    created: list = field(default_factory=list)
    reused: list = field(default_factory=list)
    pending: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


def find_header(ws, key="Test Case ID", max_row=40):
    for row in ws.iter_rows(min_row=1, max_row=min(max_row, ws.max_row)):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == key:
                cols = {}
                for c in row:
                    if isinstance(c.value, str) and c.value.strip():
                        cols[c.value.strip()] = c.column
                return cell.row, cols
    return None, None


def read_cases(wb) -> list[TestCase]:
    cases = []
    for ws in wb.worksheets:
        if ws.title in GENERATED_SHEETS:
            continue
        hrow, cols = find_header(ws)
        if not hrow:
            continue
        for r in range(hrow + 1, ws.max_row + 1):
            values = {name: ws.cell(row=r, column=c).value for name, c in cols.items()}
            if values.get("Test Case ID"):
                cases.append(TestCase(ws.title, r, cols, values))
    return cases


def read_requirements(wb) -> list[tuple[str, str]]:
    if "Requirements" not in wb.sheetnames:
        return []
    ws = wb["Requirements"]
    hrow, cols = find_header(ws, "REQ ID")
    if not hrow:
        return []
    out = []
    for r in range(hrow + 1, ws.max_row + 1):
        rid = ws.cell(row=r, column=cols["REQ ID"]).value
        if rid:
            out.append((str(rid).strip(), str(ws.cell(row=r, column=cols.get("Requirement", 2)).value or "")))
    return out


def set_value(ws, row, col, value, result: Result):
    cell = ws.cell(row=row, column=col)
    old = cell.value if cell.value not in ("",) else None
    new = value if value not in ("",) else None
    if old != new:
        cell.value = new
        result.changed = True


# ---------------------------------------------------------------- Jira

def adf(sections: list[tuple[str, str]]) -> dict:
    """Deskripsi Jira (Atlassian Document Format): judul tebal lalu isi, baris baru dipertahankan."""
    content = []
    for title, text in sections:
        content.append({"type": "paragraph", "content": [{"type": "text", "text": title, "marks": [{"type": "strong"}]}]})
        parts = []
        for i, line in enumerate((text or "-").split("\n")):
            if i:
                parts.append({"type": "hardBreak"})
            parts.append({"type": "text", "text": line or " "})
        content.append({"type": "paragraph", "content": parts})
    return {"type": "doc", "version": 1, "content": content}


class JiraClient:
    def __init__(self, base_url: str, email: str, token: str, project: str):
        import requests

        self.base = base_url.rstrip("/")
        self.project = project
        self.http = requests.Session()
        self.http.auth = (email, token)
        self.http.headers.update({"Accept": "application/json"})

    @classmethod
    def from_env(cls):
        need = ["JIRA_BASE_URL", "JIRA_EMAIL", "JIRA_API_TOKEN", "JIRA_PROJECT_KEY"]
        if not all(os.environ.get(k) for k in need):
            return None
        return cls(*(os.environ[k] for k in need))

    def browse_url(self, key: str) -> str:
        return f"{self.base}/browse/{key}"

    def _check(self, r):
        if r.status_code >= 400:
            raise RuntimeError(f"Jira {r.request.method} {r.url} -> {r.status_code}: {r.text[:400]}")
        return r.json() if r.content else {}

    def find_open_bug(self, tc_id: str) -> Issue | None:
        jql = (f'project = "{self.project}" AND labels = "{LABEL}" AND labels = "{tc_id}" '
               f"AND statusCategory != Done ORDER BY created DESC")
        data = self._check(self.http.post(f"{self.base}/rest/api/3/search/jql",
                                          json={"jql": jql, "fields": ["summary", "status"], "maxResults": 1}))
        issues = data.get("issues", [])
        return self._issue(issues[0]) if issues else None

    def create_bug(self, tc: TestCase, workbook_name: str) -> Issue:
        actual = tc.get("Current Result") or "(Current Result belum diisi)"
        summary = f"[{tc.id}] {tc.get('Test Case')}"
        if tc.get("Current Result"):
            summary += f": {actual.splitlines()[0]}"
        fields = {
            "project": {"key": self.project},
            "issuetype": {"name": "Bug"},
            "summary": summary[:250],
            "labels": [LABEL, tc.id, f"severity-{tc.severity.lower()}"],
            "description": adf([
                ("Test case", f"{tc.id} · {tc.get('Test Case')} ({tc.sheet} / {tc.get('Feature')})"),
                ("Severity", tc.severity),
                ("Requirement", ", ".join(tc.requirements) or "-"),
                ("Pre-condition", tc.get("Pre-Condition")),
                ("Langkah reproduksi", tc.get("Test Step")),
                ("Test data", tc.get("Test Data")),
                ("Hasil yang diharapkan", tc.get("Expected Result")),
                ("Hasil sebenarnya", actual),
                ("Catatan", tc.get("Notes") or "-"),
                ("Sumber", f"Dibuat otomatis dari {workbook_name}, sheet {tc.sheet}, baris {tc.row}."),
            ]),
            "priority": {"name": PRIORITY_BY_SEVERITY[tc.severity]},
        }
        r = self.http.post(f"{self.base}/rest/api/3/issue", json={"fields": fields})
        if r.status_code == 400 and "priority" in r.text:
            fields.pop("priority")  # field priority tidak ada di layar create proyek ini
            r = self.http.post(f"{self.base}/rest/api/3/issue", json={"fields": fields})
        key = self._check(r)["key"]
        return Issue(key=key, summary=summary, status="To Do", done=False)

    def attach(self, key: str, path: Path):
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        with path.open("rb") as fh:
            self._check(self.http.post(f"{self.base}/rest/api/3/issue/{key}/attachments",
                                       headers={"X-Atlassian-Token": "no-check"},
                                       files={"file": (path.name, fh, mime)}))

    def get_issue(self, key: str) -> Issue | None:
        r = self.http.get(f"{self.base}/rest/api/3/issue/{key}", params={"fields": "summary,status"})
        if r.status_code == 404:
            return None
        return self._issue(self._check(r))

    @staticmethod
    def _issue(data) -> Issue:
        f = data.get("fields", {})
        st = f.get("status") or {}
        return Issue(key=data["key"], summary=f.get("summary", ""), status=st.get("name", ""),
                     done=(st.get("statusCategory") or {}).get("key") == "done")


# ---------------------------------------------------------------- sinkronisasi

def evidence_path(tc: TestCase, root: Path) -> Path | None:
    cap = tc.get("Capture")
    if not cap or cap.startswith("http"):
        return None
    p = (root / cap).resolve()
    return p if p.is_file() and root.resolve() in p.parents else None


def sync_bugs(wb, cases, jira, root: Path, workbook_name: str, result: Result) -> dict[str, Issue]:
    issues: dict[str, Issue] = {}
    for tc in cases:
        if tc.status.startswith("INVALID"):
            result.warnings.append(f"{tc.id}: Status '{tc.get('Status')}' tidak dikenal "
                                   f"(pakai {', '.join(STATUSES)})")
        if tc.status != "Failed" or tc.get("Bug ID"):
            continue
        if jira is None:
            result.pending.append(tc.id)
            continue
        issue = jira.find_open_bug(tc.id)
        if issue:
            result.reused.append((tc.id, issue.key))
        else:
            issue = jira.create_bug(tc, workbook_name)
            result.created.append((tc.id, issue.key))
            ev = evidence_path(tc, root)
            if ev:
                try:
                    jira.attach(issue.key, ev)
                except RuntimeError as exc:
                    result.warnings.append(f"{issue.key}: lampiran {ev.name} gagal diunggah ({exc})")
        issues[issue.key] = issue
        tc.values["Bug ID"] = issue.key
        set_value(wb[tc.sheet], tc.row, tc.cols["Bug ID"], issue.key, result)

    for tc in cases:
        for key in bug_keys(tc):
            if key in issues or jira is None:
                continue
            issue = jira.get_issue(key)
            if issue is None:
                result.warnings.append(f"{tc.id}: tiket {key} tidak ditemukan di Jira")
                issue = Issue(key=key, status="Tidak ditemukan")
            issues[key] = issue
        if jira is not None and "Bug Status" in tc.cols and bug_keys(tc):
            status = " / ".join(issues[k].status for k in bug_keys(tc))
            tc.values["Bug Status"] = status
            set_value(wb[tc.sheet], tc.row, tc.cols["Bug Status"], status, result)

    if jira is None:  # mode lokal: pakai status yang sudah tertulis di Excel
        for tc in cases:
            for key in bug_keys(tc):
                st = tc.get("Bug Status")
                issues.setdefault(key, Issue(key=key, status=st, done=st.lower() in ("done", "closed", "resolved")))
    return issues


def bug_keys(tc: TestCase) -> list[str]:
    return [k.strip() for k in tc.get("Bug ID").replace(";", ",").split(",") if k.strip()]


def action_for(tc: TestCase, issues: dict[str, Issue]) -> str:
    keys = bug_keys(tc)
    if not keys:
        return "Buat tiket bug" if tc.status == "Failed" else ""
    done = all(issues[k].done for k in keys if k in issues)
    if tc.status == "Failed" and done:
        return "Perlu retest"
    if tc.status == "Passed" and not done:
        return "Lulus, tiket masih terbuka"
    if tc.status == "Failed":
        return "Menunggu perbaikan"
    return ""


def coverage_rows(cases, requirements):
    rows = []
    for rid, text in requirements:
        linked = [tc for tc in cases if rid in tc.requirements]
        counts = {s: sum(tc.status == s for tc in linked) for s in STATUSES}
        bugs = sorted({k for tc in linked for k in bug_keys(tc)})
        if not linked:
            state = "Tidak ada TC"
        elif counts["Failed"]:
            state = "Ada yang gagal"
        elif counts["Blocked"]:
            state = "Ada yang terblokir"
        elif counts["Not Run"]:
            state = "Belum selesai diuji" if counts["Passed"] else "Belum diuji"
        else:
            state = "Lulus semua"
        rows.append([rid, text, ", ".join(tc.id for tc in linked), len(linked), counts["Passed"],
                     counts["Failed"], counts["Blocked"], counts["Not Run"], ", ".join(bugs), state])
    return rows


def bug_rows(cases, issues, jira):
    by_key: dict[str, list[TestCase]] = {}
    for tc in cases:
        for k in bug_keys(tc):
            by_key.setdefault(k, []).append(tc)
    order = {s: i for i, s in enumerate(SEVERITIES)}
    rows = []
    for key, tcs in by_key.items():
        issue = issues.get(key, Issue(key=key))
        sev = min((tc.severity for tc in tcs), key=lambda s: order[s])
        title = issue.summary or f"[{tcs[0].id}] {tcs[0].get('Test Case')}"
        link = jira.browse_url(key) if jira else ""
        rows.append([key, title, ", ".join(tc.id for tc in tcs), ", ".join(sorted({tc.sheet for tc in tcs})),
                     sev, issue.status, link, "Closed" if issue.done else "Open"])
    rows.sort(key=lambda r: (order[r[4]], r[0]))
    return rows


def write_table(ws, first_row: int, rows: list[list], result: Result, width: int):
    wrap = Alignment(wrap_text=True, vertical="top")
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            set_value(ws, first_row + i, j, v, result)
            ws.cell(row=first_row + i, column=j).alignment = wrap
    for r in range(first_row + len(rows), max(ws.max_row, first_row) + 1):
        for j in range(1, width + 1):
            set_value(ws, r, j, None, result)


# ---------------------------------------------------------------- dashboard

def summarize(cases, issues, requirements, result, jira, workbook_name):
    modules = {}
    for tc in cases:
        m = modules.setdefault(tc.sheet, {s: 0 for s in STATUSES} | {"total": 0})
        m["total"] += 1
        m[tc.status if tc.status in STATUSES else "Not Run"] += 1
    totals = {s: sum(m[s] for m in modules.values()) for s in STATUSES}
    executed = totals["Passed"] + totals["Failed"] + totals["Blocked"]
    decided = totals["Passed"] + totals["Failed"]
    return {
        "generated_at": datetime.now(WIB).strftime("%d %b %Y %H:%M WIB"),
        "workbook": workbook_name,
        "jira_connected": jira is not None,
        "jira_url": jira.base if jira else "",
        "total": len(cases),
        "executed": executed,
        "pass_rate": round(100 * totals["Passed"] / decided) if decided else None,
        "totals": totals,
        "modules": modules,
        "positive": sum(tc.get("Test Type") == "Positive" for tc in cases),
        "negative": sum(tc.get("Test Type") == "Negative" for tc in cases),
        "bugs": [dict(zip(["key", "title", "cases", "module", "severity", "status", "link", "state"], r))
                 for r in bug_rows(cases, issues, jira)],
        "rtm": [dict(zip(["id", "text", "cases", "n", "passed", "failed", "blocked", "notrun", "bugs", "state"], r))
                for r in coverage_rows(cases, requirements)],
        "cases": [{"id": tc.id, "module": tc.sheet, "feature": tc.get("Feature"), "title": tc.get("Test Case"),
                   "type": tc.get("Test Type"), "priority": tc.get("Priority"), "status": tc.status,
                   "bugs": bug_keys(tc), "action": action_for(tc, issues)} for tc in cases],
        "pending": result.pending,
        "warnings": result.warnings,
    }


def render_dashboard(data: dict) -> str:
    e = html.escape
    t = data["totals"]
    total = max(data["total"], 1)
    seg = "".join(
        f'<span class="seg s-{s.replace(" ", "").lower()}" style="width:{100 * t[s] / total:.2f}%" '
        f'title="{s}: {t[s]}"></span>' for s in STATUSES if t[s])
    kpis = [("Test case", data["total"], f'{data["positive"]} positif · {data["negative"]} negatif'),
            ("Dieksekusi", data["executed"], f'{t["Not Run"]} belum dijalankan'),
            ("Pass rate", f'{data["pass_rate"]}%' if data["pass_rate"] is not None else "–", "Passed / (Passed + Failed)"),
            ("Bug terbuka", sum(b["state"] == "Open" for b in data["bugs"]), f'{len(data["bugs"])} bug tercatat')]
    kpi_html = "".join(f'<div class="kpi"><span>{e(k)}</span><b>{e(str(v))}</b><small>{e(s)}</small></div>'
                       for k, v, s in kpis)
    legend = "".join(f'<span><i class="dot s-{s.replace(" ", "").lower()}"></i>{s} {t[s]}</span>' for s in STATUSES)
    mod_rows = "".join(
        f'<tr><td>{e(name)}</td><td>{m["total"]}</td><td>{m["Passed"]}</td><td>{m["Failed"]}</td>'
        f'<td>{m["Blocked"]}</td><td>{m["Not Run"]}</td></tr>' for name, m in data["modules"].items())

    actions = [c for c in data["cases"] if c["action"]]
    act_html = ("".join(f'<li><b>{e(c["id"])}</b> {e(c["title"])} <span class="tag">{e(c["action"])}</span>'
                        f'{" · " + ", ".join(e(b) for b in c["bugs"]) if c["bugs"] else ""}</li>' for c in actions)
                or "<li>Tidak ada tindakan yang tertunda.</li>")

    def bug_link(b):
        return f'<a href="{e(b["link"])}">{e(b["key"])}</a>' if b["link"] else e(b["key"])
    bug_html = "".join(
        f'<tr><td>{bug_link(b)}</td><td>{e(b["title"])}</td><td>{e(b["cases"])}</td><td>{e(b["severity"])}</td>'
        f'<td>{e(b["status"] or "-")}</td></tr>' for b in data["bugs"]) or '<tr><td colspan="5">Belum ada bug.</td></tr>'
    rtm_html = "".join(
        f'<tr><td>{e(r["id"])}</td><td>{e(r["text"])}</td><td>{r["n"]}</td><td>{r["passed"]}</td>'
        f'<td>{r["failed"]}</td><td>{r["notrun"]}</td><td>{e(r["bugs"] or "-")}</td>'
        f'<td><span class="st">{e(r["state"])}</span></td></tr>' for r in data["rtm"])
    case_html = "".join(
        f'<tr data-status="{e(c["status"])}"><td>{e(c["id"])}</td><td>{e(c["module"])}</td><td>{e(c["title"])}</td>'
        f'<td>{e(c["type"])}</td><td><span class="pill p-{c["status"].replace(" ", "").lower()}">{e(c["status"])}</span>'
        f'</td><td>{e(", ".join(c["bugs"]) or "-")}</td></tr>' for c in data["cases"])
    notes = []
    if not data["jira_connected"]:
        notes.append("Jira belum terhubung: tiket bug belum dibuat otomatis. Atur JIRA_* di GitHub (lihat README).")
    if data["pending"]:
        notes.append("Failed tanpa tiket: " + ", ".join(data["pending"]))
    notes += data["warnings"]
    note_html = "".join(f'<p class="note">{e(n)}</p>' for n in notes)
    filters = "".join(f'<option value="{s}">{s}</option>' for s in STATUSES)

    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ParaBank · Dashboard QA Manual</title>
<style>
:root{{--g:#2E4A33;--ink:#1d221e;--mute:#5d665f;--line:#e2e5dc;--bg:#f6f7f2;--pass:#3f7d3a;--fail:#b3412f;--block:#b97a12;--nr:#9a9d95}}
*{{box-sizing:border-box}}body{{margin:0;font:15px/1.55 system-ui,-apple-system,"Segoe UI",sans-serif;color:var(--ink);background:var(--bg)}}
main{{max-width:1100px;margin:0 auto;padding:32px 20px 60px}}h1{{margin:0;font-size:26px;color:var(--g)}}
h2{{font-size:18px;margin:36px 0 12px}}.sub{{color:var(--mute);margin:4px 0 0}}a{{color:var(--g)}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-top:24px}}
.kpi{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px}}.kpi span{{color:var(--mute);font-size:13px}}
.kpi b{{display:block;font-size:28px;color:var(--g)}}.kpi small{{color:var(--mute)}}
.bar{{display:flex;height:14px;border-radius:7px;overflow:hidden;background:var(--line);margin:22px 0 8px}}.seg{{display:block}}
.s-passed{{background:var(--pass)}}.s-failed{{background:var(--fail)}}.s-blocked{{background:var(--block)}}.s-notrun{{background:var(--nr)}}
.legend{{display:flex;gap:16px;flex-wrap:wrap;color:var(--mute);font-size:13px}}.dot{{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px}}
table{{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:14px}}
th,td{{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}}th{{background:var(--g);color:#fff;font-weight:600}}
.wrap{{overflow-x:auto}}.pill{{padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600;white-space:nowrap}}
.p-passed{{background:#dcebd6;color:#25501f}}.p-failed{{background:#f6d6d0;color:#7a2618}}.p-blocked{{background:#f8e6c3;color:#6b4508}}.p-notrun{{background:#e8e8e3;color:#45473f}}
.tag{{background:#fdf0d2;color:#6b4508;border-radius:6px;padding:1px 7px;font-size:12px;margin-left:6px}}
ul.act{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 16px 12px 34px;margin:0}}ul.act li{{margin:4px 0}}
.note{{background:#fff7e6;border:1px solid #f0d9a6;border-radius:8px;padding:8px 12px;margin:12px 0 0;font-size:14px}}
.st{{font-size:12px;color:var(--mute)}}select{{font:inherit;padding:4px 8px;margin-bottom:10px}}
.links{{margin-top:14px;display:flex;gap:14px;flex-wrap:wrap}}
</style></head><body><main>
<h1>ParaBank · Dashboard QA Manual</h1>
<p class="sub">Diperbarui otomatis dari <b>{e(data["workbook"])}</b> · {e(data["generated_at"])}</p>
<div class="links"><a href="{e(data["workbook"])}" download>Unduh workbook Excel</a>
<a href="data.json">data.json</a>{f'<a href="{e(data["jira_url"])}">Jira</a>' if data["jira_url"] else ""}</div>
{note_html}
<div class="kpis">{kpi_html}</div>
<div class="bar" role="img" aria-label="Komposisi status test case">{seg}</div><div class="legend">{legend}</div>
<h2>Tindakan berikutnya</h2><ul class="act">{act_html}</ul>
<h2>Hasil per modul</h2><div class="wrap"><table><tr><th>Modul</th><th>Total</th><th>Passed</th><th>Failed</th><th>Blocked</th><th>Not Run</th></tr>{mod_rows}</table></div>
<h2>Bug</h2><div class="wrap"><table><tr><th>Bug ID</th><th>Judul</th><th>Test case</th><th>Severity</th><th>Status Jira</th></tr>{bug_html}</table></div>
<h2>Requirement traceability</h2><div class="wrap"><table><tr><th>REQ</th><th>Requirement</th><th>TC</th><th>Passed</th><th>Failed</th><th>Not Run</th><th>Bug</th><th>Status</th></tr>{rtm_html}</table></div>
<h2>Semua test case</h2><label>Filter status <select id="f"><option value="">Semua</option>{filters}</select></label>
<div class="wrap"><table id="cases"><tr><th>ID</th><th>Modul</th><th>Test case</th><th>Tipe</th><th>Status</th><th>Bug</th></tr>{case_html}</table></div>
</main><script>
document.getElementById('f').addEventListener('change',function(){{var v=this.value;
document.querySelectorAll('#cases tr[data-status]').forEach(function(r){{r.hidden=v&&r.dataset.status!==v}})}});
</script></body></html>"""


# ---------------------------------------------------------------- main

def run(workbook: Path, site: Path | None, jira=None, root: Path | None = None) -> dict:
    root = root or Path.cwd()
    wb = openpyxl.load_workbook(workbook)
    result = Result()
    cases = read_cases(wb)
    requirements = read_requirements(wb)
    issues = sync_bugs(wb, cases, jira, root, workbook.name, result)

    if "RTM" in wb.sheetnames:
        write_table(wb["RTM"], 3, coverage_rows(cases, requirements), result, 10)
    if "Bug Report" in wb.sheetnames:
        write_table(wb["Bug Report"], 2, bug_rows(cases, issues, jira), result, 8)
    if result.changed:
        wb.save(workbook)

    data = summarize(cases, issues, requirements, result, jira, workbook.name)
    if site:
        site.mkdir(parents=True, exist_ok=True)
        (site / "index.html").write_text(render_dashboard(data), encoding="utf-8")
        (site / "data.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        shutil.copy2(workbook, site / workbook.name)
    data["_result"] = result
    return data


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--workbook", default="manual/ParaBank-Manual-QA.xlsx", type=Path)
    ap.add_argument("--site", default="site", type=Path)
    ap.add_argument("--local", action="store_true", help="abaikan Jira walaupun variabel JIRA_* tersedia")
    args = ap.parse_args(argv)

    jira = None if args.local else JiraClient.from_env()
    print("Mode:", "Jira " + jira.base if jira else "lokal (tanpa Jira)")
    data = run(args.workbook, args.site, jira)
    res = data["_result"]
    t = data["totals"]
    print(f"Test case: {data['total']} | Passed {t['Passed']} · Failed {t['Failed']} · "
          f"Blocked {t['Blocked']} · Not Run {t['Not Run']}")
    for tc, key in res.created:
        print(f"  Bug dibuat: {key} untuk {tc}")
    for tc, key in res.reused:
        print(f"  Tiket terbuka dipakai ulang: {key} untuk {tc}")
    if res.pending:
        print("  Failed tanpa tiket (Jira belum terhubung):", ", ".join(res.pending))
    for w in res.warnings:
        print("  Peringatan:", w)
    print("Workbook", "diperbarui" if res.changed else "tidak berubah")
    return 0


if __name__ == "__main__":
    sys.exit(main())
