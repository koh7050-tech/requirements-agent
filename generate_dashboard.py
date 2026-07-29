#!/usr/bin/env python3
"""XLSX 요구분석 데이터팩 → HTML 대시보드 생성"""

from __future__ import annotations
import sys
from pathlib import Path
from datetime import date

try:
    import openpyxl
except ImportError:
    print("[ERROR] pip3 install openpyxl")
    sys.exit(1)


def load_sheet(wb, name: str) -> tuple[list[str], list[list]]:
    if name not in wb.sheetnames:
        return [], []
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [str(h or "") for h in rows[0]]
    data = [[str(c or "") for c in r] for r in rows[1:] if any(c for c in r)]
    return headers, data


def badge(text: str) -> str:
    colors = {
        "Must":   ("#c0392b", "#fff"),
        "Should": ("#d4a017", "#fff"),
        "Could":  ("#2980b9", "#fff"),
        "Won't":  ("#7f8c8d", "#fff"),
        "PASS":   ("#27ae60", "#fff"),
        "FAIL":   ("#e74c3c", "#fff"),
        "대기":   ("#95a5a6", "#fff"),
        "진행중": ("#f39c12", "#fff"),
        "완료":   ("#27ae60", "#fff"),
        "Active": ("#2980b9", "#fff"),
        "기능":   ("#8e44ad", "#fff"),
        "비기능": ("#16a085", "#fff"),
    }
    bg, fg = colors.get(text, ("#ecf0f1", "#2c3e50"))
    return f'<span style="background:{bg};color:{fg};padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{text}</span>'


def table_html(headers: list[str], rows: list[list], badge_cols: list[str] = None) -> str:
    badge_cols = badge_cols or []
    th = "".join(f"<th>{h}</th>" for h in headers)
    trs = []
    for row in rows:
        tds = []
        for h, v in zip(headers, row):
            if h in badge_cols and v:
                tds.append(f"<td>{badge(v)}</td>")
            else:
                tds.append(f"<td>{v}</td>")
        trs.append("<tr>" + "".join(tds) + "</tr>")
    return f"""
<div class="table-wrap">
<table>
  <thead><tr>{th}</tr></thead>
  <tbody>{"".join(trs)}</tbody>
</table>
</div>"""


def kpi_cards(headers: list[str], rows: list[list]) -> str:
    if not rows:
        return "<p>KPI 데이터 없음</p>"
    idx = {h: i for i, h in enumerate(headers)}
    cards = []
    for r in rows:
        kpi_id  = r[idx.get("KPI_ID", 0)]
        req     = r[idx.get("연결_REQ", 1)]
        name    = r[idx.get("지표명", 3)]
        target  = r[idx.get("목표값", 8)]
        unit    = r[idx.get("단위", 4)]
        period  = r[idx.get("측정주기", 6)]
        owner   = r[idx.get("오너", 10)]
        cards.append(f"""
<div class="kpi-card">
  <div class="kpi-id">{kpi_id} · {req}</div>
  <div class="kpi-name">{name}</div>
  <div class="kpi-target">{target} <span class="kpi-unit">{unit}</span></div>
  <div class="kpi-meta">{period} 측정 · {owner}</div>
</div>""")
    return '<div class="kpi-grid">' + "".join(cards) + "</div>"


def moscow_chart(headers: list[str], rows: list[list]) -> str:
    if not rows:
        return ""
    idx = {h: i for i, h in enumerate(headers)}
    counts: dict[str, int] = {"Must": 0, "Should": 0, "Could": 0, "Won't": 0}
    for r in rows:
        c = r[idx.get("분류", 1)]
        if c in counts:
            counts[c] += 1
    total = sum(counts.values()) or 1
    colors = {"Must": "#c0392b", "Should": "#d4a017", "Could": "#2980b9", "Won't": "#7f8c8d"}
    bars = ""
    for label, count in counts.items():
        pct = count / total * 100
        bars += f"""
<div class="bar-row">
  <div class="bar-label">{label}</div>
  <div class="bar-track">
    <div class="bar-fill" style="width:{pct:.1f}%;background:{colors[label]}"></div>
  </div>
  <div class="bar-count">{count}건</div>
</div>"""
    return f'<div class="bar-chart">{bars}</div>'


def action_timeline(headers: list[str], rows: list[list]) -> str:
    if not rows:
        return "<p>액션 없음</p>"
    idx = {h: i for i, h in enumerate(headers)}
    by_month: dict[str, list] = {}
    for r in rows:
        due = r[idx.get("목표일", 4)]
        month = due[:7] if len(due) >= 7 else "미정"
        by_month.setdefault(month, []).append(r)
    html = ""
    for month in sorted(by_month):
        items = ""
        for r in by_month[month]:
            act_id = r[idx.get("ACT_ID", 0)]
            action = r[idx.get("액션", 2)]
            owner  = r[idx.get("오너", 3)]
            status = r[idx.get("상태", 5)]
            pri    = r[idx.get("우선순위", 6)]
            items += f"""
<div class="tl-item">
  <span class="tl-id">{act_id}</span>
  {badge(pri)} {badge(status)}
  <span class="tl-action">{action}</span>
  <span class="tl-owner">· {owner}</span>
</div>"""
        html += f'<div class="tl-month"><div class="tl-month-label">{month}</div>{items}</div>'
    return f'<div class="timeline">{html}</div>'


def section(title: str, icon: str, content: str) -> str:
    return f"""
<section class="section">
  <h2 class="section-title">{icon} {title}</h2>
  {content}
</section>"""


def generate(xlsx_path: Path, output_path: Path, project: str, mode: str = "") -> None:
    wb = openpyxl.load_workbook(xlsx_path)
    today = date.today().isoformat()

    _, src_rows     = load_sheet(wb, "1_출처")
    fr_h, fr_rows   = load_sheet(wb, "3_기능요구사항")
    nfr_h, nfr_rows = load_sheet(wb, "4_비기능요구사항")
    mo_h, mo_rows   = load_sheet(wb, "5_분석_MoSCoW")
    kpi_h, kpi_rows = load_sheet(wb, "6_KPI")
    act_h, act_rows = load_sheet(wb, "7_액션아이템")
    tr_h, tr_rows   = load_sheet(wb, "8_추적성")
    qc_h, qc_rows   = load_sheet(wb, "9_품질Gate")

    must_count   = sum(1 for r in fr_rows + nfr_rows if len(r) > 3 and r[3] == "Must")
    action_count = len(act_rows)
    kpi_count    = len(kpi_rows)
    qc_pass      = sum(1 for r in qc_rows if len(r) > 1 and r[1] == "PASS")
    qc_total     = len(qc_rows)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{project} — 요구분석 대시보드</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f4f8;color:#2c3e50;font-size:13px}}
.topbar{{background:#1F4E79;color:#fff;padding:16px 32px;display:flex;justify-content:space-between;align-items:center}}
.topbar h1{{font-size:18px;font-weight:700}}
.topbar .meta{{font-size:12px;opacity:.8}}
.summary{{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;padding:24px 32px}}
.stat-card{{background:#fff;border-radius:10px;padding:20px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.stat-num{{font-size:36px;font-weight:800;color:#1F4E79}}
.stat-label{{font-size:12px;color:#7f8c8d;margin-top:4px}}
.main{{padding:0 32px 40px}}
.section{{background:#fff;border-radius:10px;padding:24px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.section-title{{font-size:15px;font-weight:700;color:#1F4E79;margin-bottom:16px;padding-bottom:8px;border-bottom:2px solid #e8f0fb}}
.table-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:#1F4E79;color:#fff;padding:8px 10px;text-align:left;font-weight:600;white-space:nowrap}}
td{{padding:7px 10px;border-bottom:1px solid #ecf0f1;vertical-align:top;line-height:1.5}}
tr:hover td{{background:#f8fbff}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px}}
.kpi-card{{background:#f8fbff;border:1px solid #dce8f7;border-radius:8px;padding:14px}}
.kpi-id{{font-size:11px;color:#7f8c8d;margin-bottom:4px}}
.kpi-name{{font-size:13px;font-weight:600;color:#2c3e50;margin-bottom:8px}}
.kpi-target{{font-size:22px;font-weight:800;color:#1F4E79}}
.kpi-unit{{font-size:13px;font-weight:400;color:#7f8c8d}}
.kpi-meta{{font-size:11px;color:#95a5a6;margin-top:6px}}
.bar-chart{{padding:4px 0}}
.bar-row{{display:flex;align-items:center;margin-bottom:10px;gap:10px}}
.bar-label{{width:60px;font-size:12px;font-weight:600}}
.bar-track{{flex:1;background:#ecf0f1;border-radius:4px;height:20px;overflow:hidden}}
.bar-fill{{height:100%;border-radius:4px;transition:width .4s}}
.bar-count{{width:36px;font-size:12px;color:#7f8c8d;text-align:right}}
.timeline{{}}
.tl-month{{margin-bottom:16px}}
.tl-month-label{{font-size:12px;font-weight:700;color:#1F4E79;background:#e8f0fb;padding:4px 10px;border-radius:4px;margin-bottom:8px;display:inline-block}}
.tl-item{{padding:6px 0 6px 12px;border-left:2px solid #dce8f7;margin-bottom:4px;display:flex;align-items:center;flex-wrap:wrap;gap:6px}}
.tl-id{{font-size:11px;color:#7f8c8d;width:56px}}
.tl-action{{font-size:12px;color:#2c3e50}}
.tl-owner{{font-size:11px;color:#7f8c8d}}
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:24px}}
footer{{text-align:center;padding:20px;font-size:11px;color:#bdc3c7}}
</style>
</head>
<body>
<div class="topbar">
  <h1>📊 {project} — 요구분석 대시보드</h1>
  <div class="meta">생성일: {today} · Requirements Analysis Agent{f" · {mode}" if mode else ""}</div>
</div>
{f'<div style="background:{"#1a6b3a" if "단일" in mode else "#1a4a7a"};color:#fff;padding:8px 32px;font-size:12px;font-weight:600;">{"🔵 단일 분석" if "단일" in mode else "🟢 통합 분석"} — {mode}</div>' if mode else ""}

<div class="summary">
  <div class="stat-card"><div class="stat-num">{len(fr_rows)}</div><div class="stat-label">기능 요구사항</div></div>
  <div class="stat-card"><div class="stat-num">{len(nfr_rows)}</div><div class="stat-label">비기능 요구사항</div></div>
  <div class="stat-card"><div class="stat-num">{kpi_count}</div><div class="stat-label">KPI</div></div>
  <div class="stat-card"><div class="stat-num">{action_count}</div><div class="stat-label">액션 아이템</div></div>
  <div class="stat-card"><div class="stat-num">{qc_pass}/{qc_total}</div><div class="stat-label">품질 Gate PASS</div></div>
</div>

<div class="main">

{section("기능 요구사항 (FR)", "🔧",
    table_html(fr_h, fr_rows, badge_cols=["우선순위", "상태"]))}

{section("비기능 요구사항 (NFR)", "⚙️",
    table_html(nfr_h, nfr_rows, badge_cols=["우선순위", "상태"]))}

<section class="section">
  <h2 class="section-title">📊 MoSCoW 우선순위 분석</h2>
  <div class="two-col">
    <div>{moscow_chart(mo_h, mo_rows)}</div>
    <div>{table_html(mo_h, mo_rows, badge_cols=["분류"])}</div>
  </div>
</section>

{section("📈 KPI", "📈", kpi_cards(kpi_h, kpi_rows))}

{section("✅ 액션 아이템 타임라인", "✅", action_timeline(act_h, act_rows))}

{section("🔗 추적성 매트릭스", "🔗",
    table_html(tr_h, tr_rows, badge_cols=["상태"]))}

{section("🛡️ 품질 Gate", "🛡️",
    table_html(qc_h, qc_rows, badge_cols=["결과"]))}

</div>
<footer>Requirements Analysis Agent · {today} · {project}</footer>
</body>
</html>"""

    output_path.write_text(html, encoding="utf-8")
    print(f"✓ 대시보드 생성: {output_path}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx",    default="outputs/요구분석_데이터팩.xlsx")
    parser.add_argument("--output",  default="outputs/dashboard.html")
    parser.add_argument("--project", default="DAM구축프로젝트")
    parser.add_argument("--mode",    default="")
    args = parser.parse_args()
    generate(Path(args.xlsx), Path(args.output), args.project, args.mode)
