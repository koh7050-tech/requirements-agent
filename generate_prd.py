#!/usr/bin/env python3
"""XLSX 요구분석 → PRD(Product Requirements Document) HTML 생성"""

from __future__ import annotations
from datetime import date
from pathlib import Path

try:
    import openpyxl
except ImportError:
    pass


def load_sheet(wb, name):
    if name not in wb.sheetnames:
        return [], []
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [str(h or "") for h in rows[0]]
    data = [[str(c or "") for c in r] for r in rows[1:] if any(c for c in r)]
    return headers, data


def priority_color(p: str) -> str:
    return {"Must": "#c0392b", "Should": "#d4a017", "Could": "#2980b9"}.get(p, "#888")


def generate_prd(xlsx_path: Path, project: str, mode: str = "") -> str:
    import openpyxl as ox
    wb = ox.load_workbook(xlsx_path)
    today = date.today().isoformat()

    fr_h, fr_rows   = load_sheet(wb, "3_기능요구사항")
    nfr_h, nfr_rows = load_sheet(wb, "4_비기능요구사항")
    kpi_h, kpi_rows = load_sheet(wb, "6_KPI")
    act_h, act_rows = load_sheet(wb, "7_액션아이템")
    qc_h,  qc_rows  = load_sheet(wb, "9_품질Gate")

    qc_pass  = sum(1 for r in qc_rows if len(r) > 1 and r[1] == "PASS")
    qc_total = len(qc_rows)

    # FR 테이블 행 생성
    fr_trs = ""
    for i, r in enumerate(fr_rows, 1):
        fr_id  = r[0] if len(r) > 0 else ""
        fr_req = r[1] if len(r) > 1 else ""
        fr_ac  = r[4] if len(r) > 4 else ""
        fr_pri = r[3] if len(r) > 3 else ""
        col = priority_color(fr_pri)
        fr_trs += f"""<tr>
          <td style="width:24px;color:#888;font-size:12px">{i}</td>
          <td><strong>{fr_id}</strong><br><span style="color:#555;font-size:12px">{fr_req}</span></td>
          <td style="font-size:12px;color:#555">{fr_ac}</td>
          <td><span style="background:{col};color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700">{fr_pri}</span></td>
          <td></td>
          <td></td>
        </tr>"""

    # NFR 테이블 행
    nfr_trs = ""
    for i, r in enumerate(nfr_rows, 1):
        nfr_id  = r[0] if len(r) > 0 else ""
        nfr_req = r[2] if len(r) > 2 else ""
        nfr_typ = r[1] if len(r) > 1 else ""
        nfr_pri = r[4] if len(r) > 4 else ""
        col = priority_color(nfr_pri)
        nfr_trs += f"""<tr>
          <td style="width:24px;color:#888;font-size:12px">{i}</td>
          <td><strong>{nfr_id}</strong> <span style="color:#888;font-size:11px">[{nfr_typ}]</span><br>
            <span style="color:#555;font-size:12px">{nfr_req}</span></td>
          <td></td>
          <td><span style="background:{col};color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700">{nfr_pri}</span></td>
          <td></td>
          <td></td>
        </tr>"""

    # KPI 성공지표 행
    kpi_rows_html = ""
    kpi_idx = {h: i for i, h in enumerate(kpi_h)}
    for r in kpi_rows:
        name   = r[kpi_idx.get("지표명", 3)] if len(r) > 3 else ""
        target = r[kpi_idx.get("목표값", 8)] if len(r) > 8 else ""
        unit   = r[kpi_idx.get("단위", 4)]   if len(r) > 4 else ""
        kpi_rows_html += f"<tr><td>{name}</td><td>{target} {unit}</td></tr>"

    # 액션 아이템 타임라인
    act_idx = {h: i for i, h in enumerate(act_h)}
    milestone_rows = ""
    for r in act_rows:
        act_id = r[act_idx.get("ACT_ID", 0)]    if len(r) > 0 else ""
        action = r[act_idx.get("액션", 2)]       if len(r) > 2 else ""
        owner  = r[act_idx.get("오너", 3)]       if len(r) > 3 else ""
        due    = r[act_idx.get("목표일", 4)]     if len(r) > 4 else ""
        status = r[act_idx.get("상태", 5)]       if len(r) > 5 else ""
        milestone_rows += f"<tr><td><code>{act_id}</code></td><td>{action}</td><td>{owner}</td><td>{due}</td><td>{status}</td></tr>"

    # 품질 Gate
    qc_rows_html = ""
    for r in qc_rows:
        gate   = r[0] if len(r) > 0 else ""
        result = r[1] if len(r) > 1 else ""
        desc   = r[2] if len(r) > 2 else ""
        color  = "#27ae60" if result == "PASS" else "#e74c3c"
        qc_rows_html += f'<tr><td>{gate}</td><td><span style="color:{color};font-weight:700">{result}</span></td><td>{desc}</td></tr>'

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PRD — {project}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, 'Segoe UI', 'Apple SD Gothic Neo', sans-serif;
    background: #fff;
    color: #1a1a1a;
    font-size: 14px;
    line-height: 1.7;
    padding: 48px 0 80px;
  }}
  .doc {{ max-width: 780px; margin: 0 auto; padding: 0 32px; }}

  /* 헤더 */
  .doc-header {{ margin-bottom: 36px; }}
  .doc-label {{ font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: #888; margin-bottom: 8px; }}
  .doc-title {{ font-size: 2.2rem; font-weight: 800; letter-spacing: -.5px; color: #111; line-height: 1.2; margin-bottom: 10px; }}
  .doc-meta {{ font-size: 12px; color: #888; }}

  /* 섹션 */
  h2 {{ font-size: 1.1rem; font-weight: 700; color: #111; margin: 36px 0 8px; border-bottom: 2px solid #f0f0f0; padding-bottom: 8px; }}
  h3 {{ font-size: .95rem; font-weight: 700; color: #333; margin: 20px 0 6px; }}
  p  {{ color: #555; font-size: 13px; margin-bottom: 12px; line-height: 1.7; }}

  /* 메타 테이블 */
  .meta-table {{ width: 100%; border-collapse: collapse; margin: 12px 0 24px; border: 1px solid #e5e5e5; border-radius: 6px; overflow: hidden; }}
  .meta-table tr {{ border-bottom: 1px solid #f0f0f0; }}
  .meta-table tr:last-child {{ border-bottom: none; }}
  .meta-table td {{ padding: 9px 14px; font-size: 13px; }}
  .meta-table td:first-child {{ color: #888; width: 160px; font-weight: 500; background: #fafafa; }}

  /* 일반 테이블 */
  table.data {{ width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 13px; }}
  table.data th {{ background: #f5f5f5; padding: 8px 12px; text-align: left; font-weight: 600; color: #444; font-size: 12px; border-bottom: 2px solid #e5e5e5; }}
  table.data td {{ padding: 9px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: top; color: #333; }}
  table.data tr:hover td {{ background: #fafafa; }}
  table.data code {{ font-size: 11px; background: #f0f0f0; padding: 1px 5px; border-radius: 3px; }}

  /* 통계 뱃지 */
  .stats {{ display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0; }}
  .stat {{ background: #f5f5f5; border-radius: 8px; padding: 12px 18px; text-align: center; min-width: 100px; }}
  .stat-n {{ font-size: 22px; font-weight: 800; color: #111; }}
  .stat-l {{ font-size: 11px; color: #888; margin-top: 2px; }}

  /* 구분선 */
  hr {{ border: none; border-top: 1px solid #f0f0f0; margin: 28px 0; }}

  /* 프린트 */
  @media print {{
    body {{ padding: 0; }}
    .doc {{ padding: 0 24px; }}
  }}
</style>
</head>
<body>
<div class="doc">

  <!-- 헤더 -->
  <div class="doc-header">
    <div class="doc-label">Product Requirements Document</div>
    <div class="doc-title">{project}</div>
    <div class="doc-meta">생성일: {today} · Requirements Analysis Agent{f" · {mode}" if mode else ""}</div>
  </div>

  <!-- 문서 정보 -->
  <h2>Details</h2>
  <table class="meta-table">
    <tr><td>Document Title</td><td>{project} 요구사항 정의서</td></tr>
    <tr><td>Generated Date</td><td>{today}</td></tr>
    <tr><td>Document Status</td><td><span style="background:#e8f4fd;color:#2980b9;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700">DRAFT</span></td></tr>
    <tr><td>Analysis Mode</td><td>{mode or "통합 분석"}</td></tr>
    <tr><td>Quality Gate</td><td><span style="font-weight:700;color:{'#27ae60' if qc_pass == qc_total else '#e74c3c'}">{qc_pass}/{qc_total} PASS</span></td></tr>
    <tr><td>Tool</td><td>Requirements Analysis Agent</td></tr>
  </table>

  <!-- 요약 통계 -->
  <h2>Summary</h2>
  <div class="stats">
    <div class="stat"><div class="stat-n">{len(fr_rows)}</div><div class="stat-l">기능 요구사항</div></div>
    <div class="stat"><div class="stat-n">{len(nfr_rows)}</div><div class="stat-l">비기능 요구사항</div></div>
    <div class="stat"><div class="stat-n">{len(kpi_rows)}</div><div class="stat-l">KPI</div></div>
    <div class="stat"><div class="stat-n">{len(act_rows)}</div><div class="stat-l">액션 아이템</div></div>
  </div>

  <hr>

  <!-- 목적 -->
  <h2>Objective</h2>
  <p>본 문서는 <strong>{project}</strong> 프로젝트의 요구사항을 정의하고, 이해관계자 간의 공통된 이해를 바탕으로 성공적인 시스템 구축을 위한 기준을 제공합니다.</p>

  <!-- 성공 지표 -->
  <h2>Success Metrics</h2>
  <p>프로젝트 성공 여부를 판단하기 위한 핵심 성과 지표(KPI)입니다.</p>
  <table class="data">
    <thead><tr><th>지표명</th><th>목표값</th></tr></thead>
    <tbody>
      {kpi_rows_html if kpi_rows_html else '<tr><td colspan="2" style="color:#aaa">KPI 데이터 없음</td></tr>'}
    </tbody>
  </table>

  <hr>

  <!-- 기능 요구사항 -->
  <h2>Functional Requirements</h2>
  <p>시스템이 반드시 수행해야 하는 기능적 요구사항 목록입니다. 우선순위(MoSCoW)에 따라 개발 스코프를 결정합니다.</p>
  <table class="data">
    <thead>
      <tr>
        <th style="width:24px">#</th>
        <th>요구사항</th>
        <th>수용기준</th>
        <th>우선순위</th>
        <th>Jira Issue</th>
        <th>Notes</th>
      </tr>
    </thead>
    <tbody>
      {fr_trs if fr_trs else '<tr><td colspan="6" style="color:#aaa">데이터 없음</td></tr>'}
    </tbody>
  </table>

  <hr>

  <!-- 비기능 요구사항 -->
  <h2>Non-Functional Requirements</h2>
  <p>성능·보안·가용성 등 시스템의 품질 속성을 정의한 요구사항입니다.</p>
  <table class="data">
    <thead>
      <tr>
        <th style="width:24px">#</th>
        <th>요구사항</th>
        <th>User Story</th>
        <th>우선순위</th>
        <th>Jira Issue</th>
        <th>Notes</th>
      </tr>
    </thead>
    <tbody>
      {nfr_trs if nfr_trs else '<tr><td colspan="6" style="color:#aaa">데이터 없음</td></tr>'}
    </tbody>
  </table>

  <hr>

  <!-- 마일스톤 -->
  <h2>Milestones & Action Items</h2>
  <p>프로젝트 진행을 위한 액션 아이템과 목표 일정입니다.</p>
  <table class="data">
    <thead><tr><th>ID</th><th>액션</th><th>담당자</th><th>목표일</th><th>상태</th></tr></thead>
    <tbody>
      {milestone_rows if milestone_rows else '<tr><td colspan="5" style="color:#aaa">데이터 없음</td></tr>'}
    </tbody>
  </table>

  <hr>

  <!-- 품질 Gate -->
  <h2>Quality Gate</h2>
  <p>요구사항 분석의 완성도 검증 결과입니다.</p>
  <table class="data">
    <thead><tr><th>검증 항목</th><th>결과</th><th>설명</th></tr></thead>
    <tbody>
      {qc_rows_html if qc_rows_html else '<tr><td colspan="3" style="color:#aaa">데이터 없음</td></tr>'}
    </tbody>
  </table>

  <hr>

  <!-- Open Questions -->
  <h2>Open Questions</h2>
  <table class="data">
    <thead><tr><th>Question</th><th>Answer</th><th>Date Answered</th></tr></thead>
    <tbody>
      <tr><td style="color:#aaa">추가 검토가 필요한 사항을 여기에 기록하세요.</td><td></td><td></td></tr>
    </tbody>
  </table>

  <hr>

  <!-- Out of Scope -->
  <h2>Out of Scope</h2>
  <p>이번 프로젝트 스코프에서 제외된 항목을 기록합니다. MoSCoW 분석에서 Won't로 분류된 항목을 참고하세요.</p>

</div>
</body>
</html>"""
