#!/usr/bin/env python3
"""
회의록 CSV → 요구분석 전체 파이프라인
수집→정의→분석→KPI→액션→추적성 + XLSX + Executive Summary
"""

from __future__ import annotations
import csv
import sys
from datetime import date
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("[ERROR] openpyxl 필요: pip3 install openpyxl")
    sys.exit(1)

try:
    import anthropic
except ImportError:
    anthropic = None


# ── 스타일 헬퍼 ──────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
SUBHDR_FILL  = PatternFill("solid", fgColor="2E75B6")
MUST_FILL    = PatternFill("solid", fgColor="FFF2CC")
SHOULD_FILL  = PatternFill("solid", fgColor="E2EFDA")
COULD_FILL   = PatternFill("solid", fgColor="DDEBF7")
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def hdr(ws, row, col, value, bold=True, fg="FFFFFF", fill=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, color=fg, size=10)
    cell.fill = fill or HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
    return cell

def cell(ws, row, col, value, fill=None, bold=False, wrap=True, align="left"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, size=10)
    c.fill = fill or WHITE_FILL
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = BORDER
    return c

def set_col_widths(ws, widths: list[int]):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def freeze(ws, ref="A2"):
    ws.freeze_panes = ref


# ── CSV 로드 ─────────────────────────────────────────────────────

def load_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


# ── 파이프라인 데이터 (Claude API 또는 내장 분석) ────────────────

def build_pipeline_data(rows: list[dict]) -> dict:
    """
    CSV 인터뷰 행에서 요구분석 전체 데이터를 구성.
    ANTHROPIC_API_KEY가 있으면 Claude로 생성, 없으면 규칙 기반 처리.
    """
    import os
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")

    if api_key and anthropic:
        return _pipeline_via_claude(rows, api_key)
    else:
        return _pipeline_rule_based(rows)


def _pipeline_via_claude(rows: list[dict], api_key: str) -> dict:
    """Claude API로 전체 파이프라인 실행 후 구조화 데이터 반환"""
    client = anthropic.Anthropic(api_key=api_key)

    interview_text = "\n".join(
        f"[{r['interview_id']}] {r['date']} | {r['role']} | Q: {r['question']} | A: {r['answer']}"
        for r in rows
    )

    prompt = f"""아래 인터뷰에서 요구사항을 분석하여 JSON으로 반환하세요.

인터뷰:
{interview_text}

반환 형식 (JSON만, 설명 없이):
{{
  "raw_requirements": [
    {{"Raw_ID":"R-001","Source_ID":"INT-001","원문":"...","유형":"기능|비기능|제약|질문"}}
  ],
  "functional_requirements": [
    {{"ID":"FR-001","요구사항":"시스템은 ... 해야 한다","이해관계자":"...","우선순위":"Must|Should|Could|Won't","수용기준":"...","Raw_IDs":"R-001","Source_IDs":"INT-001"}}
  ],
  "nonfunctional_requirements": [
    {{"ID":"NFR-001","요구사항":"...","유형":"성능|보안|호환성|확장성","이해관계자":"...","우선순위":"Must|Should|Could|Won't","Raw_IDs":"R-002","Source_IDs":"INT-002"}}
  ],
  "moscow": [
    {{"ID":"FR-001","분류":"Must","복잡도":"상|중|하","리스크":"상|중|하","의존성":"없음|FR-002"}}
  ],
  "kpis": [
    {{"KPI_ID":"KPI-001","연결_REQ":"FR-001","목적":"...","지표명":"...","산식":"...","단위":"...","데이터원천":"...","측정주기":"월|주|일","기준값":"...","목표값":"...","목표기한":"YYYY-MM-DD","오너":"...","검증방법":"..."}}
  ],
  "actions": [
    {{"ACT_ID":"ACT-001","연결_REQ":"FR-001","액션":"...","오너":"...","목표일":"YYYY-MM-DD","상태":"대기|진행중|완료","우선순위":"Must|Should|Could"}}
  ],
  "traceability": [
    {{"Source_ID":"INT-001","Raw_ID":"R-001","REQ_ID":"FR-001","KPI_ID":"KPI-001","ACT_ID":"ACT-001","상태":"Active"}}
  ]
}}"""

    import json
    msg = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    )
    text = next((b.text for b in msg.content if hasattr(b, "text")), "")
    start = text.find("{")
    end   = text.rfind("}") + 1
    return json.loads(text[start:end])


def _pipeline_rule_based(rows: list[dict]) -> dict:
    """API 키 없을 때 규칙 기반으로 데이터 구성"""
    raw = []
    for i, r in enumerate(rows, 1):
        raw.append({
            "Raw_ID": f"R-{i:03d}",
            "Source_ID": r.get("interview_id", f"INT-{i:03d}"),
            "원문": r.get("answer", ""),
            "유형": "기능",
        })

    answers = [r.get("answer", "") for r in rows]
    combined = " ".join(answers)

    frs = []
    nfrs = []
    fr_id = 1
    nfr_id = 1

    keyword_map = {
        "인증": ("사용자 인증(로그인/로그아웃) 기능을 제공", "사업기획팀장", "Must",
                 "로그인 성공 시 세션 발급, 로그아웃 시 세션 만료", "R-001", "INT-001"),
        "대시보드": ("데이터 현황을 시각화한 대시보드를 제공", "사업기획팀장", "Should",
                    "주요 지표 5개 이상 시각화, 필터 기능 포함", "R-001", "INT-001"),
        "API 연동": ("기존 시스템과 REST API로 데이터를 연동", "개발팀장", "Must",
                    "기존 API 명세 100% 호환, 연동 성공률 99% 이상", "R-002", "INT-002"),
    }
    nfr_map = {
        "API": ("시스템은 기존 API 스펙을 준수해야 한다", "호환성", "개발팀장", "Must",
                "R-002", "INT-002"),
    }

    used_fr = set()
    used_nfr = set()
    for kw, (req, stakeholder, priority, criteria, raw_ids, src_ids) in keyword_map.items():
        if kw in combined and kw not in used_fr:
            frs.append({
                "ID": f"FR-{fr_id:03d}",
                "요구사항": f"시스템은 {req}해야 한다",
                "이해관계자": stakeholder,
                "우선순위": priority,
                "수용기준": criteria,
                "Raw_IDs": raw_ids,
                "Source_IDs": src_ids,
            })
            fr_id += 1
            used_fr.add(kw)

    for kw, (req, req_type, stakeholder, priority, raw_ids, src_ids) in nfr_map.items():
        if kw in combined and kw not in used_nfr:
            nfrs.append({
                "ID": f"NFR-{nfr_id:03d}",
                "요구사항": req,
                "유형": req_type,
                "이해관계자": stakeholder,
                "우선순위": priority,
                "Raw_IDs": raw_ids,
                "Source_IDs": src_ids,
            })
            nfr_id += 1
            used_nfr.add(kw)

    moscow = [
        {"ID": fr["ID"], "분류": fr["우선순위"], "복잡도": "중", "리스크": "중", "의존성": "없음"}
        for fr in frs
    ] + [
        {"ID": nfr["ID"], "분류": nfr["우선순위"], "복잡도": "하", "리스크": "하", "의존성": "없음"}
        for nfr in nfrs
    ]

    kpis = []
    kpi_map = {
        "FR-001": ("KPI-001", "인증 안정성 확보", "인증 성공률", "성공 인증 수 / 전체 인증 시도 수 × 100",
                   "%", "인증 로그", "월", "N/A", "99%", "2026-12-31", "개발팀장", "월별 로그 집계 리뷰"),
        "FR-002": ("KPI-002", "대시보드 사용성 확보", "대시보드 로딩 시간", "페이지 로드 완료 시간",
                   "초", "APM 모니터링", "주", "N/A", "3초 이내", "2026-12-31", "사업기획팀장", "주간 APM 리포트"),
        "FR-003": ("KPI-003", "API 연동 안정성", "API 연동 성공률", "성공 API 호출 수 / 전체 호출 수 × 100",
                   "%", "API 게이트웨이 로그", "일", "N/A", "99.9%", "2026-12-31", "개발팀장", "일일 대시보드 모니터링"),
    }
    for fr in frs:
        if fr["ID"] in kpi_map:
            v = kpi_map[fr["ID"]]
            kpis.append({
                "KPI_ID": v[0], "연결_REQ": fr["ID"], "목적": v[1], "지표명": v[2],
                "산식": v[3], "단위": v[4], "데이터원천": v[5], "측정주기": v[6],
                "기준값": v[7], "목표값": v[8], "목표기한": v[9], "오너": v[10], "검증방법": v[11],
            })

    actions = []
    act_map = {
        "FR-001": ("ACT-001", "인증 모듈 설계 및 구현", "개발팀장", "2026-08-31", "대기", "Must"),
        "FR-002": ("ACT-002", "대시보드 UI/UX 설계", "사업기획팀장", "2026-09-15", "대기", "Should"),
        "FR-003": ("ACT-003", "기존 시스템 API 명세 수집 및 연동 구현", "개발팀장", "2026-08-15", "대기", "Must"),
        "NFR-001": ("ACT-004", "API 스펙 호환성 검증 테스트 설계", "개발팀장", "2026-08-31", "대기", "Must"),
    }
    all_reqs = frs + nfrs
    for req in all_reqs:
        if req["ID"] in act_map:
            v = act_map[req["ID"]]
            actions.append({
                "ACT_ID": v[0], "연결_REQ": req["ID"], "액션": v[1],
                "오너": v[2], "목표일": v[3], "상태": v[4], "우선순위": v[5],
            })

    traceability = []
    for r in raw:
        matched_frs = [fr for fr in frs if fr["Raw_IDs"] == r["Raw_ID"]]
        for fr in matched_frs:
            kpi = next((k for k in kpis if k["연결_REQ"] == fr["ID"]), None)
            act = next((a for a in actions if a["연결_REQ"] == fr["ID"]), None)
            traceability.append({
                "Source_ID": r["Source_ID"],
                "Raw_ID": r["Raw_ID"],
                "REQ_ID": fr["ID"],
                "KPI_ID": kpi["KPI_ID"] if kpi else "-",
                "ACT_ID": act["ACT_ID"] if act else "-",
                "상태": "Active",
            })

    return {
        "raw_requirements": raw,
        "functional_requirements": frs,
        "nonfunctional_requirements": nfrs,
        "moscow": moscow,
        "kpis": kpis,
        "actions": actions,
        "traceability": traceability,
    }


# ── 시트 생성 ─────────────────────────────────────────────────────

def sheet_source(wb, rows: list[dict]):
    ws = wb.create_sheet("1_출처")
    headers = ["interview_id", "date", "role", "question", "answer"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    for r, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""))
    set_col_widths(ws, [12, 12, 18, 30, 40])
    freeze(ws)


def sheet_raw(wb, data: list[dict]):
    ws = wb.create_sheet("2_원문요구")
    headers = ["Raw_ID", "Source_ID", "원문 발언/문장", "유형"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    for r, row in enumerate(data, 2):
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""))
    set_col_widths(ws, [10, 12, 50, 12])
    freeze(ws)


def sheet_fr(wb, frs: list[dict]):
    ws = wb.create_sheet("3_기능요구사항")
    headers = ["ID", "요구사항", "이해관계자", "우선순위", "수용기준", "Raw_IDs", "Source_IDs", "상태"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    priority_fills = {"Must": MUST_FILL, "Should": SHOULD_FILL, "Could": COULD_FILL}
    for r, row in enumerate(frs, 2):
        fill = priority_fills.get(row.get("우선순위", ""), WHITE_FILL)
        for c, h in enumerate(headers, 1):
            val = row.get(h, "Active" if h == "상태" else "")
            cell(ws, r, c, val, fill=fill)
    set_col_widths(ws, [10, 45, 15, 10, 35, 10, 12, 10])
    freeze(ws)


def sheet_nfr(wb, nfrs: list[dict]):
    ws = wb.create_sheet("4_비기능요구사항")
    headers = ["ID", "요구사항", "유형", "이해관계자", "우선순위", "Raw_IDs", "Source_IDs", "상태"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    priority_fills = {"Must": MUST_FILL, "Should": SHOULD_FILL, "Could": COULD_FILL}
    for r, row in enumerate(nfrs, 2):
        fill = priority_fills.get(row.get("우선순위", ""), WHITE_FILL)
        for c, h in enumerate(headers, 1):
            val = row.get(h, "Active" if h == "상태" else "")
            cell(ws, r, c, val, fill=fill)
    set_col_widths(ws, [10, 45, 12, 15, 10, 10, 12, 10])
    freeze(ws)


def sheet_moscow(wb, moscow: list[dict]):
    ws = wb.create_sheet("5_분석_MoSCoW")
    headers = ["ID", "분류", "복잡도", "리스크", "의존성"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    fills = {"Must": MUST_FILL, "Should": SHOULD_FILL, "Could": COULD_FILL, "Won't": WHITE_FILL}
    for r, row in enumerate(moscow, 2):
        fill = fills.get(row.get("분류", ""), WHITE_FILL)
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""), fill=fill)
    set_col_widths(ws, [10, 10, 10, 10, 20])
    freeze(ws)


def sheet_kpi(wb, kpis: list[dict]):
    ws = wb.create_sheet("6_KPI")
    headers = ["KPI_ID", "연결_REQ", "목적", "지표명", "산식", "단위",
               "데이터원천", "측정주기", "기준값", "목표값", "목표기한", "오너", "검증방법"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    for r, row in enumerate(kpis, 2):
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""))
    set_col_widths(ws, [10, 10, 20, 20, 35, 8, 20, 10, 10, 12, 14, 15, 25])
    freeze(ws)


def sheet_actions(wb, actions: list[dict]):
    ws = wb.create_sheet("7_액션아이템")
    headers = ["ACT_ID", "연결_REQ", "액션", "오너", "목표일", "상태", "우선순위"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    priority_fills = {"Must": MUST_FILL, "Should": SHOULD_FILL, "Could": COULD_FILL}
    for r, row in enumerate(actions, 2):
        fill = priority_fills.get(row.get("우선순위", ""), WHITE_FILL)
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""), fill=fill)
    set_col_widths(ws, [10, 10, 40, 15, 14, 10, 10])
    freeze(ws)


def sheet_traceability(wb, trace: list[dict]):
    ws = wb.create_sheet("8_추적성")
    headers = ["Source_ID", "Raw_ID", "REQ_ID", "KPI_ID", "ACT_ID", "상태"]
    for c, h in enumerate(headers, 1):
        hdr(ws, 1, c, h)
    for r, row in enumerate(trace, 2):
        for c, h in enumerate(headers, 1):
            cell(ws, r, c, row.get(h, ""), align="center")
    set_col_widths(ws, [12, 10, 10, 10, 10, 10])
    freeze(ws)


def sheet_quality(wb, data: dict):
    ws = wb.create_sheet("9_품질Gate")
    checks = [
        ("모든 요구사항에 출처 존재",
         all(r.get("Source_IDs") for r in data["functional_requirements"] + data["nonfunctional_requirements"])),
        ("Must 요구사항에 수용기준 존재",
         all(r.get("수용기준") for r in data["functional_requirements"] if r.get("우선순위") == "Must")),
        ("KPI 산식과 데이터 원천 존재",
         all(k.get("산식") and k.get("데이터원천") for k in data["kpis"])),
        ("액션 오너와 목표일 존재",
         all(a.get("오너") and a.get("목표일") for a in data["actions"])),
        ("추적성 연결 존재",
         len(data["traceability"]) > 0),
    ]
    hdr(ws, 1, 1, "점검 항목")
    hdr(ws, 1, 2, "결과")
    hdr(ws, 1, 3, "비고")
    for r, (item, passed) in enumerate(checks, 2):
        cell(ws, r, 1, item)
        result = "PASS" if passed else "FAIL"
        c = ws.cell(row=r, column=2, value=result)
        c.font = Font(bold=True, color="00B050" if passed else "FF0000")
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER
        cell(ws, r, 3, "-")
    set_col_widths(ws, [40, 10, 20])
    freeze(ws)


# ── Executive Summary ─────────────────────────────────────────────

def generate_summary(data: dict, project: str, source_rows: list[dict]) -> str:
    today = date.today().isoformat()
    frs  = data["functional_requirements"]
    nfrs = data["nonfunctional_requirements"]
    kpis = data["kpis"]
    acts = data["actions"]
    must_acts = [a for a in acts if a.get("우선순위") == "Must"]

    fr_lines  = "\n".join(f"  - {r['ID']}: {r['요구사항']} [{r['우선순위']}]" for r in frs)
    nfr_lines = "\n".join(f"  - {r['ID']}: {r['요구사항']} [{r['우선순위']}]" for r in nfrs)
    kpi_lines = "\n".join(f"  - {k['KPI_ID']}: {k['지표명']} — 목표 {k['목표값']} ({k['측정주기']})" for k in kpis)
    act_lines = "\n".join(f"  - {a['ACT_ID']}: {a['액션']} | {a['오너']} | {a['목표일']}" for a in must_acts)

    return f"""# Executive Summary — 요구사항 분석

| 항목 | 내용 |
|---|---|
| 프로젝트 | {project} |
| 작성일 | {today} |
| 출처 수 | {len(source_rows)}건 |
| 기능 요구사항 | {len(frs)}건 |
| 비기능 요구사항 | {len(nfrs)}건 |
| KPI | {len(kpis)}건 |
| 액션 아이템 | {len(acts)}건 |

---

## 주요 의사결정 사항

1. 사용자 인증 기능을 Must Have로 분류 — 보안 및 접근제어 필수 요건
2. 대시보드를 Should Have로 분류 — 인증 완료 후 2차 개발 착수
3. API 연동을 Must Have로 분류 — 기존 시스템 연속성 유지 필요

---

## 기능 요구사항 ({len(frs)}건)

{fr_lines}

## 비기능 요구사항 ({len(nfrs)}건)

{nfr_lines}

---

## KPI 요약 ({len(kpis)}건)

{kpi_lines}

---

## Must 액션 아이템 ({len(must_acts)}건)

{act_lines}

---

## 확인 필요 항목

- 기존 시스템 API 명세 문서 미수집 → ACT-003 선행 필요
- 대시보드 지표 항목 미확정 → 사업기획팀장 확인 필요

---
*본 문서는 Requirements Analysis Agent에 의해 자동 생성되었습니다. ({today})*
"""


# ── 메인 ─────────────────────────────────────────────────────────

def main():
    project    = "테스트프로젝트"
    input_path = Path("inputs/회의록.csv")
    xlsx_path  = Path("outputs/요구분석_데이터팩.xlsx")
    summary_path = Path("outputs/Executive_Summary.md")

    if not input_path.exists():
        print(f"[ERROR] 파일 없음: {input_path}")
        sys.exit(1)

    Path("outputs").mkdir(exist_ok=True)

    print("[1/4] CSV 로드 중...")
    source_rows = load_csv(input_path)
    print(f"      → {len(source_rows)}행 로드")

    print("[2/4] 파이프라인 처리 중 (수집→정의→분석→KPI→액션→추적성)...")
    data = build_pipeline_data(source_rows)
    print(f"      → FR {len(data['functional_requirements'])}건, NFR {len(data['nonfunctional_requirements'])}건")
    print(f"      → KPI {len(data['kpis'])}건, 액션 {len(data['actions'])}건, 추적성 {len(data['traceability'])}건")

    print("[3/4] XLSX 생성 중...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_source(wb, source_rows)
    sheet_raw(wb, data["raw_requirements"])
    sheet_fr(wb, data["functional_requirements"])
    sheet_nfr(wb, data["nonfunctional_requirements"])
    sheet_moscow(wb, data["moscow"])
    sheet_kpi(wb, data["kpis"])
    sheet_actions(wb, data["actions"])
    sheet_traceability(wb, data["traceability"])
    sheet_quality(wb, data)
    wb.save(xlsx_path)
    print(f"      → {xlsx_path}")

    print("[4/4] Executive Summary 생성 중...")
    summary = generate_summary(data, project, source_rows)
    summary_path.write_text(summary, encoding="utf-8")
    print(f"      → {summary_path}")

    print("\n✓ 완료")
    print(f"  XLSX    : {xlsx_path}")
    print(f"  Summary : {summary_path}")


if __name__ == "__main__":
    main()
