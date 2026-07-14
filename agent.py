"""
Requirements Analysis Agent — 4-stage pipeline
수집 → 정의 → 분석 → KPI 도출
"""

import anthropic
from pathlib import Path
from datetime import date

client = anthropic.Anthropic()
MODEL = "claude-opus-4-7"


def _stream(system: str, user: str, label: str) -> str:
    print(f"\n{'='*50}")
    print(f"  {label}")
    print(f"{'='*50}")
    result = ""
    with client.messages.stream(
        model=MODEL,
        max_tokens=8192,
        thinking={"type": "adaptive"},
        system=system,
        messages=[{"role": "user", "content": user}],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            result += text
    print()
    return result


# ── 파일 파싱 ────────────────────────────────────────────────────

def parse_pdf(path: Path) -> str:
    import pdfplumber
    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text)


def parse_docx(path: Path) -> str:
    from docx import Document
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def load_input(path: str | Path) -> str:
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix == ".pdf":
        return parse_pdf(p)
    if suffix in (".docx", ".doc"):
        return parse_docx(p)
    return p.read_text(encoding="utf-8")


# ── 4단계 파이프라인 ─────────────────────────────────────────────

def stage1_collect(raw_inputs: list[str]) -> str:
    combined = "\n\n---\n\n".join(raw_inputs)

    system = (
        "당신은 비즈니스 요구사항 분석 전문가입니다. "
        "주어진 문서·인터뷰 내용에서 요구사항 후보를 빠짐없이 추출합니다. "
        "응답은 한국어로 작성하세요."
    )
    user = f"""아래 자료에서 요구사항 후보를 모두 추출하세요.

=== 입력 자료 ===
{combined}

=== 출력 형식 ===
기능적 요구사항과 비기능적 요구사항을 구분하여 항목별로 나열하세요.
각 항목: [R번호] 요구사항 내용 (출처)"""

    return _stream(system, user, "1단계: 수집 (요구사항 후보 추출)")


def stage2_define(collected: str) -> str:
    system = (
        "당신은 비즈니스 요구사항 정의 전문가입니다. "
        "수집된 요구사항을 체계적으로 분류·구조화합니다. "
        "응답은 한국어로 작성하세요."
    )
    user = f"""아래 요구사항 목록을 정의하고 구조화하세요.

=== 수집된 요구사항 ===
{collected}

=== 작업 내용 ===
1. 기능적(FR) / 비기능적(NFR) 분류
2. 중복 통합 및 고유 ID 부여 (FR-001, NFR-001 형식)
3. 이해관계자 매핑

=== 출력 형식 (Markdown 표) ===
## 기능적 요구사항 (Functional Requirements)
| ID | 요구사항 | 이해관계자 | 우선순위 |
|---|---|---|---|

## 비기능적 요구사항 (Non-Functional Requirements)
| ID | 요구사항 | 유형 | 이해관계자 |
|---|---|---|---|"""

    return _stream(system, user, "2단계: 정의 (분류 및 구조화)")


def stage3_analyze(defined: str) -> str:
    system = (
        "당신은 비즈니스 요구사항 분석 전문가입니다. "
        "우선순위·실현가능성·이해관계자 영향을 평가합니다. "
        "응답은 한국어로 작성하세요."
    )
    user = f"""아래 요구사항을 분석하세요.

=== 정의된 요구사항 ===
{defined}

=== 분석 항목 ===
1. **MoSCoW 우선순위 매트릭스**
   - Must Have / Should Have / Could Have / Won't Have

2. **실현 가능성 평가**
   - 기술 복잡도(상/중/하), 비용 대비 효과, 구현 리스크

3. **이해관계자 영향도 분석**
   - 주요 이해관계자 및 각 요구사항의 영향

각 항목을 Markdown 표와 텍스트로 작성하세요."""

    return _stream(system, user, "3단계: 분석 (우선순위 및 실현가능성)")


def stage4_kpi(defined: str, analyzed: str) -> str:
    system = (
        "당신은 비즈니스 KPI 설계 전문가입니다. "
        "SMART 원칙에 따라 측정 가능한 핵심 성과 지표를 도출합니다. "
        "응답은 한국어로 작성하세요."
    )
    user = f"""아래 요구사항과 분석을 바탕으로 KPI를 도출하세요.

=== 정의된 요구사항 ===
{defined}

=== 분석 결과 ===
{analyzed}

=== KPI 도출 원칙 (SMART) ===
Specific·Measurable·Achievable·Relevant·Time-bound

=== 출력 형식 ===
## KPI 목록
| KPI ID | 목표 | 지표명 | 측정 방법 | 현재값(Baseline) | 목표값 | 측정 주기 |
|---|---|---|---|---|---|---|

## KPI 모니터링 방법
(대시보드 구조 및 리뷰 주기)"""

    return _stream(system, user, "4단계: KPI 도출")


# ── 보고서 생성 ──────────────────────────────────────────────────

def generate_report(
    project: str,
    collected: str,
    defined: str,
    analyzed: str,
    kpi: str,
) -> str:
    today = date.today().isoformat()
    return f"""# 요구사항 분석 보고서

| 항목 | 내용 |
|---|---|
| 프로젝트 | {project} |
| 작성일 | {today} |
| 도메인 | 비즈니스/사업 기획 |

---

## 1. 수집 — 요구사항 후보

{collected}

---

## 2. 정의 — 분류 및 구조화

{defined}

---

## 3. 분석 — 우선순위 및 실현가능성

{analyzed}

---

## 4. KPI 도출

{kpi}

---
*본 보고서는 Requirements Analysis Agent에 의해 자동 생성되었습니다.*
"""


# ── 공개 API ─────────────────────────────────────────────────────

def run(
    project: str,
    file_paths: list[str] | None = None,
    interview_text: str | None = None,
    output_path: str | None = None,
) -> str:
    """
    Run the full 4-stage pipeline and return the Markdown report.

    Args:
        project: 프로젝트 이름
        file_paths: 분석할 문서 경로 목록 (PDF, DOCX, TXT)
        interview_text: 인터뷰/회의록 텍스트 (직접 입력)
        output_path: 보고서 저장 경로 (기본: outputs/<project>.md)
    """
    raw_inputs: list[str] = []

    if file_paths:
        for fp in file_paths:
            print(f"[파일 로드] {fp}")
            text = load_input(fp)
            raw_inputs.append(f"[파일: {Path(fp).name}]\n{text}")

    if interview_text:
        raw_inputs.append(f"[인터뷰/회의록]\n{interview_text}")

    if not raw_inputs:
        raise ValueError("file_paths 또는 interview_text 중 하나는 필요합니다.")

    collected = stage1_collect(raw_inputs)
    defined   = stage2_define(collected)
    analyzed  = stage3_analyze(defined)
    kpi       = stage4_kpi(defined, analyzed)

    report = generate_report(project, collected, defined, analyzed, kpi)

    if output_path is None:
        safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in project)
        output_path = f"outputs/{safe_name}.md"

    Path(output_path).write_text(report, encoding="utf-8")
    print(f"\n✓ 보고서 저장: {output_path}")

    return report
