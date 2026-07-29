#!/usr/bin/env python3
"""
드롭 폴더 파일 자동 변환기
.txt / .docx / .pdf / .eml → inputs/*.csv
"""

from __future__ import annotations
import csv
import re
import sys
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Optional


DROP_DIR      = Path(__file__).parent / "drop"
PROCESSED_DIR = DROP_DIR / "processed"
FAILED_DIR    = DROP_DIR / "failed"
INPUTS_DIR    = Path(__file__).parent / "inputs"

SUPPORTED = {".txt", ".docx", ".pdf", ".eml", ".md"}


# ── 파일 읽기 ────────────────────────────────────────────────────

def read_txt(path: Path) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def read_docx(path: Path) -> str:
    from docx import Document
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def read_pdf(path: Path) -> str:
    import pdfplumber
    pages = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            t = pg.extract_text()
            if t:
                pages.append(t)
    return "\n".join(pages)


def read_eml(path: Path) -> str:
    import email
    msg = email.message_from_bytes(path.read_bytes())
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    parts.append(payload.decode("utf-8", errors="replace"))
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            parts.append(payload.decode("utf-8", errors="replace"))
    return "\n".join(parts)


def load_file(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".txt" or suffix == ".md":
        return read_txt(path)
    elif suffix in (".docx", ".doc"):
        return read_docx(path)
    elif suffix == ".pdf":
        return read_pdf(path)
    elif suffix == ".eml":
        return read_eml(path)
    raise ValueError(f"지원하지 않는 형식: {suffix}")


# ── 텍스트 → 인터뷰 CSV 변환 ────────────────────────────────────

ROLE_PATTERNS = [
    r"펜타\s*PM", r"펜타\s*담당자", r"펜타",
    r"제일기획\s*담당자", r"제일기획", r"제일",
    r"PM", r"팀장", r"담당자", r"개발팀", r"기획팀",
    r"운영자", r"관리자", r"작업자",
]

SECTION_RE = re.compile(
    r"^(?:\d+[\.\)]\s+|[■●▶◆•]\s*)(.+)$",
    re.MULTILINE
)

ROLE_RE = re.compile(
    r"\[([^\]]+)\]|【([^】]+)】|<([^>]+)>|^([가-힣\w]+(?:PM|팀장|담당자|운영자|관리자))\s*[:：]",
    re.MULTILINE
)


def detect_role(text: str, default: str = "담당자") -> str:
    m = ROLE_RE.search(text)
    if m:
        role = next(g for g in m.groups() if g)
        return role.strip()
    for pat in ROLE_PATTERNS:
        if re.search(pat, text):
            return re.search(pat, text).group().strip()
    return default


def split_sections(text: str) -> list[tuple[str, str]]:
    """번호/불릿 기반으로 섹션 분리 → (제목, 내용) 리스트"""
    lines = text.strip().split("\n")
    sections: list[tuple[str, list[str]]] = []
    current_title = ""
    current_body: list[str] = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(\d+[\.\)]\s+|[■●▶◆•]\s*)(.+)$", line)
        if m:
            if current_title:
                sections.append((current_title, current_body))
            current_title = m.group(2).strip()
            current_body = []
        else:
            if current_title:
                current_body.append(line)
            else:
                current_title = line
                current_body = []

    if current_title:
        sections.append((current_title, current_body))

    return sections


def section_to_answer(body: list[str]) -> str:
    """섹션 본문을 단일 답변 문자열로 압축"""
    cleaned = []
    for line in body:
        line = line.strip().lstrip("-•·ㄴ").strip()
        if line:
            cleaned.append(line)
    return ". ".join(cleaned[:6])  # 최대 6문장


def text_to_rows(text: str, source_name: str) -> list[dict]:
    today = date.today().isoformat()
    sections = split_sections(text)

    if not sections:
        # 섹션 분리 실패 시 단락 기반 분리
        paras = [p.strip() for p in text.split("\n\n") if p.strip()]
        sections = [(f"항목-{i+1}", [p]) for i, p in enumerate(paras)]

    rows = []
    for i, (title, body) in enumerate(sections, 1):
        if not body and not title:
            continue
        answer = section_to_answer(body) if body else title
        if len(answer) < 5:
            continue

        role = detect_role(" ".join(body), "담당자")
        rows.append({
            "interview_id": f"INT-{i:03d}",
            "date":         today,
            "role":         role,
            "question":     title[:80],
            "answer":       answer[:300],
        })

    return rows


# ── CSV 저장 ─────────────────────────────────────────────────────

def save_csv(rows: list[dict], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["interview_id", "date", "role", "question", "answer"])
        writer.writeheader()
        writer.writerows(rows)


# ── 메인 변환 로직 ────────────────────────────────────────────────

def convert_file(path: Path) -> Optional[Path]:
    print(f"[변환 시작] {path.name}")
    try:
        text = load_file(path)
        if not text.strip():
            raise ValueError("파일 내용 없음")

        rows = text_to_rows(text, path.stem)
        if not rows:
            raise ValueError("변환된 행 없음")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = re.sub(r"[^\w가-힣-]", "_", path.stem)
        out_path = INPUTS_DIR / f"{stem}_{timestamp}.csv"

        save_csv(rows, out_path)
        print(f"[완료] {len(rows)}행 → {out_path.name}")

        # 처리된 파일 이동
        shutil.move(str(path), str(PROCESSED_DIR / path.name))
        return out_path

    except Exception as e:
        print(f"[실패] {path.name}: {e}")
        shutil.move(str(path), str(FAILED_DIR / path.name))
        return None


def process_drop_folder() -> list[Path]:
    results = []
    files = [f for f in DROP_DIR.iterdir()
             if f.is_file() and f.suffix.lower() in SUPPORTED]

    if not files:
        print("[드롭폴더] 처리할 파일 없음")
        return results

    for f in sorted(files):
        result = convert_file(f)
        if result:
            results.append(result)

    return results


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 인자로 파일 직접 지정
        for arg in sys.argv[1:]:
            convert_file(Path(arg))
    else:
        # 드롭 폴더 전체 처리
        converted = process_drop_folder()
        if converted:
            print(f"\n✓ {len(converted)}개 파일 변환 완료")
            import subprocess
            pipeline = str(Path(__file__).parent / "pipeline.sh")

            # 새 파일만 단독 분석
            for csv_path in converted:
                print(f"[단일 분석] {csv_path.name}")
                subprocess.run(
                    ["bash", "-c",
                     f'source $HOME/.zshrc 2>/dev/null; INPUT_FILE="{csv_path}" ANALYSIS_MODE=단일 bash {pipeline} --no-open'],
                    shell=False
                )
