"""
CLI entry point for the Requirements Analysis Agent.

Usage:
    python run.py --project "프로젝트명" --files doc1.pdf doc2.docx --interview interview.txt
    python run.py --project "프로젝트명" --interview "회의록 직접 입력 텍스트"
"""

import argparse
import sys
from pathlib import Path

from agent import run


def main() -> None:
    parser = argparse.ArgumentParser(
        description="요구사항 분석 에이전트 — 수집/정의/분석/KPI 4단계 파이프라인"
    )
    parser.add_argument(
        "--project", "-p",
        required=True,
        help="프로젝트 이름 (보고서 제목 및 파일명에 사용)",
    )
    parser.add_argument(
        "--files", "-f",
        nargs="*",
        metavar="FILE",
        help="분석할 문서 경로 (PDF, DOCX, TXT). 여러 파일 공백으로 구분.",
    )
    parser.add_argument(
        "--interview", "-i",
        metavar="TEXT_OR_FILE",
        help="인터뷰/회의록 텍스트 또는 텍스트 파일 경로",
    )
    parser.add_argument(
        "--output", "-o",
        metavar="PATH",
        help="보고서 저장 경로 (기본: outputs/<project>.md)",
    )

    args = parser.parse_args()

    if not args.files and not args.interview:
        parser.error("--files 또는 --interview 중 하나는 필요합니다.")

    interview_text: str | None = None
    if args.interview:
        p = Path(args.interview)
        if p.exists() and p.is_file():
            interview_text = p.read_text(encoding="utf-8")
            print(f"[인터뷰 파일 로드] {p}")
        else:
            interview_text = args.interview

    Path("outputs").mkdir(exist_ok=True)

    try:
        run(
            project=args.project,
            file_paths=args.files or None,
            interview_text=interview_text,
            output_path=args.output,
        )
    except KeyboardInterrupt:
        print("\n\n[중단됨]")
        sys.exit(1)


if __name__ == "__main__":
    main()
