#!/usr/bin/env python3
"""파이프라인 완료 후 대시보드 이메일 발송"""

from __future__ import annotations
import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import date


def send_dashboard(dashboard_path: Path, recipients: list[str],
                   sources: list[str] | None = None, mode: str = "통합") -> None:
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pw   = os.environ.get("GMAIL_APP_PASSWORD")

    if not gmail_user or not gmail_pw:
        print("[이메일 SKIP] GMAIL_USER 또는 GMAIL_APP_PASSWORD 환경변수 없음")
        sys.exit(0)

    today = date.today().isoformat()
    sources = sources or []

    # 원본 파일명 목록 (경로 제거, 확장자 유지)
    source_names = [Path(s).name for s in sources]
    source_label = ", ".join(source_names) if source_names else "알 수 없음"

    # Executive Summary 본문 읽기 (첫 번째 파일 기준)
    summary_text = ""
    for src in source_names:
        stem = Path(src).stem
        summary_path = dashboard_path.parent / f"{stem}_Executive_Summary.md"
        if summary_path.exists():
            summary_text = summary_path.read_text(encoding="utf-8")
            break
    if not summary_text:
        fallback = dashboard_path.parent / "DAM_회의록_Executive_Summary.md"
        summary_text = fallback.read_text(encoding="utf-8") if fallback.exists() else ""

    msg = MIMEMultipart()
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(recipients)
    mode_label = "단일 분석" if mode == "단일" else "통합 분석"
    msg["Subject"] = f"[DAM {mode_label}] {source_label} — {today}"

    source_block = "\n".join(f"  • {n}" for n in source_names) if source_names else "  • 알 수 없음"

    body = f"""안녕하세요,

DAM 구축 프로젝트 요구분석 에이전트가 [{mode_label}]을 완료했습니다.
대시보드 HTML 파일을 첨부했습니다.

[분석 유형] {mode_label}
[분석 원본 파일]
{source_block}

{'─'*40}
{summary_text[:1500] if summary_text else '요약 없음'}
{'─'*40}

Requirements Analysis Agent
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # dashboard.html 첨부
    with dashboard_path.open("rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="dashboard_{today}.html"',
        )
        msg.attach(part)

    last_err = None
    # 포트 465(SSL) → 587(STARTTLS) 순서로 시도
    for port, use_ssl in [(465, True), (587, False)]:
        try:
            if use_ssl:
                server = smtplib.SMTP_SSL("smtp.gmail.com", port, timeout=15)
            else:
                server = smtplib.SMTP("smtp.gmail.com", port, timeout=15)
                server.ehlo()
                server.starttls()
                server.ehlo()
            server.login(gmail_user, gmail_pw)
            server.sendmail(gmail_user, recipients, msg.as_string())
            server.quit()
            print(f"[이메일 발송] 포트 {port} → {', '.join(recipients)}")
            return
        except Exception as e:
            last_err = e
            print(f"[이메일] 포트 {port} 실패: {e}")

    print(f"[이메일 발송 실패] 465/587 모두 불가: {last_err}")
    sys.exit(1)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--dashboard", required=True)
    parser.add_argument("--to",        required=True, help="수신자 (콤마 구분)")
    parser.add_argument("--sources",   default="",     help="원본 파일명 목록 (콤마 구분)")
    parser.add_argument("--mode",      default="통합", help="단일 또는 통합")
    args = parser.parse_args()

    sources = [s.strip() for s in args.sources.split(",")] if args.sources else []
    send_dashboard(
        Path(args.dashboard),
        [r.strip() for r in args.to.split(",")],
        sources=sources,
        mode=args.mode,
    )
