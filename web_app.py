#!/usr/bin/env python3
"""요구사항 분석 에이전트 웹 서비스"""

from __future__ import annotations
import csv
import io
import json
import os
import queue
import sys
import tempfile
import threading
import time
import uuid
from datetime import date, datetime
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file

# 기존 모듈 경로 추가
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from auto_convert import text_to_rows, save_csv, load_file
import process_csv as pc
import generate_dashboard as gd

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

SUPPORTED = {".txt", ".md", ".docx", ".pdf", ".eml"}

# 작업별 로그 큐 (job_id → queue)
_job_queues: dict[str, queue.Queue] = {}
# 작업별 결과 (job_id → dict)
_job_results: dict[str, dict] = {}
_lock = threading.Lock()


# ── 헬퍼 ─────────────────────────────────────────────────────────

def new_job() -> str:
    jid = str(uuid.uuid4())
    with _lock:
        _job_queues[jid] = queue.Queue()
        _job_results[jid] = {"status": "running"}
    return jid


def log(jid: str, msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with _lock:
        if jid in _job_queues:
            _job_queues[jid].put(line)


def finish_job(jid: str, result: dict) -> None:
    with _lock:
        _job_results[jid] = result
        if jid in _job_queues:
            _job_queues[jid].put("__DONE__")


# ── 파이프라인 (스레드에서 실행) ─────────────────────────────────

def run_pipeline(jid: str, text: str, filename: str) -> None:
    try:
        log(jid, f"=== 요구사항 분석 시작: {filename} ===")

        # Step 1: 텍스트 → CSV 행
        log(jid, "Step 1/4  텍스트 파싱 중...")
        rows = text_to_rows(text, Path(filename).stem)
        if not rows:
            raise ValueError("파싱된 내용이 없습니다. 파일 내용을 확인해 주세요.")
        log(jid, f"  → {len(rows)}개 항목 파싱 완료")

        # Step 2: CSV 저장 (임시)
        tmp_dir = Path(tempfile.mkdtemp())
        stem = Path(filename).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = tmp_dir / f"{stem}_{timestamp}.csv"
        save_csv(rows, csv_path)
        log(jid, "Step 2/4  CSV 변환 완료")

        # Step 3: XLSX 생성
        log(jid, "Step 3/4  요구분석 XLSX 생성 중...")
        xlsx_path = tmp_dir / f"{stem}_요구분석.xlsx"
        summary_path = tmp_dir / f"{stem}_Executive_Summary.md"

        source_rows = pc.load_csv(csv_path)
        data = pc.build_pipeline_data(source_rows)
        import openpyxl
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        pc.sheet_source(wb, source_rows)
        pc.sheet_raw(wb, data["raw_requirements"])
        pc.sheet_fr(wb, data["functional_requirements"])
        pc.sheet_nfr(wb, data["nonfunctional_requirements"])
        pc.sheet_moscow(wb, data["moscow"])
        pc.sheet_kpi(wb, data["kpis"])
        pc.sheet_actions(wb, data["actions"])
        pc.sheet_traceability(wb, data["traceability"])
        pc.sheet_quality(wb, data)
        wb.save(xlsx_path)

        doc_title = stem.replace("_", " ")
        summary_text_content = pc.generate_summary(data, doc_title, source_rows)
        summary_path.write_text(summary_text_content, encoding="utf-8")
        log(jid, f"  → XLSX {xlsx_path.name} 생성 완료")

        # Step 4: 대시보드 HTML
        log(jid, "Step 4/4  대시보드 생성 중...")
        dashboard_path = tmp_dir / "dashboard.html"
        gd.generate(
            xlsx_path=xlsx_path,
            output_path=dashboard_path,
            project=doc_title,
            mode=f"단일 분석 · {filename}",
        )
        log(jid, "  → 대시보드 생성 완료")
        log(jid, "=== 분석 완료 ===")

        dashboard_html = dashboard_path.read_text(encoding="utf-8")
        xlsx_bytes = xlsx_path.read_bytes()
        summary_text = summary_path.read_text(encoding="utf-8") if summary_path.exists() else summary_text_content

        finish_job(jid, {
            "status": "done",
            "dashboard_html": dashboard_html,
            "xlsx_bytes": xlsx_bytes,
            "summary": summary_text,
            "filename": stem,
        })

    except Exception as e:
        log(jid, f"[오류] {e}")
        finish_job(jid, {"status": "error", "message": str(e)})


# ── 라우트 ───────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    jid = new_job()

    # 파일 업로드 vs 텍스트 입력
    uploaded = request.files.get("file")
    text_input = request.form.get("text", "").strip()

    if uploaded and uploaded.filename:
        filename = uploaded.filename
        suffix = Path(filename).suffix.lower()
        if suffix not in SUPPORTED:
            return jsonify({"error": f"지원하지 않는 형식: {suffix}. ({', '.join(SUPPORTED)})"}), 400

        # 임시 파일로 저장 후 읽기
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            uploaded.save(tmp.name)
            tmp_path = Path(tmp.name)
        try:
            text = load_file(tmp_path)
        finally:
            tmp_path.unlink(missing_ok=True)

    elif text_input:
        filename = "직접입력_회의록.txt"
        text = text_input
    else:
        return jsonify({"error": "파일 또는 텍스트를 입력해 주세요."}), 400

    if not text.strip():
        return jsonify({"error": "파일 내용이 비어 있습니다."}), 400

    threading.Thread(target=run_pipeline, args=(jid, text, filename), daemon=True).start()
    return jsonify({"job_id": jid})


@app.route("/stream/<jid>")
def stream(jid: str):
    def generate():
        with _lock:
            q = _job_queues.get(jid)
        if not q:
            yield "data: [오류] 작업을 찾을 수 없습니다.\n\ndata: __DONE__\n\n"
            return
        while True:
            try:
                msg = q.get(timeout=30)
                yield f"data: {msg}\n\n"
                if msg == "__DONE__":
                    break
            except queue.Empty:
                yield "data: [대기 중...]\n\n"

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/result/<jid>")
def result(jid: str):
    with _lock:
        r = _job_results.get(jid, {})
    if not r or r.get("status") == "running":
        return jsonify({"status": "running"})
    if r.get("status") == "error":
        return jsonify({"status": "error", "message": r.get("message", "")})
    return jsonify({
        "status": "done",
        "dashboard_html": r.get("dashboard_html", ""),
        "summary": r.get("summary", ""),
        "filename": r.get("filename", "result"),
    })


@app.route("/download/<jid>")
def download(jid: str):
    with _lock:
        r = _job_results.get(jid, {})
    if not r or r.get("status") != "done":
        return "결과 없음", 404
    xlsx_bytes = r.get("xlsx_bytes", b"")
    filename = r.get("filename", "result")
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{filename}_요구분석.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
