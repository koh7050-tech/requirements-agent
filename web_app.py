#!/usr/bin/env python3
"""요구사항 분석 에이전트 웹 서비스"""

from __future__ import annotations
import io
import os
import sys
import tempfile
import threading
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

from auto_convert import text_to_rows, save_csv, load_file
import process_csv as pc
import generate_dashboard as gd

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

SUPPORTED = {".txt", ".md", ".docx", ".pdf", ".eml"}

# job_id → {"status": running|done|error, "logs": [...], ...}
_jobs: dict[str, dict] = {}
_lock = threading.Lock()


def new_job() -> str:
    jid = str(uuid.uuid4())
    with _lock:
        _jobs[jid] = {"status": "running", "logs": []}
    return jid


def log(jid: str, msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with _lock:
        if jid in _jobs:
            _jobs[jid]["logs"].append(line)


def run_pipeline(jid: str, text: str, filename: str) -> None:
    try:
        log(jid, f"=== 요구사항 분석 시작: {filename} ===")

        log(jid, "Step 1/4  텍스트 파싱 중...")
        rows = text_to_rows(text, Path(filename).stem)
        if not rows:
            raise ValueError("파싱된 내용이 없습니다. 파일 내용을 확인해 주세요.")
        log(jid, f"  → {len(rows)}개 항목 파싱 완료")

        tmp_dir = Path(tempfile.mkdtemp())
        stem = Path(filename).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = tmp_dir / f"{stem}_{timestamp}.csv"
        save_csv(rows, csv_path)
        log(jid, "Step 2/4  CSV 변환 완료")

        log(jid, "Step 3/4  요구분석 XLSX 생성 중...")
        xlsx_path = tmp_dir / f"{stem}_요구분석.xlsx"
        summary_path = tmp_dir / f"{stem}_Executive_Summary.md"

        import openpyxl
        source_rows = pc.load_csv(csv_path)
        data = pc.build_pipeline_data(source_rows)
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
        summary_content = pc.generate_summary(data, doc_title, source_rows)
        summary_path.write_text(summary_content, encoding="utf-8")
        log(jid, f"  → XLSX 생성 완료")

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

        with _lock:
            _jobs[jid].update({
                "status": "done",
                "dashboard_html": dashboard_path.read_text(encoding="utf-8"),
                "xlsx_bytes": xlsx_path.read_bytes(),
                "summary": summary_content,
                "filename": stem,
            })

    except Exception as e:
        log(jid, f"[오류] {e}")
        with _lock:
            _jobs[jid]["status"] = "error"
            _jobs[jid]["message"] = str(e)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    jid = new_job()

    uploaded = request.files.get("file")
    text_input = request.form.get("text", "").strip()

    if uploaded and uploaded.filename:
        filename = uploaded.filename
        suffix = Path(filename).suffix.lower()
        if suffix not in SUPPORTED:
            with _lock:
                del _jobs[jid]
            return jsonify({"error": f"지원하지 않는 형식: {suffix}. ({', '.join(SUPPORTED)})"}), 400

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
        with _lock:
            del _jobs[jid]
        return jsonify({"error": "파일 또는 텍스트를 입력해 주세요."}), 400

    if not text.strip():
        with _lock:
            del _jobs[jid]
        return jsonify({"error": "파일 내용이 비어 있습니다."}), 400

    threading.Thread(target=run_pipeline, args=(jid, text, filename), daemon=True).start()
    return jsonify({"job_id": jid})


@app.route("/status/<jid>")
def status(jid: str):
    with _lock:
        job = _jobs.get(jid)
    if not job:
        return jsonify({"status": "error", "message": "작업을 찾을 수 없습니다."}), 404

    resp = {"status": job["status"], "logs": job.get("logs", [])}
    if job["status"] == "error":
        resp["message"] = job.get("message", "")
    if job["status"] == "done":
        resp["dashboard_html"] = job.get("dashboard_html", "")
        resp["summary"] = job.get("summary", "")
        resp["filename"] = job.get("filename", "result")
    return jsonify(resp)


@app.route("/download/<jid>")
def download(jid: str):
    with _lock:
        job = _jobs.get(jid)
    if not job or job.get("status") != "done":
        return "결과 없음", 404
    return send_file(
        io.BytesIO(job["xlsx_bytes"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{job['filename']}_요구분석.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
