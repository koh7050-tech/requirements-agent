#!/usr/bin/env python3
"""요구사항 분석 에이전트 웹 서비스"""

from __future__ import annotations
import io
import json
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
from generate_prd import generate_prd

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

SUPPORTED = {".txt", ".md", ".docx", ".pdf", ".eml"}
TMP = Path(tempfile.gettempdir()) / "req_agent_jobs"
TMP.mkdir(exist_ok=True)


def job_dir(jid: str) -> Path:
    d = TMP / jid
    d.mkdir(exist_ok=True)
    return d


def write_state(jid: str, state: dict) -> None:
    (job_dir(jid) / "state.json").write_text(
        json.dumps(state, ensure_ascii=False), encoding="utf-8"
    )


def read_state(jid: str) -> dict | None:
    p = job_dir(jid) / "state.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


def append_log(jid: str, msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(job_dir(jid) / "run.log", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def read_logs(jid: str) -> list[str]:
    p = job_dir(jid) / "run.log"
    if not p.exists():
        return []
    return p.read_text(encoding="utf-8").splitlines()


def run_pipeline(jid: str, text: str, filename: str) -> None:
    jd = job_dir(jid)
    write_state(jid, {"status": "running"})

    try:
        append_log(jid, f"=== 요구사항 분석 시작: {filename} ===")

        append_log(jid, "Step 1/4  텍스트 파싱 중...")
        rows = text_to_rows(text, Path(filename).stem)
        if not rows:
            raise ValueError("파싱된 내용이 없습니다. 파일 내용을 확인해 주세요.")
        append_log(jid, f"  → {len(rows)}개 항목 파싱 완료")

        stem = Path(filename).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = jd / f"{stem}_{timestamp}.csv"
        save_csv(rows, csv_path)
        append_log(jid, "Step 2/4  CSV 변환 완료")

        append_log(jid, "Step 3/4  요구분석 XLSX 생성 중...")
        xlsx_path = jd / f"{stem}_요구분석.xlsx"
        summary_path = jd / f"{stem}_Executive_Summary.md"

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
        append_log(jid, "  → XLSX 생성 완료")

        append_log(jid, "Step 4/4  대시보드 생성 중...")
        dashboard_path = jd / "dashboard.html"
        gd.generate(
            xlsx_path=xlsx_path,
            output_path=dashboard_path,
            project=doc_title,
            mode=f"단일 분석 · {filename}",
        )
        append_log(jid, "  → 대시보드 생성 완료")
        append_log(jid, "=== 분석 완료 ===")

        write_state(jid, {
            "status": "done",
            "filename": stem,
            "summary": summary_content,
        })

    except Exception as e:
        append_log(jid, f"[오류] {e}")
        write_state(jid, {"status": "error", "message": str(e)})


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    jid = str(uuid.uuid4())

    uploaded = request.files.get("file")
    text_input = request.form.get("text", "").strip()

    if uploaded and uploaded.filename:
        filename = uploaded.filename
        suffix = Path(filename).suffix.lower()
        if suffix not in SUPPORTED:
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
        return jsonify({"error": "파일 또는 텍스트를 입력해 주세요."}), 400

    if not text.strip():
        return jsonify({"error": "파일 내용이 비어 있습니다."}), 400

    threading.Thread(target=run_pipeline, args=(jid, text, filename), daemon=True).start()
    return jsonify({"job_id": jid})


@app.route("/status/<jid>")
def status(jid: str):
    # jid 검증 (경로 조작 방지)
    if not jid.replace("-", "").isalnum():
        return jsonify({"status": "error", "message": "잘못된 요청"}), 400

    state = read_state(jid)
    logs = read_logs(jid)

    if state is None:
        # 아직 파일이 생성되지 않았으면 running으로 간주 (스레드 시작 직후)
        return jsonify({"status": "running", "logs": logs})

    resp = {"status": state["status"], "logs": logs}

    if state["status"] == "error":
        resp["message"] = state.get("message", "알 수 없는 오류")

    if state["status"] == "done":
        jd = job_dir(jid)
        stem = state.get("filename", "result")
        dashboard_path = jd / "dashboard.html"
        resp["dashboard_html"] = dashboard_path.read_text(encoding="utf-8") if dashboard_path.exists() else ""
        resp["summary"] = state.get("summary", "")
        resp["filename"] = stem

    return jsonify(resp)


@app.route("/download-prd/<jid>")
def download_prd(jid: str):
    if not jid.replace("-", "").isalnum():
        return "잘못된 요청", 400
    state = read_state(jid)
    if not state or state.get("status") != "done":
        return "결과 없음", 404
    jd = job_dir(jid)
    xlsx_files = list(jd.glob("*_요구분석.xlsx"))
    if not xlsx_files:
        return "XLSX 없음", 404
    stem = state.get("filename", "result")
    mode = state.get("summary", "")[:20]
    prd_html = generate_prd(xlsx_files[0], stem.replace("_", " "), "")
    return send_file(
        io.BytesIO(prd_html.encode("utf-8")),
        mimetype="text/html",
        as_attachment=True,
        download_name=f"{stem}_PRD.html",
    )


@app.route("/download/<jid>")
def download(jid: str):
    if not jid.replace("-", "").isalnum():
        return "잘못된 요청", 400

    state = read_state(jid)
    if not state or state.get("status") != "done":
        return "결과 없음", 404

    jd = job_dir(jid)
    stem = state.get("filename", "result")
    xlsx_files = list(jd.glob("*_요구분석.xlsx"))
    if not xlsx_files:
        return "XLSX 없음", 404

    return send_file(
        io.BytesIO(xlsx_files[0].read_bytes()),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{stem}_요구분석.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
