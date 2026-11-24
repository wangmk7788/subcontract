import io
import os
import re
import zipfile
from urllib.parse import quote
from flask import Flask, render_template, request, jsonify, send_file

from excel_parser import parse_project_excel
from docx_generator import generate_subcontract_approval_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB max upload


def _safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]+', "_", str(name)).strip() or "하도급업체"


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/parse", methods=["POST"])
def api_parse():
    if "file" not in request.files:
        return jsonify({"error": "엑셀 파일이 업로드되지 않았습니다."}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "파일명이 비어 있습니다."}), 400

    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": ".xlsx 또는 .xlsm 형식의 엑셀 파일만 지원합니다."}), 400

    try:
        file_bytes = io.BytesIO(file.read())
        parsed = parse_project_excel(file_bytes)
        return jsonify(parsed)
    except Exception as e:
        return jsonify({"error": f"엑셀 분석 중 오류가 발생했습니다: {str(e)}"}), 500


@app.route("/api/generate-single", methods=["POST"])
def api_generate_single():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "요청 데이터가 비어 있습니다."}), 400

    try:
        docx_buf = generate_subcontract_approval_docx(data)
        sub_name = _safe_filename(data.get("sub_company", "하도급업체"))
        filename = f"붙임1.하도급_계약_승인신청서_{sub_name}.docx"

        response = send_file(
            docx_buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        encoded = quote(filename)
        response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded}"
        return response
    except Exception as e:
        return jsonify({"error": f"워드 파일 생성 중 오류가 발생했습니다: {str(e)}"}), 500


@app.route("/api/generate-all", methods=["POST"])
def api_generate_all():
    data = request.get_json(silent=True)
    if not data or not data.get("items"):
        return jsonify({"error": "생성할 하도급 업체 목록이 비어 있습니다."}), 400

    common = data.get("common", {})
    items = data.get("items", [])

    try:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, item in enumerate(items, start=1):
                merged = {**common, **item}
                docx_buf = generate_subcontract_approval_docx(merged)
                sub_name = _safe_filename(merged.get("sub_company", f"업체{idx}"))
                fname = f"{idx:02d}_붙임1.하도급_계약_승인신청서_{sub_name}.docx"
                zf.writestr(fname, docx_buf.getvalue())

        zip_buf.seek(0)
        zip_name = "붙임1_하도급계약승인신청서_일괄생성.zip"
        response = send_file(
            zip_buf,
            as_attachment=True,
            download_name=zip_name,
            mimetype="application/zip",
        )
        encoded = quote(zip_name)
        response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded}"
        return response
    except Exception as e:
        return jsonify({"error": f"일괄 생성 중 오류가 발생했습니다: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
