"""Flask Web 应用。"""

from __future__ import annotations

import os
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from .config import AppConfig
from .constants import ALLOWED_EXTENSIONS, FIELDNAMES
from .export import build_csv_bytes, build_excel_bytes
from .importer import SUPPORTED_IMPORT_EXTENSIONS, import_table_bytes
from .parser import process_pdf_bytes

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def create_app(config: AppConfig | None = None) -> Flask:
    """创建 Flask 应用实例。"""

    app_config = config or AppConfig.from_env()
    app = Flask(
        __name__,
        template_folder=str(PROJECT_ROOT / "templates"),
        static_folder=str(PROJECT_ROOT / "static"),
    )
    app.config["MAX_CONTENT_LENGTH"] = app_config.max_content_length

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/parse", methods=["POST"])
    def api_parse():
        """批量解析接口（旧版兼容）。"""

        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "未收到文件"}), 400

        records = []
        for file in files:
            if not _is_pdf(file.filename):
                continue
            records.append(process_pdf_bytes(file.read(), file.filename))

        return jsonify({"records": records, "fields": FIELDNAMES})

    @app.route("/api/parse_one", methods=["POST"])
    def api_parse_one():
        """逐文件解析接口（旧版兼容）。"""

        file = request.files.get("file")
        if not file:
            return jsonify({"error": "未收到文件"}), 400
        if not _is_pdf(file.filename):
            return jsonify({"error": "仅支持 PDF 文件"}), 400

        record = process_pdf_bytes(file.read(), file.filename)
        return jsonify({"record": record, "fields": FIELDNAMES})

    @app.route("/api/v1/parse_one", methods=["POST"])
    def api_v1_parse_one():
        """带统一响应包裹的新接口。"""

        file = request.files.get("file")
        if not file:
            return _error("未收到文件", 400)
        if not _is_pdf(file.filename):
            return _error("仅支持 PDF 文件", 400)

        record = process_pdf_bytes(file.read(), file.filename)
        return _ok({"record": record, "fields": FIELDNAMES})

    @app.route("/api/download_csv", methods=["POST"])
    def api_download_csv():
        records = _records_from_request()
        return send_file(
            build_csv_bytes(records),
            mimetype="text/csv; charset=utf-8",
            as_attachment=True,
            download_name="tickets.csv",
        )

    @app.route("/api/download_excel", methods=["POST"])
    def api_download_excel():
        records = _records_from_request()
        try:
            excel_bytes = build_excel_bytes(records)
        except RuntimeError as exc:
            return jsonify({"error": str(exc)}), 500
        return send_file(
            excel_bytes,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="tickets.xlsx",
        )

    @app.route("/api/import_table", methods=["POST"])
    def api_import_table():
        """导入旧 CSV / Excel 表格并返回标准记录。"""

        file = request.files.get("file")
        if not file:
            return jsonify({"error": "未收到文件"}), 400
        if not _is_import_table(file.filename):
            return jsonify({"error": "仅支持 CSV、XLSX 或 XLSM 表格"}), 400

        try:
            records = import_table_bytes(file.read(), file.filename)
        except (RuntimeError, UnicodeDecodeError, ValueError) as exc:
            return jsonify({"error": str(exc)}), 400
        return jsonify({"records": records, "fields": FIELDNAMES})

    @app.route("/api/download", methods=["POST"])
    def api_download():
        return api_download_csv()

    return app


def _is_pdf(filename: str | None) -> bool:
    return bool(filename and os.path.splitext(filename.lower())[1] in ALLOWED_EXTENSIONS)


def _is_import_table(filename: str | None) -> bool:
    return bool(filename and os.path.splitext(filename.lower())[1] in SUPPORTED_IMPORT_EXTENSIONS)


def _records_from_request() -> list[dict]:
    data = request.get_json(silent=True) or {}
    records = data.get("records", [])
    return records if isinstance(records, list) else []


def _ok(data: dict, status_code: int = 200):
    return jsonify({"ok": True, "data": data, "error": None}), status_code


def _error(message: str, status_code: int):
    return jsonify({"ok": False, "data": None, "error": message}), status_code
