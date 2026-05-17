from io import BytesIO

from extract_tickets.web import create_app


def test_parse_one_rejects_non_pdf():
    app = create_app()
    client = app.test_client()

    response = client.post(
        "/api/parse_one",
        data={"file": (BytesIO(b"text"), "ticket.txt")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 400
    assert response.get_json()["error"] == "仅支持 PDF 文件"


def test_v1_parse_one_wraps_response(monkeypatch):
    monkeypatch.setattr(
        "extract_tickets.web.process_pdf_bytes",
        lambda _, filename: {"文件名": filename, "备注": ""},
    )
    app = create_app()
    client = app.test_client()

    response = client.post(
        "/api/v1/parse_one",
        data={"file": (BytesIO(b"%PDF"), "ticket.pdf")},
        content_type="multipart/form-data",
    )

    body = response.get_json()
    assert response.status_code == 200
    assert body["ok"] is True
    assert body["data"]["record"]["文件名"] == "ticket.pdf"


def test_import_table_reads_csv():
    app = create_app()
    client = app.test_client()

    response = client.post(
        "/api/import_table",
        data={"file": (BytesIO("文件名,车次\r\nold.pdf,G6989\r\n".encode()), "tickets.csv")},
        content_type="multipart/form-data",
    )

    body = response.get_json()
    assert response.status_code == 200
    assert body["records"][0]["文件名"] == "old.pdf"
    assert body["records"][0]["车次"] == "G6989"


def test_import_table_rejects_unknown_extension():
    app = create_app()
    client = app.test_client()

    response = client.post(
        "/api/import_table",
        data={"file": (BytesIO(b"data"), "tickets.txt")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 400
    assert response.get_json()["error"] == "仅支持 CSV、XLSX 或 XLSM 表格"
