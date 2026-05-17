from extract_tickets.export import build_csv_bytes, build_excel_bytes
from extract_tickets.importer import import_table_bytes


RECORD = {
    "文件名": "old.pdf",
    "发票号码": "1234567890",
    "开票日期": "",
    "出发站": "威海南海",
    "到达站": "青岛北",
    "车次": "G6989",
    "乘车日期": "2025年10月18日",
    "开车时间": "10:47",
    "车厢号": "08",
    "座位号": "10F",
    "席别": "二等座",
    "票价(元)": "110.00",
    "乘车人": "张三",
    "证件号(脱敏)": "370101****0812",
    "电子客票号": "1864780086101690009372025",
    "备注": "始发改签",
    "报销状态": "已报销",
}


def test_import_exported_csv_bytes():
    imported = import_table_bytes(build_csv_bytes([RECORD]).getvalue(), "tickets.csv")

    assert imported == [RECORD]


def test_import_exported_excel_bytes():
    imported = import_table_bytes(build_excel_bytes([RECORD]).getvalue(), "tickets.xlsx")

    assert imported == [RECORD]


def test_import_csv_defaults_reimbursement_status():
    csv_text = "文件名,车次,备注\r\nold.pdf,G6989,\r\n"

    imported = import_table_bytes(csv_text.encode("utf-8"), "tickets.csv")

    assert imported[0]["文件名"] == "old.pdf"
    assert imported[0]["车次"] == "G6989"
    assert imported[0]["报销状态"] == "未报销"
