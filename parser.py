"""
国铁电子客票 PDF 解析模块
供 Web 服务 (app.py) 和 CLI 工具 (extract_tickets.py) 共用
"""

import io
import re
import pdfplumber


# ── 字段定义 ──────────────────────────────────────────────
FIELDNAMES = [
    "文件名",
    "发票号码",
    "开票日期",
    "出发站",
    "到达站",
    "车次",
    "乘车日期",
    "开车时间",
    "车厢号",
    "座位号",
    "席别",
    "票价(元)",
    "乘车人",
    "证件号(脱敏)",
    "电子客票号",
    "备注",
]


def extract_text_from_bytes(pdf_bytes: bytes) -> str:
    """用 pdfplumber 从 PDF 字节流提取全文（合并所有页）。"""
    texts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
    return "\n".join(texts)


def extract_text_from_file(pdf_path: str) -> str:
    """用 pdfplumber 从 PDF 文件路径提取全文（合并所有页）。"""
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
    return "\n".join(texts)


def parse_guotie(text: str, filename: str) -> dict:
    """从国铁电子客票文本中解析各字段，返回字典。"""
    rec = {f: "" for f in FIELDNAMES}
    rec["文件名"] = filename

    # 发票号码
    m = re.search(r"发票号码[：:]\s*(\d+)", text)
    if m:
        rec["发票号码"] = m.group(1)

    # 开票日期
    m = re.search(r"开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", text)
    if m:
        rec["开票日期"] = m.group(1)

    # 出发站 / 车次 / 到达站
    # 车次兼容大小写字母 + 数字组合（如 G6989、c6494）
    m = re.search(
        r"^([^\n]+?)\s+([A-Za-z]\d{3,5})\s+([^\n]+?)$",
        text,
        re.MULTILINE,
    )
    if m:
        rec["出发站"] = m.group(1).replace(" ", "").strip()
        rec["车次"]   = m.group(2).upper().strip()
        rec["到达站"] = m.group(3).replace(" ", "").strip()

    # 乘车日期 + 开车时间 + 车厢 + 座位 + 席别
    # 格式1（动车/高铁）：2025年10月18日 10:47开 08车10F号 二等座
    # 格式2（卧铺）：     2026年03月22日 23:44开 09车028号上铺 新空调 软卧
    m = re.search(
        r"(\d{4}年\d{1,2}月\d{1,2}日)\s+(\d{2}:\d{2})开\s+(\d+)车(\w+?)号(\S+铺)?\s*([\S]*[座铺卧])",
        text,
    )
    if m:
        rec["乘车日期"] = m.group(1)
        rec["开车时间"] = m.group(2)
        rec["车厢号"]   = m.group(3)
        berth_part = m.group(5) or ""           # 上铺/下铺/中铺（卧铺附在座位号后）
        rec["座位号"]   = m.group(4) + berth_part
        rec["席别"]     = m.group(6).strip()

    # 票价
    m = re.search(r"[¥￥]([\d]+\.?\d*)", text)
    if m:
        rec["票价(元)"] = m.group(1)

    # 证件号（脱敏）+ 乘车人
    m = re.search(r"(\d{6}\*{4}\d{4,6})\s+([\u4e00-\u9fa5]{2,6})", text)
    if m:
        rec["证件号(脱敏)"] = m.group(1)
        rec["乘车人"]       = m.group(2)

    # 电子客票号
    m = re.search(r"电子客票号[：:]\s*(\d+)", text)
    if m:
        rec["电子客票号"] = m.group(1)

    # 备注（改签标记 / 学生票）
    notes = []
    if "始发改签" in text:
        notes.append("始发改签")
    if re.search(r"学生", text):
        notes.append("学生票")
    rec["备注"] = "；".join(notes)

    return rec


def process_pdf_bytes(pdf_bytes: bytes, filename: str) -> dict:
    """从 PDF 字节流解析，返回一条记录（供 Web 服务调用）。"""
    try:
        text = extract_text_from_bytes(pdf_bytes)
        if not text.strip():
            rec = {f: "" for f in FIELDNAMES}
            rec["文件名"] = filename
            rec["备注"] = "图片型PDF，需OCR，暂跳过"
            return rec
        if any(k in text for k in ["电子客票", "中国铁路", "发票号码"]):
            return parse_guotie(text, filename)
        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = "未识别的发票格式"
        return rec
    except Exception as e:
        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = f"解析错误: {e}"
        return rec


def process_pdf_file(pdf_path: str, filename: str = None) -> dict:
    """从 PDF 文件路径解析，返回一条记录（供 CLI 调用）。"""
    if filename is None:
        filename = os.path.basename(pdf_path)
    try:
        text = extract_text_from_file(pdf_path)
        if not text.strip():
            rec = {f: "" for f in FIELDNAMES}
            rec["文件名"] = filename
            rec["备注"] = "图片型PDF，需OCR，暂跳过"
            return rec
        if any(k in text for k in ["电子客票", "中国铁路", "发票号码"]):
            return parse_guotie(text, filename)
        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = "未识别的发票格式"
        return rec
    except Exception as e:
        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = f"解析错误: {e}"
        return rec
