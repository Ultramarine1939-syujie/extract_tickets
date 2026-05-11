"""
增值税电子普通发票 PDF 解析模块
供 Web 服务 (web_app/app.py) 和 CLI 工具 (extract_fapiao.py) 共用
"""

import io
import re
import os
import pdfplumber


FIELDNAMES = [
    "文件名",
    "发票号码",
    "发票代码",
    "开票日期",
    "购买方名称",
    "购买方税号",
    "购买方银行账号",
    "销售方名称",
    "销售方税号",
    "销售方地址电话",
    "货物或应税劳务名称",
    "规格型号",
    "计量单位",
    "数量",
    "单价",
    "金额",
    "税率",
    "税额",
    "价税合计",
    "发票状态",
    "备注",
    "校验码",
    "电子发票编号",
    "报销状态",
    "报销日期",
    "部门/项目",
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


def parse_fapiao(text: str, filename: str) -> dict:
    """从增值税电子普通发票文本中解析各字段，返回字典。"""
    rec = {f: "" for f in FIELDNAMES}
    rec["文件名"] = filename

    invoice_num_match = re.search(r"发票号码[：:]\s*(\d+)", text)
    if invoice_num_match:
        rec["发票号码"] = invoice_num_match.group(1)

    invoice_code_match = re.search(r"发票代码[：:]\s*(\d+)", text)
    if invoice_code_match:
        rec["发票代码"] = invoice_code_match.group(1)

    date_match = re.search(r"开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)", text)
    if date_match:
        rec["开票日期"] = date_match.group(1)

    buyer_name_match = re.search(r"(?:购买方|买方)[名称]+[：:]\s*([^\n\r]{2,50})", text)
    if buyer_name_match:
        rec["购买方名称"] = buyer_name_match.group(1).strip()

    buyer_tax_match = re.search(r"购买方[纳税人识别号税号]+[：:]\s*([A-Z0-9]+)", text)
    if buyer_tax_match:
        rec["购买方税号"] = buyer_tax_match.group(1).strip()

    buyer_bank_match = re.search(r"购买方开户行及账号[：:]\s*([^\n\r]{5,50})", text)
    if buyer_bank_match:
        rec["购买方银行账号"] = buyer_bank_match.group(1).strip()

    seller_name_match = re.search(r"(?:销售方|卖方)[名称]+[：:]\s*([^\n\r]{2,50})", text)
    if seller_name_match:
        rec["销售方名称"] = seller_name_match.group(1).strip()

    seller_tax_match = re.search(r"销售方[纳税人识别号税号]+[：:]\s*([A-Z0-9]+)", text)
    if seller_tax_match:
        rec["销售方税号"] = seller_tax_match.group(1).strip()

    seller_info_match = re.search(r"销售方地址电话[：:]\s*([^\n\r]{5,60})", text)
    if seller_info_match:
        rec["销售方地址电话"] = seller_info_match.group(1).strip()

    goods_match = re.search(r"(?:货物或应税劳务|项目|商品)[名称]+[：:]\s*([^\n\r]{2,50})", text)
    if goods_match:
        rec["货物或应税劳务名称"] = goods_match.group(1).strip()

    spec_match = re.search(r"规格型号[：:]\s*([^\n\r]{1,30})", text)
    if spec_match:
        rec["规格型号"] = spec_match.group(1).strip()

    unit_match = re.search(r"计量单位[：:]\s*([^\n\r]{1,10})", text)
    if unit_match:
        rec["计量单位"] = unit_match.group(1).strip()

    qty_match = re.search(r"数量[：:]\s*([\d.]+)", text)
    if qty_match:
        rec["数量"] = qty_match.group(1).strip()

    price_match = re.search(r"单价[：:]\s*([\d.]+)", text)
    if price_match:
        rec["单价"] = price_match.group(1).strip()

    amount_match = re.search(r"(?:金额|不含税金额)[：:]\s*[¥￥]?\s*([\d.]+)", text)
    if amount_match:
        rec["金额"] = amount_match.group(1).strip()

    tax_rate_match = re.search(r"税率[：:]\s*([\d.]+%)", text)
    if tax_rate_match:
        rec["税率"] = tax_rate_match.group(1).strip()

    tax_amount_match = re.search(r"税额[：:]\s*[¥￥]?\s*([\d.]+)", text)
    if tax_amount_match:
        rec["税额"] = tax_amount_match.group(1).strip()

    total_match = re.search(r"价税合计[（(小写）):\s]*[¥￥]?\s*([\d.]+)", text)
    if total_match:
        rec["价税合计"] = total_match.group(1).strip()

    if "作废" in text:
        rec["发票状态"] = "作废"
    elif "红冲" in text:
        rec["发票状态"] = "红冲"
    else:
        rec["发票状态"] = "正常"

    remarks_match = re.search(r"备注[：:]\s*([^\n\r]{1,100})", text)
    if remarks_match:
        rec["备注"] = remarks_match.group(1).strip()

    checksum_match = re.search(r"校验码[：:]\s*(\d{20})", text)
    if checksum_match:
        rec["校验码"] = checksum_match.group(1).strip()

    elec_num_match = re.search(r"电子发票编号[：:]\s*([A-Z0-9]+)", text)
    if elec_num_match:
        rec["电子发票编号"] = elec_num_match.group(1).strip()

    rec["报销状态"] = "未报销"

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

        if any(k in text for k in ["增值税电子发票", "发票号码", "发票代码", "购买方名称", "销售方名称"]):
            return parse_fapiao(text, filename)

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

        if any(k in text for k in ["增值税电子发票", "发票号码", "发票代码", "购买方名称", "销售方名称"]):
            return parse_fapiao(text, filename)

        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = "未识别的发票格式"
        return rec
    except Exception as e:
        rec = {f: "" for f in FIELDNAMES}
        rec["文件名"] = filename
        rec["备注"] = f"解析错误: {e}"
        return rec
