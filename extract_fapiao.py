"""
增值税电子发票 PDF 批量解析工具（CLI 版）
提取字段：文件名、发票号码、发票代码、开票日期、购买方名称、销售方名称、
          金额、税率、税额、价税合计、发票状态、备注等

核心解析逻辑已抽取到 fapiao_parser.py 共用。
"""

import os
import csv

from fapiao_parser import FIELDNAMES, process_pdf_file


FOLDER = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV = os.path.join(FOLDER, "fapiao.csv")


def process_folder(folder: str) -> list[dict]:
    """遍历文件夹，解析所有 PDF 文件，返回记录列表。"""
    records = []
    pdf_files = sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    )

    for filename in pdf_files:
        path = os.path.join(folder, filename)
        print(f"处理: {filename}")
        rec = process_pdf_file(path, filename)

        if rec["备注"].startswith("图片型PDF"):
            print(f"  [!] 无文字层，跳过: {filename}")
        elif rec["备注"].startswith("未识别"):
            print(f"  ? 未识别格式: {filename}")
        elif rec["备注"].startswith("解析错误"):
            print(f"  [ERR] 错误: {rec['备注']}")
        else:
            print(
                f"  [OK] {rec['发票号码']}  {rec['开票日期']}  "
                f"{rec['购买方名称']}  {rec['销售方名称']}  "
                f"¥{rec['价税合计']}  {rec['发票状态']}"
            )

        records.append(rec)

    return records


def save_csv(records: list[dict], output_path: str):
    """将记录列表保存为 CSV 文件。"""
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)
    print(f"\n已保存 {len(records)} 条记录 → {output_path}")


if __name__ == "__main__":
    records = process_folder(FOLDER)
    save_csv(records, OUTPUT_CSV)
