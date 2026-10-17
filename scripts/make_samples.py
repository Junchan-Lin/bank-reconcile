"""產生虛構的範例檔到 samples/。資料為虛構，涵蓋每一種配對結果。

用法：uv run python scripts/make_samples.py
"""

import csv
import datetime as dt
import io
import re
import zipfile
from pathlib import Path

from openpyxl import Workbook

SAMPLES = Path(__file__).resolve().parent.parent / "samples"
BANK_A = "示範銀行A_115年10月.csv"
BANK_B = "示範銀行B_115年10月.xlsx"
ORDERS = "示範電商訂單_2026年10月.xlsx"

# （訂單編號, 日, 金額, 收件人, 匯款帳號末五碼）
ORDER_ROWS = [
    ("A1001", 1, 1200, "王小明", "12345"),
    ("A1002", 1, 860, "陳美玲", None),
    ("A1003", 2, 2400, "林志豪", "54321"),
    ("A1004", 2, 2400, "張雅婷", None),
    ("A1005", 3, 1500, "李建宏", None),
    ("A1006", 3, 3200, "黃淑芬", None),
    ("A1007", 4, 990, "吳家豪", "67890"),
    ("A1008", 4, 990, "劉佳穎", "67890"),
    ("A1009", 5, 4500, "蔡宗翰", None),
    ("A1010", 6, 780, "鄭雅文", None),
    ("A1011", 6, 1350, "謝承恩", None),
    ("A1012", 7, 2050, "許庭瑋", "24680"),
]

# （民國日期, 摘要, 支出, 存入, 備註, 對方帳號）；各列對應的配對結果寫在註解
BANK_A_ROWS = [
    ("115/10/01", "跨行轉入", "", "1,200", "A1001", "822-0001-12345"),  # ① 已配對
    ("115/10/01", "提款", "3,000", "", "", ""),  # 提款，不算入帳
    ("115/10/02", "跨行轉入", "", "860", "訂單a-1002", "700-0002-33333"),  # ① 已配對
    ("115/10/03", "跨行轉入", "", "2,385", "A1003 林", "812-0003-44444"),  # ① 金額不符
    ("115/10/03", "跨行轉入", "", "2,400", "", "013-0004-54321"),  # ③ 候選 A1004
    ("115/10/04", "ATM轉入", "", "1,500", "", ""),  # ③ 候選 A1005
    ("115/10/05", "跨行轉入", "", "990", "", "808-0005-67890"),  # ② 兩筆候選
    ("115/10/05", "跨行轉入", "", "3,200", "", ""),  # ③ 候選 A1006
    ("115/10/06", "跨行轉入", "", "5,000", "退款", "822-0006-99999"),  # 不明入帳
    ("115/10/07", "跨行轉入", "", "1,350", "", "004-0007-11111"),  # ③ 候選 A1011
    ("115/10/08", "跨行轉入", "", "2,050", "", "007-0008-24680"),  # ② 已配對
    (
        "115/10/09",
        "跨行轉入",
        "",
        "1,200",
        "A1001",
        "822-0001-12345",
    ),  # ① 疑似重複付款（A1001 已配對）→ 不明入帳
    ("115/13/40", "跨行轉入", "", "500", "", ""),  # 略過：日期不存在
]


def save_fixed(workbook: Workbook, path: Path) -> None:
    """存檔並固定時間戳（openpyxl 存檔時會覆寫 modified、zip 也帶現在時間），讓重跑結果相同。"""
    stamp = dt.datetime(2026, 10, 1)
    workbook.properties.created = workbook.properties.modified = stamp
    raw = io.BytesIO()
    workbook.save(raw)
    with zipfile.ZipFile(raw) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = re.sub(
                    rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-10-01T00:00:00Z", data
                )
            out.writestr(zipfile.ZipInfo(item.filename, stamp.timetuple()[:6]), data)


def write_bank_a() -> None:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerow(["台灣示範銀行 存款交易明細"])
    writer.writerow(["帳號：012-345-678901　期間：115/10/01 至 115/10/31"])
    writer.writerow([])
    writer.writerow(["交易日期", "摘要", "支出金額", "存入金額", "餘額", "備註", "對方帳號"])
    balance = 50_000
    for date, summary, out, deposit, memo, account in BANK_A_ROWS:
        balance += int((deposit or "0").replace(",", "")) - int((out or "0").replace(",", ""))
        writer.writerow([date, summary, out, deposit, f"{balance:,}", memo, account])
    (SAMPLES / BANK_A).write_bytes(buffer.getvalue().encode("cp950"))


def write_bank_b() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["示範銀行B 交易明細表"])
    sheet.append(["帳務日", "交易說明", "提款", "存款", "結餘", "附言", "轉出帳號"])
    sheet.append(
        [dt.datetime(2026, 10, 1), "網銀轉入", None, 1200, 21200, "A1001", "0001234512345"]
    )
    sheet.append([dt.datetime(2026, 10, 2), "轉帳支出", 500, None, 20700, "", ""])
    sheet.append([dt.datetime(2026, 10, 3), "網銀轉入", None, 2050, 22750, "", "0007000824680"])
    save_fixed(workbook, SAMPLES / BANK_B)


def write_orders() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["訂單編號", "下單日期", "訂單金額", "收件人", "匯款帳號末五碼"])
    for order_no, day, amount, name, last5 in ORDER_ROWS:
        sheet.append([order_no, dt.datetime(2026, 10, day), amount, name, last5])
    save_fixed(workbook, SAMPLES / ORDERS)


def main() -> None:
    SAMPLES.mkdir(exist_ok=True)
    write_bank_a()
    write_bank_b()
    write_orders()
    print(f"已產生範例檔到 {SAMPLES}")


if __name__ == "__main__":
    main()
