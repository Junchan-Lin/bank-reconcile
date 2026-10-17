import datetime as dt
import io

import pytest
from openpyxl import Workbook

from reconcile import readers
from reconcile.readers import ReadError, read_table
from tests.conftest import csv_bytes, xlsx_bytes


def test_utf8_csv_with_bom():
    data = '﻿訂單編號,日期,金額\nA1,2026/10/01,"1,200"\n'.encode()
    table = read_table("orders.csv", data)
    assert table.header == ["訂單編號", "日期", "金額"]
    assert table.rows == [(2, ["A1", "2026/10/01", "1,200"])]


def test_big5_csv_with_title_rows():
    data = csv_bytes(
        [
            "台灣示範銀行 存款交易明細",
            "帳號：012-345-678901",
            "",
            "交易日期,摘要,存入金額",
            "115/10/01,跨行轉入,1200",
            ",,",
            "115/10/02,跨行轉入,860",
        ],
        encoding="cp950",
    )
    table = read_table("bank.csv", data)
    assert table.header == ["交易日期", "摘要", "存入金額"]
    assert table.rows == [
        (5, ["115/10/01", "跨行轉入", "1200"]),
        (7, ["115/10/02", "跨行轉入", "860"]),
    ]


def test_xlsx_keeps_cell_types():
    data = xlsx_bytes(
        [
            ["示範銀行B 交易明細表"],
            ["帳務日", "說明", "存款"],
            [dt.datetime(2026, 10, 1), " 轉入 ", 1200],
        ]
    )
    table = read_table("bank.xlsx", data)
    assert table.header == ["帳務日", "說明", "存款"]
    row_no, row = table.rows[0]
    assert row_no == 3
    assert row == [dt.datetime(2026, 10, 1), "轉入", 1200]


def test_column_lookup_tolerates_short_rows():
    table = readers.Table(["a", "b", "c"], [(2, ["1"])])
    assert table.column(table.rows[0][1], "a") == "1"
    assert table.column(table.rows[0][1], "c") is None


@pytest.mark.parametrize(
    ("filename", "data", "message"),
    [
        ("x.txt", b"abc", "只支援"),
        ("x.xlsx", b"not a zip", "無法讀取 Excel"),
        ("x.csv", b"\xff\xfe\xfa\xfb\xfc" * 3, "編碼"),
        ("x.csv", csv_bytes(["只有一欄", "1"]), "找不到表頭列"),
    ],
)
def test_unreadable_files(filename, data, message):
    with pytest.raises(ReadError, match=message):
        read_table(filename, data)


def test_size_limits(monkeypatch):
    monkeypatch.setattr(readers, "MAX_ROWS", 2)
    with pytest.raises(ReadError, match="2 萬列"):
        read_table("x.csv", csv_bytes(["a,b,c", "1,2,3", "1,2,3", "1,2,3"]))
    monkeypatch.setattr(readers, "MAX_BYTES", 10)
    with pytest.raises(ReadError, match="5 MB"):
        read_table("x.csv", csv_bytes(["a,b,c", "1,2,3"]))


def test_xlsx_wide_dimension_is_capped():
    workbook = Workbook()
    sheet = workbook.active
    sheet.cell(row=1, column=16384, value="x")
    sheet.append(["帳務日", "說明", "存款"])
    sheet.append(["2026/10/01", "轉入", 1200])
    buffer = io.BytesIO()
    workbook.save(buffer)
    table = read_table("bank.xlsx", buffer.getvalue())
    assert table.header == ["帳務日", "說明", "存款"]
    assert table.rows == [(3, ["2026/10/01", "轉入", 1200])]
    assert all(len(row) <= readers.MAX_COLUMNS for _, row in table.rows)


def test_csv_too_many_rows_aborts():
    lines = ["a,b,c"] + ["1,2,3"] * (readers.MAX_ROWS + readers.HEADER_SCAN_ROWS + 5)
    with pytest.raises(ReadError, match="2 萬列"):
        read_table("big.csv", csv_bytes(lines))


def test_xlsx_row_limit_stops_early(monkeypatch):
    monkeypatch.setattr(readers, "MAX_ROWS", 5)
    rows = [["a", "b", "c"]] + [[1, 2, 3]] * (readers.HEADER_SCAN_ROWS + 10)
    with pytest.raises(ReadError, match="2 萬列"):
        read_table("big.xlsx", xlsx_bytes(rows))
