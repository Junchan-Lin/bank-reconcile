import datetime as dt

import pytest

from reconcile.formats import BANK, ORDERS, Format
from reconcile.normalize import (
    DepositRow,
    OrderRow,
    Skipped,
    build_deposits,
    build_orders,
    last5,
    normalize_order_no,
    parse_amount,
    parse_date,
)
from reconcile.readers import Table


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("1,200", 1200),
        ("NT$1,200", 1200),
        ("$ 860 元", 860),
        (1200, 1200),
        (1200.0, 1200),
        ("1200.00", 1200),
        ("", None),
        (None, None),
        ("-300", -300),
    ],
)
def test_parse_amount(value, expected):
    assert parse_amount(value) == expected


@pytest.mark.parametrize("value", ["abc", 12.5, "1.5", True])
def test_parse_amount_rejects(value):
    with pytest.raises(ValueError):
        parse_amount(value)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("2026-10-05", dt.date(2026, 10, 5)),
        ("2026/10/5", dt.date(2026, 10, 5)),
        ("20261005", dt.date(2026, 10, 5)),
        ("115/10/05", dt.date(2026, 10, 5)),
        ("115.10.05", dt.date(2026, 10, 5)),
        ("2026/10/05 14:30", dt.date(2026, 10, 5)),
        (dt.datetime(2026, 10, 5, 9), dt.date(2026, 10, 5)),
        (dt.date(2026, 10, 5), dt.date(2026, 10, 5)),
    ],
)
def test_parse_date(value, expected):
    assert parse_date(value) == expected


@pytest.mark.parametrize("value", ["", None, "10/05", "115/13/40", "hello"])
def test_parse_date_rejects(value):
    with pytest.raises(ValueError):
        parse_date(value)


def test_last5_and_order_no():
    assert last5("822-0001-12345") == "12345"
    assert last5(12345.0) == "12345"
    assert last5("123") is None
    assert last5(None) is None
    assert normalize_order_no(" a-1002 ") == "A1002"


BANK_FMT = Format("A", BANK, {"date": "日期", "amount": "存入", "memo": "備註", "account": "帳號"})
ORDER_FMT = Format("O", ORDERS, {"order_no": "編號", "date": "日期", "amount": "金額"})


def test_build_deposits_skips_withdrawals_and_records_bad_rows():
    table = Table(
        ["日期", "存入", "備註", "帳號"],
        [
            (2, ["115/10/01", "1,200", "A1001", "822-12345"]),
            (3, ["115/10/01", "", "", ""]),  # 提款列：沒有存入
            (4, ["115/13/40", "500", "", ""]),  # 日期不存在
            (5, ["115/10/02", "abc", "", ""]),
        ],  # 金額無法解析
    )
    deposits, skipped = build_deposits(table, BANK_FMT)
    assert deposits == [DepositRow(2, dt.date(2026, 10, 1), 1200, "A1001", "12345")]
    assert [s.row_no for s in skipped] == [4, 5]
    assert "日期" in skipped[0].reason


def test_build_orders_validates_and_dedupes():
    table = Table(
        ["編號", "日期", "金額"],
        [
            (2, ["A1", "2026-10-01", 1200]),
            (3, ["", "2026-10-01", 100]),
            (4, ["A2", "2026-10-01", 0]),
            (5, ["a-1", "2026-10-02", 300]),
        ],
    )
    orders, skipped = build_orders(table, ORDER_FMT)
    assert orders == [OrderRow(2, "A1", dt.date(2026, 10, 1), 1200, "", None)]
    assert skipped == [
        Skipped(3, "缺少訂單編號"),
        Skipped(4, "金額必須大於 0"),
        Skipped(5, "訂單編號 a-1 重複"),
    ]
