import datetime as dt

from reconcile.matching import (
    AUTO,
    CANDIDATE,
    MATCHED,
    MISMATCH,
    PENDING,
    LinkResult,
    is_effective,
    match,
)
from reconcile.normalize import DepositRow, OrderRow

D = dt.date


def dep(row, day, amount, memo="", last5=None):
    return DepositRow(row, D(2026, 10, day), amount, memo, last5)


def order(row, no, day, amount, last5=None):
    return OrderRow(row, no, D(2026, 10, day), amount, "", last5)


def test_layer1_order_number_in_memo():
    results = match([dep(2, 1, 1200, "訂單 a-1001")], [order(2, "A1001", 1, 1200)], 7)
    assert results == [LinkResult(0, 0, MATCHED, 1, "備註含訂單編號 A1001", AUTO)]


def test_layer1_amount_mismatch_reports_difference():
    [result] = match([dep(2, 1, 2385, "A1003")], [order(2, "A1003", 1, 2400)], 7)
    assert (result.kind, result.layer, result.status) == (MISMATCH, 1, AUTO)
    assert "差額 -15 元" in result.reason


def test_layer1_prefers_longest_order_number():
    orders = [order(2, "A10", 1, 100), order(3, "A101", 1, 100)]
    [result] = match([dep(2, 1, 100, "A101")], orders, 7)
    assert result.order_index == 1


def test_layer2_last5_and_amount():
    [result] = match([dep(2, 3, 2050, last5="24680")], [order(2, "B1", 1, 2050, "24680")], 7)
    assert (result.kind, result.layer) == (MATCHED, 2)


def test_layer2_multiple_hits_become_candidates_and_skip_layer3():
    orders = [
        order(2, "B1", 1, 990, "67890"),
        order(3, "B2", 1, 990, "67890"),
        order(4, "B3", 1, 990),
    ]
    results = match([dep(2, 2, 990, last5="67890")], orders, 7)
    assert [(r.order_index, r.kind, r.layer, r.status) for r in results] == [
        (0, CANDIDATE, 2, PENDING),
        (1, CANDIDATE, 2, PENDING),
    ]


def test_layer3_amount_within_window():
    orders = [order(2, "C1", 1, 1500), order(3, "C2", 9, 1500), order(4, "C3", 1, 1600)]
    results = match([dep(2, 8, 1500)], orders, 7)
    assert [(r.order_index, r.layer) for r in results] == [(0, 3)]  # 10/1 + 7 天 = 10/8
    assert match([dep(2, 9, 1500)], orders[:1], 7) == []  # 超過容許天數


def test_used_orders_are_not_reused_and_order_is_deterministic():
    orders = [order(2, "A1", 1, 100)]
    deposits = [dep(3, 2, 100, "A1"), dep(2, 1, 100, "A1")]  # 列號 2 的入帳日期較早
    [result] = match(deposits, orders, 7)
    assert result.deposit_index == 1


def test_is_effective():
    assert is_effective(MATCHED, AUTO)
    assert is_effective(MISMATCH, AUTO)
    assert is_effective(CANDIDATE, "confirmed")
    assert not is_effective(CANDIDATE, PENDING)
