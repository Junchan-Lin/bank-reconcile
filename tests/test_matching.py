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
    orders = [order(2, "A1001", 1, 100)]
    deposits = [dep(3, 2, 100, "A1001"), dep(2, 1, 100, "A1001")]  # 列號 2 的入帳日期較早
    [result] = match(deposits, orders, 7)
    assert result.deposit_index == 1


def test_is_effective():
    assert is_effective(MATCHED, AUTO)
    assert is_effective(MISMATCH, AUTO)
    assert is_effective(CANDIDATE, "confirmed")
    assert not is_effective(CANDIDATE, PENDING)


def test_layer1_no_substring_match():
    assert match([dep(2, 20, 100, "A10")], [order(2, "A1", 1, 100)], 7) == []


def test_layer1_weak_key_needs_equal_amount():
    orders = [order(2, "1001", 1, 1500)]
    assert match([dep(2, 1, 2500, "轉帳1001元")], orders, 7) == []
    [result] = match([dep(2, 1, 1500, "訂單1001")], orders, 7)
    assert (result.kind, result.layer, result.status) == (MATCHED, 1, AUTO)
    assert result.reason == "備註含訂單編號 1001，金額相符"


def test_layer1_phone_number_does_not_build_order_no():
    assert match([dep(2, 20, 100, "0912-345-678")], [order(2, "345", 1, 100)], 7) == []


def test_layer1_spaces_split_tokens():
    assert match([dep(2, 20, 100, "A1 23")], [order(2, "A123", 1, 100)], 7) == []


def test_layer1_duplicate_payment_becomes_unknown():
    orders = [order(2, "A1000", 1, 100), order(3, "A1001", 1, 100), order(4, "C9", 1, 100)]
    deposits = [dep(2, 1, 100, "A1001"), dep(3, 2, 100, "A1001")]
    results = match(deposits, orders, 7)
    assert [(r.deposit_index, r.order_index, r.kind) for r in results] == [(0, 1, MATCHED)]


def test_layer1_multiple_orders_become_candidates():
    orders = [order(2, "A1001", 1, 100), order(3, "A1002", 1, 100)]
    results = match([dep(2, 1, 200, "A1001、A1002")], orders, 7)
    assert [(r.order_index, r.kind, r.layer, r.status) for r in results] == [
        (0, CANDIDATE, 1, PENDING),
        (1, CANDIDATE, 1, PENDING),
    ]
    assert results[0].reason == "備註含多個訂單編號（A1001、A1002），需確認"


def test_layer1_full_width_memo():
    [result] = match([dep(2, 1, 100, "Ａ１００１")], [order(2, "A1001", 1, 100)], 7)
    assert (result.kind, result.layer) == (MATCHED, 1)


def test_layer3_candidate_limit():
    orders = [order(2 + n, f"C{n}", 1, 700) for n in range(6)]
    assert match([dep(2, 2, 700)], orders, 7) == []
    results = match([dep(2, 2, 700)], orders[:5], 7)
    assert [r.order_index for r in results] == [0, 1, 2, 3, 4]


def test_performance_20000_by_20000():
    import random
    import time

    rng = random.Random(0)
    amounts = [100 * (n + 1) for n in range(50)]
    orders = [
        order(n + 2, f"ORD{n}", 1 + n % 28, rng.choice(amounts), f"{n % 5000:05d}")
        for n in range(20000)
    ]
    deposits = []
    for n in range(20000):
        memo = f"ORD{rng.randrange(20000)}" if n % 3 == 0 else ""
        last = f"{rng.randrange(5000):05d}" if n % 3 == 1 else None
        deposits.append(dep(n + 2, 1 + n % 28, rng.choice(amounts), memo, last))
    start = time.perf_counter()
    match(deposits, orders, 7)
    assert time.perf_counter() - start < 10


def test_order_no_with_symbol_matches_memo_with_or_without_symbol():
    orders = [order(2, "#A1001", 1, 100)]
    for memo in ("A1001", "#A1001"):
        [result] = match([dep(2, 1, 100, memo)], orders, 7)
        assert (result.kind, result.layer) == (MATCHED, 1)


def test_order_no_with_separator_matches_only_same_token_sequence():
    orders = [order(2, "SO/1001", 1, 100)]
    for memo in ("SO/1001", "SO 1001"):
        [result] = match([dep(2, 1, 100, memo)], orders, 7)
        assert (result.kind, result.layer) == (MATCHED, 1)
    # 備註「SO1001」是單一詞，和〔SO, 1001〕不同，不算命中
    assert match([dep(2, 20, 100, "SO1001")], orders, 7) == []


def test_letters_only_order_no_is_short_and_needs_confirmation():
    orders = [order(2, "TEST", 1, 100)]
    assert match([dep(2, 20, 999, "TEST PAYMENT")], orders, 7) == []
    [result] = match([dep(2, 1, 100, "TEST PAYMENT")], orders, 7)
    assert (result.kind, result.layer, result.status) == (CANDIDATE, 1, PENDING)
    assert result.reason == "備註含短編號 TEST，金額相符，需確認"


def test_short_numeric_order_no_in_date_is_only_candidate():
    [result] = match([dep(2, 1, 990, "10/12 王小明")], [order(2, "12", 1, 990)], 7)
    assert (result.kind, result.layer) == (CANDIDATE, 1)


def test_numeric_order_no_with_equal_amount_matches():
    [result] = match([dep(2, 1, 100, "1001")], [order(2, "1001", 1, 100)], 7)
    assert (result.kind, result.layer) == (MATCHED, 1)


def test_candidate_for_order_matched_by_other_deposit_is_removed():
    orders = [order(2, "A1001", 1, 100), order(3, "A1002", 1, 100)]
    deposits = [dep(2, 1, 200, "A1001 A1002"), dep(3, 2, 100, "A1001")]
    results = match(deposits, orders, 7)
    assert sorted((r.deposit_index, r.order_index, r.kind) for r in results) == [
        (0, 1, CANDIDATE),
        (1, 0, MATCHED),
    ]


def test_duplicate_order_numbers_are_listed_once_in_reason():
    orders = [order(2, "A1001", 1, 100), order(3, "A1001", 1, 100), order(4, "A1002", 1, 100)]
    results = match([dep(2, 1, 300, "A1001 A1002")], orders, 7)
    assert len(results) == 3
    assert results[0].reason == "備註含多個訂單編號（A1001、A1002），需確認"
