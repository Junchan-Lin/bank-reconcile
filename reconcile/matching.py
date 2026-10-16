"""三層配對：備註含訂單編號 → 末五碼加金額 → 金額加日期。純邏輯，不碰資料庫。"""

import datetime as dt
from dataclasses import dataclass

from reconcile.normalize import DepositRow, OrderRow, normalize_order_no

MATCHED = "matched"
MISMATCH = "mismatch"
CANDIDATE = "candidate"

AUTO = "auto"
PENDING = "pending"
CONFIRMED = "confirmed"
REJECTED = "rejected"


@dataclass(frozen=True)
class LinkResult:
    deposit_index: int
    order_index: int
    kind: str
    layer: int
    reason: str
    status: str


def is_effective(kind: str, status: str) -> bool:
    """有效配對：自動配對（含金額不符），或已確認的候選。"""
    return kind in (MATCHED, MISMATCH) or (kind == CANDIDATE and status == CONFIRMED)


def match(deposits: list[DepositRow], orders: list[OrderRow], window_days: int) -> list[LinkResult]:
    """入帳依（日期、列號）由早到晚處理；一筆訂單最多被確定配對一次。"""
    keys = {index: normalize_order_no(o.order_no) for index, o in enumerate(orders)}
    used: set[int] = set()
    results: list[LinkResult] = []
    sequence = sorted(range(len(deposits)), key=lambda i: (deposits[i].date, deposits[i].row_no))

    remaining = []
    for i in sequence:  # 層級 ①
        deposit = deposits[i]
        memo = normalize_order_no(deposit.memo)
        hits = [j for j, key in keys.items() if key and j not in used and key in memo]
        if not hits:
            remaining.append(i)
            continue
        j = max(hits, key=lambda j: (len(keys[j]), -j))
        used.add(j)
        found = orders[j]
        reason = f"備註含訂單編號 {found.order_no}"
        if deposit.amount == found.amount:
            results.append(LinkResult(i, j, MATCHED, 1, reason, AUTO))
        else:
            diff = deposit.amount - found.amount
            reason += f"，入帳 {deposit.amount:,} 元、訂單 {found.amount:,} 元，差額 {diff:+,} 元"
            results.append(LinkResult(i, j, MISMATCH, 1, reason, AUTO))

    after_layer2 = []
    for i in remaining:  # 層級 ②
        deposit = deposits[i]
        hits = [
            j
            for j, o in enumerate(orders)
            if deposit.last5
            and j not in used
            and o.last5 == deposit.last5
            and o.amount == deposit.amount
        ]
        if len(hits) == 1:
            used.add(hits[0])
            reason = f"帳號末五碼 {deposit.last5} 與金額相符"
            results.append(LinkResult(i, hits[0], MATCHED, 2, reason, AUTO))
        elif hits:
            reason = f"帳號末五碼 {deposit.last5} 與金額相符（共 {len(hits)} 筆，需確認）"
            results.extend(LinkResult(i, j, CANDIDATE, 2, reason, PENDING) for j in hits)
        else:
            after_layer2.append(i)

    window = dt.timedelta(days=window_days)
    for i in after_layer2:  # 層級 ③
        deposit = deposits[i]
        for j, o in enumerate(orders):
            if j in used or o.amount != deposit.amount:
                continue
            if o.date <= deposit.date <= o.date + window:
                days = (deposit.date - o.date).days
                reason = f"金額相符，入帳日在訂單日後 {days} 天"
                results.append(LinkResult(i, j, CANDIDATE, 3, reason, PENDING))
    return results
