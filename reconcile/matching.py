"""三層配對：備註含訂單編號 → 末五碼加金額 → 金額加日期。純邏輯，不碰資料庫。"""

import bisect
import datetime as dt
import re
import unicodedata
from dataclasses import dataclass

from reconcile.normalize import DepositRow, OrderRow

MATCHED = "matched"
MISMATCH = "mismatch"
CANDIDATE = "candidate"

AUTO = "auto"
PENDING = "pending"
CONFIRMED = "confirmed"
REJECTED = "rejected"

MAX_CANDIDATES = 5
STRONG_MIN_LENGTH = 4
MAX_KEY_TOKENS = 4
TOKEN_SPLIT = re.compile(r"[^0-9A-Z\-]+")


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


STRONG, NUMERIC, SHORT = "strong", "numeric", "short"


def _strength(joined: str) -> str:
    if len(joined) < STRONG_MIN_LENGTH:
        return SHORT
    if joined.isdigit():
        return NUMERIC
    if any(c.isdigit() for c in joined) and any(c.isalpha() for c in joined):
        return STRONG
    return SHORT


def _tokens(text: str) -> tuple[str, ...]:
    """訂單編號與備註共用的切詞：NFKC、大寫、以非英數與連字號切開、去連字號。"""
    upper = unicodedata.normalize("NFKC", text).upper()
    return tuple(t for t in (p.replace("-", "") for p in TOKEN_SPLIT.split(upper)) if t)


def match(deposits: list[DepositRow], orders: list[OrderRow], window_days: int) -> list[LinkResult]:
    """入帳依（日期、列號）由早到晚處理；一筆訂單最多被確定配對一次。"""
    by_tokens: dict[tuple[str, ...], list[int]] = {}
    max_len = 1
    for j, o in enumerate(orders):
        key = _tokens(o.order_no)
        if key and len(key) <= MAX_KEY_TOKENS:
            by_tokens.setdefault(key, []).append(j)
            max_len = max(max_len, len(key))
    strengths = {
        j: _strength("".join(_tokens(orders[j].order_no))) for js in by_tokens.values() for j in js
    }
    by_last5: dict[tuple[str, int], list[int]] = {}
    by_amount: dict[int, list[tuple[dt.date, int]]] = {}
    for j, o in enumerate(orders):
        if o.last5:
            by_last5.setdefault((o.last5, o.amount), []).append(j)
        by_amount.setdefault(o.amount, []).append((o.date, j))
    for entries in by_amount.values():
        entries.sort()

    used: set[int] = set()
    results: list[LinkResult] = []
    sequence = sorted(range(len(deposits)), key=lambda i: (deposits[i].date, deposits[i].row_no))

    remaining = []
    for i in sequence:  # 層級 ①
        deposit = deposits[i]
        memo = _tokens(deposit.memo)
        hit_orders: set[int] = set()
        for size in range(1, min(max_len, len(memo)) + 1):
            for start in range(len(memo) - size + 1):
                hit_orders.update(by_tokens.get(memo[start : start + size], ()))
        usable = sorted(
            j
            for j in hit_orders
            if j not in used and (strengths[j] == STRONG or orders[j].amount == deposit.amount)
        )
        used_strong = any(j in used and strengths[j] == STRONG for j in hit_orders)
        if len(usable) > 1:
            names = "、".join(dict.fromkeys(orders[j].order_no for j in usable))
            reason = f"備註含多個訂單編號（{names}），需確認"
            results.extend(LinkResult(i, j, CANDIDATE, 1, reason, PENDING) for j in usable)
        elif usable:
            [j] = usable
            found = orders[j]
            if strengths[j] == SHORT:
                reason = f"備註含短編號 {found.order_no}，金額相符，需確認"
                results.append(LinkResult(i, j, CANDIDATE, 1, reason, PENDING))
                continue
            used.add(j)
            reason = f"備註含訂單編號 {found.order_no}"
            if deposit.amount == found.amount:
                if strengths[j] == NUMERIC:
                    reason += "，金額相符"
                results.append(LinkResult(i, j, MATCHED, 1, reason, AUTO))
            else:
                diff = deposit.amount - found.amount
                reason += (
                    f"，入帳 {deposit.amount:,} 元、訂單 {found.amount:,} 元，差額 {diff:+,} 元"
                )
                results.append(LinkResult(i, j, MISMATCH, 1, reason, AUTO))
        elif not used_strong:
            remaining.append(i)

    after_layer2 = []
    for i in remaining:  # 層級 ②
        deposit = deposits[i]
        hits = []
        if deposit.last5:
            hits = [j for j in by_last5.get((deposit.last5, deposit.amount), []) if j not in used]
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
        entries = by_amount.get(deposit.amount, [])
        start = bisect.bisect_left(entries, (deposit.date - window, -1))
        stop = bisect.bisect_right(entries, (deposit.date, len(orders)))
        hits = sorted((j, d) for d, j in entries[start:stop] if j not in used)
        if len(hits) > MAX_CANDIDATES:
            continue
        for j, order_date in hits:
            days = (deposit.date - order_date).days
            reason = f"金額相符，入帳日在訂單日後 {days} 天"
            results.append(LinkResult(i, j, CANDIDATE, 3, reason, PENDING))
    return [r for r in results if not (r.kind == CANDIDATE and r.order_index in used)]
