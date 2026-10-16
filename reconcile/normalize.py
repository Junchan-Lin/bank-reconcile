"""把各種格式轉成統一的入帳與訂單資料。"""

import datetime as dt
import re
from dataclasses import dataclass

from reconcile.formats import Format
from reconcile.readers import Table

AMOUNT_NOISE = re.compile(r"NT\$|\$|元|,|\s", re.IGNORECASE)
AMOUNT_TEXT = re.compile(r"-?\d+(?:\.0+)?")
DATE_COMPACT = re.compile(r"\d{8}")
DATE_SEPARATED = re.compile(r"(\d{2,4})[/\-.](\d{1,2})[/\-.](\d{1,2})(?:\s.*)?")
ROC_OFFSET = 1911


@dataclass(frozen=True)
class DepositRow:
    row_no: int
    date: dt.date
    amount: int
    memo: str
    last5: str | None


@dataclass(frozen=True)
class OrderRow:
    row_no: int
    order_no: str
    date: dt.date
    amount: int
    customer: str
    last5: str | None


@dataclass(frozen=True)
class Skipped:
    row_no: int
    reason: str


def parse_amount(value) -> int | None:
    """回傳整數金額；空白回傳 None；無法解析時拋出 ValueError。"""
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"無法解析金額「{value}」")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not value.is_integer():
            raise ValueError(f"金額不是整數「{value}」")
        return int(value)
    text = AMOUNT_NOISE.sub("", str(value))
    if not text:
        return None
    if not AMOUNT_TEXT.fullmatch(text):
        raise ValueError(f"無法解析金額「{value}」")
    return int(text.split(".")[0])


def parse_date(value) -> dt.date:
    """支援西元與民國年；Excel 的日期儲存格直接使用。"""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = "" if value is None else str(value).strip()
    if DATE_COMPACT.fullmatch(text):
        year, month, day = int(text[:4]), int(text[4:6]), int(text[6:])
    else:
        match = DATE_SEPARATED.fullmatch(text)
        if not match:
            raise ValueError(f"無法解析日期「{text}」")
        year, month, day = (int(group) for group in match.groups())
        if year < ROC_OFFSET:  # 民國年
            year += ROC_OFFSET
    try:
        return dt.date(year, month, day)
    except ValueError as e:
        raise ValueError(f"日期不存在「{text}」") from e


def last5(value) -> str | None:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    digits = re.sub(r"\D", "", "" if value is None else str(value))
    return digits[-5:] if len(digits) >= 5 else None


def normalize_order_no(value) -> str:
    return re.sub(r"[\s\-]", "", "" if value is None else str(value)).upper()


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _cell(table: Table, row: list, fmt: Format, name: str):
    column = fmt.mapping.get(name)
    return table.column(row, column) if column else None


def build_deposits(table: Table, fmt: Format) -> tuple[list[DepositRow], list[Skipped]]:
    deposits, skipped = [], []
    for row_no, row in table.rows:
        try:
            amount = parse_amount(_cell(table, row, fmt, "amount"))
            if amount is None or amount <= 0:
                continue  # 沒有存入金額，例如提款列
            deposits.append(
                DepositRow(
                    row_no,
                    parse_date(_cell(table, row, fmt, "date")),
                    amount,
                    _text(_cell(table, row, fmt, "memo")),
                    last5(_cell(table, row, fmt, "account")),
                )
            )
        except ValueError as e:
            skipped.append(Skipped(row_no, str(e)))
    return deposits, skipped


def build_orders(table: Table, fmt: Format) -> tuple[list[OrderRow], list[Skipped]]:
    orders, skipped, seen = [], [], set()
    for row_no, row in table.rows:
        order_no = _text(_cell(table, row, fmt, "order_no"))
        try:
            if not order_no:
                raise ValueError("缺少訂單編號")
            amount = parse_amount(_cell(table, row, fmt, "amount"))
            if amount is None:
                raise ValueError("缺少金額")
            if amount <= 0:
                raise ValueError("金額必須大於 0")
            key = normalize_order_no(order_no)
            if key in seen:
                raise ValueError(f"訂單編號 {order_no} 重複")
            order = OrderRow(
                row_no,
                order_no,
                parse_date(_cell(table, row, fmt, "date")),
                amount,
                _text(_cell(table, row, fmt, "customer")),
                last5(_cell(table, row, fmt, "last5")),
            )
        except ValueError as e:
            skipped.append(Skipped(row_no, str(e)))
            continue
        seen.add(key)
        orders.append(order)
    return orders, skipped
