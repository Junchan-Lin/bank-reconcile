"""格式設定：記錄哪一欄是日期、金額…，並依表頭自動辨識。"""

from dataclasses import dataclass, field

BANK = "bank"
ORDERS = "orders"
KIND_LABELS = {BANK: "銀行明細", ORDERS: "訂單"}

# 欄位 → （顯示名稱，是否必要）
FIELDS: dict[str, dict[str, tuple[str, bool]]] = {
    BANK: {
        "date": ("日期", True),
        "amount": ("存入金額", True),
        "memo": ("備註", False),
        "account": ("對方帳號", False),
    },
    ORDERS: {
        "order_no": ("訂單編號", True),
        "date": ("日期", True),
        "amount": ("金額", True),
        "customer": ("客戶", False),
        "last5": ("匯款末五碼", False),
    },
}


@dataclass(frozen=True)
class Format:
    name: str
    kind: str
    # 欄位 → 欄名
    mapping: dict[str, str] = field(hash=False)
    builtin: bool = True
    id: int | None = None


BUILTIN_FORMATS = (
    Format(
        "示範銀行 A",
        BANK,
        {"date": "交易日期", "amount": "存入金額", "memo": "備註", "account": "對方帳號"},
    ),
    Format(
        "示範銀行 B",
        BANK,
        {"date": "帳務日", "amount": "存款", "memo": "附言", "account": "轉出帳號"},
    ),
    Format(
        "示範電商訂單",
        ORDERS,
        {
            "order_no": "訂單編號",
            "date": "下單日期",
            "amount": "訂單金額",
            "customer": "收件人",
            "last5": "匯款帳號末五碼",
        },
    ),
)


def detect(header: list[str], kind: str, saved: list[Format]) -> Format | None:
    """使用者建立的格式優先；表頭包含格式所需的所有欄名即視為符合。"""
    columns = set(header)
    for fmt in [*saved, *BUILTIN_FORMATS]:
        if fmt.kind == kind and set(fmt.mapping.values()) <= columns:
            return fmt
    return None


def validate_mapping(kind: str, mapping: dict[str, str], header: list[str]) -> list[str]:
    errors = []
    for name, (label, required) in FIELDS[kind].items():
        column = mapping.get(name)
        if not column:
            if required:
                errors.append(f"請選擇「{label}」對應的欄位")
        elif column not in header:
            errors.append(f"找不到欄位「{column}」")
    used = [mapping[name] for name in FIELDS[kind] if mapping.get(name)]
    for column in dict.fromkeys(used):
        if used.count(column) > 1:
            errors.append(f"「{column}」不能同時對應多個欄位")
    return errors
