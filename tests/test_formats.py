from reconcile.formats import BANK, ORDERS, Format, detect, validate_mapping

BANK_A_HEADER = ["交易日期", "摘要", "支出金額", "存入金額", "餘額", "備註", "對方帳號"]
BANK_B_HEADER = ["帳務日", "交易說明", "提款", "存款", "結餘", "附言", "轉出帳號"]
ORDERS_HEADER = ["訂單編號", "下單日期", "訂單金額", "收件人", "匯款帳號末五碼"]


def test_detect_builtin_formats():
    assert detect(BANK_A_HEADER, BANK, []).name == "示範銀行 A"
    assert detect(BANK_B_HEADER, BANK, []).name == "示範銀行 B"
    assert detect(ORDERS_HEADER, ORDERS, []).name == "示範電商訂單"


def test_detect_requires_all_mapped_columns_and_kind():
    assert detect(["交易日期", "存入金額"], BANK, []) is None
    assert detect(BANK_A_HEADER, ORDERS, []) is None


def test_saved_formats_take_priority():
    saved = Format("我的銀行", BANK, {"date": "交易日期", "amount": "存入金額"}, builtin=False)
    assert detect(BANK_A_HEADER, BANK, [saved]) is saved


def test_validate_mapping():
    header = ["日期", "金額", "說明"]
    assert validate_mapping(BANK, {"date": "日期", "amount": "金額"}, header) == []
    assert validate_mapping(BANK, {"date": "日期"}, header) == ["請選擇「存入金額」對應的欄位"]
    assert validate_mapping(BANK, {"date": "日期", "amount": "不存在"}, header) == [
        "找不到欄位「不存在」"
    ]


def test_validate_mapping_rejects_duplicate_columns():
    header = ["日期", "金額", "說明"]
    mapping = {"date": "金額", "amount": "金額", "memo": "說明"}
    assert validate_mapping(BANK, mapping, header) == ["「金額」不能同時對應多個欄位"]
