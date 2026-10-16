# 銀行入帳對帳 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 依 [design.md](design.md) 建立「銀行入帳對帳」第一版：上傳銀行明細與訂單檔，自動辨識格式、三層配對、在網頁確認待確認項目、下載 Excel 報表。

**Architecture:** FastAPI + Jinja2 伺服器端渲染；`readers`（讀檔、找表頭）→ `formats`（欄位對應、辨識）→ `normalize`（統一資料）→ `matching`（純邏輯配對）→ `service`（存成對帳工作、確認／排除、整理結果）→ `report`（openpyxl 報表）。SQLite 保存，到期自動刪除。

**Tech Stack:** Python 3.12、uv、FastAPI、Jinja2、SQLAlchemy 2 + SQLite、openpyxl；測試用 pytest、httpx（TestClient）；ruff。

**慣例（沿用 weather-alert）：**
- 工作目錄：`~/sky/codesky-works/bank-reconcile/`，git 身分已在 repo 層級設定為 `Code Sky <junchan.lin@gmail.com>`，分支 `feat/v1`。
- 每個 commit 訊息結尾加上一行 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`（以下範例省略）。
- 每個 Task 結束前執行 `uv run ruff format . && uv run ruff check . && uv run pytest`。
- 時間：資料庫以 naive UTC 存 `created_at`，顯示時轉 Asia/Taipei；業務日期（入帳日、訂單日）用 `datetime.date`。

---

## 檔案結構

```
bank-reconcile/
├── pyproject.toml / uv.lock / .python-version / .gitignore / .dockerignore
├── Dockerfile / docker-compose.yml / .env.example / README.md / CLAUDE.md
├── docs/design.md / docs/plan.md
├── scripts/make_samples.py         # 產生 samples/ 的虛構範例檔
├── samples/                        # 示範銀行A（CSV, Big5）、示範銀行B（xlsx）、示範電商訂單（xlsx）
├── reconcile/
│   ├── config.py                   # Settings
│   ├── models.py                   # 資料表、init_db
│   ├── readers.py                  # read_table：CSV/xlsx、編碼、表頭列
│   ├── formats.py                  # Format、內建格式、detect、validate_mapping
│   ├── normalize.py                # parse_amount/parse_date/last5、build_deposits/build_orders
│   ├── matching.py                 # match：三層配對（純邏輯）
│   ├── service.py                  # create_job、confirm/reject、job_view、purge_expired
│   ├── report.py                   # build_report → xlsx bytes
│   └── web/                        # app.py、uploads.py、templates/、static/
└── tests/
```

---

### Task 1: 專案骨架

**Files:**
- Create: `pyproject.toml`, `.python-version`, `.gitignore`, `.dockerignore`, `reconcile/__init__.py`, `tests/__init__.py`

- [ ] **Step 1: 建立分支與 `pyproject.toml`**

```bash
git switch -c feat/v1
```

```toml
[project]
name = "bank-reconcile"
version = "0.1.0"
description = "銀行入帳對帳"
requires-python = ">=3.12"
dependencies = [
    "fastapi>=0.115",
    "jinja2>=3.1",
    "openpyxl>=3.1",
    "python-multipart>=0.0.9",
    "sqlalchemy>=2.0",
    "tzdata>=2024.1",
    "uvicorn>=0.30",
]

[dependency-groups]
dev = [
    "httpx>=0.27",
    "pytest>=8",
    "ruff>=0.6",
]

[tool.uv]
package = false

[tool.pytest.ini_options]
testpaths = ["tests"]

[tool.ruff]
line-length = 100
target-version = "py312"
extend-exclude = ["docs"]

[tool.ruff.lint]
select = ["E", "F", "I", "UP", "B", "SIM"]
```

- [ ] **Step 2: 其他骨架檔**

`.python-version`：

```
3.12
```

`.gitignore`：

```
.venv/
__pycache__/
*.pyc
.pytest_cache/
.ruff_cache/
.env
data/
```

`.dockerignore`：

```
.venv
.git
.env
data
tests
docs
scripts
**/__pycache__
.pytest_cache
.ruff_cache
```

`reconcile/__init__.py`、`tests/__init__.py`：空檔。

- [ ] **Step 3: 安裝並確認**

Run: `uv sync && uv run pytest`
Expected: 安裝成功；`no tests ran`（exit 5）。

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml uv.lock .python-version .gitignore .dockerignore reconcile/__init__.py tests/__init__.py
git commit -m "建立專案骨架"
```

---

### Task 2: 設定與資料模型

**Files:**
- Create: `reconcile/config.py`, `reconcile/models.py`, `tests/conftest.py`
- Test: `tests/test_models.py`

- [ ] **Step 1: `tests/conftest.py`**

```python
import io

import pytest
from openpyxl import Workbook

from reconcile.models import init_db


@pytest.fixture
def session_factory(tmp_path):
    return init_db(f"sqlite:///{tmp_path / 'test.db'}")


def csv_bytes(lines: list[str], encoding: str = "utf-8") -> bytes:
    return ("\n".join(lines) + "\n").encode(encoding)


def xlsx_bytes(rows: list[list]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
```

- [ ] **Step 2: 寫失敗的測試 `tests/test_models.py`**

```python
import datetime as dt

from sqlalchemy import select

from reconcile.config import Settings
from reconcile.models import Deposit, Job, Link, Order, SkippedRow


def test_settings_from_env(monkeypatch):
    monkeypatch.setenv("RETENTION_DAYS", "30")
    monkeypatch.setenv("UPLOAD_DIR", "/tmp/up")
    settings = Settings.from_env()
    assert settings.retention_days == 30
    assert str(settings.upload_dir) == "/tmp/up"
    assert settings.samples_dir.name == "samples"


def test_job_cascade_delete(session_factory):
    with session_factory() as session:
        job = Job(bank_filename="b.csv", orders_filename="o.xlsx", bank_format="A",
                  orders_format="O", window_days=7)
        deposit = Deposit(row_no=5, date=dt.date(2026, 10, 1), amount=1200, memo="", last5=None)
        order = Order(row_no=2, order_no="A1", date=dt.date(2026, 10, 1), amount=1200,
                      customer="", last5=None)
        job.deposits.append(deposit)
        job.orders.append(order)
        job.links.append(Link(deposit=deposit, order=order, kind="matched", layer=1,
                              reason="x", status="auto"))
        job.skipped.append(SkippedRow(source="bank", row_no=9, reason="y"))
        session.add(job)
        session.commit()
        assert job.created_at is not None
        session.delete(job)
        session.commit()
        for model in (Deposit, Order, Link, SkippedRow):
            assert session.scalars(select(model)).all() == []
```

- [ ] **Step 3: 執行確認失敗**

Run: `uv run pytest tests/test_models.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.models'`

- [ ] **Step 4: `reconcile/config.py`**

```python
import os
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class Settings:
    database_url: str = "sqlite:///data/reconcile.db"
    # 對帳工作保留天數；0 表示不自動刪除
    retention_days: int = 90
    upload_dir: Path = Path("data/uploads")
    samples_dir: Path = ROOT / "samples"

    @classmethod
    def from_env(cls) -> "Settings":
        env = os.environ.get
        return cls(
            database_url=env("DATABASE_URL") or cls.database_url,
            retention_days=int(env("RETENTION_DAYS") or 90),
            upload_dir=Path(env("UPLOAD_DIR") or cls.upload_dir),
        )
```

- [ ] **Step 5: `reconcile/models.py`**

```python
import datetime as dt
from pathlib import Path

from sqlalchemy import JSON, ForeignKey, String, Text, create_engine
from sqlalchemy.engine import make_url
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, sessionmaker


def utcnow() -> dt.datetime:
    return dt.datetime.now(dt.UTC).replace(tzinfo=None)


class Base(DeclarativeBase):
    pass


class FormatProfile(Base):
    """使用者建立的欄位對應。內建格式寫在 formats.py，不存資料庫。"""

    __tablename__ = "format_profiles"

    id: Mapped[int] = mapped_column(primary_key=True)
    # bank / orders
    kind: Mapped[str] = mapped_column(String(10))
    name: Mapped[str] = mapped_column(String(100))
    # 欄位 → 欄名，例如 {"date": "交易日期", "amount": "存入金額"}
    mapping: Mapped[dict] = mapped_column(JSON)
    created_at: Mapped[dt.datetime] = mapped_column(default=utcnow)


class Job(Base):
    __tablename__ = "jobs"

    id: Mapped[int] = mapped_column(primary_key=True)
    created_at: Mapped[dt.datetime] = mapped_column(default=utcnow)
    bank_filename: Mapped[str] = mapped_column(String(255))
    orders_filename: Mapped[str] = mapped_column(String(255))
    bank_format: Mapped[str] = mapped_column(String(100))
    orders_format: Mapped[str] = mapped_column(String(100))
    window_days: Mapped[int]

    deposits: Mapped[list["Deposit"]] = relationship(
        back_populates="job", cascade="all, delete-orphan", order_by="Deposit.row_no"
    )
    orders: Mapped[list["Order"]] = relationship(
        back_populates="job", cascade="all, delete-orphan", order_by="Order.row_no"
    )
    links: Mapped[list["Link"]] = relationship(
        back_populates="job", cascade="all, delete-orphan", order_by="Link.id"
    )
    skipped: Mapped[list["SkippedRow"]] = relationship(
        back_populates="job",
        cascade="all, delete-orphan",
        order_by="(SkippedRow.source, SkippedRow.row_no)",
    )


class Deposit(Base):
    __tablename__ = "deposits"

    id: Mapped[int] = mapped_column(primary_key=True)
    job_id: Mapped[int] = mapped_column(ForeignKey("jobs.id"), index=True)
    row_no: Mapped[int]
    date: Mapped[dt.date]
    amount: Mapped[int]
    memo: Mapped[str] = mapped_column(Text, default="")
    last5: Mapped[str | None] = mapped_column(String(5))

    job: Mapped[Job] = relationship(back_populates="deposits")


class Order(Base):
    __tablename__ = "orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    job_id: Mapped[int] = mapped_column(ForeignKey("jobs.id"), index=True)
    row_no: Mapped[int]
    order_no: Mapped[str] = mapped_column(String(100))
    date: Mapped[dt.date]
    amount: Mapped[int]
    customer: Mapped[str] = mapped_column(String(200), default="")
    last5: Mapped[str | None] = mapped_column(String(5))

    job: Mapped[Job] = relationship(back_populates="orders")


class Link(Base):
    __tablename__ = "links"

    id: Mapped[int] = mapped_column(primary_key=True)
    job_id: Mapped[int] = mapped_column(ForeignKey("jobs.id"), index=True)
    deposit_id: Mapped[int] = mapped_column(ForeignKey("deposits.id"))
    order_id: Mapped[int] = mapped_column(ForeignKey("orders.id"))
    # matched / mismatch / candidate
    kind: Mapped[str] = mapped_column(String(10))
    layer: Mapped[int]
    reason: Mapped[str] = mapped_column(Text)
    # auto / pending / confirmed / rejected
    status: Mapped[str] = mapped_column(String(10))

    job: Mapped[Job] = relationship(back_populates="links")
    deposit: Mapped[Deposit] = relationship()
    order: Mapped[Order] = relationship()


class SkippedRow(Base):
    __tablename__ = "skipped_rows"

    id: Mapped[int] = mapped_column(primary_key=True)
    job_id: Mapped[int] = mapped_column(ForeignKey("jobs.id"), index=True)
    # bank / orders
    source: Mapped[str] = mapped_column(String(10))
    row_no: Mapped[int]
    reason: Mapped[str] = mapped_column(Text)

    job: Mapped[Job] = relationship(back_populates="skipped")


def init_db(database_url: str) -> sessionmaker:
    url = make_url(database_url)
    connect_args = {}
    if url.get_backend_name() == "sqlite":
        connect_args["check_same_thread"] = False
        if url.database and url.database != ":memory:":
            Path(url.database).parent.mkdir(parents=True, exist_ok=True)
    engine = create_engine(url, connect_args=connect_args)
    Base.metadata.create_all(engine)
    return sessionmaker(engine, expire_on_commit=False)
```

- [ ] **Step 6: 執行確認通過**

Run: `uv run pytest tests/test_models.py -v`
Expected: 2 passed

- [ ] **Step 7: Commit**

```bash
git add reconcile/config.py reconcile/models.py tests/conftest.py tests/test_models.py
git commit -m "加入設定與資料模型"
```

---

### Task 3: 讀取 CSV／Excel

**Files:**
- Create: `reconcile/readers.py`
- Test: `tests/test_readers.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_readers.py`**

```python
import datetime as dt

import pytest

from reconcile import readers
from reconcile.readers import ReadError, read_table
from tests.conftest import csv_bytes, xlsx_bytes


def test_utf8_csv_with_bom():
    data = "\ufeff訂單編號,日期,金額\nA1,2026/10/01,\"1,200\"\n".encode()
    table = read_table("orders.csv", data)
    assert table.header == ["訂單編號", "日期", "金額"]
    assert table.rows == [(2, ["A1", "2026/10/01", "1,200"])]


def test_big5_csv_with_title_rows():
    data = csv_bytes(
        ["台灣示範銀行 存款交易明細", "帳號：012-345-678901", "",
         "交易日期,摘要,存入金額", "115/10/01,跨行轉入,1200", ",,", "115/10/02,跨行轉入,860"],
        encoding="cp950",
    )
    table = read_table("bank.csv", data)
    assert table.header == ["交易日期", "摘要", "存入金額"]
    assert table.rows == [
        (5, ["115/10/01", "跨行轉入", "1200"]),
        (7, ["115/10/02", "跨行轉入", "860"]),
    ]


def test_xlsx_keeps_cell_types():
    data = xlsx_bytes([["示範銀行B 交易明細表"], ["帳務日", "說明", "存款"],
                       [dt.datetime(2026, 10, 1), " 轉入 ", 1200]])
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
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_readers.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.readers'`

- [ ] **Step 3: 實作 `reconcile/readers.py`**

```python
"""讀取上傳的 CSV 或 Excel，找出表頭列。"""

import csv
import io
from dataclasses import dataclass

from openpyxl import load_workbook

MAX_BYTES = 5 * 1024 * 1024
MAX_ROWS = 20_000
HEADER_SCAN_ROWS = 10
# 前幾列中，第一個至少有這麼多個非空白儲存格的列視為表頭（銀行匯出常在上方有標題列）
HEADER_MIN_COLUMNS = 3
# 台灣銀行匯出的 CSV 常是 Big5（cp950）
CSV_ENCODINGS = ("utf-8-sig", "cp950")


class ReadError(Exception):
    """檔案無法讀取；訊息直接顯示給使用者。"""


@dataclass
class Table:
    header: list[str]
    # （原始列號，儲存格）；列號從 1 起算，與 Excel／文字編輯器看到的一致
    rows: list[tuple[int, list]]

    def column(self, row: list, name: str):
        index = self.header.index(name)
        return row[index] if index < len(row) else None


def read_table(filename: str, data: bytes) -> Table:
    if len(data) > MAX_BYTES:
        raise ReadError("檔案超過 5 MB")
    name = filename.lower()
    if name.endswith(".csv"):
        raw = _read_csv(data)
    elif name.endswith((".xlsx", ".xlsm")):
        raw = _read_xlsx(data)
    else:
        raise ReadError("只支援 CSV 或 Excel（.xlsx）檔")
    return _to_table(raw)


def _read_csv(data: bytes) -> list[list]:
    for encoding in CSV_ENCODINGS:
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        return [[cell.strip() for cell in row] for row in csv.reader(io.StringIO(text))]
    raise ReadError("無法判斷 CSV 的文字編碼（支援 UTF-8 與 Big5）")


def _read_xlsx(data: bytes) -> list[list]:
    try:
        workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as e:
        raise ReadError("無法讀取 Excel 檔，檔案可能已損壞") from e
    try:
        sheet = workbook.worksheets[0]
        return [[_clean(value) for value in row] for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def _clean(value):
    return value.strip() if isinstance(value, str) else value


def _is_blank(value) -> bool:
    return value is None or value == ""


def _to_table(raw: list[list]) -> Table:
    for index, row in enumerate(raw[:HEADER_SCAN_ROWS]):
        if sum(not _is_blank(value) for value in row) >= HEADER_MIN_COLUMNS:
            header_index = index
            break
    else:
        raise ReadError("找不到表頭列（前 10 列內至少要有一列含 3 個以上的欄名）")
    header = ["" if _is_blank(value) else str(value).strip() for value in raw[header_index]]
    rows = [
        (number + 1, list(row))
        for number, row in enumerate(raw)
        if number > header_index and not all(_is_blank(value) for value in row)
    ]
    if len(rows) > MAX_ROWS:
        raise ReadError("資料超過 2 萬列")
    return Table(header, rows)
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_readers.py -v`
Expected: 9 passed

- [ ] **Step 5: Commit**

```bash
git add reconcile/readers.py tests/test_readers.py
git commit -m "加入 CSV／Excel 讀取"
```

---

### Task 4: 格式設定與辨識

**Files:**
- Create: `reconcile/formats.py`
- Test: `tests/test_formats.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_formats.py`**

```python
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
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_formats.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.formats'`

- [ ] **Step 3: 實作 `reconcile/formats.py`**

```python
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
    return errors
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_formats.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
git add reconcile/formats.py tests/test_formats.py
git commit -m "加入格式設定與辨識"
```

---

### Task 5: 正規化

**Files:**
- Create: `reconcile/normalize.py`
- Test: `tests/test_normalize.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_normalize.py`**

```python
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
    [("1,200", 1200), ("NT$1,200", 1200), ("$ 860 元", 860), (1200, 1200), (1200.0, 1200),
     ("1200.00", 1200), ("", None), (None, None), ("-300", -300)],
)
def test_parse_amount(value, expected):
    assert parse_amount(value) == expected


@pytest.mark.parametrize("value", ["abc", 12.5, "1.5", True])
def test_parse_amount_rejects(value):
    with pytest.raises(ValueError):
        parse_amount(value)


@pytest.mark.parametrize(
    ("value", "expected"),
    [("2026-10-05", dt.date(2026, 10, 5)), ("2026/10/5", dt.date(2026, 10, 5)),
     ("20261005", dt.date(2026, 10, 5)), ("115/10/05", dt.date(2026, 10, 5)),
     ("115.10.05", dt.date(2026, 10, 5)), ("2026/10/05 14:30", dt.date(2026, 10, 5)),
     (dt.datetime(2026, 10, 5, 9), dt.date(2026, 10, 5)), (dt.date(2026, 10, 5), dt.date(2026, 10, 5))],
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
        [(2, ["115/10/01", "1,200", "A1001", "822-12345"]),
         (3, ["115/10/01", "", "", ""]),          # 提款列：沒有存入
         (4, ["115/13/40", "500", "", ""]),       # 日期不存在
         (5, ["115/10/02", "abc", "", ""])],      # 金額無法解析
    )
    deposits, skipped = build_deposits(table, BANK_FMT)
    assert deposits == [DepositRow(2, dt.date(2026, 10, 1), 1200, "A1001", "12345")]
    assert [s.row_no for s in skipped] == [4, 5]
    assert "日期" in skipped[0].reason


def test_build_orders_validates_and_dedupes():
    table = Table(
        ["編號", "日期", "金額"],
        [(2, ["A1", "2026-10-01", 1200]),
         (3, ["", "2026-10-01", 100]),
         (4, ["A2", "2026-10-01", 0]),
         (5, ["a-1", "2026-10-02", 300])],
    )
    orders, skipped = build_orders(table, ORDER_FMT)
    assert orders == [OrderRow(2, "A1", dt.date(2026, 10, 1), 1200, "", None)]
    assert skipped == [
        Skipped(3, "缺少訂單編號"),
        Skipped(4, "金額必須大於 0"),
        Skipped(5, "訂單編號 a-1 重複"),
    ]
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_normalize.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.normalize'`

- [ ] **Step 3: 實作 `reconcile/normalize.py`**

```python
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
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_normalize.py -v`
Expected: 全部通過

- [ ] **Step 5: Commit**

```bash
git add reconcile/normalize.py tests/test_normalize.py
git commit -m "加入資料正規化"
```

---

### Task 6: 三層配對

**Files:**
- Create: `reconcile/matching.py`
- Test: `tests/test_matching.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_matching.py`**

```python
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
    orders = [order(2, "B1", 1, 990, "67890"), order(3, "B2", 1, 990, "67890"),
              order(4, "B3", 1, 990)]
    results = match([dep(2, 2, 990, last5="67890")], orders, 7)
    assert [(r.order_index, r.kind, r.layer, r.status) for r in results] == [
        (0, CANDIDATE, 2, PENDING), (1, CANDIDATE, 2, PENDING)
    ]


def test_layer3_amount_within_window():
    orders = [order(2, "C1", 1, 1500), order(3, "C2", 9, 1500), order(4, "C3", 1, 1600)]
    results = match([dep(2, 8, 1500)], orders, 7)
    assert [(r.order_index, r.layer) for r in results] == [(0, 3)]  # 10/1 + 7 天 = 10/8
    assert match([dep(2, 9, 1500)], orders[:1], 7) == []           # 超過容許天數


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
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_matching.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.matching'`

- [ ] **Step 3: 實作 `reconcile/matching.py`**

```python
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


def match(
    deposits: list[DepositRow], orders: list[OrderRow], window_days: int
) -> list[LinkResult]:
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
            reason += (
                f"，入帳 {deposit.amount:,} 元、訂單 {found.amount:,} 元，差額 {diff:+,} 元"
            )
            results.append(LinkResult(i, j, MISMATCH, 1, reason, AUTO))

    after_layer2 = []
    for i in remaining:  # 層級 ②
        deposit = deposits[i]
        hits = [
            j
            for j, o in enumerate(orders)
            if deposit.last5 and j not in used and o.last5 == deposit.last5
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
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_matching.py -v`
Expected: 8 passed

- [ ] **Step 5: Commit**

```bash
git add reconcile/matching.py tests/test_matching.py
git commit -m "加入三層配對"
```

---

### Task 7: 對帳工作服務

**Files:**
- Create: `reconcile/service.py`
- Test: `tests/test_service.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_service.py`**

```python
import datetime as dt

from sqlalchemy import select

from reconcile import service
from reconcile.formats import BANK, ORDERS, Format
from reconcile.models import FormatProfile, Job, utcnow
from reconcile.readers import Table

BANK_FMT = Format("A", BANK, {"date": "日期", "amount": "存入", "memo": "備註", "account": "帳號"})
ORDER_FMT = Format("O", ORDERS, {"order_no": "編號", "date": "日期", "amount": "金額"})


def make_job(session):
    bank = Table(
        ["日期", "存入", "備註", "帳號"],
        [(2, ["2026/10/01", "1200", "A1", ""]),     # ① 已配對
         (3, ["2026/10/02", "990", "", ""]),        # ③ 兩筆候選（B1、B2）
         (4, ["2026/10/03", "5000", "退款", ""]),   # 不明入帳
         (5, ["bad", "100", "", ""])],               # 略過
    )
    orders = Table(
        ["編號", "日期", "金額"],
        [(2, ["A1", "2026/10/01", "1200"]), (3, ["B1", "2026/10/01", "990"]),
         (4, ["B2", "2026/10/02", "990"]), (5, ["C1", "2026/10/02", "780"])],
    )
    return service.create_job(
        session,
        tables={BANK: bank, ORDERS: orders},
        formats={BANK: BANK_FMT, ORDERS: ORDER_FMT},
        filenames={BANK: "b.csv", ORDERS: "o.xlsx"},
        window_days=7,
    )


def test_create_job_and_view(session_factory):
    with session_factory() as session:
        job = make_job(session)
        view = service.job_view(job)
        assert [link.order.order_no for link in view.matched] == ["A1"]
        assert [link.order.order_no for link in view.pending] == ["B1", "B2"]
        assert [o.order_no for o in view.unpaid] == ["C1"]
        assert [d.amount for d in view.unknown] == [5000]
        assert [s.row_no for s in view.skipped] == [5]
        summary = view.summary
        assert summary["deposit_count"] == 3
        assert summary["order_count"] == 4
        assert summary["matched_count"] == 1
        assert summary["match_rate"] == 25
        assert summary["pending_count"] == 1
        assert summary["unpaid_total"] == 780
        assert summary["unknown_total"] == 5000


def test_confirm_rejects_competing_candidates(session_factory):
    with session_factory() as session:
        job = make_job(session)
        first, second = service.job_view(job).pending
        service.confirm_link(session, first)
        assert second.status == "rejected"
        view = service.job_view(job)
        assert [link.order.order_no for link in view.matched] == ["A1", "B1"]
        assert view.pending == []
        assert [o.order_no for o in view.unpaid] == ["B2", "C1"]


def test_reject_only_changes_that_candidate(session_factory):
    with session_factory() as session:
        job = make_job(session)
        first, second = service.job_view(job).pending
        service.reject_link(session, first)
        assert second.status == "pending"
        matched = service.job_view(job).matched[0]
        service.reject_link(session, matched)  # 自動配對不受影響
        assert matched.status == "auto"


def test_purge_expired(session_factory):
    with session_factory() as session:
        old = make_job(session)
        old.created_at = utcnow() - dt.timedelta(days=91)
        make_job(session)
        session.commit()
        assert service.purge_expired(session, 90) == 1
        assert service.purge_expired(session, 0) == 0
        assert len(session.scalars(select(Job)).all()) == 1


def test_saved_formats(session_factory):
    with session_factory() as session:
        session.add(FormatProfile(kind=BANK, name="我的銀行", mapping={"date": "D"}))
        session.commit()
        [fmt] = service.saved_formats(session, BANK)
        assert (fmt.name, fmt.builtin, fmt.mapping) == ("我的銀行", False, {"date": "D"})
        assert service.saved_formats(session, ORDERS) == []
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_service.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.service'`

- [ ] **Step 3: 實作 `reconcile/service.py`**

```python
"""對帳工作：建立、確認／排除候選、整理結果、到期刪除。"""

import datetime as dt
from dataclasses import dataclass

from sqlalchemy import select
from sqlalchemy.orm import Session

from reconcile import matching
from reconcile.formats import BANK, ORDERS, Format
from reconcile.models import Deposit, FormatProfile, Job, Link, Order, SkippedRow, utcnow
from reconcile.normalize import build_deposits, build_orders
from reconcile.readers import Table


def saved_formats(session: Session, kind: str) -> list[Format]:
    profiles = session.scalars(
        select(FormatProfile).where(FormatProfile.kind == kind).order_by(FormatProfile.id)
    )
    return [Format(p.name, p.kind, dict(p.mapping), builtin=False, id=p.id) for p in profiles]


def create_job(
    session: Session,
    *,
    tables: dict[str, Table],
    formats: dict[str, Format],
    filenames: dict[str, str],
    window_days: int,
) -> Job:
    deposits, bank_skipped = build_deposits(tables[BANK], formats[BANK])
    orders, order_skipped = build_orders(tables[ORDERS], formats[ORDERS])
    job = Job(
        bank_filename=filenames[BANK],
        orders_filename=filenames[ORDERS],
        bank_format=formats[BANK].name,
        orders_format=formats[ORDERS].name,
        window_days=window_days,
    )
    job.deposits = [
        Deposit(row_no=d.row_no, date=d.date, amount=d.amount, memo=d.memo, last5=d.last5)
        for d in deposits
    ]
    job.orders = [
        Order(row_no=o.row_no, order_no=o.order_no, date=o.date, amount=o.amount,
              customer=o.customer, last5=o.last5)
        for o in orders
    ]
    job.skipped = [SkippedRow(source=BANK, row_no=s.row_no, reason=s.reason) for s in bank_skipped]
    job.skipped += [
        SkippedRow(source=ORDERS, row_no=s.row_no, reason=s.reason) for s in order_skipped
    ]
    for r in matching.match(deposits, orders, window_days):
        job.links.append(
            Link(deposit=job.deposits[r.deposit_index], order=job.orders[r.order_index],
                 kind=r.kind, layer=r.layer, reason=r.reason, status=r.status)
        )
    session.add(job)
    session.commit()
    return job


def confirm_link(session: Session, link: Link) -> None:
    """確認候選；同一入帳或同一訂單的其他待確認候選自動排除。只處理待確認的候選。"""
    if link.status != matching.PENDING:
        return
    link.status = matching.CONFIRMED
    for other in link.job.links:
        if other is link or other.status != matching.PENDING:
            continue
        if other.deposit_id == link.deposit_id or other.order_id == link.order_id:
            other.status = matching.REJECTED
    session.commit()


def reject_link(session: Session, link: Link) -> None:
    if link.status != matching.PENDING:
        return
    link.status = matching.REJECTED
    session.commit()


@dataclass
class JobView:
    job: Job
    matched: list[Link]
    mismatched: list[Link]
    pending: list[Link]
    unpaid: list[Order]
    unknown: list[Deposit]
    skipped: list[SkippedRow]

    @property
    def summary(self) -> dict[str, int]:
        job = self.job
        return {
            "deposit_count": len(job.deposits),
            "deposit_total": sum(d.amount for d in job.deposits),
            "order_count": len(job.orders),
            "order_total": sum(o.amount for o in job.orders),
            "matched_count": len(self.matched),
            "match_rate": round(100 * len(self.matched) / len(job.orders)) if job.orders else 0,
            "pending_count": len({link.deposit_id for link in self.pending}),
            "mismatch_count": len(self.mismatched),
            "unpaid_count": len(self.unpaid),
            "unpaid_total": sum(o.amount for o in self.unpaid),
            "unknown_count": len(self.unknown),
            "unknown_total": sum(d.amount for d in self.unknown),
            "skipped_count": len(self.skipped),
        }


def job_view(job: Job) -> JobView:
    effective = [link for link in job.links if matching.is_effective(link.kind, link.status)]
    pending = [link for link in job.links if link.status == matching.PENDING]
    pending.sort(key=lambda link: (link.deposit.date, link.deposit.row_no, link.order.row_no))
    busy_deposits = {link.deposit_id for link in effective + pending}
    busy_orders = {link.order_id for link in effective + pending}
    return JobView(
        job=job,
        matched=[link for link in effective if link.kind != matching.MISMATCH],
        mismatched=[link for link in effective if link.kind == matching.MISMATCH],
        pending=pending,
        unpaid=[o for o in job.orders if o.id not in busy_orders],
        unknown=[d for d in job.deposits if d.id not in busy_deposits],
        skipped=list(job.skipped),
    )


def purge_expired(session: Session, retention_days: int, now: dt.datetime | None = None) -> int:
    if retention_days <= 0:
        return 0
    cutoff = (now or utcnow()) - dt.timedelta(days=retention_days)
    expired = session.scalars(select(Job).where(Job.created_at < cutoff)).all()
    for job in expired:
        session.delete(job)
    session.commit()
    return len(expired)
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_service.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add reconcile/service.py tests/test_service.py
git commit -m "加入對帳工作服務"
```

---

### Task 8: Excel 報表

**Files:**
- Create: `reconcile/report.py`
- Test: `tests/test_report.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_report.py`**

```python
import io

from openpyxl import load_workbook

from reconcile import service
from reconcile.report import SHEETS, build_report
from tests.test_service import make_job


def test_report_sheets_and_rows(session_factory):
    with session_factory() as session:
        view = service.job_view(make_job(session))
        workbook = load_workbook(io.BytesIO(build_report(view)))
    assert workbook.sheetnames == list(SHEETS)
    assert workbook["摘要"]["A2"].value == "銀行入帳筆數"
    assert workbook["已配對"].max_row == 2       # 標題＋1 筆
    assert workbook["待確認"].max_row == 3       # 標題＋2 筆候選
    assert workbook["未收款訂單"]["A2"].value == "C1"
    assert workbook["不明入帳"]["C2"].value == 5000
    assert workbook["不明入帳"]["C2"].number_format == "#,##0"
    assert workbook["略過的資料列"]["B2"].value == 5
    assert workbook["已配對"].freeze_panes == "A2"
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_report.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.report'`

- [ ] **Step 3: 實作 `reconcile/report.py`**

```python
"""產生 Excel 對帳報表。"""

import io

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from reconcile.formats import KIND_LABELS
from reconcile.service import JobView

SHEETS = ("摘要", "已配對", "待確認", "金額不符", "未收款訂單", "不明入帳", "略過的資料列")
RED = PatternFill("solid", fgColor="FDE2E1")
YELLOW = PatternFill("solid", fgColor="FEF3C7")
MONEY = "#,##0"
DATE = "yyyy-mm-dd"


def build_report(view: JobView) -> bytes:
    s = view.summary
    workbook = Workbook()
    workbook.remove(workbook.active)
    _sheet(
        workbook, "摘要", ["項目", "數值"],
        [("銀行入帳筆數", s["deposit_count"]), ("入帳總額", s["deposit_total"]),
         ("訂單筆數", s["order_count"]), ("訂單總額", s["order_total"]),
         ("已配對筆數", s["matched_count"]), ("配對率（%）", s["match_rate"]),
         ("待確認入帳", s["pending_count"]), ("金額不符", s["mismatch_count"]),
         ("未收款筆數", s["unpaid_count"]), ("未收款總額", s["unpaid_total"]),
         ("不明入帳筆數", s["unknown_count"]), ("不明入帳總額", s["unknown_total"]),
         ("略過的資料列", s["skipped_count"])],
        money=(2,),
    )
    pair_header = ["入帳日期", "入帳金額", "備註", "訂單編號", "訂單日期", "訂單金額", "客戶", "依據"]
    _sheet(workbook, "已配對", pair_header, [_pair(link) for link in view.matched],
           money=(2, 6), dates=(1, 5))
    _sheet(workbook, "待確認", ["入帳列號", *pair_header],
           [(link.deposit.row_no, *_pair(link)) for link in view.pending],
           fill=YELLOW, money=(3, 7), dates=(2, 6))
    _sheet(workbook, "金額不符", [*pair_header, "差額"],
           [(*_pair(link), link.deposit.amount - link.order.amount) for link in view.mismatched],
           fill=RED, money=(2, 6, 9), dates=(1, 5))
    _sheet(workbook, "未收款訂單", ["訂單編號", "訂單日期", "訂單金額", "客戶"],
           [(o.order_no, o.date, o.amount, o.customer) for o in view.unpaid],
           fill=RED, money=(3,), dates=(2,))
    _sheet(workbook, "不明入帳", ["入帳列號", "入帳日期", "入帳金額", "備註", "帳號末五碼"],
           [(d.row_no, d.date, d.amount, d.memo, d.last5 or "") for d in view.unknown],
           fill=RED, money=(3,), dates=(2,))
    _sheet(workbook, "略過的資料列", ["來源", "原始列號", "原因"],
           [(KIND_LABELS[r.source], r.row_no, r.reason) for r in view.skipped])
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _pair(link) -> tuple:
    d, o = link.deposit, link.order
    return (d.date, d.amount, d.memo, o.order_no, o.date, o.amount, o.customer, link.reason)


def _sheet(workbook, title, header, rows, *, fill=None, money=(), dates=()):
    sheet = workbook.create_sheet(title)
    sheet.append(header)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    for row in rows:
        sheet.append(list(row))
        for cell in sheet[sheet.max_row]:
            if fill:
                cell.fill = fill
            if cell.column in money:
                cell.number_format = MONEY
            elif cell.column in dates:
                cell.number_format = DATE
    for column in sheet.columns:
        width = max(len(str(cell.value or "")) for cell in column) * 2
        sheet.column_dimensions[column[0].column_letter].width = min(50, max(10, width))
    return sheet
```

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_report.py -v`
Expected: 1 passed

- [ ] **Step 5: Commit**

```bash
git add reconcile/report.py tests/test_report.py
git commit -m "加入 Excel 報表"
```

---

### Task 9: 範例資料

**Files:**
- Create: `scripts/make_samples.py`, `samples/示範銀行A_115年10月.csv`, `samples/示範銀行B_115年10月.xlsx`, `samples/示範電商訂單_2026年10月.xlsx`
- Test: `tests/test_samples.py`

- [ ] **Step 1: 實作 `scripts/make_samples.py`**

```python
"""產生虛構的範例檔到 samples/。資料為虛構，涵蓋每一種配對結果。

用法：uv run python scripts/make_samples.py
"""

import csv
import datetime as dt
import io
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
    ("115/10/01", "提款", "3,000", "", "", ""),                         # 提款，不算入帳
    ("115/10/02", "跨行轉入", "", "860", "訂單a-1002", "700-0002-33333"),  # ① 已配對
    ("115/10/03", "跨行轉入", "", "2,385", "A1003 林", "812-0003-44444"),  # ① 金額不符
    ("115/10/03", "跨行轉入", "", "2,400", "", "013-0004-54321"),      # ③ 候選 A1004
    ("115/10/04", "ATM轉入", "", "1,500", "", ""),                      # ③ 候選 A1005
    ("115/10/05", "跨行轉入", "", "990", "", "808-0005-67890"),        # ② 兩筆候選
    ("115/10/05", "跨行轉入", "", "3,200", "", ""),                     # ③ 候選 A1006
    ("115/10/06", "跨行轉入", "", "5,000", "退款", "822-0006-99999"),  # 不明入帳
    ("115/10/07", "跨行轉入", "", "1,350", "", "004-0007-11111"),      # ③ 候選 A1011
    ("115/10/08", "跨行轉入", "", "2,050", "", "007-0008-24680"),      # ② 已配對
    ("115/13/40", "跨行轉入", "", "500", "", ""),                       # 略過：日期不存在
]


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
    sheet.append([dt.datetime(2026, 10, 1), "網銀轉入", None, 1200, 21200, "A1001", "0001234512345"])
    sheet.append([dt.datetime(2026, 10, 2), "轉帳支出", 500, None, 20700, "", ""])
    sheet.append([dt.datetime(2026, 10, 3), "網銀轉入", None, 2050, 22750, "", "0007000824680"])
    workbook.save(SAMPLES / BANK_B)


def write_orders() -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["訂單編號", "下單日期", "訂單金額", "收件人", "匯款帳號末五碼"])
    for order_no, day, amount, name, last5 in ORDER_ROWS:
        sheet.append([order_no, dt.datetime(2026, 10, day), amount, name, last5])
    workbook.save(SAMPLES / ORDERS)


def main() -> None:
    SAMPLES.mkdir(exist_ok=True)
    write_bank_a()
    write_bank_b()
    write_orders()
    print(f"已產生範例檔到 {SAMPLES}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 產生範例檔**

Run: `uv run python scripts/make_samples.py && ls samples`
Expected: 三個檔案

- [ ] **Step 3: 寫測試 `tests/test_samples.py`**

```python
from reconcile import service
from reconcile.config import Settings
from reconcile.formats import BANK, ORDERS, detect
from reconcile.readers import read_table
from scripts.make_samples import BANK_A, BANK_B, ORDERS as ORDERS_FILE

SAMPLES = Settings().samples_dir


def load(name):
    return read_table(name, (SAMPLES / name).read_bytes())


def test_bank_b_sample_is_detected():
    assert detect(load(BANK_B).header, BANK, []).name == "示範銀行 B"


def test_sample_job_covers_every_outcome(session_factory):
    tables = {BANK: load(BANK_A), ORDERS: load(ORDERS_FILE)}
    formats = {kind: detect(tables[kind].header, kind, []) for kind in tables}
    assert formats[BANK].name == "示範銀行 A"
    with session_factory() as session:
        job = service.create_job(session, tables=tables, formats=formats,
                                 filenames={BANK: BANK_A, ORDERS: ORDERS_FILE}, window_days=7)
        view = service.job_view(job)
        assert [link.order.order_no for link in view.matched] == ["A1001", "A1002", "A1012"]
        assert [link.order.order_no for link in view.mismatched] == ["A1003"]
        assert [link.order.order_no for link in view.pending] == [
            "A1004", "A1005", "A1007", "A1008", "A1006", "A1011"
        ]
        assert [o.order_no for o in view.unpaid] == ["A1009", "A1010"]
        assert [d.amount for d in view.unknown] == [5000]
        assert [s.reason for s in view.skipped] == ["日期不存在「115/13/40」"]
```

`scripts/__init__.py`：空檔（讓測試可以 import `scripts.make_samples`）。

- [ ] **Step 4: 執行確認通過**

Run: `uv run pytest tests/test_samples.py -v`
Expected: 2 passed。若配對結果與預期不同，先檢查範例資料與 `matching.py` 的規則，說明差異原因後再決定修哪一邊。

- [ ] **Step 5: Commit**

```bash
git add scripts samples tests/test_samples.py
git commit -m "加入虛構範例資料"
```

---

### Task 10: 網頁：上傳、結果頁、確認、報表、範例

**Files:**
- Create: `reconcile/web/__init__.py`（空檔）, `reconcile/web/uploads.py`, `reconcile/web/app.py`, `reconcile/web/templates/{base,index,job}.html`, `reconcile/web/static/style.css`
- Test: `tests/test_web.py`

- [ ] **Step 1: 寫失敗的測試 `tests/test_web.py`**

```python
import pytest
from fastapi.testclient import TestClient

from reconcile.config import Settings
from reconcile.web.app import create_app
from scripts.make_samples import BANK_A, ORDERS


@pytest.fixture
def settings(tmp_path):
    return Settings(database_url=f"sqlite:///{tmp_path / 'web.db'}", upload_dir=tmp_path / "up")


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings)) as client:
        yield client


def sample_files(settings):
    return {
        "bank_file": (BANK_A, (settings.samples_dir / BANK_A).read_bytes()),
        "orders_file": (ORDERS, (settings.samples_dir / ORDERS).read_bytes()),
    }


def test_index_opens(client):
    response = client.get("/")
    assert response.status_code == 200
    assert "載入範例資料" in response.text


def test_upload_detected_formats_creates_job(client, settings):
    response = client.post("/jobs", files=sample_files(settings), data={"window_days": "7"},
                           follow_redirects=False)
    assert response.status_code == 303
    assert response.headers["location"] == "/jobs/1"
    page = client.get("/jobs/1").text
    assert "A1003" in page and "待確認" in page and "日期不存在" in page


def test_upload_rejects_unreadable_file(client):
    response = client.post("/jobs", files={"bank_file": ("x.txt", b"x"),
                                           "orders_file": ("o.txt", b"x")})
    assert response.status_code == 400
    assert "銀行明細：只支援" in response.text


def test_upload_rejects_bad_window(client, settings):
    response = client.post("/jobs", files=sample_files(settings), data={"window_days": "99"})
    assert response.status_code == 400


def test_samples_confirm_and_report(client):
    response = client.post("/samples", follow_redirects=False)
    assert response.headers["location"] == "/jobs/1"
    # 範例資料的 Link id：4、5 為 A1007／A1008（層級 ② 兩筆候選），7 為 A1004（層級 ③）
    response = client.post("/links/4/confirm", follow_redirects=False)
    assert response.status_code == 303
    assert response.headers["location"] == "/jobs/1#pending"
    response = client.post("/links/7/reject", follow_redirects=False)
    assert response.status_code == 303
    report = client.get("/jobs/1/report.xlsx")
    assert report.status_code == 200
    assert report.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "filename*=UTF-8''" in report.headers["content-disposition"]


def test_delete_job_and_404s(client):
    client.post("/samples")
    assert client.post("/jobs/1/delete", follow_redirects=False).status_code == 303
    assert client.get("/jobs/1").status_code == 404
    assert client.post("/links/1/confirm").status_code == 404
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_web.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'reconcile.web'`

- [ ] **Step 3: `reconcile/web/uploads.py`**（欄位對應頁需要暫存上傳檔）

```python
"""暫存等待欄位對應的上傳檔；以隨機 token 識別，超過 1 小時刪除。"""

import json
import re
import secrets
import shutil
import time
from pathlib import Path

TTL_SECONDS = 3600
TOKEN = re.compile(r"[A-Za-z0-9_-]{16,64}")


def save_pending(upload_dir: Path, files: dict[str, tuple[str, bytes]], window_days: int) -> str:
    token = secrets.token_urlsafe(16)
    folder = upload_dir / token
    folder.mkdir(parents=True)
    for kind, (_, data) in files.items():
        (folder / kind).write_bytes(data)
    meta = {"window_days": window_days, "filenames": {k: name for k, (name, _) in files.items()}}
    (folder / "meta.json").write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    return token


def load_pending(upload_dir: Path, token: str) -> tuple[dict, dict[str, tuple[str, bytes]]] | None:
    if not TOKEN.fullmatch(token):  # 防止路徑穿越
        return None
    folder = upload_dir / token
    if not (folder / "meta.json").is_file():
        return None
    meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
    files = {kind: (name, (folder / kind).read_bytes()) for kind, name in meta["filenames"].items()}
    return meta, files


def remove_pending(upload_dir: Path, token: str) -> None:
    if TOKEN.fullmatch(token):
        shutil.rmtree(upload_dir / token, ignore_errors=True)


def purge_pending(upload_dir: Path) -> None:
    if not upload_dir.is_dir():
        return
    for folder in upload_dir.iterdir():
        if folder.is_dir() and time.time() - folder.stat().st_mtime > TTL_SECONDS:
            shutil.rmtree(folder, ignore_errors=True)
```

- [ ] **Step 4: `reconcile/web/app.py`**

```python
import datetime as dt
from contextlib import asynccontextmanager
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy import select
from sqlalchemy.orm import Session

from reconcile import service
from reconcile.config import Settings
from reconcile.formats import BANK, KIND_LABELS, ORDERS, detect
from reconcile.models import Job, Link, init_db
from reconcile.readers import MAX_BYTES, ReadError, read_table
from reconcile.report import build_report
from reconcile.web import uploads

HERE = Path(__file__).parent
DISPLAY_TZ = ZoneInfo("Asia/Taipei")
SAMPLE_FILES = {BANK: "示範銀行A_115年10月.csv", ORDERS: "示範電商訂單_2026年10月.xlsx"}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def localtime(value: dt.datetime | None) -> str:
    if value is None:
        return "—"
    return value.replace(tzinfo=dt.UTC).astimezone(DISPLAY_TZ).strftime("%Y-%m-%d %H:%M")


templates = Jinja2Templates(directory=HERE / "templates")
templates.env.filters["localtime"] = localtime
templates.env.filters["money"] = lambda v: f"{v:,}" if v is not None else "—"
templates.env.globals["KIND_LABELS"] = KIND_LABELS


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or Settings.from_env()
    session_factory = init_db(settings.database_url)
    upload_dir = settings.upload_dir

    def housekeeping() -> None:
        with session_factory() as session:
            service.purge_expired(session, settings.retention_days)
        uploads.purge_pending(upload_dir)

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        housekeeping()
        yield

    app = FastAPI(title="銀行入帳對帳", lifespan=lifespan)
    app.mount("/static", StaticFiles(directory=HERE / "static"), name="static")
    app.state.session_factory = session_factory

    def render_index(request: Request, session: Session, error=None, status_code=200):
        jobs = session.scalars(select(Job).order_by(Job.id.desc())).all()
        rows = [(job, service.job_view(job).summary) for job in jobs]
        return templates.TemplateResponse(
            request, "index.html",
            {"rows": rows, "error": error, "retention_days": settings.retention_days},
            status_code=status_code,
        )

    def load_job(session: Session, job_id: int) -> Job:
        job = session.get(Job, job_id)
        if job is None:
            raise HTTPException(404)
        return job

    def load_link(session: Session, link_id: int) -> Link:
        link = session.get(Link, link_id)
        if link is None:
            raise HTTPException(404)
        return link

    def start_job(session: Session, files: dict, window_days: int):
        """讀檔並辨識格式；都辨識得出來就建立對帳工作，否則轉到欄位對應頁。"""
        tables = {}
        for kind, (filename, data) in files.items():
            try:
                tables[kind] = read_table(filename, data)
            except ReadError as e:
                raise ReadError(f"{KIND_LABELS[kind]}：{e}") from e
        formats = {
            kind: detect(tables[kind].header, kind, service.saved_formats(session, kind))
            for kind in tables
        }
        if all(formats.values()):
            job = service.create_job(
                session, tables=tables, formats=formats,
                filenames={kind: name for kind, (name, _) in files.items()},
                window_days=window_days,
            )
            return RedirectResponse(f"/jobs/{job.id}", status_code=303)
        token = uploads.save_pending(upload_dir, files, window_days)
        return RedirectResponse(f"/mapping/{token}", status_code=303)

    @app.get("/", response_class=HTMLResponse)
    def index(request: Request):
        housekeeping()
        with session_factory() as session:
            return render_index(request, session)

    @app.post("/jobs", response_class=HTMLResponse)
    async def upload(
        request: Request,
        bank_file: UploadFile = File(...),
        orders_file: UploadFile = File(...),
        window_days: int = Form(7),
    ):
        files = {
            BANK: (bank_file.filename or "", await bank_file.read(MAX_BYTES + 1)),
            ORDERS: (orders_file.filename or "", await orders_file.read(MAX_BYTES + 1)),
        }
        with session_factory() as session:
            if not 0 <= window_days <= 60:
                return render_index(request, session, "日期容許天數須介於 0 到 60", 400)
            try:
                return start_job(session, files, window_days)
            except ReadError as e:
                return render_index(request, session, str(e), 400)

    @app.post("/samples")
    def load_samples(request: Request):
        files = {
            kind: (name, (settings.samples_dir / name).read_bytes())
            for kind, name in SAMPLE_FILES.items()
        }
        with session_factory() as session:
            return start_job(session, files, 7)

    @app.get("/jobs/{job_id}", response_class=HTMLResponse)
    def job_page(request: Request, job_id: int):
        with session_factory() as session:
            job = load_job(session, job_id)
            return templates.TemplateResponse(
                request, "job.html", {"job": job, "view": service.job_view(job)}
            )

    @app.post("/links/{link_id}/confirm")
    def confirm(link_id: int):
        with session_factory() as session:
            link = load_link(session, link_id)
            service.confirm_link(session, link)
            return RedirectResponse(f"/jobs/{link.job_id}#pending", status_code=303)

    @app.post("/links/{link_id}/reject")
    def reject(link_id: int):
        with session_factory() as session:
            link = load_link(session, link_id)
            service.reject_link(session, link)
            return RedirectResponse(f"/jobs/{link.job_id}#pending", status_code=303)

    @app.get("/jobs/{job_id}/report.xlsx")
    def report(job_id: int):
        with session_factory() as session:
            job = load_job(session, job_id)
            data = build_report(service.job_view(job))
            filename = f"對帳報表_工作{job.id}_{job.created_at:%Y%m%d}.xlsx"
        disposition = (
            f'attachment; filename="report_{job_id}.xlsx"; filename*=UTF-8\'\'{quote(filename)}'
        )
        return Response(data, media_type=XLSX, headers={"Content-Disposition": disposition})

    @app.post("/jobs/{job_id}/delete")
    def delete_job(job_id: int):
        with session_factory() as session:
            session.delete(load_job(session, job_id))
            session.commit()
        return RedirectResponse("/", status_code=303)

    return app
```

- [ ] **Step 5: 模板**

`reconcile/web/templates/base.html`：

```html
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{% block title %}銀行入帳對帳{% endblock %}</title>
  <link rel="stylesheet" href="/static/style.css">
</head>
<body>
  <header>
    <a class="brand" href="/">🧾 銀行入帳對帳</a>
    <nav><a href="/">對帳工作</a><a href="/formats">格式設定</a></nav>
  </header>
  <main>{% block content %}{% endblock %}</main>
  <footer>Code Sky 逅斯蓋工作室</footer>
</body>
</html>
```

`reconcile/web/templates/index.html`：

```html
{% extends "base.html" %}
{% block content %}
<section class="card">
  <h2>新增對帳</h2>
  <p class="muted">上傳銀行入帳明細與訂單匯出檔（CSV 或 Excel），系統會自動辨識格式並配對；格式辨識不出來時，會請你指定欄位。</p>
  {% if error %}<p class="error">{{ error }}</p>{% endif %}
  <form class="upload" method="post" action="/jobs" enctype="multipart/form-data">
    <label>銀行明細 <input type="file" name="bank_file" accept=".csv,.xlsx" required></label>
    <label>訂單 <input type="file" name="orders_file" accept=".csv,.xlsx" required></label>
    <label>日期容許天數 <input type="number" name="window_days" min="0" max="60" value="7"> 天</label>
    <button>開始對帳</button>
  </form>
  <form method="post" action="/samples"><button class="secondary">載入範例資料</button></form>
</section>

<section>
  <h2>對帳工作</h2>
  <p class="muted">對帳工作保留 {{ retention_days }} 天，到期自動刪除（含銀行資料）。</p>
  <table>
    <thead><tr><th>#</th><th>建立時間</th><th>銀行明細</th><th>訂單</th><th>配對率</th><th>待確認</th><th>未收款</th><th>不明入帳</th></tr></thead>
    <tbody>
    {% for job, s in rows %}
      <tr>
        <td><a href="/jobs/{{ job.id }}">{{ job.id }}</a></td>
        <td>{{ job.created_at | localtime }}</td>
        <td>{{ job.bank_filename }}</td>
        <td>{{ job.orders_filename }}</td>
        <td>{{ s.match_rate }}%</td>
        <td>{{ s.pending_count }}</td>
        <td>{{ s.unpaid_count }}</td>
        <td>{{ s.unknown_count }}</td>
      </tr>
    {% else %}
      <tr><td colspan="8" class="muted">尚無對帳工作</td></tr>
    {% endfor %}
    </tbody>
  </table>
</section>
{% endblock %}
```

`reconcile/web/templates/job.html`：

```html
{% extends "base.html" %}
{% block title %}對帳工作 #{{ job.id }}｜銀行入帳對帳{% endblock %}
{% block content %}
{% set s = view.summary %}
<p><a href="/">← 回到對帳工作</a></p>
<h2>對帳工作 #{{ job.id }}</h2>
<p class="muted">{{ job.created_at | localtime }}｜銀行明細：{{ job.bank_filename }}（{{ job.bank_format }}）｜訂單：{{ job.orders_filename }}（{{ job.orders_format }}）｜日期容許 {{ job.window_days }} 天</p>
<p class="actions">
  <a class="button" href="/jobs/{{ job.id }}/report.xlsx">下載 Excel 報表</a>
  <form method="post" action="/jobs/{{ job.id }}/delete" onsubmit="return confirm('確定刪除這個對帳工作？')"><button class="danger">刪除</button></form>
</p>

<div class="cards">
  <div class="stat"><b>{{ s.match_rate }}%</b><span>配對率（{{ s.matched_count }} / {{ s.order_count }} 筆訂單）</span></div>
  <div class="stat warn"><b>{{ s.pending_count }}</b><span>待確認入帳</span></div>
  <div class="stat bad"><b>{{ s.mismatch_count }}</b><span>金額不符</span></div>
  <div class="stat bad"><b>{{ s.unpaid_count }}</b><span>未收款（{{ s.unpaid_total | money }} 元）</span></div>
  <div class="stat bad"><b>{{ s.unknown_count }}</b><span>不明入帳（{{ s.unknown_total | money }} 元）</span></div>
  <div class="stat"><b>{{ s.deposit_count }}</b><span>入帳（{{ s.deposit_total | money }} 元）</span></div>
</div>

<h3 id="pending">待確認</h3>
<table>
  <thead><tr><th>入帳</th><th>金額</th><th>備註</th><th>候選訂單</th><th>訂單日期</th><th>客戶</th><th>依據</th><th></th></tr></thead>
  <tbody>
  {% for link in view.pending %}
    <tr class="row-warn">
      <td>第 {{ link.deposit.row_no }} 列｜{{ link.deposit.date }}</td>
      <td>{{ link.deposit.amount | money }}</td>
      <td>{{ link.deposit.memo }}</td>
      <td>{{ link.order.order_no }}</td>
      <td>{{ link.order.date }}</td>
      <td>{{ link.order.customer }}</td>
      <td>{{ link.reason }}</td>
      <td class="nowrap">
        <form method="post" action="/links/{{ link.id }}/confirm"><button>確認</button></form>
        <form method="post" action="/links/{{ link.id }}/reject"><button class="secondary">排除</button></form>
      </td>
    </tr>
  {% else %}
    <tr><td colspan="8" class="muted">沒有待確認項目</td></tr>
  {% endfor %}
  </tbody>
</table>

<h3>金額不符</h3>
<table>
  <thead><tr><th>入帳日期</th><th>入帳金額</th><th>訂單</th><th>訂單金額</th><th>依據</th></tr></thead>
  <tbody>
  {% for link in view.mismatched %}
    <tr class="row-bad"><td>{{ link.deposit.date }}</td><td>{{ link.deposit.amount | money }}</td><td>{{ link.order.order_no }}</td><td>{{ link.order.amount | money }}</td><td>{{ link.reason }}</td></tr>
  {% else %}
    <tr><td colspan="5" class="muted">無</td></tr>
  {% endfor %}
  </tbody>
</table>

<h3>未收款訂單</h3>
<table>
  <thead><tr><th>訂單編號</th><th>訂單日期</th><th>金額</th><th>客戶</th></tr></thead>
  <tbody>
  {% for o in view.unpaid %}
    <tr class="row-bad"><td>{{ o.order_no }}</td><td>{{ o.date }}</td><td>{{ o.amount | money }}</td><td>{{ o.customer }}</td></tr>
  {% else %}
    <tr><td colspan="4" class="muted">無</td></tr>
  {% endfor %}
  </tbody>
</table>

<h3>不明入帳</h3>
<table>
  <thead><tr><th>列號</th><th>日期</th><th>金額</th><th>備註</th><th>帳號末五碼</th></tr></thead>
  <tbody>
  {% for d in view.unknown %}
    <tr class="row-bad"><td>{{ d.row_no }}</td><td>{{ d.date }}</td><td>{{ d.amount | money }}</td><td>{{ d.memo }}</td><td>{{ d.last5 or "" }}</td></tr>
  {% else %}
    <tr><td colspan="5" class="muted">無</td></tr>
  {% endfor %}
  </tbody>
</table>

<h3>已配對</h3>
<table>
  <thead><tr><th>入帳日期</th><th>金額</th><th>訂單</th><th>客戶</th><th>依據</th></tr></thead>
  <tbody>
  {% for link in view.matched %}
    <tr><td>{{ link.deposit.date }}</td><td>{{ link.deposit.amount | money }}</td><td>{{ link.order.order_no }}</td><td>{{ link.order.customer }}</td><td>{{ link.reason }}{% if link.status == "confirmed" %}（已確認）{% endif %}</td></tr>
  {% else %}
    <tr><td colspan="5" class="muted">無</td></tr>
  {% endfor %}
  </tbody>
</table>

<h3>略過的資料列</h3>
<table>
  <thead><tr><th>來源</th><th>原始列號</th><th>原因</th></tr></thead>
  <tbody>
  {% for r in view.skipped %}
    <tr><td>{{ KIND_LABELS[r.source] }}</td><td>{{ r.row_no }}</td><td>{{ r.reason }}</td></tr>
  {% else %}
    <tr><td colspan="3" class="muted">無</td></tr>
  {% endfor %}
  </tbody>
</table>
{% endblock %}
```

- [ ] **Step 6: `reconcile/web/static/style.css`**

```css
:root { --fg: #1f2937; --muted: #6b7280; --line: #e5e7eb; --accent: #2563eb; --bad: #dc2626; --warn: #d97706; }
* { box-sizing: border-box; }
body { margin: 0; font-family: -apple-system, "PingFang TC", "Noto Sans TC", sans-serif; color: var(--fg); background: #f9fafb; }
header { display: flex; align-items: center; gap: 2rem; padding: .8rem 2rem; background: #111827; }
header a { color: #e5e7eb; text-decoration: none; margin-right: 1rem; }
header .brand { color: #fff; font-weight: bold; font-size: 1.1rem; }
main { max-width: 1150px; margin: 0 auto; padding: 1.5rem 2rem; }
footer { text-align: center; color: var(--muted); font-size: .85rem; padding: 2rem; }
section { margin-bottom: 2rem; }
.card { background: #fff; border: 1px solid var(--line); padding: 1rem 1.5rem; }
a { color: var(--accent); }
table { width: 100%; border-collapse: collapse; background: #fff; border: 1px solid var(--line); margin-bottom: 1.5rem; }
th, td { padding: .5rem .7rem; border-bottom: 1px solid var(--line); text-align: left; font-size: .9rem; vertical-align: top; }
th { background: #f3f4f6; font-weight: 600; }
tr.row-bad td { background: #fef2f2; }
tr.row-warn td { background: #fffbeb; }
.muted { color: var(--muted); }
.error { color: var(--bad); font-weight: 600; }
.nowrap { white-space: nowrap; }
.nowrap form { display: inline; }
form.upload { display: flex; gap: 1rem; align-items: center; flex-wrap: wrap; margin-bottom: .75rem; }
input, select, button, .button { font: inherit; padding: .4rem .7rem; border: 1px solid #d1d5db; border-radius: 6px; }
input[type=number] { width: 5rem; }
button, .button { background: var(--accent); color: #fff; border-color: var(--accent); cursor: pointer; text-decoration: none; display: inline-block; }
button.secondary { background: #fff; color: var(--fg); border-color: #d1d5db; }
button.danger { background: var(--bad); border-color: var(--bad); }
.actions { display: flex; gap: .75rem; align-items: center; }
.actions form { margin: 0; }
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: .75rem; margin-bottom: 1.5rem; }
.stat { background: #fff; border: 1px solid var(--line); padding: .8rem 1rem; }
.stat b { display: block; font-size: 1.6rem; }
.stat span { color: var(--muted); font-size: .85rem; }
.stat.warn b { color: var(--warn); }
.stat.bad b { color: var(--bad); }
.mapping { display: grid; grid-template-columns: 10rem 1fr; gap: .5rem 1rem; align-items: center; max-width: 640px; margin-bottom: 1rem; }
.preview { font-size: .8rem; }
```

- [ ] **Step 7: 執行確認通過**

Run: `uv run pytest tests/test_web.py -v`
Expected: 6 passed。注意 `test_samples_confirm_and_report` 依賴範例資料的 Link id 順序（id 4、5、7 為待確認候選）；若不同，先用 `uv run python -c` 列出 links 確認順序，再調整測試中的 id 並說明原因。

- [ ] **Step 8: Commit**

```bash
git add reconcile/web tests/test_web.py
git commit -m "加入網頁：上傳、結果頁、確認與報表"
```

---

### Task 11: 網頁：欄位對應與格式設定

**Files:**
- Modify: `reconcile/web/app.py`
- Create: `reconcile/web/templates/mapping.html`, `reconcile/web/templates/formats.html`
- Test: `tests/test_web.py`

- [ ] **Step 1: 在 `tests/test_web.py` 末尾加入失敗的測試**

在 import 區加入 `from reconcile.models import FormatProfile` 與 `from tests.conftest import csv_bytes`。

```python
CUSTOM_BANK = csv_bytes(["入帳日,金額,說明,帳號", "2026/10/01,1200,A1001,822-12345"])


def upload_custom(client, settings):
    files = sample_files(settings)
    files["bank_file"] = ("my_bank.csv", CUSTOM_BANK)
    return client.post("/jobs", files=files, follow_redirects=False)


def test_unknown_format_goes_to_mapping(client, settings):
    response = upload_custom(client, settings)
    assert response.status_code == 303
    location = response.headers["location"]
    assert location.startswith("/mapping/")
    page = client.get(location).text
    assert "入帳日" in page and "示範電商訂單" in page  # 訂單檔已辨識


def test_mapping_creates_job_and_saves_format(client, settings):
    location = upload_custom(client, settings).headers["location"]
    response = client.post(location, data={
        "bank_date": "入帳日", "bank_amount": "金額", "bank_memo": "說明",
        "bank_account": "帳號", "bank_save_name": "我的銀行",
    }, follow_redirects=False)
    assert response.headers["location"] == "/jobs/1"
    assert "A1001" in client.get("/jobs/1").text
    with client.app.state.session_factory() as session:
        [profile] = session.query(FormatProfile).all()
        assert profile.name == "我的銀行"
    # 同樣格式再上傳一次，直接辨識
    assert upload_custom(client, settings).headers["location"] == "/jobs/2"
    assert client.get(location).status_code == 404  # 暫存檔已刪除


def test_mapping_requires_required_fields(client, settings):
    location = upload_custom(client, settings).headers["location"]
    response = client.post(location, data={"bank_date": "入帳日"})
    assert response.status_code == 400
    assert "請選擇「存入金額」對應的欄位" in response.text


def test_formats_page_and_delete(client, settings):
    location = upload_custom(client, settings).headers["location"]
    client.post(location, data={"bank_date": "入帳日", "bank_amount": "金額",
                                "bank_save_name": "我的銀行"})
    page = client.get("/formats").text
    assert "我的銀行" in page and "示範銀行 A" in page
    assert client.post("/formats/1/delete", follow_redirects=False).status_code == 303
    assert "我的銀行" not in client.get("/formats").text


def test_mapping_token_is_validated(client):
    assert client.get("/mapping/..%2F..%2Fetc").status_code == 404
    assert client.get("/mapping/unknown-token-1234567890").status_code == 404
```

- [ ] **Step 2: 執行確認失敗**

Run: `uv run pytest tests/test_web.py -v`
Expected: 新增的測試失敗（404 或 405）

- [ ] **Step 3: 修改 `reconcile/web/app.py`**

import 區：`from reconcile.formats import ...` 改為：

```python
from reconcile.formats import (
    BANK,
    BUILTIN_FORMATS,
    FIELDS,
    KIND_LABELS,
    ORDERS,
    Format,
    detect,
    validate_mapping,
)
```

並把 `from reconcile.models import Job, Link, init_db` 改為 `from reconcile.models import FormatProfile, Job, Link, init_db`。

在 `return app` 之前加入：

```python
    def load_pending_or_404(token: str):
        pending = uploads.load_pending(upload_dir, token)
        if pending is None:
            raise HTTPException(404)
        meta, files = pending
        return meta, files, {kind: read_table(*files[kind]) for kind in files}

    def render_mapping(request, session, token, meta, tables, errors=(), values=None,
                       status_code=200):
        sides = []
        for kind in (BANK, ORDERS):
            table = tables[kind]
            sides.append({
                "kind": kind,
                "label": KIND_LABELS[kind],
                "filename": meta["filenames"][kind],
                "detected": detect(table.header, kind, service.saved_formats(session, kind)),
                "columns": [column for column in table.header if column],
                "fields": FIELDS[kind],
                "header": table.header,
                "preview": table.rows[:3],
            })
        return templates.TemplateResponse(
            request, "mapping.html",
            {"token": token, "sides": sides, "errors": list(errors), "values": values or {}},
            status_code=status_code,
        )

    @app.get("/mapping/{token}", response_class=HTMLResponse)
    def mapping_page(request: Request, token: str):
        meta, _, tables = load_pending_or_404(token)
        with session_factory() as session:
            return render_mapping(request, session, token, meta, tables)

    @app.post("/mapping/{token}", response_class=HTMLResponse)
    async def submit_mapping(request: Request, token: str):
        meta, files, tables = load_pending_or_404(token)
        values = {key: str(value).strip() for key, value in (await request.form()).items()}
        with session_factory() as session:
            formats, errors, profiles = {}, [], []
            for kind in (BANK, ORDERS):
                fmt = detect(tables[kind].header, kind, service.saved_formats(session, kind))
                if fmt is None:
                    mapping = {
                        name: values[f"{kind}_{name}"]
                        for name in FIELDS[kind]
                        if values.get(f"{kind}_{name}")
                    }
                    kind_errors = validate_mapping(kind, mapping, tables[kind].header)
                    errors += [f"{KIND_LABELS[kind]}：{e}" for e in kind_errors]
                    name = values.get(f"{kind}_save_name", "")
                    fmt = Format(name or f"自訂{KIND_LABELS[kind]}格式", kind, mapping,
                                 builtin=False)
                    if name:
                        profiles.append(FormatProfile(kind=kind, name=name, mapping=mapping))
                formats[kind] = fmt
            if errors:
                return render_mapping(request, session, token, meta, tables, errors, values, 400)
            session.add_all(profiles)
            job = service.create_job(
                session, tables=tables, formats=formats, filenames=meta["filenames"],
                window_days=meta["window_days"],
            )
        uploads.remove_pending(upload_dir, token)
        return RedirectResponse(f"/jobs/{job.id}", status_code=303)

    @app.get("/formats", response_class=HTMLResponse)
    def formats_page(request: Request):
        with session_factory() as session:
            saved = session.query(FormatProfile).order_by(FormatProfile.id).all()
            return templates.TemplateResponse(
                request, "formats.html",
                {"builtin": BUILTIN_FORMATS, "saved": saved, "fields": FIELDS},
            )

    @app.post("/formats/{profile_id}/delete")
    def delete_format(profile_id: int):
        with session_factory() as session:
            profile = session.get(FormatProfile, profile_id)
            if profile is None:
                raise HTTPException(404)
            session.delete(profile)
            session.commit()
        return RedirectResponse("/formats", status_code=303)
```

- [ ] **Step 4: `reconcile/web/templates/mapping.html`**

```html
{% extends "base.html" %}
{% block title %}指定欄位｜銀行入帳對帳{% endblock %}
{% block content %}
<p><a href="/">← 取消</a></p>
<h2>指定欄位</h2>
<p class="muted">有檔案的格式無法自動辨識，請指定每個欄位對應哪一欄。填上名稱就會儲存這個格式，下次上傳同樣格式的檔案會自動辨識。</p>
{% for e in errors %}<p class="error">{{ e }}</p>{% endfor %}
<form method="post" action="/mapping/{{ token }}">
{% for side in sides %}
  <section class="card">
    <h3>{{ side.label }}：{{ side.filename }}</h3>
    {% if side.detected %}
      <p>已自動辨識為「{{ side.detected.name }}」。</p>
    {% else %}
      <div class="mapping">
      {% for name, spec in side.fields.items() %}
        {% set key = side.kind ~ "_" ~ name %}
        <label for="{{ key }}">{{ spec[0] }}{% if spec[1] %} *{% endif %}</label>
        <select id="{{ key }}" name="{{ key }}">
          <option value="">{{ "（請選擇）" if spec[1] else "（不使用）" }}</option>
          {% for column in side.columns %}
          <option value="{{ column }}" {{ 'selected' if values.get(key) == column }}>{{ column }}</option>
          {% endfor %}
        </select>
      {% endfor %}
        <label for="{{ side.kind }}_save_name">儲存為格式</label>
        <input id="{{ side.kind }}_save_name" name="{{ side.kind }}_save_name" placeholder="例如：第一銀行（可留空）" value="{{ values.get(side.kind ~ '_save_name', '') }}">
      </div>
      <p class="muted">前 3 列資料：</p>
      <table class="preview">
        <thead><tr><th>列號</th>{% for column in side.header %}<th>{{ column }}</th>{% endfor %}</tr></thead>
        <tbody>
        {% for row_no, row in side.preview %}
          <tr><td>{{ row_no }}</td>{% for value in row %}<td>{{ value if value is not none else "" }}</td>{% endfor %}</tr>
        {% endfor %}
        </tbody>
      </table>
    {% endif %}
  </section>
{% endfor %}
  <p><button>開始對帳</button></p>
</form>
{% endblock %}
```

- [ ] **Step 5: `reconcile/web/templates/formats.html`**

```html
{% extends "base.html" %}
{% block title %}格式設定｜銀行入帳對帳{% endblock %}
{% block content %}
<h2>格式設定</h2>
<p class="muted">上傳的檔案依表頭自動辨識格式；你建立的格式優先於內建格式。</p>
<table>
  <thead><tr><th>名稱</th><th>類型</th><th>欄位對應</th><th></th></tr></thead>
  <tbody>
  {% for fmt in saved %}
    <tr>
      <td>{{ fmt.name }}</td>
      <td>{{ KIND_LABELS[fmt.kind] }}</td>
      <td>{% for name, column in fmt.mapping.items() %}{{ fields[fmt.kind][name][0] }} ← {{ column }}{% if not loop.last %}、{% endif %}{% endfor %}</td>
      <td><form method="post" action="/formats/{{ fmt.id }}/delete" onsubmit="return confirm('確定刪除這個格式？')"><button class="danger">刪除</button></form></td>
    </tr>
  {% endfor %}
  {% for fmt in builtin %}
    <tr>
      <td>{{ fmt.name }}<span class="muted">（內建）</span></td>
      <td>{{ KIND_LABELS[fmt.kind] }}</td>
      <td>{% for name, column in fmt.mapping.items() %}{{ fields[fmt.kind][name][0] }} ← {{ column }}{% if not loop.last %}、{% endif %}{% endfor %}</td>
      <td></td>
    </tr>
  {% endfor %}
  </tbody>
</table>
{% endblock %}
```

- [ ] **Step 6: 執行確認通過**

Run: `uv run ruff format . && uv run ruff check . && uv run pytest`
Expected: 全部通過

- [ ] **Step 7: Commit**

```bash
git add reconcile/web tests/test_web.py
git commit -m "加入欄位對應與格式設定"
```

---

### Task 12: Docker、README、CLAUDE.md

**Files:**
- Create: `Dockerfile`, `docker-compose.yml`, `.env.example`, `README.md`, `CLAUDE.md`

- [ ] **Step 1: `Dockerfile`**

```dockerfile
FROM python:3.12-slim

COPY --from=ghcr.io/astral-sh/uv:0.11 /uv /bin/uv

WORKDIR /app
ENV UV_COMPILE_BYTECODE=1 UV_LINK_MODE=copy PATH="/app/.venv/bin:$PATH" PYTHONUNBUFFERED=1 TZ=Asia/Taipei

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

COPY reconcile ./reconcile
COPY samples ./samples

CMD ["uvicorn", "--factory", "reconcile.web.app:create_app", "--host", "0.0.0.0", "--port", "8000"]
```

- [ ] **Step 2: `docker-compose.yml`**

```yaml
services:
  bank-reconcile:
    build: .
    image: bank-reconcile
    ports:
      - "8000:8000"
    env_file:
      - path: .env
        required: false
    environment:
      DATABASE_URL: sqlite:////data/reconcile.db
      UPLOAD_DIR: /data/uploads
    volumes:
      - reconcile-data:/data

volumes:
  reconcile-data:
```

- [ ] **Step 3: `.env.example`**

```
# 對帳工作保留天數，到期自動刪除（含銀行資料）；0 表示不自動刪除
RETENTION_DAYS=90
```

- [ ] **Step 4: `README.md`（給客戶看）**

```markdown
# 銀行入帳對帳

> 逅斯蓋工作室（Code Sky）作品

## 解決什麼問題

採用銀行轉帳收款的店家，每週或每月都要把銀行入帳明細和訂單一筆筆對起來：哪些訂單收到錢了、哪些還沒收、哪些金額不對、哪筆入帳不知道是誰匯的。人工比對費時，也容易漏。

這套工具把兩份檔案上傳後，自動完成比對；不確定的配對交給你在畫面上確認，最後下載一份可以直接交給會計的 Excel 報表。

## 功能

- **直接用銀行與平台匯出的檔案**：支援 CSV（UTF-8、Big5）與 Excel；自動略過表頭上方的標題列；支援民國年日期。
- **自動辨識格式**：內建示範銀行與電商訂單格式；其他格式第一次指定欄位後即可儲存，下次自動辨識。
- **三層配對，每筆都附依據**
  1. 匯款備註含訂單編號 → 直接配對（金額不同時標為「金額不符」並列出差額）
  2. 轉帳帳號末五碼與金額相符 → 直接配對（多筆符合時交給你確認）
  3. 只有金額相符、且在訂單日後幾天內入帳 → 列為「待確認」
- **待確認項目一鍵確認或排除**：確認後，同一筆入帳或訂單的其他候選自動排除。
- **Excel 報表**：摘要、已配對、待確認、金額不符、未收款訂單、不明入帳、略過的資料列，並以顏色標示。
- **資料不外流**：所有處理在本機進行，不連任何外部服務；對帳工作到期自動刪除。

## 快速體驗

```bash
docker compose up --build
```

打開 http://localhost:8000 ，按「載入範例資料」，就能看到每一種配對結果；在「待確認」按確認或排除，再下載 Excel 報表。範例檔在 `samples/`，資料均為虛構。

## 截圖

（待補：上傳頁、對帳結果頁、Excel 報表）

## 使用的技術

Python 3.12、FastAPI、Jinja2、SQLAlchemy + SQLite、openpyxl、Docker Compose。

## 對應的服務方案

**報表自動化**：把每週、每月重複的 Excel 整理與比對工作自動化。銀行、電商平台、ERP 匯出的各種格式都能接。

## 第一版不處理的情況

- 一筆匯款付多張訂單、一張訂單分多次付款：會分別出現在「不明入帳」與「未收款」，由人工處理。
- 客戶自行扣手續費（例如少匯 15 元）：只有備註含訂單編號時，會以「金額不符」列出差額。
- 發票比對、直接串接銀行或電商平台 API。
```

- [ ] **Step 5: `CLAUDE.md`**

```markdown
# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

逅斯蓋工作室（Code Sky）作品集作品 2：銀行入帳對帳。設計依據是 [docs/design.md](docs/design.md)，偏離設計前先提出。README 以客戶角度撰寫。

## 指令

- `uv sync`：安裝套件（Python 3.12，`package = false`）
- `uv run pytest`；單一測試：`uv run pytest tests/test_matching.py::test_layer1_order_number_in_memo`
- `uv run ruff format . && uv run ruff check .`
- `docker compose up --build`：http://localhost:8000
- 本機單獨執行：`uv run uvicorn --factory reconcile.web.app:create_app --port 8000`
- 重新產生範例檔：`uv run python scripts/make_samples.py`（修改後要同步更新 `tests/test_samples.py` 的預期結果）

## 架構重點

- **資料流**：`readers.read_table`（CSV 依序試 UTF-8、cp950；前 10 列中第一個有 3 個以上非空白儲存格的列是表頭）→ `formats.detect`（使用者格式優先，表頭包含格式所需所有欄名即符合）→ `normalize.build_deposits/build_orders`（銀行只取存入 > 0 的列；解析失敗的列記為 Skipped）→ `matching.match`（純邏輯）→ `service.create_job` 存成 Job。
- **配對**（`matching.py`）：入帳依（日期、列號）排序；層級 ① 備註含訂單編號（正規化：大寫、去空白與連字號，取最長編號）；② 末五碼＋金額，多筆即候選且不再進 ③；③ 金額相同且入帳日在訂單日到 N 天後之間。用掉的訂單不再使用。
- **有效配對**：`matching.is_effective`；「未收款」與「不明入帳」是沒有有效配對也沒有 pending 候選者（`service.job_view`）。確認候選會自動排除同入帳、同訂單的其他 pending 候選。
- **欄位對應暫存**：`web/uploads.py` 以隨機 token 存在 `UPLOAD_DIR`，token 格式受驗證以防路徑穿越；1 小時後或完成對帳後刪除。
- **到期刪除**：`service.purge_expired` 在啟動與開啟首頁時執行。
- **時間**：`created_at` 存 naive UTC，顯示用 `localtime` filter 轉 Asia/Taipei；業務日期是 `datetime.date`。

## 測試

- 不需要網路；檔案用 `tests/conftest.py` 的 `csv_bytes`、`xlsx_bytes` 產生。
- `tests/test_samples.py` 鎖定範例資料的每一種結果；`tests/test_web.py` 的部分測試依賴範例資料的 Link id 順序。
```

- [ ] **Step 6: 確認 Docker 可以啟動**

Run: `docker compose up --build -d && sleep 5 && curl -s -o /dev/null -w '%{http_code}\n' http://localhost:8000/ && curl -s -o /dev/null -w '%{http_code}\n' -X POST http://localhost:8000/samples && docker compose down`
Expected: `200`、`303`。Docker 沒有在執行時跳過此步並在回報中說明。

- [ ] **Step 7: Commit**

```bash
git add Dockerfile docker-compose.yml .env.example README.md CLAUDE.md
git commit -m "加入 Docker、README 與 CLAUDE.md"
```

---

### Task 13: 整體驗證

- [ ] **Step 1:** `uv run ruff format --check . && uv run ruff check . && uv run pytest` 全部通過。
- [ ] **Step 2:** 本機啟動後，用 HTTP 請求走一遍：首頁 → 載入範例資料 → 結果頁 → 確認一筆、排除一筆 → 下載報表並用 openpyxl 開啟確認工作表；上傳一個自訂格式的 CSV → 欄位對應 → 建立對帳工作。
- [ ] **Step 3:** 回報完成功能、測試數量、未完成項目（README 截圖）與可調整方向。

## Self-Review 紀錄

- 設計各節對應：讀檔（Task 3）、格式與辨識（Task 4、11）、正規化與民國年（Task 5）、三層配對與決定性（Task 6）、有效配對定義、確認／排除、摘要、到期刪除（Task 7）、Excel 報表（Task 8）、範例資料（Task 9）、網頁流程與暫存檔（Task 10、11）、錯誤處理（Task 3、5、10、11）、設定與部署（Task 2、12）。
- 名稱一致：`read_table`、`Table.column`、`Format(name, kind, mapping, builtin, id)`、`detect(header, kind, saved)`、`build_deposits/build_orders`、`match(deposits, orders, window_days)`、`LinkResult`、`service.create_job(session, *, tables, formats, filenames, window_days)`、`job_view(job).summary`。
