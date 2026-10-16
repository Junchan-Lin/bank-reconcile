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
