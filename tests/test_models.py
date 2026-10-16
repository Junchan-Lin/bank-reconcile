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
        job = Job(
            bank_filename="b.csv",
            orders_filename="o.xlsx",
            bank_format="A",
            orders_format="O",
            window_days=7,
        )
        deposit = Deposit(row_no=5, date=dt.date(2026, 10, 1), amount=1200, memo="", last5=None)
        order = Order(
            row_no=2, order_no="A1", date=dt.date(2026, 10, 1), amount=1200, customer="", last5=None
        )
        job.deposits.append(deposit)
        job.orders.append(order)
        job.links.append(
            Link(deposit=deposit, order=order, kind="matched", layer=1, reason="x", status="auto")
        )
        job.skipped.append(SkippedRow(source="bank", row_no=9, reason="y"))
        session.add(job)
        session.commit()
        assert job.created_at is not None
        session.delete(job)
        session.commit()
        for model in (Deposit, Order, Link, SkippedRow):
            assert session.scalars(select(model)).all() == []
