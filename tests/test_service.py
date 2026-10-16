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
        [
            (2, ["2026/10/01", "1200", "A1001", ""]),  # ① 已配對
            (3, ["2026/10/02", "990", "", ""]),  # ③ 兩筆候選（B1001、B1002）
            (4, ["2026/10/03", "5000", "退款", ""]),  # 不明入帳
            (5, ["bad", "100", "", ""]),
        ],  # 略過
    )
    orders = Table(
        ["編號", "日期", "金額"],
        [
            (2, ["A1001", "2026/10/01", "1200"]),
            (3, ["B1001", "2026/10/01", "990"]),
            (4, ["B1002", "2026/10/02", "990"]),
            (5, ["C1001", "2026/10/02", "780"]),
        ],
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
        assert [link.order.order_no for link in view.matched] == ["A1001"]
        assert [link.order.order_no for link in view.pending] == ["B1001", "B1002"]
        assert [o.order_no for o in view.unpaid] == ["C1001"]
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
        assert [link.order.order_no for link in view.matched] == ["A1001", "B1001"]
        assert view.pending == []
        assert [o.order_no for o in view.unpaid] == ["B1002", "C1001"]


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
