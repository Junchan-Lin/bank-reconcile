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
        Order(
            row_no=o.row_no,
            order_no=o.order_no,
            date=o.date,
            amount=o.amount,
            customer=o.customer,
            last5=o.last5,
        )
        for o in orders
    ]
    job.skipped = [SkippedRow(source=BANK, row_no=s.row_no, reason=s.reason) for s in bank_skipped]
    job.skipped += [
        SkippedRow(source=ORDERS, row_no=s.row_no, reason=s.reason) for s in order_skipped
    ]
    for r in matching.match(deposits, orders, window_days):
        job.links.append(
            Link(
                deposit=job.deposits[r.deposit_index],
                order=job.orders[r.order_index],
                kind=r.kind,
                layer=r.layer,
                reason=r.reason,
                status=r.status,
            )
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
