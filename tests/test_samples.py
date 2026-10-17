from reconcile import service
from reconcile.config import Settings
from reconcile.formats import BANK, ORDERS, detect
from reconcile.readers import read_table
from scripts.make_samples import BANK_A, BANK_B
from scripts.make_samples import ORDERS as ORDERS_FILE

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
        job = service.create_job(
            session,
            tables=tables,
            formats=formats,
            filenames={BANK: BANK_A, ORDERS: ORDERS_FILE},
            window_days=7,
        )
        view = service.job_view(job)
        assert [link.order.order_no for link in view.matched] == ["A1001", "A1002", "A1012"]
        assert [link.order.order_no for link in view.mismatched] == ["A1003"]
        assert [link.order.order_no for link in view.pending] == [
            "A1004",
            "A1005",
            "A1007",
            "A1008",
            "A1006",
            "A1011",
        ]
        assert [o.order_no for o in view.unpaid] == ["A1009", "A1010"]
        assert [d.amount for d in view.unknown] == [5000]
        assert [s.reason for s in view.skipped] == ["日期不存在「115/13/40」"]
