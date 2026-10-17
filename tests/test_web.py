import pytest
from fastapi.testclient import TestClient
from sqlalchemy import select

from reconcile.config import Settings
from reconcile.models import Link
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
    response = client.post(
        "/jobs", files=sample_files(settings), data={"window_days": "7"}, follow_redirects=False
    )
    assert response.status_code == 303
    assert response.headers["location"] == "/jobs/1"
    page = client.get("/jobs/1").text
    assert "A1003" in page and "待確認" in page and "日期不存在" in page


def test_upload_rejects_unreadable_file(client):
    response = client.post(
        "/jobs", files={"bank_file": ("x.txt", b"x"), "orders_file": ("o.txt", b"x")}
    )
    assert response.status_code == 400
    assert "銀行明細：只支援" in response.text


def test_upload_rejects_bad_window(client, settings):
    response = client.post("/jobs", files=sample_files(settings), data={"window_days": "99"})
    assert response.status_code == 400


def test_samples_confirm_and_report(client):
    response = client.post("/samples", follow_redirects=False)
    assert response.headers["location"] == "/jobs/1"
    with client.app.state.session_factory() as session:
        pending = session.scalars(select(Link).where(Link.status == "pending")).all()
        ids = {link.order.order_no: link.id for link in pending}
    response = client.post(f"/links/{ids['A1007']}/confirm", follow_redirects=False)
    assert response.status_code == 303
    assert response.headers["location"] == "/jobs/1#pending"
    response = client.post(f"/links/{ids['A1004']}/reject", follow_redirects=False)
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


def test_upload_work_runs_off_event_loop(client, settings, monkeypatch):
    import asyncio
    import threading

    from reconcile import service

    seen = {}
    original = service.create_job

    def spy(*args, **kwargs):
        seen["main"] = threading.current_thread() is threading.main_thread()
        try:
            asyncio.get_running_loop()
            seen["loop"] = True
        except RuntimeError:
            seen["loop"] = False
        return original(*args, **kwargs)

    monkeypatch.setattr(service, "create_job", spy)
    client.post("/jobs", files=sample_files(settings), follow_redirects=False)
    assert seen == {"main": False, "loop": False}
