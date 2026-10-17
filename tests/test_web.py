import pytest
from fastapi.testclient import TestClient
from sqlalchemy import select

from reconcile.config import Settings
from reconcile.models import FormatProfile, Link
from reconcile.web.app import create_app
from scripts.make_samples import BANK_A, ORDERS
from tests.conftest import csv_bytes


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
    response = client.post(
        location,
        data={
            "bank_date": "入帳日",
            "bank_amount": "金額",
            "bank_memo": "說明",
            "bank_account": "帳號",
            "bank_save_name": "我的銀行",
        },
        follow_redirects=False,
    )
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
    client.post(
        location,
        data={"bank_date": "入帳日", "bank_amount": "金額", "bank_save_name": "我的銀行"},
    )
    page = client.get("/formats").text
    assert "我的銀行" in page and "示範銀行 A" in page
    assert client.post("/formats/1/delete", follow_redirects=False).status_code == 303
    assert "我的銀行" not in client.get("/formats").text


def test_mapping_token_is_validated(client):
    assert client.get("/mapping/..%2F..%2Fetc").status_code == 404
    assert client.get("/mapping/unknown-token-1234567890").status_code == 404


def test_mapping_rejects_same_column_for_two_fields(client, settings):
    location = upload_custom(client, settings).headers["location"]
    response = client.post(
        location,
        data={"bank_date": "金額", "bank_amount": "金額", "bank_save_name": "壞格式"},
    )
    assert response.status_code == 400
    assert "「金額」不能同時對應多個欄位" in response.text
    with client.app.state.session_factory() as session:
        assert session.query(FormatProfile).count() == 0


def test_index_retention_zero_text(tmp_path):
    settings = Settings(
        database_url=f"sqlite:///{tmp_path / 'r.db'}", upload_dir=tmp_path / "up", retention_days=0
    )
    with TestClient(create_app(settings)) as client:
        text = client.get("/").text
    assert "對帳工作不會自動刪除。" in text
    assert "保留 0 天" not in text


def test_job_page_shows_order_total(client, settings):
    response = client.post("/samples", follow_redirects=True)
    assert "訂單（" in response.text
    assert '<p class="actions">' not in response.text
