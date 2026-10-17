import datetime as dt
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Annotated
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
            request,
            "index.html",
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
                session,
                tables=tables,
                formats=formats,
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
        bank_file: Annotated[UploadFile, File()],
        orders_file: Annotated[UploadFile, File()],
        window_days: Annotated[int, Form()] = 7,
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
            f"attachment; filename=\"report_{job_id}.xlsx\"; filename*=UTF-8''{quote(filename)}"
        )
        return Response(data, media_type=XLSX, headers={"Content-Disposition": disposition})

    @app.post("/jobs/{job_id}/delete")
    def delete_job(job_id: int):
        with session_factory() as session:
            session.delete(load_job(session, job_id))
            session.commit()
        return RedirectResponse("/", status_code=303)

    return app
