FROM python:3.12-slim

COPY --from=ghcr.io/astral-sh/uv:0.11 /uv /bin/uv

WORKDIR /app
ENV UV_COMPILE_BYTECODE=1 UV_LINK_MODE=copy PATH="/app/.venv/bin:$PATH" PYTHONUNBUFFERED=1 TZ=Asia/Taipei

COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

COPY reconcile ./reconcile
COPY samples ./samples

CMD ["uvicorn", "--factory", "reconcile.web.app:create_app", "--host", "0.0.0.0", "--port", "8000"]
