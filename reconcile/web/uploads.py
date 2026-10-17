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
