import os
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class Settings:
    database_url: str = "sqlite:///data/reconcile.db"
    # 對帳工作保留天數；0 表示不自動刪除
    retention_days: int = 90
    upload_dir: Path = Path("data/uploads")
    samples_dir: Path = ROOT / "samples"

    @classmethod
    def from_env(cls) -> "Settings":
        env = os.environ.get
        return cls(
            database_url=env("DATABASE_URL") or cls.database_url,
            retention_days=int(env("RETENTION_DAYS") or 90),
            upload_dir=Path(env("UPLOAD_DIR") or cls.upload_dir),
        )
