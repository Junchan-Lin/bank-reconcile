import io

import pytest
from openpyxl import Workbook

from reconcile.models import init_db


@pytest.fixture
def session_factory(tmp_path):
    return init_db(f"sqlite:///{tmp_path / 'test.db'}")


def csv_bytes(lines: list[str], encoding: str = "utf-8") -> bytes:
    return ("\n".join(lines) + "\n").encode(encoding)


def xlsx_bytes(rows: list[list]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
