import io

from openpyxl import load_workbook

from reconcile import service
from reconcile.report import SHEETS, build_report
from tests.test_service import make_job


def test_report_sheets_and_rows(session_factory):
    with session_factory() as session:
        view = service.job_view(make_job(session))
        workbook = load_workbook(io.BytesIO(build_report(view)))
    assert workbook.sheetnames == list(SHEETS)
    assert workbook["摘要"]["A2"].value == "銀行入帳筆數"
    assert workbook["已配對"].max_row == 2  # 標題＋1 筆
    assert workbook["待確認"].max_row == 3  # 標題＋2 筆候選
    assert workbook["未收款訂單"]["A2"].value == "C1001"
    assert workbook["不明入帳"]["C2"].value == 5000
    assert workbook["不明入帳"]["C2"].number_format == "#,##0"
    assert workbook["略過的資料列"]["B2"].value == 5
    assert workbook["已配對"].freeze_panes == "A2"
