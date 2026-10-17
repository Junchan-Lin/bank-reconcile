# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

逅斯蓋工作室（Code Sky）作品集作品 2：銀行入帳對帳。設計依據是 [docs/design.md](docs/design.md)，偏離設計前先提出。README 以客戶角度撰寫。

## 指令

- `uv sync`：安裝套件（Python 3.12，`package = false`）
- `uv run pytest`；單一測試：`uv run pytest tests/test_matching.py::test_layer1_order_number_in_memo`
- `uv run ruff format . && uv run ruff check .`
- `docker compose up --build`：http://localhost:8000
- 本機單獨執行：`uv run uvicorn --factory reconcile.web.app:create_app --port 8000`
- 重新產生範例檔：`uv run python scripts/make_samples.py`（修改後要同步更新 `tests/test_samples.py` 的預期結果）

## 架構重點

- **資料流**：`readers.read_table`（CSV 依序試 UTF-8、cp950；前 10 列中第一個有 3 個以上非空白儲存格的列是表頭）→ `formats.detect`（使用者格式優先，表頭包含格式所需所有欄名即符合）→ `normalize.build_deposits/build_orders`（銀行只取存入 > 0 的列；解析失敗的列記為 Skipped）→ `matching.match`（純邏輯）→ `service.create_job` 存成 Job。
- **配對**（`matching.py`）：入帳依（日期、列號）排序。層級 ① 以切詞序列比對：備註與訂單編號共用 `_tokens`（NFKC、大寫、以非英數與連字號切開、去連字號），編號切詞後超過 `MAX_KEY_TOKENS`（4）段者不參與。編號分 strong（英數混合、4 碼以上）、numeric（純數字、4 碼以上）、short（其餘）；numeric 等於入帳金額或落在 1900–2099 者降為 short。strong 命中即配對（金額不同為 mismatch）；numeric 需金額相同才配對；只命中 short 的只產生待確認候選（金額相同才列），且該入帳繼續進 ②。多筆命中產生候選；命中的 strong 訂單都已被用掉時，該入帳不再進 ②③（成為不明入帳）。② 末五碼＋金額，多筆即候選且不再進 ③。③ 金額相同且入帳日在訂單日到 N 天後之間，符合者超過 `MAX_CANDIDATES`（5）筆時不列候選。用掉的訂單不再使用，最後以已用訂單為對象的候選一律移除。
- **有效配對**：`matching.is_effective`；「未收款」與「不明入帳」是沒有有效配對也沒有 pending 候選者（`service.job_view`）。確認候選會自動排除同入帳、同訂單的其他 pending 候選。
- **欄位對應暫存**：`web/uploads.py` 以隨機 token 存在 `UPLOAD_DIR`，token 格式受驗證以防路徑穿越；1 小時後或完成對帳後刪除。
- **非同步**：上傳與欄位對應的重工作（讀檔、配對、寫入資料庫）以 `run_in_threadpool` 執行，避免卡住 event loop。
- **報表安全**：`report.py` 將 `data_type` 為 `"f"` 的儲存格強制改成字串，避免使用者資料成為公式（公式注入）。
- **欄位對應驗證**：`formats.validate_mapping` 會拒絕同一欄被對應到多個欄位。
- **到期刪除**：`service.purge_expired` 在啟動與開啟首頁時執行。
- **時間**：`created_at` 存 naive UTC，顯示用 `localtime` filter 轉 Asia/Taipei；業務日期是 `datetime.date`。

## 測試

- 不需要網路；檔案用 `tests/conftest.py` 的 `csv_bytes`、`xlsx_bytes` 產生。
- `tests/test_samples.py` 鎖定範例資料的每一種結果；`tests/test_web.py` 的部分測試依賴範例資料的 Link id 順序。
