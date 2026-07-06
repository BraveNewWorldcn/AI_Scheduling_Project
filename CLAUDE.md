# AI Scheduling Engine (AI排单调度引擎)

Sales order scheduling engine with Feishu Bitable integration. Reads orders, SKU data, and inventory, computes availability, generates schedules, and pushes results back to Feishu. Includes finance reconciliation, SF Express logistics tracking, and AI-powered daily report bots.

## Architecture

```
run.py                  → Entry point: starts scheduler_api:app via uvicorn on port 8000
scheduler_api.py        → Main FastAPI app: scheduling engine, finance module, webhook handler
customer_agent.py       → Standalone dual-bot daemon (order query + supply chain AI)
ai_daily_agent.py       → Standalone daily report generator (DeepSeek-powered)
import_orders.py        → Excel importer for OA orders and inventory snapshots
sf_shipping.py          → SF Express logistics tracking, mounted as FastAPI routes
shared.py               → Column name mappings and audit function (single source of truth)
ai_service.py           → DeepSeek API wrapper for order risk analysis
audit_rules.py          → Standalone finance rule auditor
verify_patch.py         → Local verification script with 12 test cases for finance module
run_production_update.py→ Production write-back workflow for finance tables
test_generate_data.py   → Test data generator for Feishu Bitable
```

## How to run

```bash
# 1. Copy .env.example to .env and fill in all values
cp .env.example .env

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the main API service
python run.py
# Or: uvicorn scheduler_api:app --host 0.0.0.0 --port 8000

# 4. (Optional) Start the customer bot in a separate terminal
python customer_agent.py

# 5. (Optional) Run the daily report bot
python ai_daily_agent.py
```

`customer_agent.py` requires `lark-cli` (not on PyPI, must be installed separately).

## Key endpoints (scheduler_api.py)

- `POST /schedule` — Run scheduling engine
- `POST /import` — Import Excel orders/inventory
- `POST /finance/sync` / `POST /finance/calculate` — Finance reconciliation
- `POST /tracking/import` / `GET /tracking/status` — SF Express logistics
- `POST /webhook/feishu` — Feishu event webhook
- `GET /daily-report` — Get daily report

## Configuration (.env)

All config via environment variables. Required: `FEISHU_APP_ID`, `FEISHU_APP_SECRET`, `BITABLE_APP_TOKEN`, table IDs for orders/SKU/inventory/detail/reservation. Optional: `DEEPSEEK_API_KEY`, SF Express credentials, finance table IDs. See `.env.example` for full list.

## Key patterns

- **No hardcoded credentials**: All secrets via `os.getenv()`. Never commit real values.
- **shared.py is canonical**: Column name mappings (`OA_COLUMN_MAP`, `MAIN_FIELD_MAP`, etc.) and `audit_schedule_results()` live here. Both `scheduler_api.py` and `import_orders.py` import from it.
- **Lazy imports for cross-dependencies**: `scheduler_api` and `customer_agent` import each other — both use lazy imports inside functions to avoid circular dependency at startup.
- **Global scheduling lock**: File-based lock at `cache/schedule_global.lock` prevents concurrent scheduling runs.
- **Cache directory**: `cache/` is gitignored. Used for schedule output CSVs, SQLite DB, pickle files.
- **Line endings**: `.gitattributes` enforces LF for Python/Markdown/JSON, CRLF for Batch/PowerShell. Use `git blame --ignore-revs-file .git-blame-ignore-revs` to skip the normalization commit.

## Feishu Bitable tables

Data stored in Feishu Bitable (accessed via `BITABLE_APP_TOKEN`):
- **销售订单主表** — Contract/order master records
- **销售订单明细表** — Line items per contract
- **SKU标准表** — Product SKU master with production cycles
- **库存快照表** — Inventory snapshots
- **AI排单总表** — Scheduling results (summary)
- **AI排单明细表** — Scheduling results (detail)
- **财务汇总表** / **财务明细表** / **计费规则表** — Finance module
- **AI排单日报表** — Daily report data
- **发货总表** — SF Express shipment records

Full table structure documented in `docs/飞书多维表格结构总览.md`.

## Git history context

Project was merged from Mac and Windows codebases. Recent cleanup (2026-07-06) removed hardcoded credentials, debug prints, untracked build artifacts, and normalized line endings. See `docs/superpowers/specs/` for cleanup plans.
