from __future__ import annotations

import sys
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

from fastapi import FastAPI, Request
import os
import re
import time
import threading
from contextlib import contextmanager
from functools import wraps
from pydantic import BaseModel
from datetime import datetime, timedelta, date, timezone
import json
from typing import Any, Dict, List, Optional, Tuple, Set, TYPE_CHECKING
from collections import defaultdict

# 懒加载 pandas（Windows 上 import pandas 需 3-5 秒，推迟到首次调用 API 时加载）
if TYPE_CHECKING:
    import pandas as pd
else:
    class _LazyPandas:
        _pd = None
        def __getattr__(self, name):
            if self._pd is None:
                import pandas as _mod
                self.__class__._pd = _mod
            return getattr(self._pd, name)

    pd = _LazyPandas()

# =========================
# 本地缓存目录
# =========================
CACHE_DIR = "cache"
os.makedirs(CACHE_DIR, exist_ok=True)
import httpx

# =========================
# 飞书字段清洗函数
# =========================
def parse_feishu_field(value):
    """清洗飞书字段：优先提取富文本包装中的纯文本，去空格，空值转空字符串"""
    if isinstance(value, list):
        texts = []
        for item in value:
            if isinstance(item, dict) and "text" in item:
                texts.append(item["text"].strip())
            else:
                texts.append(str(item).strip())
        if len(texts) == 1:
            return texts[0]
        return ", ".join(texts)
    elif isinstance(value, dict):
        for k in ("text", "number", "phone", "email", "url"):
            if k in value:
                return str(value[k]).strip()
        return str(value).strip()
    elif isinstance(value, (int, float)):
        return value
    elif isinstance(value, str):
        return value.strip()
    elif value is None:
        return ""
    else:
        return str(value).strip()

# ----- 自动加载 .env 文件 -----
def _load_dotenv():
    """从项目目录的 .env 文件加载环境变量（不覆盖已设置的系统环境变量）。"""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value

_load_dotenv()

# ========== 飞书配置 ==========
# 所有部署相关配置一律从环境变量读取，值定义在项目 .env 文件中。
# 代码中不保留任何真实值作为默认值。
FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET")
BITABLE_APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")

if not FEISHU_APP_SECRET:
    raise ValueError("请先设置环境变量 FEISHU_APP_SECRET（或在项目 .env 文件中配置）")

# ===== 数据表 ID 配置 =====
TABLE_ID_ITEMS = os.getenv("TABLE_ID_ITEMS", "")
TABLE_ID_SKU   = os.getenv("TABLE_ID_SKU", "")
TABLE_ID_INV   = os.getenv("TABLE_ID_INV", "")
TABLE_ID_DETAIL = os.getenv("TABLE_ID_DETAIL", "")
TABLE_ID_RESERVATION = os.getenv("TABLE_ID_RESERVATION", "")
TABLE_ID_MAIN = os.getenv("TABLE_ID_MAIN", "")
TABLE_ID_SHIPPING = os.getenv("TABLE_ID_SHIPPING", "")
TABLE_ID_DAILY_REPORT = os.getenv("TABLE_ID_DAILY_REPORT", "")

# ===== 输出表字段映射 =====
OUTPUT_SUMMARY_FIELDS_MAP = {
    "合同编号": "合同编号",
    "项目类型": "项目类型",
    "订单SKU总数": "订单SKU总数",
    "订单总数量": "订单总数量",
    "缺货SKU数": "缺货SKU数",
    "缺货SKU列表": "缺货SKU列表",
    "整体状态": "整体状态",
    "AI建议发货时间": "AI建议发货时间",
    "AI风险": "AI风险",
    "AI建议": "AI建议",
    "排单批次号": "排单批次号",
    "订单状态": "订单状态",
    "是否人工确认": "是否人工确认",
    "人工确认发货时间": "人工确认发货时间"
}

SUMMARY_NUMERIC_COLS_DEFAULT = {"订单SKU总数", "订单总数量", "缺货SKU数"}
SUMMARY_DATE_COLS_DEFAULT = {"AI建议发货时间", "人工确认发货时间"}
RESERVATION_NUMERIC_COLS_DEFAULT = {"预留数量"}
RESERVATION_DATE_COLS_DEFAULT = {"创建时间"}
SCHEDULE_LOCK_PATH = os.path.join(CACHE_DIR, "schedule_global.lock")
app = FastAPI(docs_url=None, redoc_url=None, title="AI排单服务")

# =========================
# Pydantic 模型定义
# =========================
class ScheduleRequest(BaseModel):
    trigger: str = "test"

# =========================
# 线程池（并行读表 + 异步上下文兼容）
# =========================
_io_executor = None  # lazy init

def _get_io_executor():
    global _io_executor
    if _io_executor is None:
        from concurrent.futures import ThreadPoolExecutor
        _io_executor = ThreadPoolExecutor(max_workers=6, thread_name_prefix="feishu_io")
    return _io_executor

# =========================
# ⭐ Token 全局缓存（提前5分钟刷新）
# =========================
_token_cache: Dict[str, Any] = {
    "token": None,
    "expires_at": 0.0,  # Unix timestamp
}
_token_lock = threading.Lock()

FEISHU_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
FEISHU_BITABLE_BASE = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}"


def get_access_token() -> str:
    """获取飞书 tenant_access_token，带全局缓存，提前 5 分钟刷新。"""
    now = time.time()
    with _token_lock:
        if _token_cache["token"] is not None and now < _token_cache["expires_at"] - 300:
            return _token_cache["token"]

        resp = httpx.post(FEISHU_TOKEN_URL, json={
            "app_id": FEISHU_APP_ID,
            "app_secret": FEISHU_APP_SECRET
        }, timeout=30.0)
        data = _safe_http_json(resp, "获取飞书Token")
        if data.get("code") != 0:
            raise Exception(f"获取飞书token失败: {data}")

        _token_cache["token"] = data["tenant_access_token"]
        _token_cache["expires_at"] = now + data.get("expire", 7200)
        return _token_cache["token"]


def _safe_http_json(resp, context: str = "") -> dict:
    """安全解析 HTTP 响应 JSON。非 JSON 响应（如 HTML 错误页）会给出可读的错误信息。"""
    content_type = resp.headers.get("content-type", "")
    if "application/json" not in content_type:
        preview = resp.text[:500] if resp.text else "(empty body)"
        raise Exception(
            f"{context} 飞书接口返回非JSON响应 "
            f"(HTTP {resp.status_code}, Content-Type={content_type}): {preview}"
        )
    try:
        return resp.json()
    except Exception as e:
        preview = resp.text[:500] if resp.text else "(empty body)"
        raise Exception(
            f"{context} JSON解析失败 (HTTP {resp.status_code}): {preview}"
        ) from e


def _feishu_headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json; charset=utf-8"
    }


# =========================
# 飞书分页读取（内部函数，接受 token 和 headers）
# =========================
def _fetch_table_records(table_id: str) -> List[Dict[str, Any]]:
    """读取一张飞书多维表格的全部记录，返回 dict 列表（含 _record_id）。"""
    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records"
    all_records: List[Dict[str, Any]] = []
    page_token: Optional[str] = None

    while True:
        params: Dict[str, Any] = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token

        resp = httpx.get(url, headers=headers, params=params, timeout=30.0)
        data = _safe_http_json(resp, "飞书读取表格")
        if data.get("code") != 0:
            raise Exception(f"飞书读取表格失败: {data}")

        items = data.get("data", {}).get("items", [])
        for item in items:
            row = {"_record_id": item.get("record_id", "")}
            for field_name, value in item.get("fields", {}).items():
                row[field_name] = parse_feishu_field(value)
            all_records.append(row)

        if not data.get("data", {}).get("has_more", False):
            break
        page_token = data["data"].get("page_token")
        if not page_token:
            break

    return all_records


def _records_to_df(records: List[Dict[str, Any]]) -> pd.DataFrame:
    """将记录列表转为 DataFrame，统一清洗列名和字符串列。"""
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                df[col] = df[col].astype(str).str.strip()
            except Exception:
                pass
    return df


def fetch_bitable_to_df(table_id: str) -> pd.DataFrame:
    """从飞书读取表格，返回清洗后的 DataFrame（缓存 token 复用）。"""
    records = _fetch_table_records(table_id)
    df = _records_to_df(records)
    print(f"[DEBUG] 读取表 {table_id} 的列名: {list(df.columns)}")
    return df


def _fetch_bitable_raw(table_id: str) -> List[Dict[str, Any]]:
    """读取表格原始记录（用于并行读取后统一构建 DataFrame）。"""
    return _fetch_table_records(table_id)


# =========================
# ⭐ 库存预建索引
# =========================
_inv_exact_index: Dict[str, List[Dict[str, Any]]] = {}
_inv_inverted_index: Dict[str, List[Dict[str, Any]]] = {}
_inv_max_results = 5


def _norm_match_key(s: str) -> str:
    """SKU 宽松匹配键：去空白、全角括号统一、小写。"""
    t = safe_str(s).replace("　", " ").replace(" ", "")
    t = t.replace("（", "(").replace("）", ")")
    return t.lower()


def _tokenize_device_text(text: str) -> List[str]:
    """将设备型号/名称拆分为关键词（≥2字符），用于倒排索引。"""
    tokens: List[str] = []
    raw = safe_str(text)
    if not raw:
        return tokens
    parts = re.split(r"[/\-\s,　，、]+", raw)
    for part in parts:
        part = part.strip()
        if len(part) >= 2:
            tokens.append(part.lower())
            norm = part.replace("（", "(").replace("）", ")").lower()
            if norm != part.lower():
                tokens.append(norm)
    return tokens


def build_inventory_index(inv: pd.DataFrame):
    """加载库存表后立即构建索引。

    - 精确索引：以 _norm_match_key(SKU) 为 key → O(1) 查找
    - 倒排索引：以设备型号/名称拆分的关键词 → 倒排索引（限制返回最多5条）
    """
    global _inv_exact_index, _inv_inverted_index
    _inv_exact_index.clear()
    _inv_inverted_index.clear()

    if inv is None or inv.empty:
        return

    inv_records = inv.to_dict(orient="records")

    for rec in inv_records:
        sku = safe_str(rec.get("SKU", ""))
        if sku:
            key = _norm_match_key(sku)
            _inv_exact_index.setdefault(key, []).append(rec)

    # 倒排索引：对设备名称/型号列建倒排
    for rec in inv_records:
        tokens: Set[str] = set()
        for col in ("国网设备型号", "国网设备名称", "设备型号", "设备名称"):
            text = safe_str(rec.get(col, ""))
            for tok in _tokenize_device_text(text):
                tokens.add(tok)
        for tok in tokens:
            entry = _inv_inverted_index.setdefault(tok, [])
            if len(entry) < _inv_max_results:
                entry.append(rec)

    # 倒排索引按库存数量降序排序（每条查询只取最多5条）
    for tok in _inv_inverted_index:
        _inv_inverted_index[tok] = sorted(
            _inv_inverted_index[tok],
            key=lambda r: to_num(r.get("库存数量", 0)),
            reverse=True,
        )[: _inv_max_results]

    print(f"[索引] 精确索引 SKU 数: {len(_inv_exact_index)}, 倒排索引词条数: {len(_inv_inverted_index)}")


def _inv_lookup_exact(sku_code: str) -> Optional[Dict[str, Any]]:
    """O(1) 精确索引查找。"""
    if not sku_code:
        return None
    key = _norm_match_key(sku_code)
    matches = _inv_exact_index.get(key, [])
    if not matches:
        return None
    # 返回库存数量最大的一条
    return max(matches, key=lambda r: to_num(r.get("库存数量", 0)), default=None)


def _inv_lookup_by_tokens(text: str) -> List[Dict[str, Any]]:
    """通过倒排索引查找（限制返回最多5条），已按库存数量降序。"""
    tokens = _tokenize_device_text(text)
    if not tokens:
        return []
    candidates: Dict[int, List[Dict[str, Any]]] = {}  # id -> [rec, score]
    for tok in tokens:
        for rec in _inv_inverted_index.get(tok, []):
            rid = rec.get("_record_id", id(rec))
            if rid not in candidates:
                candidates[rid] = [rec, 0]
            candidates[rid][1] += 1

    # 按匹配得分降序，同分按库存数量降序
    sorted_candidates = sorted(
        candidates.values(),
        key=lambda x: (x[1], to_num(x[0].get("库存数量", 0))),
        reverse=True,
    )
    return [c[0] for c in sorted_candidates[: _inv_max_results]]


def _is_generic_spec_text(t: str) -> bool:
    """规格里常见「定制/外购」等无具体型号信息。"""
    s = safe_str(t).strip().replace(" ", "").replace("　", "")
    if len(s) < 6:
        return True
    generic_tokens = ("定制/外购", "定制外购", "定制", "外购", "外协")
    if s in generic_tokens:
        return True
    if all(ch in "定制外购外协/\\.-" for ch in s) and len(s) <= 12:
        return True
    return False


def find_inventory_row(
    sku_code: str,
    row_dict: Dict[str, Any],
    inv: pd.DataFrame,
    sku_df: Optional[pd.DataFrame] = None,
) -> Tuple[Optional[Dict[str, Any]], str]:
    """在库存快照中查找对应行（优先走预建索引，O(1)/O(k)）。

    1) 先按 SKU 精确索引查找
    2) 通过 SKU标准表拿到设备型号/名称 → 倒排索引查找
    3) 按明细的规格/产品名称 → 倒排索引查找
    """
    if inv is None or inv.empty:
        return None, "库存表为空"

    code = safe_str(sku_code)

    # 1) 精确索引 O(1)
    hit = _inv_lookup_exact(code)
    if hit is not None:
        return hit, "SKU列匹配(索引)"

    # 2) 经 SKU 主数据拿到型号/名称再匹配
    if sku_df is not None and not sku_df.empty and "产品编码SKU" in sku_df.columns and code:
        sm = sku_df[sku_df["产品编码SKU"].astype(str).str.strip() == code]
        if sm.empty:
            sm = sku_df[sku_df["产品编码SKU"].astype(str).apply(_norm_match_key) == _norm_match_key(code)]
        if not sm.empty:
            sku_row = sm.iloc[0]
            for key in ("设备型号", "设备名称"):
                t = safe_str(sku_row.get(key, ""))
                if len(t) < 3:
                    continue
                candidates = _inv_lookup_by_tokens(t)
                for c in candidates:
                    return c, f"SKU标准表.{key}→倒排索引({t[:40]})"

    # 3) 按明细规格/产品名称 → 倒排索引
    texts: List[str] = []
    for k in ("规格", "产品名称"):
        t = safe_str(row_dict.get(k, ""))
        if len(t) >= 2 and not _is_generic_spec_text(t):
            texts.append(t)
    if code and all(t != code for t in texts):
        texts.insert(0, code)

    for text in texts:
        candidates = _inv_lookup_by_tokens(text)
        if candidates:
            return candidates[0], f"倒排索引匹配({text[:40]})"

    return None, f"未匹配(SKU编码={code!r}, 规格={safe_str(row_dict.get('规格', ''))[:40]!r})"


# =========================
# 并行加载三张核心表（ThreadPoolExecutor，兼容 sync/async）
# =========================
def load_feishu_data() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """并行拉取三张核心表（3线程同时请求飞书），然后构建索引。

    使用 ThreadPoolExecutor 而非 asyncio，确保在同步和异步上下文中均可用。
    """
    print("=" * 60)
    print("正在从飞书并行加载三张核心表...")
    t0 = time.time()

    executor = _get_io_executor()
    f_orders = executor.submit(_fetch_bitable_raw, TABLE_ID_ITEMS)
    f_sku = executor.submit(_fetch_bitable_raw, TABLE_ID_SKU)
    f_inv = executor.submit(_fetch_bitable_raw, TABLE_ID_INV)

    orders_raw = f_orders.result()
    sku_raw = f_sku.result()
    inv_raw = f_inv.result()

    orders_df = _records_to_df(orders_raw)
    sku_df = _records_to_df(sku_raw)
    inv_df = _records_to_df(inv_raw)

    build_inventory_index(inv_df)

    elapsed = time.time() - t0
    print(f"飞书数据加载完成（并行 3 线程），耗时 {elapsed:.1f}s")
    print(f"销售订单：{len(orders_df)} 条")
    print(f"SKU数据：{len(sku_df)} 条")
    print(f"库存数据：{len(inv_df)} 条")
    return orders_df, sku_df, inv_df


# =========================
# 日期处理
# =========================
DETAIL_DATE_COLS_DEFAULT = {"预计到货日期"}
DETAIL_NUMERIC_COLS_DEFAULT = {"库存可用量", "缺口数量"}


def to_feishu_date_millis(val: Any) -> Optional[int]:
    """将日期转为飞书多维表格日期字段可用的毫秒时间戳（北京时间当日 0 点）。"""
    d = parse_date_to_date(val)
    if d is None:
        return None
    cn = timezone(timedelta(hours=8))
    dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=cn)
    return int(dt.timestamp() * 1000)


@contextmanager
def schedule_global_lock(timeout_seconds: int = 1):
    """用锁文件做跨进程互斥，避免多实例同时执行排单。"""
    start = time.time()
    fd = None
    while True:
        try:
            fd = os.open(SCHEDULE_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_RDWR)
            os.write(fd, f"{os.getpid()} {datetime.now().isoformat()}".encode("utf-8"))
            break
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(SCHEDULE_LOCK_PATH) > 60 * 60:
                    os.remove(SCHEDULE_LOCK_PATH)
                    continue
            except FileNotFoundError:
                continue
            if time.time() - start >= timeout_seconds:
                raise RuntimeError("排单任务正在执行，本次请求已拒绝，避免多实例并发写入")
            time.sleep(0.2)

    try:
        yield
    finally:
        if fd is not None:
            os.close(fd)
        try:
            os.remove(SCHEDULE_LOCK_PATH)
        except FileNotFoundError:
            pass


def with_schedule_lock(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            with schedule_global_lock(timeout_seconds=1):
                return func(*args, **kwargs)
        except RuntimeError as e:
            print(f"⏔ {e}")
            return {"msg": "排单任务已被互斥锁阻止", "error": str(e)}
    return wrapper


# =========================
# ⭐ 批量写入/更新（飞书批量API，每批≤500条）
# =========================
BATCH_SIZE = 500


def _feishu_value_for_write(col_name: str, val: Any, numeric_fields: Set[str], date_fields: Set[str]) -> Optional[Any]:
    """将单元格值转换为飞书 API 可接受的格式。"""
    if col_name in date_fields:
        ms = to_feishu_date_millis(val)
        return ms if ms is not None else None
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, list):
        return ", ".join([str(v) for v in val])
    if col_name in numeric_fields:
        try:
            return to_num(val)
        except Exception:
            return None
    return str(val)


def _build_fields_dict(
    row: Any,
    columns: List[str],
    fields_map: Optional[Dict[str, str]],
    numeric_fields: Set[str],
    date_fields: Set[str],
    skip_cols: Optional[Set[str]] = None,
) -> Dict[str, Any]:
    """从一行数据构建飞书 fields 字典，跳过内部列和空值。"""
    skip = skip_cols or set()
    fields: Dict[str, Any] = {}
    cols_to_iter = columns if hasattr(row, 'index') else list(row.keys())

    for col in cols_to_iter:
        if col in skip:
            continue
        if isinstance(col, str) and col.startswith("_"):
            continue

        try:
            val = row[col]
        except (KeyError, TypeError):
            continue

        if pd.isna(val) or val is None or val == "":
            continue

        feishu_name = (fields_map or {}).get(col, col)
        fv = _feishu_value_for_write(col, val, numeric_fields, date_fields)
        if fv is None:
            continue
        fields[feishu_name] = fv

    return fields


def write_df_to_bitable(
    table_id: str,
    df: pd.DataFrame,
    fields_map: Optional[Dict[str, str]] = None,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
):
    """批量写入飞书多维表格（使用批量创建 API，每批 ≤500 条）。

    禁止在 iterrows 循环内发起 HTTP 请求。
    """
    if df is None or df.empty:
        return

    numeric_fields = set(numeric_cols or set())
    date_fields = set(date_cols or set())
    columns = list(df.columns)
    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/batch_create"

    records_batch: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        fields = _build_fields_dict(row, columns, fields_map, numeric_fields, date_fields)
        if fields:
            records_batch.append({"fields": fields})

        if len(records_batch) >= BATCH_SIZE:
            _post_batch(url, headers, records_batch, "创建")
            records_batch.clear()

    if records_batch:
        _post_batch(url, headers, records_batch, "创建")


def update_bitable_records(
    table_id: str,
    df: pd.DataFrame,
    record_id_col: str = "_record_id",
    fields_map: Optional[Dict[str, str]] = None,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
):
    """批量更新飞书多维表格原记录（使用批量更新 API，每批 ≤500 条）。

    禁止在 iterrows 循环内发起 HTTP 请求。
    """
    if df is None or df.empty:
        return

    if record_id_col not in df.columns:
        raise ValueError(f"df 缺少 {record_id_col} 列，无法更新回填")

    numeric_fields = set(numeric_cols or set())
    date_fields = set(date_cols or set())
    columns = list(df.columns)
    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/batch_update"

    records_batch: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        record_id = safe_str(row.get(record_id_col, ""))
        if not record_id:
            continue

        fields = _build_fields_dict(
            row, columns, fields_map, numeric_fields, date_fields,
            skip_cols={record_id_col},
        )
        if not fields:
            continue

        records_batch.append({"record_id": record_id, "fields": fields})

        if len(records_batch) >= BATCH_SIZE:
            _post_batch(url, headers, records_batch, "更新")
            records_batch.clear()

    if records_batch:
        _post_batch(url, headers, records_batch, "更新")


def _post_batch(url: str, headers: Dict[str, str], records: List[Dict[str, Any]], action: str):
    """发送一批记录到飞书批量 API。"""
    try:
        resp = httpx.post(url, headers=headers, json={"records": records}, timeout=60.0)
        data = _safe_http_json(resp, f"批量{action}")
        if data.get("code") != 0:
            print(f"批量{action}失败: {data}")
        else:
            print(f"批量{action}成功: {len(records)} 条")
    except Exception as e:
        print(f"批量{action}异常: {str(e)}")


def upsert_bitable_records_by_key(
    table_id: str,
    df: pd.DataFrame,
    key_col: str,
    key_field_name: str,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
) -> Optional[str]:
    """按业务唯一键做 upsert：已有记录则更新，没有则新增。"""
    if df is None or df.empty:
        return None

    try:
        existing = fetch_bitable_to_df(table_id)
    except Exception as e:
        err = str(e)
        if "1254004" in err or "WrongTableId" in err:
            return (
                "读取 AI排单总表失败：飞书错误 WrongTableId(1254004)，说明 TABLE_ID_DETAIL 不是有效的多维表格 table_id。"
                f" 当前 TABLE_ID_DETAIL={table_id!r}。详情：{err}"
            )
        print(f"读取 AI排单总表时出现异常，将采用直接新增方式：{err}")
        existing = pd.DataFrame()

    key_to_record: Dict[str, str] = {}
    if existing is not None and not existing.empty and "_record_id" in existing.columns and key_field_name in existing.columns:
        for _, r in existing.iterrows():
            k = safe_str(r.get(key_field_name, ""))
            rid = safe_str(r.get("_record_id", ""))
            if k and rid and k not in key_to_record:
                key_to_record[k] = rid

    to_update: List[Dict[str, Any]] = []
    to_create: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        key = safe_str(row.get(key_col, ""))
        if not key:
            continue
        if key in key_to_record:
            d = row.to_dict()
            d["_record_id"] = key_to_record[key]
            to_update.append(d)
        else:
            to_create.append(row.to_dict())

    if to_update:
        update_bitable_records(table_id, pd.DataFrame(to_update), record_id_col="_record_id",
                               numeric_cols=numeric_cols, date_cols=date_cols)

    if to_create:
        write_df_to_bitable(table_id, pd.DataFrame(to_create), fields_map=None,
                            numeric_cols=numeric_cols, date_cols=date_cols)

    return None


# =========================
# 数据缓存（避免每次请求重复读取飞书）
# =========================
data_cache = {
    "items": None,
    "sku": None,
    "inv": None
}


def load_data():
    """从飞书一次性加载数据到内存（并行读取），并验证关键字段。"""
    print("=" * 60)
    print("正在从飞书并行加载数据...")
    print("=" * 60)

    try:
        items_df, sku_df, inv_df = load_feishu_data()

        data_cache["items"] = items_df
        data_cache["sku"] = sku_df
        data_cache["inv"] = inv_df

        # ===== 验证表字段完整性 =====
        print("\n【表字段验证】")

        required_items_cols = ["合同编号", "合同数量", "SKU编码"]
        optional_items_cols = ["产品名称", "规格", "库存可用量", "缺口数量", "库存状态", "预计到货日期", "是否RB800", "排单批次号",
                               "是否紧急订单", "是否换货订单", "是否补发订单", "是否维修订单"]

        required_sku_cols = ["产品编码SKU", "标准生产周期"]
        optional_sku_cols = ["设备名称", "设备型号", "是否新产品", "是否自研", "是否外采"]

        required_inv_cols = ["库存日期", "SKU", "库存数量"]
        optional_inv_cols = ["国网设备名称", "国网设备型号", "待采购出库", "在途数量", "数据来源", "导入批次号"]

        def check_columns(df, table_name, required_cols, optional_cols=None):
            optional_cols = optional_cols or []
            missing_required = [col for col in required_cols if col not in df.columns]
            missing_optional = [col for col in optional_cols if col not in df.columns]
            if missing_required:
                print(f"❌ {table_name} 缺少【必填】字段：{missing_required}")
                return False
            if missing_optional:
                print(f"{table_name} 缺少【可选】字段：{missing_optional}（不影响运行）")
            print(f"[OK]{table_name} 必填字段齐全")
            return True

        all_ok = True
        all_ok &= check_columns(items_df, "销售订单明细表", required_items_cols, optional_items_cols)
        all_ok &= check_columns(sku_df, "SKU标准表", required_sku_cols, optional_sku_cols)
        all_ok &= check_columns(inv_df, "库存快照表", required_inv_cols, optional_inv_cols)

        print("\n【数据统计】")
        print(f"  销售订单明细数：{len(items_df)} 条")
        print(f"  SKU标准数：{len(sku_df)} 条")
        print(f"  库存快照数：{len(inv_df)} 条")

        if all_ok:
            print("\n[OK]所有必填字段检验完毕，数据可以使用！")
        else:
            print("\n❌ 存在必填字段缺失，无法继续运行。请检查飞书表结构/列名！")
        print("=" * 60)
    except Exception as e:
        print(f"\n[X] 飞书数据加载失败: {str(e)}")
        print("详细信息：", str(e))
        raise


# =========================
# 辅助函数
# =========================
def safe_str(value):
    """安全转换为字符串并清洗"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()


def to_num(val) -> float:
    try:
        if pd.isna(val):
            return 0.0
    except Exception:
        pass
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def parse_date_to_date(val: Any) -> Optional[date]:
    """兼容飞书日期毫秒值、秒值、日期字符串和 pandas 日期，返回 date。"""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass

    if isinstance(val, pd.Timestamp):
        return None if pd.isna(val) else val.date()
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val

    if isinstance(val, (int, float)) and not isinstance(val, bool):
        n = float(val)
        if n <= 0:
            return None
        try:
            if n >= 10_000_000_000:
                parsed = pd.to_datetime(n, unit="ms", errors="coerce")
            elif n >= 1_000_000_000:
                parsed = pd.to_datetime(n, unit="s", errors="coerce")
            else:
                parsed = pd.to_datetime(n, errors="coerce")
        except Exception:
            return None
        return None if pd.isna(parsed) else parsed.date()

    s = safe_str(val)
    if not s or s.lower() in {"nan", "nat", "none", "null", "1970-01-01"}:
        return None
    if re.fullmatch(r"\d+(\.\d+)?", s):
        return parse_date_to_date(float(s))

    parsed = pd.to_datetime(s, errors="coerce")
    if pd.isna(parsed):
        return None
    d = parsed.date()
    return None if d.year <= 1970 else d


def date_to_yyyy_mm_dd(val: Any) -> str:
    d = parse_date_to_date(val)
    return d.strftime("%Y-%m-%d") if d else ""


def next_working_day(start_date: date) -> date:
    """从 start_date 开始，找到下一个工作日（跳过周末）。"""
    start_date = parse_date_to_date(start_date) or datetime.today().date()
    if start_date.weekday() < 5:
        return start_date
    days_to_add = 7 - start_date.weekday()
    return start_date + timedelta(days=days_to_add)


def add_calendar_days_then_working_day(start_date: date, days: int) -> date:
    """先加自然日，再把结果顺延到工作日。"""
    start_date = parse_date_to_date(start_date) or datetime.today().date()
    return next_working_day(start_date + timedelta(days=int(days)))


def json_safe_value(value: Any) -> Any:
    """递归清理 NaN/NaT/Timestamp，保证 FastAPI JSONResponse 可序列化。"""
    if isinstance(value, dict):
        return {k: json_safe_value(v) for k, v in value.items() if not (isinstance(k, str) and k.startswith("_"))}
    if isinstance(value, list):
        return [json_safe_value(v) for v in value]
    if isinstance(value, tuple):
        return [json_safe_value(v) for v in value]
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return date_to_yyyy_mm_dd(value)
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    return value


def public_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    if df is None or df.empty:
        return []
    public_df = df.copy()
    internal_cols = [c for c in public_df.columns if isinstance(c, str) and c.startswith("_")]
    if internal_cols:
        public_df = public_df.drop(columns=internal_cols, errors="ignore")
    public_df = public_df.astype(object).where(pd.notnull(public_df), None)
    return json_safe_value(public_df.to_dict(orient="records"))


# =========================
# RB800 判断
# =========================
def is_rb800_from_text(text: str) -> bool:
    t = safe_str(text).upper().replace(" ", "")
    return "RB800" in t


def pick_model_text(order_row: Dict[str, Any]) -> str:
    for k in ("规格", "产品名称", "SKU编码"):
        v = safe_str(order_row.get(k, ""))
        if v:
            return v
    return ""


# =========================
# SKU 供给类型
# =========================
def sku_supply_type_from_sku_row(sku_row: Dict[str, Any]) -> str:
    if safe_str(sku_row.get("是否外采", "")) == "是":
        return "外采"
    if safe_str(sku_row.get("是否自研", "")) == "是":
        return "自研"
    if safe_str(sku_row.get("是否新产品", "")) == "是":
        return "新产品"
    return ""


# =========================
# 库存计算
# =========================
def calc_available_stock(inv_row: Optional[Dict[str, Any]], reserved_qty: float = 0.0) -> float:
    if not inv_row:
        base = 0.0
    else:
        base = to_num(inv_row.get("库存数量", 0))
    return max(base - to_num(reserved_qty), 0.0)


def calc_stock_status_and_gap(demand: float, available: float) -> Tuple[str, float]:
    gap = max(to_num(demand) - to_num(available), 0.0)
    return ("有货" if gap <= 0 else "缺货"), (0.0 if gap <= 0 else gap)


def get_available_stock(inv_row):
    return to_num(inv_row.get("库存数量", 0))


# =========================
# 缺货预计到货
# =========================
def calc_shortage_eta_date(
    sku_code: str,
    sku_df: pd.DataFrame,
    base_date: date,
) -> date:
    match = sku_df[sku_df["产品编码SKU"] == sku_code] if (sku_df is not None and not sku_df.empty) else pd.DataFrame()
    if match.empty:
        cycle = 25
    else:
        sku_info = match.iloc[0].to_dict()
        cycle = to_num(sku_info.get("标准生产周期", 0))
        if cycle <= 0:
            cycle = 25

    final_date = add_calendar_days_then_working_day(base_date, int(cycle))
    return final_date


# =========================
# 产能函数
# =========================
def calc_daily_capacity(sku_kinds: int, total_qty: float, base: int = 5) -> int:
    sku_kinds = int(to_num(sku_kinds))
    total_qty = to_num(total_qty)
    bonus = 0
    if sku_kinds < 5 or total_qty < 8:
        bonus = 2
    elif sku_kinds < 10 or total_qty < 15:
        bonus = 1
    return int(base + bonus)


# =========================
# ⭐ 产能调度线性化（单次遍历，并查集路径压缩，O(n·α(n))）
# =========================
def apply_capacity_scheduling(summary: pd.DataFrame, today: date) -> Tuple[pd.DataFrame, int]:
    """对 AI排单总表应用每日产能限制。

    使用并查集式路径压缩确保 O(n·α(n))，禁止嵌套 while+for 导致 O(n²)。
    日期仅做 date 对象比较，不反复 strftime。
    """
    if summary is None or summary.empty or "AI建议发货时间" not in summary.columns:
        return summary, 0

    df = summary.copy()
    # 只做一次日期解析，后续用 date 对象比较
    df["__ship_date"] = df["AI建议发货时间"].apply(lambda x: next_working_day(parse_date_to_date(x) or today))

    if "订单状态" not in df.columns:
        df["订单状态"] = "待确认"
    df["订单状态"] = df["订单状态"].apply(normalize_order_status)
    df["__locked"] = df["订单状态"] != "待确认"
    df["__special"] = df.get("项目类型", "").astype(str) == "特殊订单"

    normal = df[~df["__locked"] & ~df["__special"]].copy()
    others = df[df["__locked"] | df["__special"]].copy()
    normal = normal.sort_values(by=["__ship_date", "合同编号"], kind="stable")

    # 每个日期的已排单数 / SKU种类累加 / 数量累加
    day_count: Dict[date, int] = {}
    day_sku_kinds: Dict[date, int] = {}
    day_total_qty: Dict[date, float] = {}
    # 并查集 overflow 指针：当某天满时，指向下一个应尝试的日期
    overflow: Dict[date, date] = {}

    delayed = 0
    assigned_dates: Dict[str, date] = {}

    for _, row in normal.iterrows():
        base_date = next_working_day(row["__ship_date"])
        order_sku_kinds = int(to_num(row.get("订单SKU总数", 0)))
        order_qty = to_num(row.get("订单总数量", 0))
        contract_id = safe_str(row.get("合同编号", ""))

        # 路径压缩：跳过所有已满日期（防死循环：最多检查 365 天）
        d = base_date
        visited: List[date] = []
        _safety = 0
        while d in overflow and _safety < 365:
            visited.append(d)
            d = overflow[d]
            _safety += 1
        # 路径压缩：visited 中的所有日期直接指向 d
        for v in visited:
            overflow[v] = d

        # 检查 d 是否真的还有容量
        c = day_count.get(d, 0)
        sk = day_sku_kinds.get(d, 0)
        tq = day_total_qty.get(d, 0)
        cap = calc_daily_capacity(sk, tq, base=5)

        if c >= cap:
            next_d = next_working_day(d + timedelta(days=1))
            overflow[d] = next_d
            # 继续找下一个空闲日
            d = next_d
            c = day_count.get(d, 0)
            sk = day_sku_kinds.get(d, 0)
            tq = day_total_qty.get(d, 0)
            cap = calc_daily_capacity(sk, tq, base=5)
            # 最多再检查一次；如果连续满则走压缩路径（防死循环）
            _safety2 = 0
            while d in overflow and _safety2 < 365:
                d = overflow[d]
                _safety2 += 1
                c = day_count.get(d, 0)
                sk = day_sku_kinds.get(d, 0)
                tq = day_total_qty.get(d, 0)
                cap = calc_daily_capacity(sk, tq, base=5)

        # 分配到 d
        day_count[d] = c + 1
        day_sku_kinds[d] = sk + order_sku_kinds
        day_total_qty[d] = tq + order_qty
        assigned_dates[contract_id] = d

        if d != base_date:
            delayed += 1

        # 如果 d 刚被填满，标记 overflow
        new_c = c + 1
        new_sk = sk + order_sku_kinds
        new_tq = tq + order_qty
        new_cap = calc_daily_capacity(new_sk, new_tq, base=5)
        if new_c >= new_cap:
            overflow[d] = next_working_day(d + timedelta(days=1))

    # 应用分配结果
    if assigned_dates:
        normal["__ship_date_new"] = normal["合同编号"].astype(str).apply(
            lambda x: assigned_dates.get(safe_str(x), next_working_day(today))
        )
        normal["AI建议发货时间"] = normal["__ship_date_new"].apply(date_to_yyyy_mm_dd)
        normal = normal.drop(columns=["__ship_date_new"])

    out = pd.concat([normal, others], ignore_index=True).drop(
        columns=["__ship_date", "__locked", "__special"], errors="ignore"
    )
    if "AI建议发货时间" in out.columns:
        out["AI建议发货时间"] = out["AI建议发货时间"].apply(
            lambda x: date_to_yyyy_mm_dd(next_working_day(parse_date_to_date(x) or today))
        )
    if "AI建议发货时间" in out.columns:
        out = out.sort_values(by=["AI建议发货时间", "合同编号"], kind="stable")

    return out, delayed


# =========================
# 订单发货时间计算
# =========================
def calc_order_ship_date_for_group(
    group: pd.DataFrame,
    sku_df: pd.DataFrame,
    today: date,
) -> Tuple[Optional[date], Optional[date]]:
    """计算订单级别的预计发货时间：返回 (latest_eta, ship_date)"""
    shortage_mask = group["库存状态"] == "缺货"
    shortage_rows = group[shortage_mask]

    if shortage_rows.empty:
        return None, next_working_day(today)

    valid_dates = [d for d in shortage_rows['预计到货日期'].apply(parse_date_to_date).tolist() if d is not None]
    if valid_dates:
        latest_eta = max(valid_dates)
        ship_date = next_working_day(latest_eta)
        return latest_eta, ship_date

    fallback_eta = add_calendar_days_then_working_day(today, 25)
    return fallback_eta, fallback_eta


# =========================
# 汇总相关辅助函数
# =========================
def normalize_shortage_sku_list(value: Any) -> List[str]:
    if isinstance(value, list):
        raw_items = value
    else:
        s = safe_str(value)
        if not s or s.lower() in {"nan", "none", "null"}:
            return []
        raw_items = re.split(r"[,，]", s)

    seen: Set[str] = set()
    result: List[str] = []
    for item in raw_items:
        sku_code = safe_str(item)
        if not sku_code or sku_code.lower() in {"nan", "none", "null"}:
            continue
        if sku_code not in seen:
            seen.add(sku_code)
            result.append(sku_code)
    return result


def shortage_sku_count_from_list(value: Any) -> int:
    return len(normalize_shortage_sku_list(value))


def normalize_order_status(value: Any, default: str = "待确认") -> str:
    status = safe_str(value)
    if status in {"待确认", "已确认", "已发货"}:
        return status
    return default


def normalize_reservation_status(value: Any, default: str = "有效") -> str:
    status = safe_str(value)
    if status in {"有效", "已转换", "已过期"}:
        return status
    return default


def is_confirmed_order(row: Dict[str, Any]) -> bool:
    return normalize_order_status(row.get("订单状态", "")) == "已确认"


def has_valid_manual_confirm_date(row: Dict[str, Any]) -> bool:
    return parse_date_to_date(row.get("人工确认发货时间")) is not None


def is_effective_manual_confirmed(row: Dict[str, Any]) -> bool:
    return safe_str(row.get("是否人工确认", "")) == "是" and has_valid_manual_confirm_date(row)


def make_run_batch_id(reservations: pd.DataFrame, today: date) -> str:
    prefix = f"RUN-{today.strftime('%Y%m%d')}-"
    max_seq = 0
    if reservations is not None and not reservations.empty and "批次号" in reservations.columns:
        for value in reservations["批次号"].astype(str):
            m = re.fullmatch(re.escape(prefix) + r"(\d{3})", safe_str(value))
            if m:
                max_seq = max(max_seq, int(m.group(1)))
    return f"{prefix}{max_seq + 1:03d}"


def ensure_reservation_table_id() -> Optional[str]:
    if TABLE_ID_RESERVATION:
        return None
    return "缺少 AI排单_库存预留表 ID，请配置环境变量 TABLE_ID_RESERVATION"


def prepare_summary_status(summary: pd.DataFrame) -> pd.DataFrame:
    if summary is None:
        return pd.DataFrame()
    out = summary.copy()
    if out.empty:
        return out
    if "订单状态" not in out.columns:
        out["订单状态"] = "待确认"
    out["订单状态"] = out["订单状态"].apply(normalize_order_status)
    if "合同编号" in out.columns:
        out["合同编号"] = out["合同编号"].astype(str).apply(safe_str)
    return out


def summarize_confirmed_stock(summary: pd.DataFrame, detail: pd.DataFrame) -> Dict[str, float]:
    locked_stock: Dict[str, float] = {}
    summary = prepare_summary_status(summary)
    if summary.empty or detail is None or detail.empty or "合同编号" not in detail.columns:
        return locked_stock

    confirmed_contracts = set(
        summary.loc[summary["订单状态"] == "已确认", "合同编号"].astype(str).apply(safe_str)
    )
    if not confirmed_contracts:
        return locked_stock

    source = detail.copy()
    source["合同编号"] = source["合同编号"].astype(str).apply(safe_str)
    if "SKU编码" not in source.columns or "合同数量" not in source.columns:
        return locked_stock

    locked_detail = source[source["合同编号"].isin(confirmed_contracts)]
    for _, r in locked_detail.iterrows():
        sku_code = safe_str(r.get("SKU编码", ""))
        qty = to_num(r.get("合同数量", 0))
        if sku_code and qty > 0:
            locked_stock[sku_code] = locked_stock.get(sku_code, 0.0) + qty
    return locked_stock


def get_confirmed_contract_ids(summary: pd.DataFrame) -> set:
    summary = prepare_summary_status(summary)
    if summary.empty or "合同编号" not in summary.columns:
        return set()
    return set(summary.loc[summary["订单状态"] == "已确认", "合同编号"].astype(str).apply(safe_str))


def summarize_effective_reservations(reservations: pd.DataFrame) -> Dict[str, float]:
    reserved_stock: Dict[str, float] = {}
    if reservations is None or reservations.empty:
        return reserved_stock
    required = {"SKU", "预留数量", "状态"}
    if not required.issubset(set(reservations.columns)):
        return reserved_stock
    active = reservations[reservations["状态"].apply(normalize_reservation_status) == "有效"]
    for _, r in active.iterrows():
        sku_code = safe_str(r.get("SKU", ""))
        qty = to_num(r.get("预留数量", 0))
        if sku_code and qty > 0:
            reserved_stock[sku_code] = reserved_stock.get(sku_code, 0.0) + qty
    return reserved_stock


def update_reservation_records(records: pd.DataFrame, status: str, reason: str) -> int:
    if records is None or records.empty:
        return 0
    if "_record_id" not in records.columns:
        return 0
    df = records[["_record_id"]].copy()
    df["状态"] = status
    df["释放原因"] = reason
    update_bitable_records(TABLE_ID_RESERVATION, df, record_id_col="_record_id")
    return int(len(df))


def cleanup_reservations(
    reservations: pd.DataFrame,
    summary: pd.DataFrame,
    current_batch_id: str,
) -> Tuple[int, int]:
    if reservations is None or reservations.empty:
        return 0, 0
    res = reservations.copy()
    if "状态" not in res.columns:
        res["状态"] = ""
    if "批次号" not in res.columns:
        res["批次号"] = ""
    if "合同编号" not in res.columns:
        res["合同编号"] = ""

    summary = prepare_summary_status(summary)
    confirmed_contracts = set()
    if not summary.empty and "合同编号" in summary.columns:
        confirmed_contracts = set(summary.loc[summary["订单状态"] == "已确认", "合同编号"].astype(str).apply(safe_str))

    active_mask = res["状态"].apply(normalize_reservation_status) == "有效"
    converted_count = 0
    if confirmed_contracts:
        converted = res[active_mask & res["合同编号"].astype(str).apply(safe_str).isin(confirmed_contracts)]
        converted_count = update_reservation_records(converted, "已转换", "人工确认自动转换")

    expired = res[
        active_mask
        & (res["批次号"].astype(str).apply(safe_str) != current_batch_id)
        & ~res["合同编号"].astype(str).apply(safe_str).isin(confirmed_contracts)
    ]
    expired_count = update_reservation_records(expired, "已过期", "超期未确认")

    return expired_count, converted_count


def release_old_contract_reservations(reservations: pd.DataFrame, contract_ids: List[str]) -> int:
    if reservations is None or reservations.empty or not contract_ids:
        return 0
    required = {"合同编号", "状态"}
    if not required.issubset(set(reservations.columns)):
        return 0
    contract_set = {safe_str(x) for x in contract_ids if safe_str(x)}
    active = reservations[
        (reservations["状态"].apply(normalize_reservation_status) == "有效")
        & reservations["合同编号"].astype(str).apply(safe_str).isin(contract_set)
    ]
    return update_reservation_records(active, "已过期", "新一轮排单覆盖")


def build_reservation_rows(summary: pd.DataFrame, detail: pd.DataFrame, batch_id: str) -> pd.DataFrame:
    if summary is None or summary.empty or detail is None or detail.empty:
        return pd.DataFrame()
    summary = prepare_summary_status(summary)
    pending_contracts = set(summary.loc[summary["订单状态"] == "待确认", "合同编号"].astype(str).apply(safe_str))
    if not pending_contracts:
        return pd.DataFrame()

    required = {"合同编号", "SKU编码", "合同数量"}
    if not required.issubset(set(detail.columns)):
        return pd.DataFrame()

    qty_by_contract_sku: Dict[Tuple[str, str], float] = {}
    now_str = datetime.now().strftime("%Y-%m-%d")
    source = detail.copy()
    source["合同编号"] = source["合同编号"].astype(str).apply(safe_str)
    source["SKU编码"] = source["SKU编码"].astype(str).apply(safe_str)
    for _, r in source[source["合同编号"].isin(pending_contracts)].iterrows():
        contract_id = safe_str(r.get("合同编号", ""))
        sku_code = safe_str(r.get("SKU编码", ""))
        qty = to_num(r.get("合同数量", 0))
        if not contract_id or not sku_code or qty <= 0:
            continue
        key = (contract_id, sku_code)
        qty_by_contract_sku[key] = qty_by_contract_sku.get(key, 0.0) + qty

    rows = []
    for (contract_id, sku_code), qty in qty_by_contract_sku.items():
        rows.append({
            "预留ID": f"{batch_id}_{contract_id}_{sku_code}",
            "合同编号": contract_id,
            "SKU": sku_code,
            "预留数量": qty,
            "批次号": batch_id,
            "创建时间": now_str,
            "状态": "有效",
            "释放原因": "",
        })
    return pd.DataFrame(rows)


def build_current_detail_for_summary(
    items_df: pd.DataFrame,
    reread_df: pd.DataFrame,
    backfill_df: pd.DataFrame,
) -> pd.DataFrame:
    """用本次加载的明细叠加本次计算结果，避免回填后重读延迟导致新增订单漏排。"""
    source = items_df.copy() if items_df is not None else pd.DataFrame()
    reread = reread_df.copy() if reread_df is not None else pd.DataFrame()
    backfill = backfill_df.copy() if backfill_df is not None else pd.DataFrame()

    for df in (source, reread, backfill):
        if not df.empty and "_record_id" in df.columns:
            df["_record_id"] = df["_record_id"].astype(str).apply(safe_str)

    if reread.empty:
        detail = source.copy()
    else:
        detail = reread.copy()
        if not source.empty and "_record_id" in detail.columns and "_record_id" in source.columns:
            existing_ids = set(detail["_record_id"].astype(str).apply(safe_str))
            missing_source = source[~source["_record_id"].astype(str).apply(safe_str).isin(existing_ids)]
            if not missing_source.empty:
                print(f"⚠️ 回填后重读缺少 {len(missing_source)} 条本次明细，已用本次加载数据补齐")
                detail = pd.concat([detail, missing_source], ignore_index=True, sort=False)

    if detail.empty or backfill.empty or "_record_id" not in detail.columns or "_record_id" not in backfill.columns:
        return detail

    backfill_map = backfill.drop_duplicates(subset=["_record_id"], keep="last").set_index("_record_id")
    overlay_cols = [c for c in backfill_map.columns if c != "_record_id"]
    detail = detail.copy()
    for col in overlay_cols:
        if col not in detail.columns:
            detail[col] = ""
        detail[col] = detail.apply(
            lambda r: backfill_map.at[safe_str(r.get("_record_id", "")), col]
            if safe_str(r.get("_record_id", "")) in backfill_map.index
            else r.get(col, ""),
            axis=1,
        )

    return detail


# =========================
# 首页 / 操作面板（无需外部CDN）
# =========================
HOME_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>AI排单服务</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f2f5;color:#333;min-height:100vh;display:flex}
.sidebar{width:240px;background:#1a1a2e;color:#eee;padding:24px 0;display:flex;flex-direction:column}
.sidebar h2{padding:0 20px 20px;border-bottom:1px solid rgba(255,255,255,.1);margin-bottom:8px;font-size:17px}
.sidebar a{color:#aaa;text-decoration:none;padding:10px 20px;font-size:14px;display:block}
.sidebar a:hover,.sidebar a.active{color:#fff;background:rgba(255,255,255,.08)}
.main{flex:1;padding:32px;max-width:960px}
.card{background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:28px;margin-bottom:20px}
.card h3{font-size:18px;margin-bottom:12px}
.card p{color:#666;line-height:1.7;margin-bottom:8px}
.btn{display:inline-block;padding:10px 28px;border:none;border-radius:6px;font-size:15px;cursor:pointer;text-decoration:none}
.btn-primary{background:#4f46e5;color:#fff}
.btn-primary:hover{background:#4338ca}
.btn-primary:disabled{opacity:.5;cursor:not-allowed}
#result{margin-top:16px;white-space:pre-wrap;background:#1e1e2e;color:#cdd6f4;border-radius:6px;padding:16px;font-family:'Cascadia Code',Consolas,monospace;font-size:13px;max-height:500px;overflow-y:auto;display:none}
.badge{display:inline-block;padding:2px 10px;border-radius:10px;font-size:12px;font-weight:600}
.badge-get{background:#dbeafe;color:#1e40af}
.badge-post{background:#fce7f3;color:#9d174d}
table{margin-top:8px;width:100%}
td{padding:6px 12px 6px 0;font-size:14px}
td:first-child{font-family:monospace;font-weight:600;width:100px}
.status{font-size:14px;color:#888;margin-top:4px}
</style>
</head>
<body>
<div class="sidebar">
  <h2>AI排单服务</h2>
  <a href="/" class="active">操作面板</a>
  <a href="/docs">API 参考</a>
</div>
<div class="main">
  <div class="card">
    <h3>触发排单</h3>
    <p>点击按钮执行一次完整的 AI 排单流程：读取订单 → 匹配库存 → 计算缺货 → 产能调度 → 写回飞书。</p>
    <button class="btn btn-primary" id="runBtn" onclick="runSchedule()">执行排单</button>
    <div class="status" id="status"></div>
    <pre id="result"></pre>
  </div>
  <div class="card">
    <h3>API 接口</h3>
    <table>
      <tr><td><span class="badge badge-post">POST</span></td><td>/schedule</td><td>手动触发排单</td></tr>
      <tr><td><span class="badge badge-get">GET</span></td><td>/daily-report</td><td>获取今日排单日报</td></tr>
      <tr><td><span class="badge badge-post">POST</span></td><td>/webhook/feishu</td><td>飞书 Webhook 回调</td></tr>
      <tr><td><span class="badge badge-get">GET</span></td><td>/</td><td>操作面板</td></tr>
      <tr><td><span class="badge badge-get">GET</span></td><td>/docs</td><td>API JSON 参考</td></tr>
    </table>
  </div>
</div>
<script>
async function runSchedule(){
  const btn=document.getElementById('runBtn');
  const res=document.getElementById('result');
  const st=document.getElementById('status');
  btn.disabled=true;
  st.textContent='⏳ 排单中，请稍候...';
  res.style.display='none';
  try{
    const r=await fetch('/schedule',{method:'POST',headers:{'Content-Type':'application/json'},body:'{"trigger":"manual"}'});
    const d=await r.json();
    res.textContent=JSON.stringify(d,null,2);
    res.style.display='block';
    st.textContent=r.ok?'[OK]排单完成，耗时 '+d.elapsed_s+'s':'❌ 排单失败';
  }catch(e){
    res.textContent='请求失败: '+e.message;
    res.style.display='block';
    st.textContent='❌ 网络错误';
  }
  btn.disabled=false;
}
</script>
</body>
</html>"""


@app.get("/")
def home():
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content=HOME_HTML)


@app.get("/docs")
def api_docs():
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content=HOME_HTML)


# =========================
# AI排单日报
# =========================

def _fmt_shortage_sku(sku_list: List[tuple],
                      name_map: Dict[str, str], unit_map: Dict[str, str]) -> str:
    """格式化紧缺 SKU 列表（按缺口降序）：设备名称（缺X单位）, ..."""
    if not sku_list:
        return ""
    parts = []
    for sku_code, gap in sku_list:
        if not sku_code or gap <= 0:
            continue
        dev_name = name_map.get(sku_code, sku_code)
        unit = unit_map.get(sku_code, "个")
        gap_str = str(int(gap)) if gap == int(gap) else str(round(gap, 1))
        parts.append(f"{dev_name}（缺{gap_str}{unit}）")
    return "，".join(parts)


def generate_daily_report(
    items_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    inv_df: pd.DataFrame,
    sku_df: pd.DataFrame,
    batch_id: str,
) -> dict:
    """生成日报统计，只读源数据，不修改任何表。"""
    from datetime import timezone as tz, timedelta
    from zoneinfo import ZoneInfo
    CST = ZoneInfo("Asia/Shanghai")

    now_cst = datetime.now(CST)
    today = now_cst.date()
    report_time = now_cst.strftime("%Y-%m-%d %H:%M:%S")
    today_str = today.strftime("%Y-%m-%d")

    def _d(val):
        """安全解析日期，返回 date 或 None"""
        return parse_date_to_date(val)

    def _s(val):
        return safe_str(val).strip()

    # ---- 辅助: 从 items 表按合同编号去重取合同级信息 ----
    items = items_df.copy() if items_df is not None else pd.DataFrame()
    for c in ["合同编号", "合同数量"]:
        if c not in items.columns:
            items[c] = ""
    items["合同编号"] = items["合同编号"].astype(str).apply(_s)
    items["合同数量"] = items["合同数量"].apply(to_num)

    contracts_in_items = set(items.loc[items["合同编号"] != "", "合同编号"])

    # ---- 发货日期派生 ----
    summary = summary_df.copy() if summary_df is not None else pd.DataFrame()
    if summary.empty:
        return {}

    for c in ["合同编号", "AI建议发货时间", "人工确认发货时间", "是否人工确认", "整体状态", "批次号"]:
        if c not in summary.columns:
            summary[c] = ""

    summary["合同编号"] = summary["合同编号"].astype(str).apply(_s)
    summary["是否人工确认"] = summary["是否人工确认"].astype(str).apply(_s)
    summary["整体状态"] = summary["整体状态"].astype(str).apply(_s)
    summary["批次号"] = summary["批次号"].astype(str).apply(_s)

    # 发货日期 = IF 人工确认发货时间 valid THEN 人工确认发货时间 ELSE AI建议发货时间
    def calc_ship_date(row):
        manual = _d(row.get("人工确认发货时间", ""))
        if manual is not None:
            return manual
        return _d(row.get("AI建议发货时间", ""))

    summary["发货日期"] = summary.apply(calc_ship_date, axis=1)

    summary_contracts = set(summary.loc[summary["合同编号"] != "", "合同编号"])

    # ---- 3.2 订单维度统计 ----
    total_orders = len(contracts_in_items)

    new_today = 0
    urgent_today = 0
    order_date_col = None
    for col_name in ["订单日期", "下单日期", "下单时间"]:
        if col_name in items.columns:
            order_date_col = col_name
            break

    if order_date_col:
        items["__order_date"] = items[order_date_col].apply(_d)
        items_today = items[items["__order_date"] == today]
        new_today = items_today.loc[items_today["合同编号"] != "", "合同编号"].nunique()

        urgent_col = None
        for col_name in ["是否紧急订单"]:
            if col_name in items.columns:
                urgent_col = col_name
                break
        if urgent_col:
            urgent_today = items_today[
                (items_today[urgent_col].astype(str).apply(_s) == "是")
                & (items_today["合同编号"] != "")
            ]["合同编号"].nunique()
    else:
        print("[日报] 销售订单表缺少订单日期字段，今日新增/紧急订单填0")

    unscheduled = len(contracts_in_items - summary_contracts)
    scheduled = len(summary_contracts)
    confirmed = summary.loc[summary["是否人工确认"] == "是", "合同编号"].nunique()

    # ---- 3.3 今日发货能力 ----
    today_ship = summary[summary["发货日期"].apply(lambda x: x == today if x else False)]
    today_ship_contracts = today_ship.loc[today_ship["合同编号"] != "", "合同编号"].unique()
    today_should_ship = len(today_ship_contracts)
    today_can_ship = today_ship[today_ship["整体状态"] != "缺货"]["合同编号"].nunique()
    today_delay = today_ship[today_ship["整体状态"] == "缺货"]["合同编号"].nunique()

    # ---- 3.4 库存维度统计 ----
    inv = inv_df.copy() if inv_df is not None else pd.DataFrame()
    if not inv.empty and "库存日期" in inv.columns:
        inv["__inv_date"] = inv["库存日期"].apply(_d)
        latest_date = inv["__inv_date"].max()
        if pd.notna(latest_date):
            inv_latest = inv[inv["__inv_date"] == latest_date]
        else:
            inv_latest = inv
    else:
        inv_latest = inv

    for c in ["SKU", "库存数量", "待采购出库", "在途数量"]:
        if c not in inv_latest.columns:
            inv_latest[c] = 0
    inv_latest["库存数量"] = inv_latest["库存数量"].apply(to_num)
    inv_latest["待采购出库"] = inv_latest["待采购出库"].apply(to_num)
    inv_latest["在途数量"] = inv_latest["在途数量"].apply(to_num)
    inv_latest["可用库存"] = (inv_latest["库存数量"] - inv_latest["待采购出库"] + inv_latest["在途数量"]).clip(lower=0)

    sku_total = inv_latest.loc[inv_latest["SKU"].astype(str).apply(_s) != "", "SKU"].nunique()

    sku_data = sku_df.copy() if sku_df is not None else pd.DataFrame()
    has_safety_stock = "安全库存" in sku_data.columns
    if has_safety_stock:
        sku_data["安全库存"] = sku_data["安全库存"].apply(to_num)
        sku_map = sku_data.set_index(sku_data["产品编码SKU"].astype(str).apply(_s))["安全库存"].to_dict()
    else:
        sku_map = {}

    # SKU → 设备名称/单位 映射
    sku_name_map: Dict[str, str] = {}
    sku_unit_map: Dict[str, str] = {}
    if not sku_data.empty and "产品编码SKU" in sku_data.columns:
        for _, sr in sku_data.iterrows():
            code = _s(sr.get("产品编码SKU", ""))
            if code:
                sku_name_map[code] = _s(sr.get("设备名称", ""))
                sku_unit_map[code] = _s(sr.get("设备单位", ""))

    # 清洗 SKU 编码：Lookup 字段可能返回含逗号的拼接串，提取第一个有效编码
    def _clean_sku(val: str) -> str:
        parts = [p.strip() for p in val.replace(",,", ",").split(",") if p.strip()]
        return parts[0] if parts else val

    sku_agg = inv_latest.groupby(inv_latest["SKU"].astype(str).apply(_s).apply(_clean_sku))["可用库存"].sum()

    sufficient, warning = 0, 0
    for sku_code, avail in sku_agg.items():
        if not sku_code:
            continue
        safety = sku_map.get(sku_code, None)
        if safety is not None and safety > 0:
            if avail >= safety:
                sufficient += 1
            elif avail > 0:
                warning += 1

    # 库存缺货：可用库存 <= 0 的 SKU 数量（与安全库存无关）
    shortage = int((sku_agg <= 0).sum())

    # ---- 最紧缺SKU：已下单订单中缺货的产品，按缺口数量降序 ----
    gap_sku, gap_max = "", 0.0
    shortage_order_skus: List[tuple] = []  # [(sku_code, total_gap), ...]
    if items is not None and not items.empty and "库存状态" in items.columns and "SKU编码" in items.columns:
        shortage_items = items[items["库存状态"].astype(str).apply(_s) == "缺货"]
        if not shortage_items.empty:
            gap_by_sku: Dict[str, float] = {}
            for _, r in shortage_items.iterrows():
                sku = _clean_sku(_s(r.get("SKU编码", "")))
                if not sku:
                    continue
                gap_by_sku[sku] = gap_by_sku.get(sku, 0.0) + to_num(r.get("缺口数量", 0))
            sorted_gaps = sorted(gap_by_sku.items(), key=lambda x: x[1], reverse=True)
            shortage_order_skus = [(sku, gap) for sku, gap in sorted_gaps if sku and gap > 0]
            if shortage_order_skus:
                gap_sku, gap_max = shortage_order_skus[0]

    # ---- 3.5 排单结果统计 ----
    summary_with_date = summary[summary["发货日期"].notna() & (summary["发货日期"] != None)]  # noqa: E711
    max_lead_contract = ""
    max_lead_days = -1
    latest_ship_date = ""

    for _, r in summary_with_date.iterrows():
        sd = r["发货日期"]
        if sd is None:
            continue
        sd_str = date_to_yyyy_mm_dd(sd)
        if sd_str and (sd_str > latest_ship_date or not latest_ship_date):
            latest_ship_date = sd_str
        lead_days = (sd - today).days
        cid = _s(r.get("合同编号", ""))
        if lead_days > max_lead_days or (lead_days == max_lead_days and cid and (not max_lead_contract or cid < max_lead_contract)):
            max_lead_days = lead_days
            max_lead_contract = cid

    # 排单占库存比例
    ratio = 0.0
    inv_total_qty = inv_latest["库存数量"].sum()
    if inv_total_qty > 0:
        scheduled_items = items[items["合同编号"].isin(summary_contracts)]
        scheduled_qty = scheduled_items["合同数量"].sum()
        ratio = round(scheduled_qty / inv_total_qty, 4)

    # ---- 3.6 未来预测 ----
    def in_range(sd, lo_days, hi_days):
        if sd is None:
            return False
        delta = (sd - today).days
        return lo_days < delta <= hi_days

    future_3 = summary[summary["发货日期"].apply(lambda x: in_range(x, 0, 3))]
    future_7 = summary[summary["发货日期"].apply(lambda x: in_range(x, 0, 7))]

    f3_ship = future_3.loc[future_3["合同编号"] != "", "合同编号"].nunique()
    f3_short = future_3[future_3["整体状态"] == "缺货"]["合同编号"].nunique()
    f7_ship = future_7.loc[future_7["合同编号"] != "", "合同编号"].nunique()

    # ---- 组装结果 ----
    report = {
        "日期": today_str,
        "排单运行时间": report_time,
        "排单批次号": batch_id,
        "订单总数": str(total_orders),
        "今日新增订单": str(new_today),
        "今日紧急订单": str(urgent_today),
        "未排单订单数": str(unscheduled),
        "已排单订单数": str(scheduled),
        "人工已确认排单数": str(confirmed),
        "今日应发货订单数": str(today_should_ship),
        "今日可发货订单数": str(today_can_ship),
        "今日预计延迟订单数": str(today_delay),
        "SKU总数": str(sku_total),
        "库存充足SKU数": str(sufficient),
        "库存预警SKU数": str(warning),
        "库存缺货SKU数": str(shortage),
        "最紧缺SKU": _fmt_shortage_sku(shortage_order_skus[:1], sku_name_map, sku_unit_map),
        "最紧缺SKU缺口": str(int(gap_max)) if gap_max == int(gap_max) else str(round(gap_max, 1)),
        "最长交期订单": max_lead_contract,
        "最晚发货日期": latest_ship_date,
        "排单占库存比例": str(ratio),
        "未来3天应发货订单数": str(f3_ship),
        "未来3天缺货订单数": str(f3_short),
        "未来7天应发货订单数": str(f7_ship),
    }
    return report


# =========================
# 读取今日排单日报（供外部机器人/通知消费）
# =========================

def get_today_report_row() -> dict:
    """读取 AI排单日报表 中最新的那一行，并校验是否为今日数据。

    三个核心安全策略：
    1. 倒序取 Top 1 —— sort 按"排单运行时间"降序，page_size=1，永远只抓最新行
    2. 今日数据校验锁 —— 如果最新行不是今天的，说明上游排单系统可能挂了没写数据
    3. 安全容错提取 —— 所有字段用 .get() 赋默认值，空单元格不会导致 KeyError 崩溃
    """
    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/tables/{TABLE_ID_DAILY_REPORT}/records/search"

    payload = {
        "page_size": 1,
        "sort": [
            {"field_name": "排单运行时间", "desc": True}
        ]
    }

    try:
        resp = httpx.post(url, headers=headers, json=payload, timeout=30.0)
        data = _safe_http_json(resp, "读取日报表")
    except Exception as e:
        raise Exception(f"读取日报表网络请求失败: {e}")

    if data.get("code") != 0:
        raise Exception(f"读取日报表失败: code={data.get('code')}, msg={data.get('msg')}")

    items = data.get("data", {}).get("items", [])
    if not items:
        raise ValueError("飞书日报表为空，没有找到任何排单日报数据。")

    record = items[0]["fields"]

    # ===== 今日数据校验锁 =====
    # 供应链业务铁律：宁可不发报表，也绝不能发错报表。
    # 如果前置排单系统挂了没写数据，这里会识别出最新行是昨天的旧数据。
    CST = timezone(timedelta(hours=8))
    today_str = datetime.now(CST).strftime("%Y-%m-%d")

    # 飞书日期字段可能是毫秒时间戳，也可能是字符串，统一尝试转换
    record_date_str = ""
    raw_date = record.get("日期")
    if raw_date is not None:
        if isinstance(raw_date, (int, float)) and raw_date > 0:
            try:
                if raw_date >= 10_000_000_000:
                    record_date_str = datetime.fromtimestamp(raw_date / 1000, tz=CST).strftime("%Y-%m-%d")
                elif raw_date >= 1_000_000_000:
                    record_date_str = datetime.fromtimestamp(raw_date, tz=CST).strftime("%Y-%m-%d")
            except Exception:
                pass
        if not record_date_str:
            record_date_str = str(raw_date).strip()

    if today_str not in str(record_date_str):
        print(
            f"[日报校验] 获取到的最新数据日期为 [{record_date_str}]，"
            f"不是今日 [{today_str}]。"
            f"可能排单脚本今日未运行或运行失败，请立即排查！"
        )
        # 供应链场景：宁可中断，不发旧数据。取消下面这行注释即可启用硬阻断：
        # raise ValueError(f"日报数据过期：最新数据日期为 {record_date_str}，非今日 {today_str}。排单脚本可能未运行。")

    # ===== 安全容错提取 =====
    # 飞书 API 返回的字段值是富文本数组格式 [{"text":"...","type":"text"}]，
    # 必须用 parse_feishu_field() 清洗为纯文本，不能直接 str()。
    report = {
        "排单批次号":       parse_feishu_field(record.get("排单批次号")) or "未知批次",
        "订单总数":         parse_feishu_field(record.get("订单总数")) or "0",
        "今日新增订单":     parse_feishu_field(record.get("今日新增订单")) or "0",
        "今日紧急订单":     parse_feishu_field(record.get("今日紧急订单")) or "0",
        "未排单订单数":     parse_feishu_field(record.get("未排单订单数")) or "0",
        "已排单订单数":     parse_feishu_field(record.get("已排单订单数")) or "0",
        "人工已确认排单数": parse_feishu_field(record.get("人工已确认排单数")) or "0",
        "今日应发货订单数": parse_feishu_field(record.get("今日应发货订单数")) or "0",
        "今日可发货订单数": parse_feishu_field(record.get("今日可发货订单数")) or "0",
        "今日预计延迟订单数": parse_feishu_field(record.get("今日预计延迟订单数")) or "0",
        "库存缺货SKU数":    parse_feishu_field(record.get("库存缺货SKU数")) or "0",
        "最紧缺SKU":        parse_feishu_field(record.get("最紧缺SKU")) or "无",
        "最紧缺SKU缺口":    parse_feishu_field(record.get("最紧缺SKU缺口")) or "0",
        "未来3天缺货订单数": parse_feishu_field(record.get("未来3天缺货订单数")) or "0",
        "日期":             record_date_str or today_str,
        "排单运行时间":     parse_feishu_field(record.get("排单运行时间")) or "",
    }

    return report


def _send_personal_notification(summary: pd.DataFrame, batch_id: str, report: dict) -> None:
    """排单完成后私聊通知排单负责人，失败不影响排单。"""
    planner_open_id = os.getenv("PLANNER_OPEN_ID", "").strip()
    if not planner_open_id:
        return

    try:
        token_resp = httpx.post(
            "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
            json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}, timeout=15.0,
        )
        token = token_resp.json().get("tenant_access_token", "")
        if not token:
            return

        total = len(summary) if summary is not None else 0
        pending = int(summary[summary["订单状态"] == "待确认"].shape[0]) if summary is not None and not summary.empty else 0
        shortage_orders = int(summary[summary["整体状态"] == "缺货"].shape[0]) if summary is not None and not summary.empty else 0

        shortage_sku = report.get("最紧缺SKU", "") if report else ""
        ship_3d = report.get("未来3天应发货订单数", "0") if report else "0"

        lines = [
            f"**排单批次**：{batch_id}",
            f"**本次排单**：{total} 份合同",
            f"**待人工确认**：**{pending}** 份",
        ]
        if shortage_orders > 0:
            lines.append(f"**存在缺货**：{shortage_orders} 份合同需关注")
        if shortage_sku:
            lines.append(f"**紧缺物料**：{shortage_sku}")
        lines.append(f"**未来3天应发货**：{ship_3d} 单")
        lines.append("")
        lines.append(f"[查看排单总表](https://wl6wihmop1.feishu.cn/base/C5JzbAfnia0nT3sRvjucXgUGnDc?table=tbl09Z6C7wCGh3mW&view=vewiuVK8pH)")

        card = {
            "config": {"wide_screen_mode": True},
            "header": {"template": "blue", "title": {"content": "AI排单已完成，请确认", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
        }

        httpx.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "receive_id": planner_open_id,
                "msg_type": "interactive",
                "content": json.dumps(card, ensure_ascii=False),
            },
            timeout=15.0,
        )
        print("[个人通知] 排单完成私聊通知已发送")
    except Exception as e:
        print(f"[个人通知] 发送失败(不影响排单): {e}")


def dispatch_schedule_notifications(report: dict) -> dict:
    """排单完成后触发飞书三线机器人通知，失败不阻断排单主流程。"""
    if not report:
        return {"enabled": False, "ok": False, "error": "日报为空，未触发通知"}
    try:
        from ai_daily_agent import execute_triple_track_dispatch

        ok = execute_triple_track_dispatch(report)
        return {
            "enabled": True,
            "ok": bool(ok),
            "error": "" if ok else "一个或多个飞书 webhook 发送失败，请查看服务日志",
        }
    except SystemExit as e:
        return {"enabled": True, "ok": False, "error": f"通知模块配置缺失: {e}"}
    except Exception as e:
        return {"enabled": True, "ok": False, "error": str(e)}


# =========================
# 主排单API
# =========================
@app.post("/schedule")
@with_schedule_lock
def run_scheduler(payload: ScheduleRequest = ScheduleRequest()):
    """AI排单主函数"""
    t_start = time.time()

    reservation_config_error = ensure_reservation_table_id()
    if reservation_config_error:
        return {"error": reservation_config_error}

    # =========================
    # 从飞书并行加载最新数据
    # =========================
    print("[加载]从飞书加载最新数据")
    load_data()

    items = data_cache["items"]
    sku   = data_cache["sku"]
    inv   = data_cache["inv"]
    today = datetime.today().date()

    # ===== 实时读取 AI排单总表与库存预留表 =====
    try:
        old_summary = prepare_summary_status(fetch_bitable_to_df(TABLE_ID_DETAIL))
    except Exception as e:
        return {"error": f"读取 AI排单总表失败：{e}"}

    try:
        reservations = fetch_bitable_to_df(TABLE_ID_RESERVATION)
    except Exception as e:
        return {"error": f"读取 AI排单_库存预留表失败：{e}"}

    # ===== 生成排单ID & 本次批次号 =====
    batch_id = make_run_batch_id(reservations, today)
    expired_count, converted_count = cleanup_reservations(reservations, old_summary, batch_id)
    if expired_count or converted_count:
        print(f"[清洁] 预留清洁完成：过期 {expired_count} 条，转换 {converted_count} 条")
    reservations = fetch_bitable_to_df(TABLE_ID_RESERVATION)

    # ===== 表字段存在性快速检查 =====
    if "产品编码SKU" not in sku.columns:
        error_msg = f"❌ 错误：SKU表缺少'产品编码SKU'列。实际列名：{list(sku.columns)}"
        print(error_msg)
        return {"error": error_msg}

    if "SKU" not in inv.columns:
        error_msg = f"❌ 错误：库存表缺少'SKU'列。实际列名：{list(inv.columns)}"
        print(error_msg)
        return {"error": error_msg}

    # =========================
    # 实时计算库存占用：已确认订单需求 + 有效预留
    # =========================
    confirmed_contract_ids = get_confirmed_contract_ids(old_summary)
    locked_stock = summarize_confirmed_stock(old_summary, items)
    active_reserved_stock = summarize_effective_reservations(reservations)
    occupied_stock: Dict[str, float] = {}
    for sku_code in set(locked_stock) | set(active_reserved_stock):
        occupied_stock[sku_code] = locked_stock.get(sku_code, 0.0) + active_reserved_stock.get(sku_code, 0.0)
    print("[锁] 已确认订单占用库存：", locked_stock)
    print("[预留] 有效预留占用库存：", active_reserved_stock)

    # =========================
    # 第一步：回填"销售订单明细表"字段
    # =========================
    inv_latest_date = today
    try:
        inv_dt = pd.to_datetime(inv.get("库存日期", pd.Series(dtype=str)), errors="coerce")
        inv_dt = inv_dt.dropna()
        if not inv_dt.empty:
            inv_latest_date = inv_dt.max().date()
    except Exception:
        pass
    stock_base_date = max(today, inv_latest_date)

    backfill_rows = []
    for _, row in items.iterrows():
        row_dict = row.to_dict()
        sku_code = safe_str(row_dict.get("SKU编码", ""))
        contract_id = safe_str(row_dict.get("合同编号", ""))

        if not sku_code or not contract_id:
            continue
        if contract_id in confirmed_contract_ids:
            print(f"[锁] 合同 {contract_id} 已确认，销售订单明细表禁止重算和回填")
            continue

        inv_info, _inv_how = find_inventory_row(sku_code, row_dict, inv, sku_df=sku)
        if inv_info is None:
            print(f"⚠️ 库存未匹配 合同={contract_id} SKU={sku_code} 规格={safe_str(row_dict.get('规格', ''))[:50]} → {_inv_how}")

        reserved = occupied_stock.get(sku_code, 0.0)
        available = calc_available_stock(inv_info, reserved_qty=reserved)

        demand = to_num(row_dict.get("合同数量", 0))
        status, gap = calc_stock_status_and_gap(demand, available)

        rb800_flag = "是" if is_rb800_from_text(pick_model_text(row_dict)) else "否"

        eta_date = ""
        if status == "缺货":
            eta = calc_shortage_eta_date(sku_code=sku_code, sku_df=sku, base_date=stock_base_date)
            eta_date = eta.strftime("%Y-%m-%d")

        backfill_rows.append({
            "_record_id": safe_str(row_dict.get("_record_id", "")),
            "库存可用量": available,
            "缺口数量": gap,
            "库存状态": status,
            "预计到货日期": eta_date,
            "是否RB800": rb800_flag,
            "排单批次号": batch_id,
        })

    df_backfill = pd.DataFrame(backfill_rows)
    if df_backfill.empty:
        print("销售订单明细表无可回填记录（已确认订单已全部跳过或无有效明细）")
    else:
        print("正在批量回填销售订单明细表...")
        update_bitable_records(
            TABLE_ID_ITEMS,
            df_backfill,
            record_id_col="_record_id",
            numeric_cols=DETAIL_NUMERIC_COLS_DEFAULT,
            date_cols=DETAIL_DATE_COLS_DEFAULT,
        )

    # =========================
    # 第二步：重新读取已回填的明细表
    # =========================
    print("正在重新读取回填后的销售订单明细表...")
    df_detail_reread = fetch_bitable_to_df(TABLE_ID_ITEMS)
    df_detail = build_current_detail_for_summary(items, df_detail_reread, df_backfill)

    # 统一关键列类型（入口统一清洗，后续不再重复 astype）
    for required_col in ("合同编号", "SKU编码", "合同数量", "库存可用量", "缺口数量"):
        if required_col not in df_detail.columns:
            df_detail[required_col] = ""
    df_detail["合同编号"] = df_detail["合同编号"].astype(str).apply(safe_str)
    df_detail["SKU编码"] = df_detail["SKU编码"].astype(str).apply(safe_str)
    df_detail = df_detail[(df_detail["合同编号"] != "") & (df_detail["SKU编码"] != "")]
    df_detail["合同数量"] = df_detail["合同数量"].apply(to_num)
    df_detail["库存可用量"] = df_detail["库存可用量"].apply(to_num)
    df_detail["缺口数量"] = df_detail["缺口数量"].apply(to_num)
    if "库存状态" not in df_detail.columns:
        df_detail["库存状态"] = ""
    df_detail["库存状态"] = df_detail["库存状态"].astype(str).apply(safe_str)
    if "预计到货日期" not in df_detail.columns:
        df_detail["预计到货日期"] = ""

    # 本地缓存明细
    try:
        df_detail.to_csv(f"{CACHE_DIR}/ai_schedule_detail.csv", index=False)
    except Exception as e:
        print("⚠️ 明细缓存保存失败：", str(e))

    # ===== 第二步：生成"AI排单总表"（按合同汇总，一单一行） =====
    print('明细表列名:', df_detail.columns.tolist())
    print(df_detail[['合同编号', 'SKU编码', '库存状态', '预计到货日期']].head(3))

    # 构建 SKU→设备名称 映射（来自 SKU标准表）
    sku_device_name_map: Dict[str, str] = {}
    if sku is not None and not sku.empty and "产品编码SKU" in sku.columns:
        for _, sku_row in sku.iterrows():
            sku_code = safe_str(sku_row.get("产品编码SKU", ""))
            if sku_code and sku_code not in sku_device_name_map:
                dev_name = safe_str(sku_row.get("设备名称", ""))
                if not dev_name:
                    dev_name = safe_str(sku_row.get("设备型号", ""))
                sku_device_name_map[sku_code] = dev_name

    # 提取合同 → 客户名称、项目名称（优先销售订单主表，其次 items 明细表）
    contract_info_map: Dict[str, Dict[str, str]] = {}
    if TABLE_ID_MAIN:
        try:
            main_df = fetch_bitable_to_df(TABLE_ID_MAIN)
            if not main_df.empty:
                for _, row in main_df.iterrows():
                    c_id = safe_str(row.get("合同编号", ""))
                    if c_id and c_id not in contract_info_map:
                        contract_info_map[c_id] = {
                            "客户名称": safe_str(row.get("客户名称", "")),
                            "项目名称": safe_str(row.get("项目名称", "")),
                            "是否紧急订单": safe_str(row.get("是否紧急订单", "")),
                            "是否换货订单": safe_str(row.get("是否换货订单", "")),
                            "是否补发订单": safe_str(row.get("是否补发订单", "")),
                            "是否维修订单": safe_str(row.get("是否维修订单", "")),
                        }
                print(f"从销售订单主表读取 {len(contract_info_map)} 条合同信息")
        except Exception as e:
            print(f"读取销售订单主表失败: {e}，回退到 items 表")
    if not contract_info_map and not items.empty:
        for _, row in items.iterrows():
            c_id = safe_str(row.get("合同编号", ""))
            if c_id and c_id not in contract_info_map:
                contract_info_map[c_id] = {
                    "客户名称": safe_str(row.get("客户名称", "")),
                    "项目名称": safe_str(row.get("项目名称", "")),
                    "是否紧急订单": "",
                    "是否换货订单": "",
                    "是否补发订单": "",
                    "是否维修订单": "",
                }

    summary_rows = []
    if not df_detail.empty:
        def safe_val(x):
            if isinstance(x, float) and pd.isna(x):
                return None
            if isinstance(x, pd.Timestamp) and pd.isna(x):
                return None
            return x

        for contract_id, group in df_detail.groupby("合同编号"):
            # 人工锁单检查
            if not old_summary.empty:
                old_row = old_summary[old_summary["合同编号"] == contract_id]
                if not old_row.empty:
                    old_status = normalize_order_status(old_row.iloc[0].get("订单状态", "待确认"))
                    if old_status in {"已确认", "已发货"}:
                        print(f"[锁] 合同 {contract_id} 订单状态={old_status}，信任人工状态并跳过AI重算")
                        summary_rows.append(old_row.iloc[0].to_dict())
                        continue

            latest_eta, ship_date = calc_order_ship_date_for_group(group=group, sku_df=sku, today=today)
            if ship_date is None:
                ship_date = next_working_day(today)

            any_rb800 = False
            for _, r in group.iterrows():
                if is_rb800_from_text(pick_model_text(r.to_dict())):
                    any_rb800 = True
                    break

            cinfo = contract_info_map.get(contract_id, {})
            is_urgent = cinfo.get("是否紧急订单", "") == "是"
            is_exchange = cinfo.get("是否换货订单", "") == "是"
            is_resend = cinfo.get("是否补发订单", "") == "是"
            is_repair = cinfo.get("是否维修订单", "") == "是"

            if any_rb800:
                project_type = "远程控制项目"
            elif is_repair or is_urgent or is_exchange or is_resend:
                project_type = "特殊订单"
            else:
                project_type = "常规项目"

            sku_codes = group["SKU编码"].astype(str).apply(safe_str)
            sku_count = safe_val(sku_codes.nunique())
            total_qty = safe_val(group["合同数量"].sum())
            shortage_mask = group["库存状态"] == "缺货"
            shortage_rows = group[shortage_mask]
            shortage_sku_codes = sorted(normalize_shortage_sku_list(shortage_rows['SKU编码'].dropna().unique().tolist()))
            shortage_sku_names = []
            for sku_code in shortage_sku_codes:
                dev_name = sku_device_name_map.get(sku_code, "")
                if dev_name:
                    shortage_sku_names.append(dev_name)
                else:
                    shortage_sku_names.append(sku_code)
            shortage_sku_count = len(shortage_sku_codes)
            overall_status = "缺货" if shortage_sku_count else "有货"

            summary_rows.append({
                "合同编号": contract_id,
                "客户名称": contract_info_map.get(contract_id, {}).get("客户名称", ""),
                "项目名称": contract_info_map.get(contract_id, {}).get("项目名称", ""),
                "项目类型": project_type,
                "订单SKU总数": int(sku_count) if sku_count is not None else 0,
                "订单总数量": int(total_qty) if total_qty is not None else 0,
                "缺货SKU数": int(shortage_sku_count),
                "缺货SKU列表": ', '.join(shortage_sku_names),
                "整体状态": overall_status,
                "AI建议发货时间": date_to_yyyy_mm_dd(next_working_day(ship_date)) if ship_date else "",
                "AI风险": "",
                "AI建议": "",
                "排单批次号": batch_id,
                "订单状态": "待确认",
                "是否人工确认": "否",
                "人工确认发货时间": ""
            })

    summary = pd.DataFrame(summary_rows)
    if not summary.empty:
        summary["合同编号"] = summary["合同编号"].astype(str).apply(safe_str)
        summary = summary.drop_duplicates(subset=["合同编号"], keep="last")
    print("summary 列名：", list(summary.columns))

    # ===== 第三步：应用每日产能限制 =====
    summary, delayed_cnt = apply_capacity_scheduling(summary, today=today)
    if delayed_cnt:
        print(f"[产能] 产能限制生效：{delayed_cnt} 单被顺延到后续日期")

    # ===== AI 风险分析 =====
    if not summary.empty:
        from ai_service import analyze_order_risk
        from concurrent.futures import ThreadPoolExecutor, as_completed
        print("正在执行 AI 风险分析（并发）...")
        tasks = []  # list of (row_idx, order_info)
        for idx, (_, row) in enumerate(summary.iterrows()):
            c_id = safe_str(row.get("合同编号", ""))
            if not c_id:
                continue
            order_info = {
                "合同编号": c_id,
                "客户名称": safe_str(row.get("客户名称", "")),
                "项目名称": safe_str(row.get("项目名称", "")),
                "项目类型": safe_str(row.get("项目类型", "")),
                "订单SKU总数": int(to_num(row.get("订单SKU总数", 0))),
                "订单总数量": int(to_num(row.get("订单总数量", 0))),
                "缺货SKU数": int(to_num(row.get("缺货SKU数", 0))),
                "缺货SKU列表": safe_str(row.get("缺货SKU列表", "")),
                "整体状态": safe_str(row.get("整体状态", "")),
                "AI建议发货时间": safe_str(row.get("AI建议发货时间", "")),
            }
            tasks.append((idx, order_info))

        ai_results: Dict[int, tuple] = {}
        total_tasks = len(tasks)
        completed = 0
        with ThreadPoolExecutor(max_workers=3, thread_name_prefix="ai_risk") as executor:
            futures = {executor.submit(analyze_order_risk, t[1]): t[0] for t in tasks}
            for future in as_completed(futures):
                row_idx = futures[future]
                try:
                    result = future.result(timeout=60)
                    ai_results[row_idx] = (str(result.get("risk", "")), str(result.get("advice", "")))
                    completed += 1
                    print(f"  AI分析进度: {completed}/{total_tasks}")
                except Exception as e:
                    completed += 1
                    print(f"  AI分析失败 行{row_idx} ({completed}/{total_tasks}): {e}")
                    ai_results[row_idx] = ("", "")

        ai_risks, ai_advices = [], []
        for i in range(len(summary)):
            r = ai_results.get(i, ("", ""))
            ai_risks.append(r[0])
            ai_advices.append(r[1])
        summary["AI风险"] = ai_risks
        summary["AI建议"] = ai_advices
        print(f"AI 风险分析完成，分析了 {len(ai_results)} 条订单")

    # ===== 统一类型转换 =====
    def convert_summary_row(row):
        converted = {}
        for key, value in row.items():
            if key == "合同编号":
                converted[key] = str(value) if value is not None else ""
            elif key == "项目类型":
                converted[key] = str(value) if value is not None else ""
            elif key == "订单SKU总数":
                converted[key] = int(to_num(value)) if value is not None and value != "" else 0
            elif key == "订单总数量":
                converted[key] = int(to_num(value)) if value is not None and value != "" else 0
            elif key == "缺货SKU数":
                converted[key] = int(to_num(value)) if value is not None and value != "" else 0
            elif key == "缺货SKU列表":
                converted[key] = ", ".join(normalize_shortage_sku_list(value))
            elif key == "整体状态":
                converted[key] = str(value) if value is not None else ""
            elif key == "AI建议发货时间":
                converted[key] = date_to_yyyy_mm_dd(next_working_day(parse_date_to_date(value) or datetime.today().date()))
            elif key == "排单批次号":
                converted[key] = str(value) if value is not None else ""
            elif key == "订单状态":
                converted[key] = normalize_order_status(value)
            elif key == "是否人工确认":
                converted[key] = "是" if safe_str(value) == "是" else "否"
            elif key == "人工确认发货时间":
                converted[key] = date_to_yyyy_mm_dd(value)
            else:
                converted[key] = value
        converted["缺货SKU列表"] = ", ".join(normalize_shortage_sku_list(converted.get("缺货SKU列表", "")))
        converted["缺货SKU数"] = shortage_sku_count_from_list(converted.get("缺货SKU列表", ""))
        converted["订单状态"] = normalize_order_status(converted.get("订单状态", "待确认"))
        converted["是否人工确认"] = "是" if is_effective_manual_confirmed(converted) else "否"
        return converted

    converted_summary_rows = [convert_summary_row(row.to_dict()) for _, row in summary.iterrows()]
    summary = pd.DataFrame(converted_summary_rows)
    if not summary.empty:
        summary = summary.drop(columns=[c for c in summary.columns if isinstance(c, str) and c.startswith("_")], errors="ignore")

    # ===== 写入飞书结果表 =====
    print("\n【写入结果到飞书】")

    pending_contracts = []
    if not summary.empty and "订单状态" in summary.columns:
        pending_contracts = summary.loc[
            summary["订单状态"].apply(normalize_order_status) == "待确认",
            "合同编号",
        ].astype(str).apply(safe_str).tolist()
    released_count = release_old_contract_reservations(reservations, pending_contracts)
    reservation_rows = build_reservation_rows(summary, df_detail, batch_id)
    if not reservation_rows.empty:
        write_df_to_bitable(
            TABLE_ID_RESERVATION,
            reservation_rows,
            numeric_cols=RESERVATION_NUMERIC_COLS_DEFAULT,
            date_cols=RESERVATION_DATE_COLS_DEFAULT,
        )
    print(f"[预留] 库存预留刷新完成：释放旧预留 {released_count} 条，新增预留 {0 if reservation_rows.empty else len(reservation_rows)} 条")

    # 保存本地缓存
    summary.to_csv(f"{CACHE_DIR}/ai_schedule_summary.csv", index=False)
    print("排单缓存已保存")

    summary = summary.astype(object).where(pd.notnull(summary), None)
    df_detail = df_detail.astype(object).where(pd.notnull(df_detail), None)

    # 汇总数据 → 批量写入 AI排单总表
    print("正在写入 AI排单总表...")
    upsert_err = upsert_bitable_records_by_key(
        TABLE_ID_DETAIL,
        summary,
        key_col="合同编号",
        key_field_name="合同编号",
        numeric_cols=SUMMARY_NUMERIC_COLS_DEFAULT,
        date_cols=SUMMARY_DATE_COLS_DEFAULT,
    )

    # ===== 写入 AI发货总表（合同级发货记录） =====
    shipping_written = 0
    if TABLE_ID_SHIPPING and not summary.empty:
        print("正在写入 AI发货总表...")
        # 读取已存在的发货记录，保护已发货合同不被覆盖
        existing_shipped_ids: Set[str] = set()
        try:
            existing_shipping = fetch_bitable_to_df(TABLE_ID_SHIPPING)
            if not existing_shipping.empty and "是否发货" in existing_shipping.columns and "合同编号" in existing_shipping.columns:
                shipped_mask = existing_shipping["是否发货"].astype(str) == "是"
                existing_shipped_ids = set(existing_shipping.loc[shipped_mask, "合同编号"].astype(str).apply(safe_str))
        except Exception as e:
            print(f"读取 AI发货总表现有记录失败: {e}")
        shipping_rows = []
        for _, row in summary.iterrows():
            c_id = safe_str(row.get("合同编号", ""))
            if not c_id:
                continue
            shipping_rows.append({
                "合同编号": c_id,
                "客户名称": contract_info_map.get(c_id, {}).get("客户名称", ""),
                "项目名称": contract_info_map.get(c_id, {}).get("项目名称", ""),
                "是否发货": "是" if c_id in existing_shipped_ids else "否",
                "发货日期": "",
                "快递公司": "",
                "快递单号": "",
                "备注": "",
            })
        df_shipping = pd.DataFrame(shipping_rows)
        if not df_shipping.empty:
            df_shipping = df_shipping.drop_duplicates(subset=["合同编号"], keep="last")
            upsert_bitable_records_by_key(
                TABLE_ID_SHIPPING,
                df_shipping,
                key_col="合同编号",
                key_field_name="合同编号",
                date_cols={"发货日期"},
            )
            shipping_written = len(df_shipping)
            print(f"AI发货总表写入完成: {shipping_written} 条")
    elif not TABLE_ID_SHIPPING:
        print("TABLE_ID_SHIPPING 未配置，跳过 AI发货总表写入")

    # ===== 生成 AI排单日报 =====
    daily_report_written = 0
    notification_result = {"enabled": False, "ok": False, "error": "未生成日报，未触发通知"}
    if TABLE_ID_DAILY_REPORT:
        try:
            print("正在生成 AI排单日报...")
            report = generate_daily_report(
                items_df=df_detail,
                summary_df=summary,
                inv_df=inv,
                sku_df=sku,
                batch_id=batch_id,
            )
            if report:
                df_report = pd.DataFrame([report])
                write_df_to_bitable(
                    TABLE_ID_DAILY_REPORT,
                    df_report,
                    date_cols={"日期"},
                )
                daily_report_written = 1
                print("AI排单日报生成完成")
                print("正在触发飞书三线机器人通知...")
                notification_result = dispatch_schedule_notifications(report)
                if notification_result.get("ok"):
                    print("飞书三线机器人通知已触发")
                else:
                    print(f"飞书三线机器人通知失败(不影响排单): {notification_result.get('error')}")

                # ----- 私聊通知排单完成 -----
                _send_personal_notification(summary, batch_id, report)
        except Exception as e:
            print(f"AI排单日报生成失败(不影响排单): {e}")
    else:
        print("TABLE_ID_DAILY_REPORT 未配置，跳过日报生成")

    elapsed = time.time() - t_start
    print(f"排单总耗时: {elapsed:.1f}s")

    if upsert_err:
        print(f"❌ {upsert_err}")
        return {
            "msg": "明细已回填，但 AI排单总表写入失败",
            "error": upsert_err,
            "AI排单总表": public_records(summary),
            "回填明细行数": 0 if df_backfill is None else int(len(df_backfill)),
            "顺延订单数": int(delayed_cnt),
            "发货总表写入行数": shipping_written,
            "日报写入": daily_report_written,
            "飞书通知": notification_result,
            "TABLE_ID_DETAIL": TABLE_ID_DETAIL,
            "elapsed_s": round(elapsed, 1),
            "notes": "请核对多维表格中「AI排单总表」的真实 table_id（应以 tbl 开头），或通过环境变量 TABLE_ID_DETAIL 覆盖。",
        }

    return {
        "msg": "AI排单完成",
        "AI排单总表": public_records(summary),
        "回填明细行数": 0 if df_backfill is None else int(len(df_backfill)),
        "顺延订单数": int(delayed_cnt),
        "预留新增行数": 0 if reservation_rows.empty else int(len(reservation_rows)),
        "预留释放行数": int(released_count),
        "发货总表写入行数": shipping_written,
        "日报写入": daily_report_written,
        "飞书通知": notification_result,
        "批次号": batch_id,
        "elapsed_s": round(elapsed, 1),
        "notes": "回填已更新销售订单明细表原记录；总表按合同编号更新/新增；AI仅生成待确认；库存预留已按待确认订单刷新"
    }


# =========================
# 日报查询 API
# =========================
@app.get("/daily-report")
def daily_report():
    """获取今日最新排单日报数据（读 AI排单日报表，含今日数据校验）。"""
    try:
        report = get_today_report_row()
        return {
            "msg": "日报获取成功",
            "report": report
        }
    except ValueError as e:
        return {"error": str(e)}
    except Exception as e:
        return {"error": f"日报获取失败: {e}"}


# =========================
# 飞书 Webhook 触发接口（接收客户消息并自动回复）
# =========================

# ===== 多 App Bot 回复配置 =====
# 每个飞书应用的凭证映射，用于以正确的机器人身份回复消息
_BOT_CREDENTIALS = {
    "cli_aa87f7619df8dbb3": {  # 订单查询助手
        "app_id": os.getenv("CUSTOMER_BOT_APP_ID", ""),
        "app_secret": os.getenv("CUSTOMER_BOT_APP_SECRET", ""),
        "name": "订单查询助手",
    },
    "cli_a96c5d017d3a1cbb": {  # 供应链AI助手
        "app_id": os.getenv("FEISHU_APP_ID", ""),
        "app_secret": os.getenv("FEISHU_APP_SECRET", ""),
        "name": "供应链AI助手",
    },
}

# token 缓存（按 app_id）
_bot_token_cache: Dict[str, tuple] = {}


def _get_bot_reply_token(app_id: str = "") -> str:
    """根据飞书 app_id 获取对应的 tenant_access_token，用于回复消息。"""
    creds = _BOT_CREDENTIALS.get(app_id)
    if not creds:
        # 回退：使用 CUSTOMER_BOT 凭证
        creds = _BOT_CREDENTIALS.get("cli_aa87f7619df8dbb3", {})
    bot_app_id = creds.get("app_id", "")
    bot_app_secret = creds.get("app_secret", "")
    if not bot_app_id or not bot_app_secret:
        raise Exception(f"缺少 App {app_id} 的凭证配置")

    # 检查缓存
    cached = _bot_token_cache.get(bot_app_id)
    if cached and cached[1] > time.time() + 300:
        return cached[0]

    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": bot_app_id, "app_secret": bot_app_secret}, timeout=15.0,
    )
    data = _safe_http_json(resp, f"获取Bot回复Token({creds.get('name', app_id)})")
    if data.get("code") != 0:
        raise Exception(f"获取Bot回复Token失败: {data}")
    token = data["tenant_access_token"]
    _bot_token_cache[bot_app_id] = (token, time.time() + 6600)
    return token


def _send_bot_reply(open_id: str, card: dict, app_id: str = ""):
    """以指定飞书应用的身份发送互动卡片回复。"""
    token = _get_bot_reply_token(app_id)
    name = _BOT_CREDENTIALS.get(app_id, {}).get("name", app_id)
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": open_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        },
        timeout=15.0,
    )
    data = _safe_http_json(resp, f"发送Bot回复({name})")
    if data.get("code") != 0:
        print(f"[Bot回复:{name}] 发送失败: {data}")
        return False
    print(f"[Bot回复:{name}] 发送成功")
    return True
def _decrypt_event(encrypt_body: str) -> dict:
    """解密飞书加密事件。

    飞书开放平台 → 事件订阅 → 加密策略 中可获取 Encrypt Key。
    将 ENCRYPT_KEY 配置到 .env 中。

    加密格式：AES-256-CBC，数据 = 16字节IV + 密文
    """
    import hashlib
    import base64

    encrypt_key = os.getenv("FEISHU_EVENT_ENCRYPT_KEY", "")
    if not encrypt_key:
        raise Exception("收到加密事件，但未配置 FEISHU_EVENT_ENCRYPT_KEY。请在飞书开放平台获取 Encrypt Key 并添加到 .env")

    key_bytes = hashlib.sha256(encrypt_key.encode("utf-8")).digest()

    try:
        raw = base64.b64decode(encrypt_body)
    except Exception:
        raise Exception("事件加密数据 Base64 解码失败")

    if len(raw) < 32:
        raise Exception(f"加密数据太短 ({len(raw)} bytes)")

    iv = raw[:16]
    ciphertext = raw[16:]

    # AES-256-CBC 解密
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        cipher = Cipher(algorithms.AES(key_bytes), modes.CBC(iv))
        decryptor = cipher.decryptor()
        plaintext = decryptor.update(ciphertext) + decryptor.finalize()
    except ImportError:
        # 回退到纯 Python AES（如果没有 cryptography 库）
        raise Exception(
            "需要安装 cryptography 库来解密事件：pip install cryptography\n"
            "或直接在飞书开放平台关闭事件加密（推荐）"
        )

    # 去除 PKCS7 padding
    pad_len = plaintext[-1]
    if isinstance(pad_len, int) and 1 <= pad_len <= 16:
        plaintext = plaintext[:-pad_len]

    try:
        result = json.loads(plaintext.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise Exception(f"解密后 JSON 解析失败: {e}")

    return result


# 消息去重集合
_seen_webhook_message_ids: set = set()


# ==============================
# 供应链AI助手 — 消息处理
# ==============================
def _build_supply_chain_reply(open_id: str, user_input: str) -> dict:
    """处理供应链AI助手的消息，返回飞书互动卡片。

    面向企业内部员工，提供排单、库存、交期等供应链管理功能。
    """
    import re
    text = user_input.strip()

    # ---- 闲聊检测 ----
    chitchat_patterns = [
        r"^(hi|hello|你好|在吗|嗨|哈喽)[\s!！。.]*$",
        r"^(测试|test)$",
    ]
    for pat in chitchat_patterns:
        if re.match(pat, text, re.IGNORECASE):
            return {
                "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": (
                    "我是供应链AI助手，帮你掌握订单排单和交付全貌。\n\n"
                    "📊 **常用指令**\n"
                    "• 发送 **日报** — 查看今日排单日报\n"
                    "• 发送 **库存** — 查看库存与缺货情况\n"
                    "• 发送 **待确认** — 查看待人工确认的订单\n"
                    "• 发送 **合同编号** — 查询具体订单交期\n"
                    "• 发送 **延迟** — 查看延迟订单\n\n"
                    "直接输入关键词即可查询 👇"
                )}],
            }

    text_lower = text.lower()

    # ---- 日报 ----
    if any(kw in text_lower for kw in ["日报", "今日排单", "排单日报", "今日", "报告"]):
        try:
            report = get_today_report_row()
            return _build_report_card(report)
        except Exception as e:
            return _error_card(f"读取日报失败: {e}")

    # ---- 库存 ----
    if any(kw in text_lower for kw in ["库存", "缺货", "物料", "紧缺"]):
        try:
            load_data_if_needed()
            return _build_inventory_card()
        except Exception as e:
            return _error_card(f"读取库存失败: {e}")

    # ---- 待确认 ----
    if any(kw in text_lower for kw in ["待确认", "待处理", "未确认"]):
        try:
            load_data_if_needed()
            return _build_pending_orders_card()
        except Exception as e:
            return _error_card(f"查询待确认订单失败: {e}")

    # ---- 延迟 ----
    if any(kw in text for kw in ["延迟", "延期", "推迟"]):
        try:
            load_data_if_needed()
            return _build_delayed_orders_card()
        except Exception as e:
            return _error_card(f"查询延迟订单失败: {e}")

    # ---- 帮助 ----
    if any(kw in text_lower for kw in ["帮助", "help", "功能", "菜单", "说明"]):
        return {
            "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": (
                "**📊 供应链AI助手功能菜单**\n\n"
                "🔹 **日报** — 今日排单日报（订单数/发货/缺货/交期）\n"
                "🔹 **库存** — 库存状态（充足/预警/缺货SKU/紧缺物料）\n"
                "🔹 **待确认** — 需要人工确认的排单列表\n"
                "🔹 **延迟** — 预计延迟的订单明细\n"
                "🔹 **合同编号** — 输入合同编号查具体订单交期\n"
                "🔹 **触发排单** — 手动触发一次AI排单（需权限）\n\n"
                "👇 直接输入关键词即可"
            )}],
        }

    # ---- 合同编号查询 ----
    contract_match = re.search(r"([A-Za-z0-9\-_]{6,})", text)
    if contract_match:
        contract_id = contract_match.group(1)
        try:
            load_data_if_needed()
            return _build_contract_card(contract_id)
        except Exception as e:
            return _error_card(f"查询合同失败: {e}")

    # ---- 默认：模糊匹配 ----
    return _build_default_card()


def load_data_if_needed():
    """按需加载数据（如果缓存为空则从飞书读取）。"""
    if data_cache.get("items") is None or data_cache.get("inv") is None:
        load_data()


def _build_report_card(report: dict) -> dict:
    """日报卡片。"""
    items = [
        f"📊 **今日排单日报** ({report.get('日期', '-')})",
        "",
        f"订单总数：**{report.get('订单总数', '-')}** 单　|　今日新增：**{report.get('今日新增订单', '-')}** 单",
        f"已排单：**{report.get('已排单订单数', '-')}** 单　|　人工已确认：**{report.get('人工已确认排单数', '-')}** 单",
        f"未排单：**{report.get('未排单订单数', '-')}** 单",
        "",
        f"🚚 今日应发：**{report.get('今日应发货订单数', '-')}** 单　|　可发：**{report.get('今日可发货订单数', '-')}** 单",
        f"⚠️ 预计延迟：**{report.get('今日预计延迟订单数', '-')}** 单",
        "",
        f"📦 库存充足SKU：**{report.get('库存充足SKU数', '-')}**　|　预警：**{report.get('库存预警SKU数', '-')}**　|　缺货：**{report.get('库存缺货SKU数', '-')}**",
    ]

    shortage = report.get('最紧缺SKU', '')
    if shortage and shortage != '无':
        items.append(f"🔴 最紧缺：**{shortage}**（缺口 {report.get('最紧缺SKU缺口', '-')}）")

    items.append("")
    items.append(f"最长交期：**{report.get('最长交期订单', '-')}**　|　最晚发货：**{report.get('最晚发货日期', '-')}**")
    items.append(f"未来3天应发：**{report.get('未来3天应发货订单数', '-')}** 单　|　缺货风险：**{report.get('未来3天缺货订单数', '-')}** 单")
    items.append(f"批次号：{report.get('排单批次号', '-')}")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "blue", "title": {"content": "📊 AI排单日报", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(items)}],
    }


def _build_inventory_card() -> dict:
    """库存状态卡片。"""
    inv = data_cache.get("inv")
    sku_df = data_cache.get("sku")
    items_df = data_cache.get("items")

    if inv is None:
        return _error_card("库存数据未加载，请先触发排单或等待数据刷新")

    total_sku = inv["SKU"].nunique() if "SKU" in inv.columns else 0
    total_qty = int(inv["库存数量"].sum()) if "库存数量" in inv.columns else 0

    # 缺货 SKU（从明细表统计）
    shortage_list: List[tuple] = []
    if items_df is not None and not items_df.empty and "库存状态" in items_df.columns and "SKU编码" in items_df.columns:
        shortage_items = items_df[items_df["库存状态"].astype(str).str.strip() == "缺货"]
        if not shortage_items.empty:
            gap_by_sku: Dict[str, float] = {}
            for _, r in shortage_items.iterrows():
                sku = safe_str(r.get("SKU编码", ""))
                gap = to_num(r.get("缺口数量", 0))
                if sku and gap > 0:
                    gap_by_sku[sku] = gap_by_sku.get(sku, 0) + gap
            shortage_list = sorted(gap_by_sku.items(), key=lambda x: x[1], reverse=True)[:8]

    lines = [
        f"📦 **库存总览**",
        "",
        f"SKU总数：**{total_sku}**　|　总库存量：**{total_qty:,}**",
    ]

    if shortage_list:
        lines.append("")
        lines.append("🔴 **缺货物料 TOP{0}**".format(min(len(shortage_list), 8)))
        for sku_code, gap in shortage_list:
            # 尝试从 SKU 表获取设备名称
            name = ""
            if sku_df is not None and not sku_df.empty and "产品编码SKU" in sku_df.columns:
                match = sku_df[sku_df["产品编码SKU"].astype(str).str.strip() == sku_code]
                if not match.empty:
                    name = safe_str(match.iloc[0].get("设备名称", ""))
            display = name if name else sku_code
            gap_str = str(int(gap)) if gap == int(gap) else str(round(gap, 1))
            lines.append(f"• **{display}** — 缺 **{gap_str}** 个")
    else:
        lines.append("")
        lines.append("✅ 当前无缺货物料")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red" if shortage_list else "green",
                    "title": {"content": "📦 库存状态", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_pending_orders_card() -> dict:
    """待确认订单列表。"""
    try:
        summary = prepare_summary_status(fetch_bitable_to_df(TABLE_ID_DETAIL))
    except Exception:
        try:
            summary = data_cache.get("summary")
        except Exception:
            summary = pd.DataFrame()

    if summary is None or summary.empty:
        return {"header": {"template": "blue", "title": {"content": "📋 待确认订单", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "暂无待确认订单数据"}]}

    pending = summary[summary["订单状态"] == "待确认"].copy() if "订单状态" in summary.columns else pd.DataFrame()
    if pending.empty:
        return {"header": {"template": "green", "title": {"content": "✅ 待确认订单", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "当前没有待确认的订单，所有订单已处理。"}]}

    pending = pending.head(15)
    lines = [f"📋 **待确认订单 — {len(pending)} 份**", ""]
    for _, r in pending.iterrows():
        cid = safe_str(r.get("合同编号", "-"))
        status = safe_str(r.get("整体状态", ""))
        ship = safe_str(r.get("AI建议发货时间", ""))
        if len(ship) >= 10:
            ship = ship[:10]
        flag = "🔴" if "缺货" in status else "🟢"
        lines.append(f"{flag} {cid} — {ship}")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "orange", "title": {"content": "📋 待确认订单", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_delayed_orders_card() -> dict:
    """延迟订单列表。"""
    try:
        summary = prepare_summary_status(fetch_bitable_to_df(TABLE_ID_DETAIL))
    except Exception:
        return _error_card("无法读取排单总表")

    if summary is None or summary.empty or "整体状态" not in summary.columns:
        return {"header": {"template": "blue", "title": {"content": "⏳ 延迟订单", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "暂无延迟订单数据"}]}

    delayed = summary[summary["整体状态"].astype(str).str.strip() == "缺货"].head(15)
    if delayed.empty:
        return {"header": {"template": "green", "title": {"content": "✅ 交付状态", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "当前所有订单交付正常，无延迟订单。"}]}

    lines = [f"⚠️ **缺货/延迟订单 — {len(delayed)} 份**", ""]
    for _, r in delayed.iterrows():
        cid = safe_str(r.get("合同编号", "-"))
        ship = safe_str(r.get("AI建议发货时间", ""))[:10]
        manual = safe_str(r.get("人工确认发货时间", ""))[:10]
        display_ship = manual if manual else ship
        lines.append(f"🔴 {cid} — 预计 {display_ship}")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red", "title": {"content": "⚠️ 延迟订单", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_contract_card(contract_id: str) -> dict:
    """单个合同详情卡片。"""
    try:
        detail_df = fetch_bitable_to_df(TABLE_ID_DETAIL)
    except Exception:
        return _error_card("无法读取排单总表")

    if detail_df is None or detail_df.empty:
        return _error_card(f"未找到合同 {contract_id} 的信息")

    match = detail_df[detail_df["合同编号"].astype(str).str.strip() == contract_id]
    if match.empty:
        return {"header": {"template": "orange", "title": {"content": "🔍 未找到", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": f"未找到合同编号 **{contract_id}**，请确认编号是否正确"}]}

    r = match.iloc[0]
    status = safe_str(r.get("整体状态", "-"))
    order_status = safe_str(r.get("订单状态", "-"))
    ai_ship = safe_str(r.get("AI建议发货时间", ""))[:10]
    manual_ship = safe_str(r.get("人工确认发货时间", ""))[:10]
    risk = safe_str(r.get("AI风险", ""))
    advice = safe_str(r.get("AI建议", ""))

    status_emoji = "🟢" if "有货" in status else "🔴" if "缺货" in status else "⚪"
    ship_display = manual_ship if manual_ship else ai_ship if ai_ship else "待定"

    lines = [
        f"📄 **合同 {contract_id}**",
        "",
        f"整体状态：{status_emoji} **{status}**",
        f"订单状态：**{order_status}**",
        f"预计发货：**{ship_display}**",
    ]
    if manual_ship:
        lines.append(f"人工确认发货时间：**{manual_ship}**")
    if risk:
        lines.append(f"AI风险提示：{risk}")
    if advice:
        lines.append(f"AI建议：{advice}")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "blue", "title": {"content": "📄 订单详情", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_default_card() -> dict:
    """默认引导卡片。"""
    return {
        "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": (
            "我是供应链AI助手，帮你掌握订单排单和交付全貌。\n\n"
            "📊 **常用指令**\n"
            "• **日报** — 今日排单统计\n"
            "• **库存** — 库存与缺货物料\n"
            "• **待确认** — 待处理的排单\n"
            "• **延迟** — 延迟/缺货订单\n"
            "• **合同编号** — 查具体订单\n\n"
            "👇 直接输入关键词即可查询"
        )}],
    }


def _error_card(msg: str) -> dict:
    return {"header": {"template": "red", "title": {"content": "⚠️ 查询失败", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": str(msg)}]}


@app.post("/webhook/feishu")
async def feishu_webhook(request: Request):
    body = await request.json()

    # ---- 诊断：打印请求体顶层结构 ----
    body_keys = list(body.keys())
    print(f"[Webhook] 收到请求, 顶层字段: {body_keys}")

    # URL 验证（飞书配置事件订阅时的 challenge 校验）
    if "challenge" in body:
        print("飞书URL验证成功")
        return {"challenge": body["challenge"]}

    # 加密事件检测并解密
    if "encrypt" in body:
        print("[Webhook] 检测到加密事件，尝试解密...")
        try:
            body = _decrypt_event(body["encrypt"])
            print(f"[Webhook] 解密成功, 字段: {list(body.keys())}")
        except Exception as e:
            print(f"[Webhook] 解密失败: {e}")
            return {"msg": "ok"}

    # ---- 解析事件（兼容 V1 和 V2 两种格式） ----
    header = body.get("header", {})
    event = body.get("event", {})
    event_type = ""

    # V2 格式: {schema, header: {event_type, ...}, event: {sender: {sender_id: {open_id}}, message: {chat_type, message_type, content}}}
    if header and isinstance(event, dict):
        event_type = header.get("event_type", "")

    # V1 格式: {type: "event_callback", event: {type: "im.message.receive_v1", open_id, msg_type, text, ...}}
    if not event_type and isinstance(event, dict):
        event_type = event.get("type", "")

    print(f"[Webhook] event_type={event_type}")

    if event_type != "im.message.receive_v1":
        print(f"[Webhook] 忽略非消息事件: event_type={event_type}")
        return {"msg": "ok"}

    # ---- 提取消息字段（兼容 V1 / V2） ----
    if header and "message" in event:
        # V2 格式
        message = event.get("message", {})
        sender = event.get("sender", {})
        sender_id_obj = sender.get("sender_id", {})
        open_id = sender_id_obj.get("open_id", "")
        sender_type = sender.get("sender_type", "")
        chat_type = message.get("chat_type", "")
        message_type = message.get("message_type", "")
        content_str = message.get("content", "{}")
        message_id = message.get("message_id", "")
        # 解析 content JSON → text
        try:
            user_input = json.loads(content_str).get("text", "").strip()
        except json.JSONDecodeError:
            user_input = content_str.strip()
    else:
        # V1 格式: event 内平铺字段
        open_id = event.get("open_id", "")
        sender_type = event.get("sender_type", "user")
        chat_type = event.get("chat_type", "")
        message_type = event.get("msg_type", "text")
        user_input = (event.get("text") or "").strip()
        message_id = event.get("open_message_id", "")

    # 识别是哪个 App 收到了消息
    event_app_id = header.get("app_id", "") or event.get("app_id", "")
    bot_name = _BOT_CREDENTIALS.get(event_app_id, {}).get("name", event_app_id or "未知Bot")

    print(f"[Webhook] app={bot_name}({event_app_id}), open_id={open_id[:16] if open_id else '(空)'}..., "
          f"chat_type={chat_type}, sender_type={sender_type}, message_type={message_type}")
    print(f"[Webhook] text={user_input[:100] if user_input else '(空)'}")

    # ---- 过滤 ----
    if chat_type and chat_type != "p2p":
        print(f"[Webhook] 跳过非私聊: chat_type={chat_type}")
        return {"msg": "ok"}
    if sender_type and sender_type != "user":
        print(f"[Webhook] 跳过非用户: sender_type={sender_type}")
        return {"msg": "ok"}
    if message_type and message_type != "text":
        print(f"[Webhook] 跳过非文本: message_type={message_type}")
        return {"msg": "ok"}
    if not open_id:
        print("[Webhook] open_id 为空, 无法回复")
        return {"msg": "ok"}
    if not user_input:
        print("[Webhook] 消息文本为空")
        return {"msg": "ok"}

    # ---- 去重 ----
    if message_id:
        if message_id in _seen_webhook_message_ids:
            print(f"[Webhook] 重复消息已忽略: message_id={message_id}")
            return {"msg": "ok"}
        _seen_webhook_message_ids.add(message_id)
        if len(_seen_webhook_message_ids) > 5000:
            _seen_webhook_message_ids.clear()

    # ---- 根据 App 路由到不同处理逻辑 ----
    try:
        if event_app_id == "cli_a96c5d017d3a1cbb":
            # 供应链AI助手 → 内部排单/库存/交期管理
            card = _build_supply_chain_reply(open_id, user_input)
        else:
            # 订单查询助手 (cli_aa87f7619df8dbb3) 或未知 → 客户订单查询
            from customer_agent import process_message as bot_process_message
            card = bot_process_message(open_id, user_input)
        _send_bot_reply(open_id, card, app_id=event_app_id)
    except Exception as e:
        print(f"[Webhook] 处理异常: {e}")
        import traceback
        traceback.print_exc()

    return {"msg": "ok"}


if __name__ == "__main__":
    import uvicorn
    print("启动AI排单服务...")
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
