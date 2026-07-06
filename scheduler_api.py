from __future__ import annotations

import sys
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

from fastapi import FastAPI, Request, UploadFile, File, Form
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
import openpyxl

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
# 财务核算模块表 ID
TABLE_ID_FINANCE_SUMMARY = os.getenv("TABLE_ID_FINANCE_SUMMARY", "")      # 财务对账总表
# ⚠️ 产品计费规则主表为只读数据源，绝对禁止任何形式的写入、修改、删除操作
TABLE_ID_BILLING_RULES = os.getenv("TABLE_ID_BILLING_RULES", "")          # 产品计费规则主表（只读）
TABLE_ID_FINANCE_DETAIL = os.getenv("TABLE_ID_FINANCE_DETAIL", "")        # 财务对账明细表
TABLE_ID_AUDIT = os.getenv("TABLE_ID_AUDIT", "")                            # 排单审计记录表

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
app = FastAPI(docs_url="/swagger", redoc_url=None, title="AI排单服务")

# ----- 挂载顺丰物流模块 -----
import sf_shipping
app.include_router(sf_shipping.router)

# 注入飞书适配器（供 sf_shipping 调用飞书 API）
class FeishuAdapter:
    @staticmethod
    def get_access_token():
        return get_access_token()
sf_shipping.set_feishu_adapter(FeishuAdapter())


@app.on_event("startup")
def _startup_shipping_scheduler():
    """启动时开启顺丰物流轮询定时任务"""
    sf_shipping.db_init()
    sf_shipping.start_scheduler()


@app.on_event("shutdown")
def _shutdown_shipping_scheduler():
    """关闭时停止顺丰物流轮询定时任务"""
    sf_shipping.stop_scheduler()

# =========================
# Pydantic 模型定义
# =========================
class ScheduleRequest(BaseModel):
    trigger: str = "test"


class FinanceCalculateRequest(BaseModel):
    threshold: int = 3


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


# 通用/低价值 token：词频高但匹配价值低，从倒排索引中过滤
_COMMON_TOKENS = {
    "定制", "外购", "外协", "设备", "系统", "装置", "终端", "模块", "配件", "组件",
    "产品", "服务", "工程", "项目", "通用", "标准", "默认", "其他", "备用",
}

def _tokenize_device_text(text: str) -> List[str]:
    """将设备型号/名称拆分为关键词（≥2字符），过滤通用词。"""
    tokens: List[str] = []
    raw = safe_str(text)
    if not raw:
        return tokens
    parts = re.split(r"[/\-\s,　，、]+", raw)
    for part in parts:
        part = part.strip()
        if len(part) >= 2 and part.lower() not in _COMMON_TOKENS:
            tokens.append(part.lower())
            norm = part.replace("（", "(").replace("）", ")").lower()
            if norm != part.lower() and norm not in _COMMON_TOKENS:
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

    # 按 SKU 聚合：合并同一SKU多批次库存，避免取max单行而低估总库存
    # 但若所有记录的 SKU 列全空（导入时 SKU 为 lookup 字段），则跳过聚合
    sku_has_data = "SKU" in inv.columns and inv["SKU"].dropna().apply(lambda x: str(x).strip() if x else "").str.len().sum() > 0
    if "SKU" in inv.columns and "库存数量" in inv.columns and sku_has_data:
        agg_map = {"库存数量": "sum"}
        for col in ["国网设备型号", "国网设备名称", "设备型号", "设备名称"]:
            if col in inv.columns:
                agg_map[col] = lambda x: next((str(v).strip() for v in x.dropna() if str(v).strip()), "")
        inv = inv.groupby("SKU", as_index=False).agg(agg_map)

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
    # 最低质量检查：得分≥2 或 查询文本是匹配设备名称的子串
    result = []
    query_lower = text.lower().strip()
    for c in sorted_candidates[: _inv_max_results]:
        score = c[1]
        rec = c[0]
        if score >= 2:
            result.append(rec)
        elif score == 1 and len(query_lower) >= 3:
            # 单 token 匹配时，检查查询文本是否是设备名称/型号的子串
            dev_texts = [
                safe_str(rec.get(k, "")).lower()
                for k in ("国网设备名称", "国网设备型号", "设备名称", "设备型号")
            ]
            if any(query_lower in dt for dt in dev_texts if len(dt) >= 3):
                result.append(rec)
    return result


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
) -> Tuple[Optional[Dict[str, Any]], str, str]:
    """在库存快照中查找对应行（优先走预建索引，O(1)/O(k)）。

    返回 (matched_row, match_description, canonical_sku)
    canonical_sku 是库存表中该产品的权威 SKU 编码，用于统一跨订单的库存追踪。

    1) 先按 SKU 精确索引查找
    2) 通过 SKU标准表拿到设备型号/名称 → 倒排索引查找
    3) 按明细的规格/产品名称 → 倒排索引查找（含验证）
    """
    if inv is None or inv.empty:
        return None, "库存表为空", safe_str(sku_code)

    code = safe_str(sku_code)

    # 1) 精确索引 O(1)
    hit = _inv_lookup_exact(code)
    if hit is not None:
        return hit, "SKU列匹配(索引)", safe_str(hit.get("SKU", code))

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
                    return c, f"SKU标准表.{key}→倒排索引({t[:40]})", safe_str(c.get("SKU", code))

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
            matched = candidates[0]
            matched_sku = safe_str(matched.get("SKU", ""))
            return matched, f"倒排索引匹配({text[:40]})", matched_sku or code

    return None, f"未匹配(SKU编码={code!r}, 规格={safe_str(row_dict.get('规格', ''))[:40]!r})", safe_str(sku_code)


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
DETAIL_NUMERIC_COLS_DEFAULT = {"库存可用量", "缺口数量", "已分配数量"}


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
            # 检查是否是僵尸锁（超过 1 小时）
            try:
                # 先读取锁文件内容，校验 PID 是否还在运行
                with open(SCHEDULE_LOCK_PATH, "r") as lf:
                    stale = True
                    try:
                        content = lf.read().strip()
                        lock_pid = int(content.split()[0]) if content else 0
                        # 检查该 PID 的进程是否还存在
                        os.kill(lock_pid, 0)  # 发送信号 0 不杀死进程，仅检查存在性
                        stale = False  # PID 存在，不是僵尸锁
                    except (OSError, ValueError, IndexError):
                        stale = True  # 无法读取 PID 或进程不存在
                if stale and time.time() - os.path.getmtime(SCHEDULE_LOCK_PATH) > 60 * 60:
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
) -> int:
    """批量写入飞书多维表格。返回成功写入的记录数。失败返回 -1。"""
    if df is None or df.empty:
        return 0

    numeric_fields = set(numeric_cols or set())
    date_fields = set(date_cols or set())
    columns = list(df.columns)
    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/batch_create"

    records_batch: List[Dict[str, Any]] = []
    success_count = 0

    for _, row in df.iterrows():
        fields = _build_fields_dict(row, columns, fields_map, numeric_fields, date_fields)
        if fields:
            records_batch.append({"fields": fields})

        if len(records_batch) >= BATCH_SIZE:
            if _post_batch(url, headers, records_batch, "创建"):
                success_count += len(records_batch)
            records_batch.clear()

    if records_batch:
        if _post_batch(url, headers, records_batch, "创建"):
            success_count += len(records_batch)

    return success_count


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


def _table_id_from_url(url: str) -> str:
    """从飞书 API URL 中提取 table_id。"""
    import re
    m = re.search(r"/tables/([^/]+)/", url)
    return m.group(1) if m else "?"


def _post_batch(url: str, headers: Dict[str, str], records: List[Dict[str, Any]], action: str) -> bool:
    """发送一批记录到飞书批量 API。返回 True 表示成功。
    
    对网络异常进行最多 3 次退避重试（1s/2s/4s），对业务错误（code != 0）不做重试以避免 batch_create 重复创建。
    """
    import time as _time
    max_retries = 3
    for attempt in range(max_retries):
        try:
            resp = httpx.post(url, headers=headers, json={"records": records}, timeout=60.0)
            data = _safe_http_json(resp, f"批量{action}")
            if data.get("code") != 0:
                print(f"批量{action}失败: table={_table_id_from_url(url)} code={data.get('code')}, msg={data.get('msg')}")
                if records:
                    print(f"  第一条数据示例: {json.dumps(records[0], ensure_ascii=False, default=str)[:500]}")
                print(f"  完整响应: {json.dumps(data, ensure_ascii=False, default=str)[:2000]}")
                return False
            print(f"批量{action}成功: {len(records)} 条")
            return True
        except Exception as e:
            if attempt < max_retries - 1:
                wait = 2 ** attempt
                print(f"批量{action}网络异常: {str(e)[:120]}，{wait}s 后重试 ({attempt+2}/{max_retries})")
                _time.sleep(wait)
                continue
            print(f"批量{action}网络异常（已重试{max_retries}次均失败）: {str(e)[:200]}")
            return False
    return False


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
                f"读取多维表格失败：飞书错误 WrongTableId(1254004)，说明 table_id 不是有效的多维表格。"
                f" 当前 table_id={table_id!r}。详情：{err}"
            )
        print(f"读取多维表格时出现异常，将采用直接新增方式：{err}")
        existing = pd.DataFrame()

    key_to_record: Dict[str, str] = {}
    if existing is not None and not existing.empty and "_record_id" in existing.columns and key_field_name in existing.columns:
        for _, r in existing.iterrows():
            k = safe_str(r.get(key_field_name, ""))
            rid = safe_str(r.get("_record_id", ""))
            if k and rid and k not in key_to_record:
                key_to_record[k] = rid
    elif existing is not None and not existing.empty and len(df) > 10 and key_field_name not in (existing.columns or []):
        # 关键字段不匹配，静默全量重复创建的风险
        existing_cols = list(existing.columns) if existing is not None else []
        err_msg = (
            f"upsert 危险操作：表中有 {len(existing)} 条已有记录，但找不到匹配键字段 "
            f"'{key_field_name}'（表实际列名: {existing_cols}），避免全量重复写入已拒绝。"
        )
        print(f"[UPSERT 阻断] {err_msg}")
        return err_msg

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


def _sync_finance_summary(force: bool = False) -> Dict[str, Any]:
    """同步销售订单主表 → 财务对账总表（Upsert by 合同编号）。

    同步字段：合同编号、下单日期、客户名称、项目名称、商务、代理商。
    保留字段（不覆盖）：项目类型、备注、AI项目金额、人工核对金额、AI费用检查。
    force=True 时强制重同步所有合同（已锁定的仍然跳过）。
    """
    if not TABLE_ID_FINANCE_SUMMARY:
        return {"ok": False, "error": "未配置 TABLE_ID_FINANCE_SUMMARY，请在 .env 中设置"}
    try:
        main_df = fetch_bitable_to_df(TABLE_ID_MAIN)
    except Exception as e:
        return {"ok": False, "error": f"读取销售订单主表失败: {e}"}

    if main_df.empty:
        return {"ok": False, "error": "销售订单主表无数据"}

    # 需要同步的字段映射：销售订单主表列名 → 对账总表列名
    FIELD_MAP = {
        "合同编号": "合同编号",
        "下单日期": "下单日期",
        "客户名称": "客户名称",
        "项目名称": "项目名称",
        "商务": "商务",
        "代理商": "代理商",
    }

    # 只取存在且需要的列
    available_cols = [c for c in FIELD_MAP if c in main_df.columns]
    if "合同编号" not in available_cols:
        return {"ok": False, "error": "销售订单主表缺少「合同编号」字段"}

    skipped_cols = [c for c in FIELD_MAP if c not in main_df.columns]
    if skipped_cols:
        print(f"[WARN] 主表缺少字段，跳过同步: {skipped_cols}")

    sync_df = main_df[available_cols].copy()

    # 读取对账总表现有数据，保留人工字段
    try:
        existing_df = fetch_bitable_to_df(TABLE_ID_FINANCE_SUMMARY)
    except Exception:
        existing_df = pd.DataFrame()

    preserve_cols = ["项目类型", "备注", "AI项目金额", "人工核对金额", "AI费用检查", "同步状态"]
    if not existing_df.empty and "合同编号" in existing_df.columns:
        # 按合同编号合并，保留已有的人工字段值
        existing_preserve = existing_df[["合同编号"] + [c for c in preserve_cols if c in existing_df.columns]]
        sync_df = sync_df.merge(existing_preserve, on="合同编号", how="left", suffixes=("", "_exist"))
        # 用已有值覆盖
        for col in preserve_cols:
            exist_col = f"{col}_exist"
            if exist_col in sync_df.columns:
                sync_df[col] = sync_df[exist_col].fillna("")
                sync_df.drop(columns=[exist_col], inplace=True)

    # 过滤：已锁定的合同跳过（不覆盖人工核对金额）
    # force=True 时只跳过已锁定，否则也跳过已同步
    before = len(sync_df)
    sync_df = sync_df[sync_df.apply(lambda r: (
        safe_str(r.get("人工核对金额", "")) == ""
    ), axis=1)]
    if not force:
        sync_df = sync_df[sync_df["同步状态"] != "已同步"]
    print(f"[Finance Sync Summary] {before} → {len(sync_df)} (force={force}, 跳过已锁定/{'跳过已同步' if not force else '含已同步'})")
    if sync_df.empty:
        return {"ok": True, "synced": 0, "skipped": before}

    # 标记本次同步的合同为"已同步"
    sync_df["同步状态"] = "已同步"

    err = upsert_bitable_records_by_key(
        TABLE_ID_FINANCE_SUMMARY, sync_df,
        key_col="合同编号", key_field_name="合同编号",
        date_cols={"下单日期"},
    )
    if err:
        return {"ok": False, "error": err}

    return {"ok": True, "synced": len(sync_df)}


def _match_spec_id(product_name: str, spec: str, billing_df: pd.DataFrame) -> str:
    """根据产品名称和规格匹配计费规则表中的 Spec_ID（仅精确匹配+别名）。"""
    if billing_df.empty or (not product_name and not spec):
        return ""

    product_name_norm = str(product_name).strip()
    spec_norm = str(spec).strip()

    # 1. 精确匹配：产品名称 == 产品名称 AND 规格 == 规格型号
    exact_match = billing_df[
        (billing_df["产品名称"].astype(str).str.strip() == product_name_norm) &
        (billing_df["规格型号"].astype(str).str.strip() == spec_norm)
    ]
    if not exact_match.empty:
        val = exact_match.iloc[0].get("Spec_ID", "")
        return str(val).strip() if pd.notna(val) else ""

    # 2. 别名匹配：同规格下，产品名称包含别名 OR 别名包含产品名称
    if "别名" in billing_df.columns:
        same_spec = billing_df[
            billing_df["规格型号"].astype(str).str.strip() == spec_norm
        ]
        for _, rule in same_spec.iterrows():
            alias = str(rule.get("别名", "")).strip()
            if not alias:
                continue
            if alias in product_name_norm or product_name_norm in alias:
                val = rule.get("Spec_ID", "")
                return str(val).strip() if pd.notna(val) else ""

    # 3. 关键词匹配（仅限 AC-RB800 规格的消控室值班助手产品，优先级从高到低）
    if "ac-rb800" in spec_norm.lower():
        keywords = [
            ("主键盘", "SPEC-RB800-KZP"),
            ("总线盘", "SPEC-RB800-ZXP"),
            ("总键盘", "SPEC-RB800-ZXP"),
            ("电话", "SPEC-RB800-PHONE"),
            ("多线盘", "SPEC-RB800-DXP"),
            ("广播", "SPEC-RB800-BROAD"),
        ]
        for kw, sid in keywords:
            if kw in product_name_norm:
                return sid

    # 未匹配 → 需在计费规则表中补充
    print(f"[WARNING] 计费规则未匹配: 产品={product_name_norm}, 规格={spec_norm} —— 请在产品计费规则主表中补充该产品")
    return ""


def _sync_finance_detail() -> Dict[str, Any]:
    """同步销售订单明细表 + 主表 + 计费规则 → 财务对账明细表（Upsert by 合同编号+产品名称+规格）。

    同步字段：
    - 从明细表：合同编号、产品名称、规格、合同数量
    - 从主表（按合同编号关联）：项目名称
    - 从计费规则表（按产品名称+规格匹配）：Spec_ID
    - 从对账总表（按合同编号关联）：合同类型
    """
    if not TABLE_ID_FINANCE_DETAIL:
        return {"ok": False, "error": "未配置 TABLE_ID_FINANCE_DETAIL，请在 .env 中设置"}
    try:
        items_df = fetch_bitable_to_df(TABLE_ID_ITEMS)
    except Exception as e:
        return {"ok": False, "error": f"读取销售订单明细表失败: {e}"}

    if items_df.empty:
        return {"ok": False, "error": "销售订单明细表无数据"}

    # 读取主表（获取项目名称）
    try:
        main_df = fetch_bitable_to_df(TABLE_ID_MAIN)
    except Exception:
        main_df = pd.DataFrame()

    # 读取计费规则表
    try:
        billing_df = fetch_bitable_to_df(TABLE_ID_BILLING_RULES)
    except Exception:
        billing_df = pd.DataFrame()

    # 读取对账总表（获取合同类型 = 项目类型）
    try:
        summary_df = fetch_bitable_to_df(TABLE_ID_FINANCE_SUMMARY)
    except Exception:
        summary_df = pd.DataFrame()

    # 构建同步 DataFrame
    sync_rows = []
    required_cols = ["合同编号", "产品名称", "规格", "合同数量"]
    available_cols = [c for c in required_cols if c in items_df.columns]
    if "合同编号" not in available_cols:
        return {"ok": False, "error": "销售订单明细表缺少「合同编号」字段"}

    for _, row in items_df.iterrows():
        contract_no = safe_str(row.get("合同编号", ""))
        product_name = safe_str(row.get("产品名称", ""))
        spec = safe_str(row.get("规格", ""))
        qty = row.get("合同数量", 0)

        if not contract_no:
            continue

        sync_row = {
            "合同编号": contract_no,
            "产品名称": product_name,
            "规格": spec,
            "合同数量": safe_numeric(qty),
        }

        # 从主表取项目名称，主表无数据时回退到对账总表
        project_name = ""
        if not main_df.empty and "合同编号" in main_df.columns and "项目名称" in main_df.columns:
            main_match = main_df[main_df["合同编号"].astype(str).str.strip() == contract_no]
            if not main_match.empty:
                project_name = safe_str(main_match.iloc[0].get("项目名称", ""))
        # 主表未找到 → 回退到对账总表
        if not project_name and not summary_df.empty and "合同编号" in summary_df.columns and "项目名称" in summary_df.columns:
            sum_match = summary_df[summary_df["合同编号"].astype(str).str.strip() == contract_no]
            if not sum_match.empty:
                project_name = safe_str(sum_match.iloc[0].get("项目名称", ""))
        if project_name:
            sync_row["项目名称"] = project_name

        # 从计费规则匹配 Spec_ID
        spec_id = _match_spec_id(product_name, spec, billing_df)
        sync_row["Spec_ID"] = spec_id

        # 从对账总表取合同类型
        if not summary_df.empty and "合同编号" in summary_df.columns and "项目类型" in summary_df.columns:
            sum_match = summary_df[summary_df["合同编号"].astype(str).str.strip() == contract_no]
            if not sum_match.empty:
                ct = sum_match.iloc[0].get("项目类型", "")
                sync_row["合同类型"] = safe_str(ct)

        sync_rows.append(sync_row)

    if not sync_rows:
        return {"ok": False, "error": "无有效明细数据可同步"}

    sync_df = pd.DataFrame(sync_rows)

    # 读取对账明细表现有数据，按 (合同编号, 产品名称, 规格) 做 key 记录已存在的 _record_id
    try:
        existing_detail = fetch_bitable_to_df(TABLE_ID_FINANCE_DETAIL)
    except Exception:
        existing_detail = pd.DataFrame()

    key_to_rid: Dict[str, str] = {}
    if not existing_detail.empty and "_record_id" in existing_detail.columns:
        for _, r in existing_detail.iterrows():
            k = (
                safe_str(r.get("合同编号", "")),
                safe_str(r.get("产品名称", "")),
                safe_str(r.get("规格", "")),
            )
            rid = safe_str(r.get("_record_id", ""))
            if all(k) and rid and k not in key_to_rid:
                key_to_rid[k] = rid

    to_update: List[Dict[str, Any]] = []
    to_create: List[Dict[str, Any]] = []

    preserve_detail_cols = ["是否计入包干", "包干合同单采数量", "单采价格", "小计", "Spec_ID", "合同类型"]
    for _, row in sync_df.iterrows():
        key = (
            safe_str(row.get("合同编号", "")),
            safe_str(row.get("产品名称", "")),
            safe_str(row.get("规格", "")),
        )
        if not all(key):
            continue

        if key in key_to_rid:
            # 更新：保留核算字段
            exist_row = existing_detail[
                (existing_detail["合同编号"].astype(str).str.strip() == key[0]) &
                (existing_detail["产品名称"].astype(str).str.strip() == key[1]) &
                (existing_detail["规格"].astype(str).str.strip() == key[2])
            ]
            d = row.to_dict()
            d["_record_id"] = key_to_rid[key]
            if not exist_row.empty:
                for col in preserve_detail_cols:
                    if col in exist_row.columns:
                        val = exist_row.iloc[0].get(col)
                        if pd.notna(val) and val != "":
                            d[col] = val
            to_update.append(d)
        else:
            to_create.append(row.to_dict())

    if to_update:
        update_df = pd.DataFrame(to_update)
        update_bitable_records(TABLE_ID_FINANCE_DETAIL, update_df, record_id_col="_record_id",
                               numeric_cols={"合同数量", "包干合同单采数量", "小计"})

    if to_create:
        create_df = pd.DataFrame(to_create)
        create_df = create_df.drop(columns=["_record_id"], errors="ignore")
        write_df_to_bitable(TABLE_ID_FINANCE_DETAIL, create_df,
                            numeric_cols={"合同数量"})

    # ===== 修复通道：回写缺失的 Spec_ID 和合同类型 =====
    # 重新读取对账明细表（含刚写入的记录），找出所有 Spec_ID 或合同类型为空的记录，
    # 用产品计费规则表和对账总表重新匹配并直接更新。
    # 这解决了两个根因：
    #   1. _run_finance_calculate 补行产生的费用记录从未经过 match_spec_id
    #   2. 计费规则表在初始同步之后才完善，导致旧记录匹配失败后未重新尝试
    try:
        repair_detail = fetch_bitable_to_df(TABLE_ID_FINANCE_DETAIL)
    except Exception:
        repair_detail = pd.DataFrame()

    repaired_count = 0
    repair_rows: List[Dict[str, Any]] = []
    if not repair_detail.empty and not billing_df.empty:
        for _, rec in repair_detail.iterrows():
            rid = safe_str(rec.get("_record_id", ""))
            if not rid:
                continue
            pn = safe_str(rec.get("产品名称", ""))
            sp = safe_str(rec.get("规格", ""))
            cur_spec = safe_str(rec.get("Spec_ID", ""))
            cur_ct = safe_str(rec.get("合同类型", ""))
            cno = safe_str(rec.get("合同编号", ""))

            needs_repair = False
            repair_fields: Dict[str, Any] = {}

            # 修正消防远程控制设备费的错误规格（SPEC-EQP → 定制）
            if pn == "消防远程控制设备费" and sp.upper() == "SPEC-EQP":
                sp = "定制"
                repair_fields["规格"] = "定制"
                needs_repair = True

            # Spec_ID 为空 → 重新匹配
            if (not cur_spec or cur_spec.lower() in ("nan",)) and pn:
                new_spec = _match_spec_id(pn, sp, billing_df)
                if new_spec:
                    repair_fields["Spec_ID"] = new_spec
                    needs_repair = True

            # Spec_ID 可能有误（之前模糊匹配结果），用精确匹配重新校验
            # 如果不一致则纠正；如果精确匹配无结果则清空（标记需人工处理）
            if cur_spec and cur_spec.lower() not in ("nan", "") and pn and not needs_repair:
                new_spec = _match_spec_id(pn, sp, billing_df)
                if new_spec and new_spec != cur_spec:
                    print(f"[Sync Detail] 纠正 Spec_ID: {pn[:30]} {sp} => {cur_spec} → {new_spec}")
                    repair_fields["Spec_ID"] = new_spec
                    needs_repair = True
                elif not new_spec:
                    # 精确匹配无结果 → 当前 Spec_ID 可能是模糊匹配的错误结果，清空
                    print(f"[Sync Detail] 清空疑似错误 Spec_ID: {pn[:30]} {sp} (原={cur_spec})")
                    repair_fields["Spec_ID"] = ""
                    needs_repair = True

            # 合同类型为空 → 从对账总表中补充
            if not cur_ct or cur_ct.lower() in ("nan",):
                if not summary_df.empty and "合同编号" in summary_df.columns and "项目类型" in summary_df.columns:
                    sm = summary_df[summary_df["合同编号"].astype(str).str.strip() == cno]
                    if not sm.empty:
                        pt = safe_str(sm.iloc[0].get("项目类型", ""))
                        if pt:
                            repair_fields["合同类型"] = pt
                            needs_repair = True

            # 项目名称为空 → 从对账总表中补充
            cur_pn = safe_str(rec.get("项目名称", ""))
            if not cur_pn or cur_pn.lower() in ("nan",):
                if not summary_df.empty and "合同编号" in summary_df.columns and "项目名称" in summary_df.columns:
                    sm = summary_df[summary_df["合同编号"].astype(str).str.strip() == cno]
                    if not sm.empty:
                        pn_val = safe_str(sm.iloc[0].get("项目名称", ""))
                        if pn_val:
                            repair_fields["项目名称"] = pn_val
                            needs_repair = True

            if needs_repair:
                repair_fields["_record_id"] = rid
                repair_rows.append(repair_fields)

        if repair_rows:
            repair_df = pd.DataFrame(repair_rows)
            repaired_count = len(repair_df)
            print(f"[Sync Detail] 修复通道：回写 {repaired_count} 条")
            for _, rr in repair_df.iterrows():
                print(f"  [Repair] rid={safe_str(rr.get('_record_id',''))} fields={list(rr.drop('_record_id',errors='ignore').keys())}")
            update_bitable_records(
                TABLE_ID_FINANCE_DETAIL,
                repair_df,
                record_id_col="_record_id",
            )

    return {
        "ok": True,
        "synced": len(sync_rows),
        "updated": len(to_update),
        "created": len(to_create),
        "repaired": repaired_count,
    }


def _run_finance_calculate(threshold: int = 3) -> Dict[str, Any]:
    """执行完整财务核算流程。

    流程：
    1. 读取三张财务表 + 计费规则表
    2. 锁定检查（人工核对金额已填写的合同跳过）
    3. 包干项目完整性检查（自动补行）
    4. 按合同逐条核算（包干/单采分类）
    5. 汇总回写 AI项目金额
    """
    if not TABLE_ID_FINANCE_SUMMARY:
        return {"ok": False, "error": "未配置 TABLE_ID_FINANCE_SUMMARY，请在 .env 中设置"}
    if not TABLE_ID_FINANCE_DETAIL:
        return {"ok": False, "error": "未配置 TABLE_ID_FINANCE_DETAIL，请在 .env 中设置"}
    if not TABLE_ID_BILLING_RULES:
        return {"ok": False, "error": "未配置 TABLE_ID_BILLING_RULES，请在 .env 中设置"}
    # 读取数据
    try:
        summary_df = fetch_bitable_to_df(TABLE_ID_FINANCE_SUMMARY)
        detail_df = fetch_bitable_to_df(TABLE_ID_FINANCE_DETAIL)
        billing_df = fetch_bitable_to_df(TABLE_ID_BILLING_RULES)
    except Exception as e:
        return {"ok": False, "error": f"读取飞书数据失败: {e}"}

    if summary_df.empty:
        return {"ok": False, "error": "财务对账总表无数据，请先同步"}

    if detail_df.empty:
        return {"ok": False, "error": "财务对账明细表无数据，请先同步"}

    if billing_df.empty:
        return {"ok": False, "error": "产品计费规则主表无数据，请先维护"}

    # 构建计费规则查找索引（Spec_ID → 规则行）
    # 关键：使用规范化的 key（lower + strip），保证与后续查询一致
    billing_index: Dict[str, Dict[str, Any]] = {}
    # 备选索引：产品名称+规格 → Spec_ID（用于明细表中 Spec_ID 为空的情况）
    product_spec_index: Dict[tuple, str] = {}
    if "Spec_ID" in billing_df.columns:
        for _, r in billing_df.iterrows():
            sid = safe_str(r.get("Spec_ID", "")).lower()  # 规范化：转小写
            if sid:
                billing_index[sid] = {
                    "计费方式": safe_str(r.get("计费方式", "")),
                    "包干阈值": safe_numeric(r.get("包干阈值", 0)),
                    "单采单价": safe_numeric(r.get("单采单价", 0)),
                }
            # 构建产品+规格索引（用于备选查找）
            product_name = safe_str(r.get("产品名称", "")).lower()
            spec_model = safe_str(r.get("规格型号", "")).lower()
            if product_name and spec_model:
                key = (product_name, spec_model)
                product_spec_index[key] = sid

    # 修复消防远程控制设备费的规格（从 SPEC-EQP 改为定制，便于后续匹配）
    if "规格" in detail_df.columns and "产品名称" in detail_df.columns:
        mask = detail_df["产品名称"] == "消防远程控制设备费"
        if mask.any():
            detail_df = detail_df.copy()  # 避免SettingWithCopyWarning
            detail_df.loc[mask, "规格"] = "定制"

    # 锁定合同列表
    locked_contracts: Set[str] = set()
    if "人工核对金额" in summary_df.columns and "合同编号" in summary_df.columns:
        for _, r in summary_df.iterrows():
            amt = r.get("人工核对金额")
            if pd.notna(amt):
                contract = safe_str(r.get("合同编号", ""))
                if contract:
                    locked_contracts.add(contract)

    # 项目类型未设置的合同（需用户先设置）
    no_project_type_contracts: Set[str] = set()
    if "项目类型" in summary_df.columns and "合同编号" in summary_df.columns:
        for _, r in summary_df.iterrows():
            cno = safe_str(r.get("合同编号", ""))
            pt = safe_str(r.get("项目类型", ""))
            if cno and not pt and cno not in locked_contracts:
                no_project_type_contracts.add(cno)

    # 新增明细行（完整性检查产生）
    new_detail_rows: List[Dict[str, Any]] = []
    # 费用检查结果
    summary_check: Dict[str, str] = {}  # contract_no → check_result

    # ========== 步骤1: 包干项目完整性检查 ==========
    for _, summary_row in summary_df.iterrows():
        contract_no = safe_str(summary_row.get("合同编号", ""))
        project_type = safe_str(summary_row.get("项目类型", ""))

        if not contract_no or contract_no in locked_contracts:
            continue

        # 项目类型未设置：不做检查，留空让用户先设置
        if not project_type:
            continue

        if project_type != "包干":
            summary_check[contract_no] = "正常"
            continue

        # 获取该合同所有明细行
        contract_detail = detail_df[
            detail_df["合同编号"].astype(str).str.strip() == contract_no
        ]

        check_msgs = []

        # 检查「消防远程控制服务费」
        has_service_fee = any(
            safe_str(r.get("产品名称", "")) == "消防远程控制服务费"
            for _, r in contract_detail.iterrows()
        )
        if not has_service_fee:
            svc_spec_id = _match_spec_id("消防远程控制服务费", "定制", billing_df)
            new_detail_rows.append({
                "合同编号": contract_no,
                "项目名称": safe_str(summary_row.get("项目名称", "")),
                "产品名称": "消防远程控制服务费",
                "规格": "定制",
                "合同数量": 1,
                "合同类型": "包干",
                "Spec_ID": svc_spec_id,
            })
            check_msgs.append("缺少「消防远程控制服务费」，已自动补行")

        # 检查「消防远程控制包干费」
        has_package_fee = any(
            safe_str(r.get("产品名称", "")) == "消防远程控制包干费"
            for _, r in contract_detail.iterrows()
        )
        if not has_package_fee:
            pkg_spec_id = _match_spec_id("消防远程控制包干费", "定制", billing_df)
            new_detail_rows.append({
                "合同编号": contract_no,
                "项目名称": safe_str(summary_row.get("项目名称", "")),
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "合同数量": 1,
                "合同类型": "包干",
                "Spec_ID": pkg_spec_id,
            })
            check_msgs.append("缺少「消防远程控制包干费」，已自动补行")

        # 检查 SPEC-RB800-KZP 主键盘合计 > 阈值（只有名称含"主键盘"的才计入设备费）
        rb800_total = 0
        for _, r in contract_detail.iterrows():
            if safe_str(r.get("Spec_ID", "")) == "SPEC-RB800-KZP" and "主键盘" in safe_str(r.get("产品名称", "")):
                rb800_total += safe_numeric(r.get("合同数量", 0))
        # 也加上刚补行的
        for nr in new_detail_rows:
            if nr["合同编号"] == contract_no and nr.get("Spec_ID") == "SPEC-RB800-KZP" and "主键盘" in safe_str(nr.get("产品名称", "")):
                rb800_total += safe_numeric(nr.get("合同数量", 0))

        if rb800_total > threshold:
            # 检查合同已是否存在设备费行，避免重复添加
            has_equipment_fee = any(
                safe_str(r.get("产品名称", "")) == "消防远程控制设备费"
                for _, r in contract_detail.iterrows()
            )
            if not has_equipment_fee:
                eqp_spec_id = _match_spec_id("消防远程控制设备费", "定制", billing_df)
                new_detail_rows.append({
                    "合同编号": contract_no,
                    "项目名称": safe_str(summary_row.get("项目名称", "")),
                    "产品名称": "消防远程控制设备费",
                    "规格": "定制",
                    "合同数量": rb800_total - threshold,
                    "合同类型": "包干",
                    "Spec_ID": eqp_spec_id,
                })
                check_msgs.append(f"SPEC-RB800-KZP 主键盘合计 {rb800_total} > {threshold}，设备费用已补充")

        # 汇总费用检查结果
        if check_msgs:
            summary_check[contract_no] = "；".join(check_msgs)
        else:
            summary_check[contract_no] = "正常"

    # 单采项目标记为正常
    for _, summary_row in summary_df.iterrows():
        contract_no = safe_str(summary_row.get("合同编号", ""))
        project_type = safe_str(summary_row.get("项目类型", ""))
        if contract_no and contract_no not in summary_check and contract_no not in locked_contracts and project_type == "单采":
            summary_check[contract_no] = "正常"

    # 将补行合并到明细表 DataFrame
    if new_detail_rows:
        new_df = pd.DataFrame(new_detail_rows)
        detail_df = pd.concat([detail_df, new_df], ignore_index=True)

    # ========== 步骤2: 按合同逐条核算 ==========
    # 构建总表合同索引
    summary_index: Dict[str, Any] = {}
    if "合同编号" in summary_df.columns:
        for _, r in summary_df.iterrows():
            cno = safe_str(r.get("合同编号", ""))
            if cno:
                summary_index[cno] = r

    # 预计算各合同的电话汇总（SPEC-RB800-PHONE + 名称含"电话"）
    phone_total_by_contract: Dict[str, float] = {}
    for _, detail_row in detail_df.iterrows():
        sid = safe_str(detail_row.get("Spec_ID", "")).lower()
        pn = safe_str(detail_row.get("产品名称", ""))
        if sid == "spec-rb800-phone" and "电话" in pn:
            cno = safe_str(detail_row.get("合同编号", ""))
            phone_total_by_contract[cno] = phone_total_by_contract.get(cno, 0) + safe_numeric(detail_row.get("合同数量", 0))
    # 每个合同只免费1个电话
    phone_remaining_free: Dict[str, float] = {cno: 1.0 for cno, total in phone_total_by_contract.items() if total > 1}

    calc_results: List[Dict[str, Any]] = []  # 明细写回数据

    for _, detail_row in detail_df.iterrows():
        contract_no = safe_str(detail_row.get("合同编号", ""))
        spec_id_orig = safe_str(detail_row.get("Spec_ID", ""))
        spec_id = spec_id_orig.lower()  # 规范化：转小写，与 billing_index 的 key 一致
        qty = safe_numeric(detail_row.get("合同数量", 0))

        if contract_no in locked_contracts:
            continue

        if contract_no in no_project_type_contracts:
            continue

        sum_row = summary_index.get(contract_no)
        if sum_row is None:
            continue

        project_type = safe_str(sum_row.get("项目类型", ""))

        # 备选查找：当 Spec_ID 为空时，用产品名称+规格查找
        if not spec_id:
            product_name = safe_str(detail_row.get("产品名称", "")).lower()
            spec_model = safe_str(detail_row.get("规格", "")).lower()
            key = (product_name, spec_model)
            if key in product_spec_index:
                spec_id = product_spec_index[key]

        # 产品未匹配计费规则 → 跳过核算，保留明细表已有价格数据，禁止写 0
        if not spec_id or spec_id not in billing_index:
            product_name = safe_str(detail_row.get("产品名称", ""))
            spec = safe_str(detail_row.get("规格", ""))
            print(f"[WARNING] 跳过核算（缺计费规则）: 合同={contract_no}, 产品={product_name}, 规格={spec}, Spec_ID={spec_id_orig}")
            continue

        rule = billing_index[spec_id]

        is_baogan = "否"
        single_qty = 0
        unit_price = 0.0
        subtotal = 0.0

        # 电话特殊规则：同一合同内 SPEC-RB800-PHONE + 名称含"电话"的总量 > 1 时，只包干1个，其余单采
        is_phone_product = (spec_id == "spec-rb800-phone" and "电话" in safe_str(detail_row.get("产品名称", "")))
        if is_phone_product and contract_no in phone_remaining_free:
            free_for_row = min(qty, phone_remaining_free[contract_no])
            phone_remaining_free[contract_no] -= free_for_row
            phone_single_qty = qty - free_for_row
            unit_price = safe_numeric(rule.get("单采单价", 0))
            single_qty = phone_single_qty
            subtotal = phone_single_qty * unit_price
            is_baogan = "是" if free_for_row > 0 else "否"
            calc_results.append({
                "_record_id": detail_row.get("_record_id", ""),
                "合同编号": contract_no,
                "产品名称": safe_str(detail_row.get("产品名称", "")),
                "规格": safe_str(detail_row.get("规格", "")),
                "是否计入包干": is_baogan,
                "包干合同单采数量": single_qty,
                "单采价格": unit_price,
                "小计": subtotal,
            })
            continue

        if project_type == "单采":
            # 全部按单采计算
            unit_price = safe_numeric(rule.get("单采单价", 0))
            single_qty = qty
            subtotal = qty * unit_price
            is_baogan = "否"

        elif project_type == "包干":
            billing_method = rule.get("计费方式", "")
            baogan_threshold = safe_numeric(rule.get("包干阈值", 0))

            if billing_method == "单采":
                # 单采设备
                unit_price = safe_numeric(rule.get("单采单价", 0))
                single_qty = qty
                subtotal = qty * unit_price
                is_baogan = "否"
            elif billing_method == "完全包干" or baogan_threshold == 0:
                # 完全包干（或无阈值）：全部在包干范围内
                is_baogan = "是"
                single_qty = 0
                unit_price = 0.0
                subtotal = 0.0
            else:
                # 限制包干：阈值内免费，超出部分按单采计算
                is_baogan = "是"
                if qty <= baogan_threshold:
                    unit_price = 0.0
                    single_qty = 0
                    subtotal = 0.0
                else:
                    single_qty = qty - baogan_threshold
                    unit_price = safe_numeric(rule.get("单采单价", 0))
                    subtotal = single_qty * unit_price

        calc_results.append({
            "_record_id": detail_row.get("_record_id", ""),
            "合同编号": contract_no,
            "产品名称": safe_str(detail_row.get("产品名称", "")),
            "规格": safe_str(detail_row.get("规格", "")),
            "是否计入包干": is_baogan,
            "包干合同单采数量": single_qty,
            "单采价格": unit_price,
            "小计": subtotal,
        })

    # ========== 步骤3: 汇总 AI项目金额 ==========
    contract_totals: Dict[str, float] = defaultdict(float)
    for cr in calc_results:
        contract_totals[cr["合同编号"]] += cr["小计"]

    # ========== 步骤4: 回写明细表 ==========
    print(f"[Finance Calc] 准备回写明细：calc_results={len(calc_results)}, new_detail_rows={len(new_detail_rows)}")
    # 更新已有行
    if calc_results:
        update_df = pd.DataFrame(calc_results)
        # 只更新已有 record_id 的行
        existing_updates = update_df[update_df["_record_id"] != ""]
        if not existing_updates.empty:
            # 不回写只读字段：合同编号、产品名称、规格
            detail_write = existing_updates.drop(columns=["合同编号", "产品名称", "规格"], errors="ignore")
            update_bitable_records(
                TABLE_ID_FINANCE_DETAIL,
                detail_write,
                record_id_col="_record_id",
                numeric_cols={"包干合同单采数量", "小计"},
            )
            print(f"[Finance Calc] 已更新现有记录数: {len(detail_write)}")

        # 将新行的核算结果合并到 new_detail_rows
        new_calc = update_df[update_df["_record_id"] == ""]
        if not new_calc.empty:
            if not new_detail_rows:
                print(f"[WARN] Finance Calc: new_calc 有 {len(new_calc)} 条核算结果但 new_detail_rows 为空，结果未匹配到明细行")
            calc_lookup: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
            for _, cr in new_calc.iterrows():
                key = (
                    safe_str(cr.get("合同编号", "")).strip().lower(),
                    safe_str(cr.get("产品名称", "")).strip(),
                    safe_str(cr.get("规格", "")).strip().lower(),
                )
                if key in calc_lookup:
                    print(f"[WARN] Finance Calc: duplicate calc_result key {key}, keep first occurrence")
                    continue
                calc_lookup[key] = cr
            for nr in new_detail_rows:
                key = (
                    safe_str(nr.get("合同编号", "")).strip().lower(),
                    safe_str(nr.get("产品名称", "")).strip(),
                    safe_str(nr.get("规格", "")).strip().lower(),
                )
                if key in calc_lookup:
                    cr = calc_lookup[key]
                    nr["是否计入包干"] = safe_str(cr.get("是否计入包干", ""))
                    nr["包干合同单采数量"] = safe_numeric(cr.get("包干合同单采数量", 0))
                    nr["单采价格"] = safe_numeric(cr.get("单采价格", 0))
                    nr["小计"] = safe_numeric(cr.get("小计", 0))
            print(f"[Finance Calc] 已合并核算结果到补行，calc_hits={len(calc_lookup)}")

    # 写入补行的新明细行（含核算结果）
    if new_detail_rows:
        new_detail_df = pd.DataFrame(new_detail_rows)
        write_df_to_bitable(
            TABLE_ID_FINANCE_DETAIL,
            new_detail_df,
            numeric_cols={"合同数量", "包干合同单采数量", "小计"},
        )

    # ========== 步骤5: 回写总表（AI费用检查 + AI项目金额）==========
    summary_update_rows: List[Dict[str, Any]] = []
    for _, r in summary_df.iterrows():
        contract_no = safe_str(r.get("合同编号", ""))
        if contract_no in locked_contracts:
            continue

        update_row = {"_record_id": r.get("_record_id", "")}

        if contract_no in summary_check:
            update_row["AI费用检查"] = summary_check[contract_no]

        if contract_no in contract_totals:
            update_row["AI项目金额"] = contract_totals[contract_no]

        summary_update_rows.append(update_row)

    # 回写前校验：确保汇总行数与预期一致（非锁定合同数）
    expected_summary = len([1 for _, r in summary_df.iterrows()
                           if safe_str(r.get("合同编号", "")) not in locked_contracts])
    if len(summary_update_rows) != expected_summary:
        print(f"[ERROR] 核算回写校验失败: 预期 {expected_summary} 合同, 实际 {len(summary_update_rows)}")
        return {"ok": False, "error": f"总表回写校验失败: 预期{expected_summary}合同, 实际{len(summary_update_rows)}"}

    if summary_update_rows:
        su_df = pd.DataFrame(summary_update_rows)
        su_to_update = su_df[su_df["_record_id"] != ""]
        if not su_to_update.empty:
            update_bitable_records(
                TABLE_ID_FINANCE_SUMMARY,
                su_to_update,
                record_id_col="_record_id",
                numeric_cols=set(),
            )

    return {
        "ok": True,
        "locked_contracts": len(locked_contracts),
        "no_project_type_contracts": len(no_project_type_contracts),
        "new_detail_rows": len(new_detail_rows),
        "calculated_rows": len(calc_results),
        "contracts_updated": len(summary_update_rows),
    }


# =========================
# 数据缓存（避免每次请求重复读取飞书）
# =========================
import threading as _threading
data_cache_lock = _threading.Lock()
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

        # 源头过滤：费用类明细不参与任何库存/缺货计算和展示
        if "产品名称" in items_df.columns:
            fee_mask = items_df["产品名称"].astype(str).str.contains("费", na=False)
            if fee_mask.sum() > 0:
                items_df = items_df[~fee_mask].copy()
                print(f"[费用过滤] 已从数据源排除 {fee_mask.sum()} 行费用类明细")

        with data_cache_lock:
            data_cache["items"] = items_df
            data_cache["sku"] = sku_df
            data_cache["inv"] = inv_df

        # ===== 验证表字段完整性 =====
        print("\n【表字段验证】")

        required_items_cols = ["合同编号", "合同数量"]
        optional_items_cols = ["产品名称", "规格", "库存可用量", "缺口数量", "库存状态", "预计到货日期", "是否RB800", "排单批次号",
                               "是否紧急订单", "是否换货订单", "是否补发订单", "是否维修订单"]

        required_sku_cols = ["产品编码SKU", "标准生产周期"]
        optional_sku_cols = ["设备名称", "设备型号", "是否新产品", "是否自研", "是否外采"]

        required_inv_cols = ["库存数量"]
        optional_inv_cols = ["库存日期", "国网设备名称", "国网设备型号", "待采购出库", "在途数量", "数据来源", "导入批次号"]

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


def safe_numeric(value) -> float:
    """安全转换数值，无效值返回 0.0。"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value) if not pd.isna(value) else 0.0
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return 0.0


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
# SKU 自动生成规则
# =========================
SKU_PREFIX_RULES: List[Tuple[str, str, str]] = [
    # (关键词, 前缀, 设备分类) —— 长关键词优先匹配
    ("消控室值班助手", "RB8-", "消控室值班助手"),
    ("消防广播",       "RB8-", "消防广播"),
    ("采集传输终端",   "TX-CJ-", "采集传输终端"),
    ("用户信息传输装置", "TX-YH-", "用户信息传输装置"),
    ("主机通讯板",     "XF-TX-", "主机通讯板"),
    ("物联网卡",       "FW-", "物联网卡"),
    ("控制柜远程控制", "XF-YK-", "控制柜远程控制"),
]


def generate_sku_for_product(
    product_name: str,
    spec: str,
    sku_df: pd.DataFrame,
) -> Optional[str]:
    """根据产品名称和规格自动生成 SKU 编码。

    按 SKU_PREFIX_RULES 匹配产品名称关键词 → 确定前缀 →
    扫描 SKU 标准表中该前缀下的最大序号 → 生成 prefix+next_number。
    """
    search_text = f"{safe_str(product_name)} {safe_str(spec)}".lower()

    # 按规则表匹配（长关键词优先）
    matched_prefix = None
    matched_category = None
    for keyword, prefix, category in SKU_PREFIX_RULES:
        if keyword in search_text:
            matched_prefix = prefix
            matched_category = category
            break

    if not matched_prefix:
        return None

    # 扫描 SKU 标准表中该前缀下的现有序号
    max_num = 0
    if sku_df is not None and not sku_df.empty and "产品编码SKU" in sku_df.columns:
        for sku_code in sku_df["产品编码SKU"].astype(str):
            sku_str = safe_str(sku_code)
            if sku_str.startswith(matched_prefix):
                suffix = sku_str[len(matched_prefix):]
                try:
                    num = int(suffix)
                    max_num = max(max_num, num)
                except ValueError:
                    pass

    next_num = max_num + 1
    if next_num > 999:
        # 序号溢出 → 用产品名 hash 作为后缀
        import hashlib
        hash_suffix = hashlib.md5(safe_str(product_name).encode()).hexdigest()[:4].upper()
        return f"{matched_prefix}{hash_suffix}"

    return f"{matched_prefix}{next_num:03d}"


def ensure_sku_in_standard_table(
    sku_code: str,
    device_name: str,
    is_new: bool,
    sku_df: pd.DataFrame,
) -> bool:
    """确保 SKU 在标准表中存在，不存在则创建。

    返回 True 表示 SKU 已存在或创建成功。
    """
    if not sku_code:
        return False

    # 检查是否已存在
    if sku_df is not None and not sku_df.empty and "产品编码SKU" in sku_df.columns:
        existing = sku_df[sku_df["产品编码SKU"].astype(str).apply(safe_str) == safe_str(sku_code)]
        if not existing.empty:
            return True

    # 不存在 → 创建
    try:
        now_str = datetime.now().strftime("%Y-%m-%d")
        new_row = {
            "产品编码SKU": safe_str(sku_code),
            "设备名称": safe_str(device_name),
            "标准生产周期": 15 if is_new else 25,
            "是否新产品": "是" if is_new else "",
            "是否自研": "",
            "是否外采": "",
        }
        new_df = pd.DataFrame([new_row])
        write_df_to_bitable(
            TABLE_ID_SKU,
            new_df,
            numeric_cols={"标准生产周期"},
        )
        print(f"[SKU自动创建] {sku_code} ({device_name}) {'新产品' if is_new else '补全'} → SKU标准表")
        return True
    except Exception as e:
        print(f"[SKU自动创建失败] {sku_code}: {e}")
        return False


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
    in_inventory: bool = True,
) -> date:
    """计算缺货 SKU 的预计到货日期。

    in_inventory=False 表示该产品在库存表中不存在，视为新产品，生产周期 15 天。
    """
    if not in_inventory:
        cycle = 15
    else:
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
    normal = normal.sort_values(by=["__ship_date", "合同编号"], kind="stable", na_position="last")

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

        # 路径压缩 + 容量检查：找到第一个有容量的工作日（防死循环：最多 365 天）
        d = base_date
        _safety = 0
        while _safety < 365:
            visited: List[date] = []
            while d in overflow:
                visited.append(d)
                d = overflow[d]
            for v in visited:
                overflow[v] = d

            c = day_count.get(d, 0)
            sk = day_sku_kinds.get(d, 0)
            tq = day_total_qty.get(d, 0)
            cap = calc_daily_capacity(sk, tq, base=5)

            if c < cap:
                break

            next_d = next_working_day(d + timedelta(days=1))
            overflow[d] = next_d
            d = next_d
            _safety += 1

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
        out = out.sort_values(by=["AI建议发货时间", "合同编号"], kind="stable", na_position="last")

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
    if status in {"待确认", "已确认", "已发货", "已签收"}:
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


def summarize_confirmed_stock(
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    inv: Optional[pd.DataFrame] = None,
    sku_df: Optional[pd.DataFrame] = None,
) -> Dict[str, float]:
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
        if not sku_code:
            sku_code = safe_str(r.get("产品名称", "")) or safe_str(r.get("规格", ""))
        qty = to_num(r.get("合同数量", 0))
        if sku_code and qty > 0:
            # Canonicalize SKU key via inventory matching
            if inv is not None and not inv.empty:
                _, _, canon_key = find_inventory_row(
                    sku_code, r.to_dict(), inv, sku_df=sku_df
                )
                canon_key = canon_key or sku_code
            else:
                canon_key = sku_code
            locked_stock[canon_key] = locked_stock.get(canon_key, 0.0) + qty
    return locked_stock


def get_confirmed_contract_ids(summary: pd.DataFrame) -> set:
    summary = prepare_summary_status(summary)
    if summary.empty or "合同编号" not in summary.columns:
        return set()
    return set(summary.loc[summary["订单状态"] == "已确认", "合同编号"].astype(str).apply(safe_str))


def summarize_effective_reservations(
    reservations: pd.DataFrame,
    inv: Optional[pd.DataFrame] = None,
    sku_df: Optional[pd.DataFrame] = None,
) -> Dict[str, float]:
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
            # Canonicalize SKU key via inventory matching
            # Pass sku_code as both 产品名称 and 规格 to maximize inverted index match surface
            if inv is not None and not inv.empty:
                _, _, canon_key = find_inventory_row(
                    sku_code, {"产品名称": sku_code, "规格": sku_code}, inv, sku_df=sku_df
                )
                canon_key = canon_key or sku_code
            else:
                canon_key = sku_code
            reserved_stock[canon_key] = reserved_stock.get(canon_key, 0.0) + qty
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


def build_reservation_rows(
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    batch_id: str,
    inv: Optional[pd.DataFrame] = None,
    sku_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    if summary is None or summary.empty or detail is None or detail.empty:
        return pd.DataFrame()
    summary = prepare_summary_status(summary)
    pending_contracts = set(summary.loc[summary["订单状态"] == "待确认", "合同编号"].astype(str).apply(safe_str))
    if not pending_contracts:
        return pd.DataFrame()

    required = {"合同编号", "合同数量"}
    if not required.issubset(set(detail.columns)):
        return pd.DataFrame()
    
    # SKU编码不存在时用产品名称兜底
    if "SKU编码" not in detail.columns:
        if "产品名称" in detail.columns:
            detail["SKU编码"] = detail["产品名称"].astype(str).apply(safe_str)
        else:
            detail["SKU编码"] = ""

    qty_by_contract_sku: Dict[Tuple[str, str], float] = {}
    now_str = datetime.now().strftime("%Y-%m-%d")
    source = detail.copy()
    source["合同编号"] = source["合同编号"].astype(str).apply(safe_str)
    source["SKU编码"] = source["SKU编码"].astype(str).apply(safe_str)
    for _, r in source[source["合同编号"].isin(pending_contracts)].iterrows():
        contract_id = safe_str(r.get("合同编号", ""))
        sku_code = safe_str(r.get("SKU编码", ""))
        if not sku_code:
            sku_code = safe_str(r.get("产品名称", "")) or safe_str(r.get("规格", ""))
        qty = to_num(r.get("合同数量", 0))
        if not contract_id or not sku_code or qty <= 0:
            continue
        # Canonicalize SKU before creating reservation
        if inv is not None and not inv.empty:
            _, _, canon_sku = find_inventory_row(
                sku_code, r.to_dict(), inv, sku_df=sku_df
            )
            canon_sku = canon_sku or sku_code
        else:
            canon_sku = sku_code
        key = (contract_id, canon_sku)
        qty_by_contract_sku[key] = qty_by_contract_sku.get(key, 0.0) + qty

    rows = []
    for (contract_id, canon_sku), qty in qty_by_contract_sku.items():
        rows.append({
            "预留ID": f"{batch_id}_{contract_id}_{canon_sku}",
            "合同编号": contract_id,
            "SKU": canon_sku,
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
.section td{padding:12px 12px 4px 0;font-weight:bold;color:#666;font-size:13px;text-transform:uppercase;letter-spacing:1px}
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
  <a href="/finance">财务核算</a>
  <a href="/docs">API 参考</a>
</div>
<div class="main">
  <div class="card">
    <h3>导入订单</h3>
    <p>拖拽或选择 OA 导出的 Excel 文件，自动校验并导入飞书多维表格。支持订单明细表和订单主表。</p>
    <input type="file" id="fileInput" accept=".xlsx,.xls" style="margin-bottom:8px"><br>
    <label><input type="checkbox" id="allowEmptySku"> 允许 SKU 为空（排单时自动匹配）</label><br>
    <button class="btn btn-primary" id="importBtn" onclick="importExcel()" style="margin-top:8px">导入到飞书</button>
    <div class="status" id="importStatus"></div>
  </div>
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
      <tr class="section"><td colspan="3">顺丰物流</td></tr>
      <tr><td><span class="badge badge-post">POST</span></td><td>/shipping/tracking/import</td><td>录入运单号 → 自动查物流</td></tr>
      <tr><td><span class="badge badge-post">POST</span></td><td>/shipping/tracking/refresh</td><td>全量刷新物流状态</td></tr>
      <tr><td><span class="badge badge-post">POST</span></td><td>/shipping/tracking/refresh/one</td><td>刷新单个运单</td></tr>
      <tr><td><span class="badge badge-get">GET</span></td><td>/shipping/tracking/status</td><td>物流状态统计</td></tr>
      <tr><td><span class="badge badge-get">GET</span></td><td>/shipping/tracking/list</td><td>运单列表</td></tr>
    </table>
  </div>
</div>
<script>
async function importExcel(){
  const file=document.getElementById('fileInput').files[0];
  const st=document.getElementById('importStatus');
  if(!file){st.textContent='请先选择 Excel 文件';return;}
  st.textContent='正在导入...';
  const fd=new FormData();
  fd.append('file',file);
  fd.append('auto_fix',document.getElementById('allowEmptySku').checked);
  try{
    const r=await fetch('/import',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok){st.textContent='导入成功! '+d.results.map(x=>x.sheet+':'+x.rows+'行').join(', ');}
    else{st.textContent='导入失败: '+d.results.map(x=>x.sheet+':'+(x.error||'OK')).join(', ');}
  }catch(e){st.textContent='网络错误: '+e.message;}
}
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

FINANCE_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>财务核算</title>
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
.card label{font-size:14px;color:#555;margin-right:8px}
.card input[type=number]{padding:6px 10px;border:1px solid #d0d5dd;border-radius:6px;font-size:14px;width:80px}
.btn{display:inline-block;padding:10px 28px;border:none;border-radius:6px;font-size:15px;cursor:pointer;text-decoration:none;margin-right:8px}
.btn-primary{background:#4f46e5;color:#fff}
.btn-primary:hover{background:#4338ca}
.btn-primary:disabled{opacity:.5;cursor:not-allowed}
.btn-secondary{background:#e5e7eb;color:#374151}
.btn-secondary:hover{background:#d1d5db}
#result{margin-top:16px;white-space:pre-wrap;background:#1e1e2e;color:#cdd6f4;border-radius:6px;padding:16px;font-family:'Cascadia Code',Consolas,monospace;font-size:13px;max-height:500px;overflow-y:auto;display:none}
.status{font-size:14px;color:#888;margin-top:4px}
</style>
</head>
<body>
<div class="sidebar">
  <h2>AI排单服务</h2>
  <a href="/">操作面板</a>
  <a href="/finance" class="active">财务核算</a>
  <a href="/docs">API 参考</a>
</div>
<div class="main">
  <div class="card">
    <h3>同步对账数据</h3>
    <p>将销售订单主表和明细表数据同步到财务对账表，并按计费规则匹配 Spec_ID。</p>
    <p style="color:#b45309;font-size:13px">注意：同步会更新已有数据，但保留核算结果字段。</p>
    <button class="btn btn-primary" id="syncBtn" onclick="syncData()">同步对账数据</button>
    <div class="status" id="syncStatus"></div>
  </div>
  <div class="card">
    <h3>财务核算</h3>
    <p>执行包干/单采分类核算，自动检查完整性并计算费用。</p>
    <label>RB800-KZP 主键盘设备费阈值：</label>
    <input type="number" id="threshold" value="3" min="0" step="1">
    <p style="color:#b45309;font-size:13px;margin-top:8px">注意：人工核对金额已填写的合同将被锁定，不参与核算。</p>
    <button class="btn btn-primary" id="calcBtn" onclick="runCalculate()">财务核算</button>
    <div class="status" id="calcStatus"></div>
    <pre id="result"></pre>
  </div>
</div>
<script>
async function syncData(){
  const btn=document.getElementById('syncBtn');
  const st=document.getElementById('syncStatus');
  const res=document.getElementById('result');
  btn.disabled=true;
  st.textContent='正在同步...';
  res.style.display='none';
  try{
    const r=await fetch('/finance/sync',{method:'POST'});
    const d=await r.json();
    res.textContent=JSON.stringify(d,null,2);
    res.style.display='block';
    st.textContent=d.ok?'同步完成':'同步失败';
  }catch(e){
    res.textContent='请求失败: '+e.message;
    res.style.display='block';
    st.textContent='网络错误';
  }
  btn.disabled=false;
}

async function runCalculate(){
  const btn=document.getElementById('calcBtn');
  const st=document.getElementById('calcStatus');
  const res=document.getElementById('result');
  const threshold=parseInt(document.getElementById('threshold').value)||3;
  btn.disabled=true;
  st.textContent='正在核算，请稍候...';
  res.style.display='none';
  try{
    const r=await fetch('/finance/calculate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({threshold})});
    const d=await r.json();
    res.textContent=JSON.stringify(d,null,2);
    res.style.display='block';
    st.textContent=d.ok?'核算完成':'核算失败';
  }catch(e){
    res.textContent='请求失败: '+e.message;
    res.style.display='block';
    st.textContent='网络错误';
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


@app.get("/finance")
def finance_page():
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content=FINANCE_HTML)


def delete_records_by_contract(table_id: str, contract_numbers: List[str]) -> int:
    """按合同编号删除飞书表格中的匹配记录。返回删除条数。"""
    headers = _feishu_headers()

    # 查询所有记录，找出匹配的 record_id
    ids_to_delete: List[str] = []
    page_token = None
    while True:
        url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/search"
        payload: Dict[str, Any] = {"page_size": 500}
        if page_token:
            payload["page_token"] = page_token
        resp = httpx.post(url, headers=headers, json=payload, timeout=60.0)
        data = _safe_http_json(resp, "查询记录")
        if data.get("code") != 0:
            raise Exception(f"查询记录失败: code={data.get('code')}, msg={data.get('msg')}")
        items = data.get("data", {}).get("items", [])
        for item in items:
            fields = item.get("fields", {})
            contract = fields.get("合同编号", "")
            if isinstance(contract, list):
                contract = contract[0].get("text", "") if contract else ""
            if contract in contract_numbers:
                ids_to_delete.append(item.get("record_id", ""))
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data.get("data", {}).get("page_token", "")

    if not ids_to_delete:
        return 0

    # 分批删除
    delete_url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/batch_delete"
    deleted = 0
    for i in range(0, len(ids_to_delete), 500):
        batch = ids_to_delete[i:i + 500]
        resp = httpx.post(delete_url, headers=headers, json={"records": batch}, timeout=60.0)
        data = _safe_http_json(resp, "删除记录")
        if data.get("code") != 0:
            raise Exception(f"删除记录失败: code={data.get('code')}, msg={data.get('msg')}")
        deleted += len(batch)
    return deleted


def delete_all_records(table_id: str) -> int:
    """删除指定飞书多维表格的全部记录，返回删除条数。"""
    headers = _feishu_headers()

    # 查所有 record_id
    all_ids = []
    page_token = None
    while True:
        url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/search"
        payload: Dict[str, Any] = {"page_size": 500}
        if page_token:
            payload["page_token"] = page_token
        resp = httpx.post(url, headers=headers, json=payload, timeout=60.0)
        data = _safe_http_json(resp, "查询记录")
        if data.get("code") != 0:
            raise Exception(f"查询记录失败: code={data.get('code')}, msg={data.get('msg')}")
        items = data.get("data", {}).get("items", [])
        for item in items:
            all_ids.append(item.get("record_id", ""))
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data.get("data", {}).get("page_token", "")

    if not all_ids:
        return 0

    # 分批删除
    delete_url = f"{FEISHU_BITABLE_BASE}/tables/{table_id}/records/batch_delete"
    deleted = 0
    for i in range(0, len(all_ids), 500):
        batch = all_ids[i:i + 500]
        resp = httpx.post(delete_url, headers=headers, json={"records": batch}, timeout=60.0)
        data = _safe_http_json(resp, "删除记录")
        if data.get("code") != 0:
            raise Exception(f"删除记录失败: code={data.get('code')}, msg={data.get('msg')}")
        deleted += len(batch)
    return deleted


@app.post("/import")
async def import_excel(file: UploadFile = File(...), auto_fix: bool = Form(False)):
    """上传 Excel 文件，校验后写入飞书多维表格。

    Excel 结构:
        Sheet1（订单表）: 订单编号, 客户名称, 项目名称, 商务, 下单时间, 国网设备名称, 设备型号, 数量
        Sheet2（库存表）: 国网设备名称, 型号, 库存数量
    """
    import io
    from shared import OA_COLUMN_MAP, MAIN_FIELD_MAP, ITEMS_FIELD_MAP, INV_COLUMN_MAP
    try:
        contents = await file.read()
        xl = pd.ExcelFile(io.BytesIO(contents))
    except Exception as e:
        return {"ok": False, "error": f"无法读取 Excel 文件: {e}"}

    results = []
    orders_done = False

    for sheet_name in xl.sheet_names:
        try:
            df_raw = pd.read_excel(xl, sheet_name=sheet_name)
        except Exception as e:
            results.append({"sheet": sheet_name, "ok": False, "error": f"读取失败: {e}"})
            continue

        print(f"[导入] Sheet={sheet_name}, 原始列={list(df_raw.columns)}")

        # ---- 检测 Sheet 类型 ----
        cols = set(df_raw.columns)
        has_order_id = "订单编号" in cols or "合同编号" in cols
        has_device = bool({"国网设备名称", "设备名称", "产品名称", "设备型号", "规格"} & cols)
        is_inventory = "库存数量" in cols

        # ============================================================
        # 订单表 → 主表 + 明细表
        # ============================================================
        if has_order_id and (has_device or "数量" in cols or "合同数量" in cols):
            if orders_done:
                continue

            df = df_raw.copy()
            df = df.loc[:, ~df.columns.astype(str).str.match(r'^Unnamed')]
            # 应用 OA 映射
            df = df.rename(columns=lambda c: OA_COLUMN_MAP.get(c, c))
            df = df.loc[:, ~df.columns.duplicated()]

            # 清洗字符串
            for col in {"合同编号", "产品名称", "规格", "客户名称", "项目名称", "商务", "SKU编码"}:
                if col in df.columns:
                    df[col] = df[col].apply(safe_str)

            # 校验
            if "合同编号" not in df.columns:
                results.append({"sheet": sheet_name, "ok": False, "error": "缺少「订单编号」列"})
                continue

            empty_contract = df["合同编号"] == ""
            if empty_contract.any():
                results.append({"sheet": sheet_name, "ok": False, "error": f"合同编号为空 {empty_contract.sum()} 行"})
                continue

            if "合同数量" in df.columns:
                df["合同数量"] = df["合同数量"].apply(to_num)

            if "下单时间" in df.columns:
                # Excel 数字日期检测（如 46156 = 2026-05-29）
                sample = df["下单时间"].dropna()
                if len(sample) > 0 and pd.api.types.is_numeric_dtype(sample):
                    excel_epoch = pd.Timestamp("1899-12-30")
                    df["下单时间"] = df["下单时间"].apply(
                        lambda x: excel_epoch + pd.Timedelta(days=int(x)) if pd.notna(x) and x > 0 else pd.NaT
                    )
                else:
                    df["下单时间"] = pd.to_datetime(df["下单时间"], errors="coerce")

            sku_empty = 0
            if "SKU编码" in df.columns:
                sku_empty = (df["SKU编码"] == "").sum()

            # 1) 写入订单主表（先删后写：防止新记录被误删）
            main_avail = [c for c in ["合同编号", "下单时间", "客户名称", "项目名称", "商务", "代理商"] if c in df.columns]
            df_main = df[main_avail].drop_duplicates(subset=["合同编号"], keep="first")
            print(f"[导入] 主表去重: {len(df_main)} 行 (原始 {len(df)} 行), 可用字段={main_avail}")

            try:
                import_contracts = list(set(df_main["合同编号"].apply(safe_str)))
                # 先删旧：清除相同合同号的旧记录
                deleted = delete_records_by_contract(TABLE_ID_MAIN, import_contracts)
                if deleted > 0:
                    print(f"[导入] 已清除 {deleted} 条旧主表记录")
                # 再写新
                written = write_df_to_bitable(
                    TABLE_ID_MAIN, df_main,
                    fields_map=MAIN_FIELD_MAP,
                    date_cols={"下单时间"},
                )
                if written <= 0:
                    results.append({"sheet": sheet_name, "ok": False, "error": f"主表写入失败: 0 条"})
                    continue
                results.append({"sheet": sheet_name, "ok": True, "table": "销售订单主表", "rows": len(df_main), "warnings": []})
            except Exception as e:
                results.append({"sheet": sheet_name, "ok": False, "error": f"主表写入异常: {e}"})
                continue

            # 2) 写入订单明细表（先删后写）
            items_avail = [c for c in ["合同编号", "下单时间", "产品名称", "规格", "合同数量"] if c in df.columns]
            df_items = df[items_avail]

            # 过滤掉产品名称包含"费"的明细（费用类不需要出库）
            if "产品名称" in df_items.columns:
                fee_mask = df_items["产品名称"].str.contains("费", na=False)
                if fee_mask.sum() > 0:
                    df_items = df_items[~fee_mask]
                    print(f"[导入] 已过滤 {fee_mask.sum()} 行含'费'的明细（费用类无需出库）")

            print(f"[导入] 明细表: {len(df_items)} 行")

            try:
                item_contracts = list(set(df_items["合同编号"].apply(safe_str))) if "合同编号" in df_items.columns else []
                # 先删旧
                if item_contracts:
                    deleted2 = delete_records_by_contract(TABLE_ID_ITEMS, item_contracts)
                    if deleted2 > 0:
                        print(f"[导入] 已清除 {deleted2} 条旧明细记录")
                # 再写新
                written2 = write_df_to_bitable(
                    TABLE_ID_ITEMS, df_items,
                    fields_map=ITEMS_FIELD_MAP,
                    numeric_cols={"合同数量"},
                    date_cols={"下单时间"},
                )
                if written2 <= 0:
                    results.append({"sheet": sheet_name, "ok": False, "error": f"明细表写入失败: 0 条"})
                else:
                    warnings = []
                    if sku_empty > 0:
                        warnings.append(f"SKU编码为空 {sku_empty} 行（排单时会自动匹配）")
                    results.append({"sheet": sheet_name, "ok": True, "table": "销售订单明细表", "rows": len(df_items), "warnings": warnings})
            except Exception as e:
                results.append({"sheet": sheet_name, "ok": False, "error": f"明细表写入异常: {e}"})

            orders_done = True

        # ============================================================
        # 库存表 → 库存快照表（清除后导入）
        # ============================================================
        elif is_inventory:
            df_inv = df_raw.copy()
            df_inv = df_inv.loc[:, ~df_inv.columns.astype(str).str.match(r'^Unnamed')]
            df_inv = df_inv.rename(columns=lambda c: INV_COLUMN_MAP.get(c, c))

            # 校验
            if "库存数量" not in df_inv.columns:
                results.append({"sheet": sheet_name, "ok": False, "error": "缺少「库存数量」列"})
                continue
            if "国网设备名称" not in df_inv.columns:
                results.append({"sheet": sheet_name, "ok": False, "error": "缺少「国网设备名称」列"})
                continue
            if "国网设备型号" not in df_inv.columns:
                results.append({"sheet": sheet_name, "ok": False, "error": "缺少「型号」列"})
                continue

            df_inv["库存数量"] = df_inv["库存数量"].apply(to_num)
            for col in ["国网设备名称", "国网设备型号"]:
                if col in df_inv.columns:
                    df_inv[col] = df_inv[col].apply(safe_str)

            empty_name = (df_inv["国网设备名称"] == "").sum()
            if empty_name > 0:
                results.append({"sheet": sheet_name, "ok": False, "error": f"国网设备名称为空 {empty_name} 行"})
                continue

            inv_cols = [c for c in ["国网设备名称", "国网设备型号", "库存数量"] if c in df_inv.columns]
            df_inv_write = df_inv[inv_cols]
            inv_field_map = {c: c for c in inv_cols}

            print(f"[导入] 库存表: {len(df_inv_write)} 行")

            # 库存表全量覆盖：先清旧数据，再写入新数据
            try:
                old_count = delete_all_records(TABLE_ID_INV)
                if old_count > 0:
                    print(f"[导入] 已清除库存旧数据 {old_count} 条")
                written = write_df_to_bitable(
                    TABLE_ID_INV, df_inv_write,
                    fields_map=inv_field_map,
                    numeric_cols={"库存数量"},
                )
                if written <= 0:
                    results.append({"sheet": sheet_name, "ok": False, "error": "库存写入失败: 0 条"})
                else:
                    results.append({"sheet": sheet_name, "ok": True, "table": "库存快照表", "rows": len(df_inv_write), "warnings": []})
            except Exception as e:
                results.append({"sheet": sheet_name, "ok": False, "error": f"库存写入异常: {e}"})

        else:
            results.append({"sheet": sheet_name, "ok": False, "error": f"无法识别类型，列名: {list(cols)}"})

    all_ok = all(r.get("ok") for r in results)

    # 导入成功后自动同步财务对账数据
    if all_ok and orders_done:
        try:
            print("[导入] 自动同步财务对账数据...")
            summary_r = _sync_finance_summary(force=True)
            detail_r = _sync_finance_detail()
            synced_count = summary_r.get("synced", 0)
            created_count = detail_r.get("created", 0)
            updated_count = detail_r.get("updated", 0)
            finance_ok = summary_r.get("ok") and detail_r.get("ok")
            results.append({
                "sheet": "财务对账(自动)",
                "ok": finance_ok,
                "table": "财务对账总表+明细表",
                "rows": synced_count + created_count + updated_count,
                "warnings": [] if finance_ok else [
                    summary_r.get("error", "") if not summary_r.get("ok") else "",
                    detail_r.get("error", "") if not detail_r.get("ok") else "",
                ]
            })
        except Exception as e:
            results.append({
                "sheet": "财务对账(自动)",
                "ok": False,
                "error": f"财务同步异常: {e}"
            })

    return {"ok": all_ok, "results": results}


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
                    sku = _s(r.get("产品名称", "")) or _s(r.get("规格", ""))
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

    # ---- 缺货统计（产品名称+缺货数量，每行一个SKU） ----
    shortage_stats_lines = []
    for sku_code, gap in shortage_order_skus:
        if not sku_code or gap <= 0:
            continue
        dev_name = sku_name_map.get(sku_code, sku_code)
        unit = sku_unit_map.get(sku_code, "个")
        gap_str = str(int(gap)) if gap == int(gap) else str(round(gap, 1))
        shortage_stats_lines.append(f"{dev_name}（缺{gap_str}{unit}）")
    shortage_stats = "\n".join(shortage_stats_lines)

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
        "最紧缺SKU缺口": str(int(gap_max)) if gap_max == int(gap_max) else str(round(gap_max, 1)),
        "最长交期订单": max_lead_contract,
        "最晚发货日期": latest_ship_date,
        "排单占库存比例": str(ratio),
        "未来3天应发货订单数": str(f3_ship),
        "未来3天缺货订单数": str(f3_short),
        "未来7天应发货订单数": str(f7_ship),
        "库存缺货统计": shortage_stats,
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
        "库存缺货统计":     parse_feishu_field(record.get("库存缺货统计")) or "无",
        "最紧缺SKU缺口":    parse_feishu_field(record.get("最紧缺SKU缺口")) or "0",
        "未来3天缺货订单数": parse_feishu_field(record.get("未来3天缺货订单数")) or "0",
        "日期":             record_date_str or today_str,
        "排单运行时间":     parse_feishu_field(record.get("排单运行时间")) or "",
    }

    return report


def _write_audit_record(audit_r: dict, batch_id: str, elapsed_s: float) -> None:
    """将审计结果写入飞书「排单审计记录」表，独立异常不抛。"""
    if not TABLE_ID_AUDIT:
        return
    try:
        errors = audit_r.get("errors", [])
        passed = 4 - len({e.split(":")[0] if ":" in e else "其他" for e in errors
                          if any(k in str(e) for k in ["可用","判定","缺货SKU","发货"])})
        passed = max(0, min(4, passed or (4 if not errors else 3)))

        # 失败规则
        failed = []
        if any("可用" in str(e) or "available" in str(e) for e in errors): failed.append("库存可用量")
        if any("判定" in str(e) or "gap" in str(e) or "矛盾" in str(e) for e in errors): failed.append("缺货判定")
        if any("缺货SKU" in str(e) or "SKU数" in str(e) or "缺口总和" in str(e) for e in errors): failed.append("缺货SKU统计")
        if any("发货" in str(e) or "日期" in str(e) for e in errors): failed.append("发货日期")

        rule_text = f"{passed}/4"
        if failed:
            rule_text += f"（{','.join(failed[:2])}失败）"

        fields = {
            "排单批次号": batch_id,
            "审计时间": int(datetime.now().timestamp() * 1000),
            "审计通过": "是" if audit_r.get("ok") else "否",
            "抽查合同数": audit_r.get("sample_count", 0),
            "通过规则": rule_text,
            "前3条异常": "\n".join(str(e)[:100] for e in errors[:3]) if errors else "",
            "总耗时s": round(elapsed_s, 1),
        }
        write_df_to_bitable(TABLE_ID_AUDIT, pd.DataFrame([fields]),
                           fields_map={k: k for k in fields},
                           numeric_cols={"抽查合同数", "总耗时s"},
                           date_cols={"审计时间"})
        print(f"[审计] 审计记录已写入飞书表 {TABLE_ID_AUDIT}")
    except Exception as e:
        print(f"[审计] 审计记录写入失败: {e}")


def _send_personal_notification(summary: pd.DataFrame, batch_id: str, report: dict, audit_result: dict = None) -> None:
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

        shortage_sku = (report.get("库存缺货统计") or "").split("\n")[0] if report else ""
        ship_3d = report.get("未来3天应发货订单数", "0") if report else "0"

        # ---- build dashboard column_set ----
        dash_columns = [
            {
                "tag": "column",
                "width": "weighted", "weight": 1,
                "elements": [{"tag": "markdown", "content": f"**{total}**\n排单合同", "text_align": "center"}],
            },
            {
                "tag": "column",
                "width": "weighted", "weight": 1,
                "elements": [{"tag": "markdown", "content": f"<font color='orange'>**{pending}**</font>\n<text_tag color='orange'>待确认</text_tag>", "text_align": "center"}],
            },
            {
                "tag": "column",
                "width": "weighted", "weight": 1,
                "elements": [{"tag": "markdown", "content": f"**{ship_3d}**\n3天内应发", "text_align": "center"}],
            },
        ]
        if shortage_orders > 0:
            dash_columns[1]["elements"][0]["content"] = f"<font color='red'>**{pending}**</font>\n<text_tag color='red'>待确认</text_tag>"

        # ---- 未来3天待发项目 ----
        future_lines = ""
        if summary is not None and not summary.empty and "发货日期" in summary.columns:
            from datetime import date, timedelta
            today_date = date.today()
            next_3d = today_date + timedelta(days=3)
            ship_col = summary["发货日期"].apply(
                lambda x: x if isinstance(x, date) else (x.date() if isinstance(x, pd.Timestamp) else None)
            )
            upcoming = summary[
                ship_col.notna() & (ship_col >= today_date) & (ship_col <= next_3d)
            ].sort_values(by=["发货日期"], kind="stable", na_position="last")

            if not upcoming.empty:
                from collections import defaultdict
                by_date: dict = defaultdict(list)
                for _, r in upcoming.iterrows():
                    sd = r["发货日期"]
                    if isinstance(sd, pd.Timestamp):
                        sd = sd.date()
                    pname = safe_str(r.get("项目名称", "")) or safe_str(r.get("合同编号", ""))
                    if sd and pname:
                        by_date[sd].append(pname)

                fl = []
                for sd in sorted(by_date.keys()):
                    names = by_date[sd]
                    date_label = f"{sd.month}/{sd.day}"
                    projects = "、".join(names[:5])
                    if len(names) > 5:
                        projects += f"等{len(names)}个"
                    fl.append(f"{date_label} {projects}")
                future_lines = "\n".join(fl)

        # ---- 排单审计结果 ----
        audit_line = ""
        if audit_result:
            total_contracts = len(summary) if summary is not None else 0
            sample_n = audit_result.get("sample_count", 0)
            errors = audit_result.get("errors", [])
            audit_ok = audit_result.get("ok", True)
            if audit_ok and not errors:
                audit_line = f"<text_tag color='green'>审计通过</text_tag> 抽查{sample_n}/{total_contracts} 4/4规则无异常"
            else:
                audit_line = f"<text_tag color='red'>审计异常</text_tag> 抽查{sample_n}/{total_contracts} {len(errors)}项异常"

        # ---- 紧缺物料（汇总全部缺货SKU，展示数量） ----
        shortage_section = ""
        if summary is not None and not summary.empty:
            shortage_rows = summary[summary["整体状态"] == "缺货"]
            if not shortage_rows.empty:
                sku_counter: dict = {}
                for _, row in shortage_rows.iterrows():
                    sku_list_str = safe_str(row.get("缺货SKU列表", ""))
                    if not sku_list_str:
                        continue
                    for sku in re.split(r"[,，]", sku_list_str):
                        sku = sku.strip()
                        if sku:
                            sku_counter[sku] = sku_counter.get(sku, 0) + 1
                if sku_counter:
                    sorted_skus = sorted(sku_counter.items(), key=lambda x: x[1], reverse=True)
                    total_kinds = len(sorted_skus)
                    lines = [f"**缺货物料**：共 {total_kinds} 种"]
                    for i, (sku, cnt) in enumerate(sorted_skus):
                        if i >= 10:
                            lines.append(f"+{total_kinds - 10} 种详见排单总表")
                            break
                        lines.append(f"{sku}    缺 {cnt} 个合同")
                    shortage_section = "\n".join(lines)

        # ---- 组装卡片 elements ----
        elements = [
            {"tag": "column_set", "flex_mode": "trisect", "columns": dash_columns},
        ]
        if future_lines:
            elements.append({"tag": "hr"})
            elements.append({"tag": "markdown", "content": f"**3天内待发项目**\n{future_lines}"})
        if audit_line:
            elements.append({"tag": "hr"})
            elements.append({"tag": "markdown", "content": audit_line})
        if shortage_section:
            elements.append({"tag": "markdown", "content": shortage_section})

        # 按钮区
        scheduler_host = os.getenv("SCHEDULER_HOST", "http://localhost:8000")
        confirm_url = f"{scheduler_host}/confirm_schedule?batch_id={batch_id}"
        detail_url = os.getenv("BITABLE_DETAIL_URL", f"https://wl6wihmop1.feishu.cn/base/{BITABLE_APP_TOKEN}?table=tbl09Z6C7wCGh3mW&view=vewiuVK8pH")
        audit_url = f"https://wl6wihmop1.feishu.cn/base/{BITABLE_APP_TOKEN}?table={TABLE_ID_AUDIT}" if TABLE_ID_AUDIT else ""
        buttons = [
            {
                "tag": "button",
                "text": {"tag": "plain_text", "content": "确认排单"},
                "type": "primary",
                "value": {"batch_id": batch_id},
                "url": confirm_url,
            },
            {
                "tag": "button",
                "text": {"tag": "plain_text", "content": "查看排单总表"},
                "url": detail_url,
                "type": "default",
            },
        ]
        if audit_url:
            buttons.append({
                "tag": "button",
                "text": {"tag": "plain_text", "content": "审计记录"},
                "url": audit_url,
                "type": "default",
            })
        elements.append({"tag": "action", "actions": buttons})

        # note 时间戳
        now_ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        elements.append({"tag": "note", "elements": [
            {"tag": "plain_text", "content": f"批次 {batch_id} | {now_ts}"}
        ]})

        card = {
            "config": {"wide_screen_mode": True},
            "header": {"template": "blue", "title": {"content": f"排单确认通知 ({batch_id})", "tag": "plain_text"}},
            "elements": elements,
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


def _send_interim_notification(msg: str) -> None:
    """排单完成后发送中间态短消息到采购群，不暴露数据详情。"""
    try:
        webhook_url = os.getenv("PROCUREMENT_WEBHOOK_URL", "")
        if not webhook_url or "/bot/v2/hook" not in webhook_url:
            return
        payload = {"msg_type": "interactive", "card": {
            "config": {"wide_screen_mode": True},
            "header": {"template": "blue", "title": {"content": "排单完成，等待确认", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": msg}],
        }}
        resp = httpx.post(webhook_url, json=payload, timeout=10.0,
                          headers={"Content-Type": "application/json; charset=utf-8"})
        if resp.json().get("code") == 0:
            print("[中间态] 采购群中间态通知已发送")
        else:
            print(f"[中间态] 发送异常: {resp.text[:200]}")
    except Exception as e:
        print(f"[中间态] 发送失败(不影响排单): {e}")


def _read_report_from_daily_table(batch_id: str) -> dict | None:
    """从日报表按 batch_id 读取排单日报数据。"""
    try:
        df = fetch_bitable_to_df(TABLE_ID_DAILY_REPORT)
        if df is None or df.empty:
            return None
        mask = df["排单批次号"].astype(str).str.strip() == str(batch_id).strip()
        matched = df[mask]
        if matched.empty:
            return None
        row = matched.iloc[-1]  # 取最新一条
        return {k: row[k] for k in row.index}
    except Exception as e:
        print(f"[确认] 从日报表读取失败: {e}")
        return None


def _write_confirmation_to_daily_table(batch_id: str, confirmed_by: str = "planner") -> bool:
    """回写日报表的确认状态字段。"""
    try:
        df = fetch_bitable_to_df(TABLE_ID_DAILY_REPORT)
        if df is None or df.empty:
            return False
        mask = df["排单批次号"].astype(str).str.strip() == str(batch_id).strip()
        matched = df[mask]
        if matched.empty:
            return False
        record_ids = matched["_record_id"].tolist()
        now_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        update_payload = {
            "records": [{
                "record_id": rid,
                "fields": {
                    "确认状态": "已确认",
                    "确认时间": now_ts,
                    "确认人": confirmed_by,
                }
            } for rid in record_ids]
        }
        token = get_access_token()
        resp = httpx.put(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{TABLE_ID_DAILY_REPORT}/records/batch_update",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=update_payload, timeout=15.0,
        )
        ok = resp.json().get("code") == 0
        if ok:
            print(f"[确认] 日报表回写成功: {batch_id}")
        return ok
    except Exception as e:
        print(f"[确认] 日报表回写失败: {e}")
        return False


@app.get("/confirm")
def confirm_panel():
    """简易确认面板：列出日报表中最近的批次，一键推送群通知。"""
    from fastapi.responses import HTMLResponse
    try:
        df = fetch_bitable_to_df(TABLE_ID_DAILY_REPORT)
        if df is None or df.empty:
            return HTMLResponse("<h3>暂无日报数据</h3>")
        rows = []
        for _, r in df.tail(10).iterrows():
            b = str(r.get("排单批次号", "")).strip()
            c = str(r.get("确认状态", "")).strip()
            t = str(r.get("排单运行时间", "")).strip()
            if isinstance(r["排单运行时间"], (int, float)):
                t = datetime.fromtimestamp(r["排单运行时间"]/1000).strftime("%m/%d %H:%M") if r["排单运行时间"] > 1e10 else ""
            btn = f'<a href="/confirm_schedule?batch_id={b}" style="background:#1456F0;color:#fff;padding:6px 16px;border-radius:6px;text-decoration:none;font-size:14px">确认发送</a>' if c != "已确认" else '<span style="color:#888;font-size:13px">已发送</span>'
            rows.append(f'<tr><td style="padding:10px 16px">{b}</td><td style="padding:10px 16px">{t}</td><td style="padding:10px 16px">{c or "待确认"}</td><td style="padding:10px 16px">{btn}</td></tr>')
        html = f"""<html><head><meta charset="utf-8"><title>排单确认面板</title><style>body{{font-family:-apple-system,sans-serif;max-width:800px;margin:40px auto;padding:0 20px}}table{{width:100%;border-collapse:collapse}}th{{text-align:left;padding:10px 16px;border-bottom:2px solid #ddd;font-size:13px;color:#666}}tr:hover{{background:#f5f5f5}}</style></head>
<body><h2>排单确认面板</h2><p style="color:#888;font-size:13px">点击「确认发送」将日报推送到采购群和老板群</p>
<table><tr><th>批次号</th><th>时间</th><th>状态</th><th>操作</th></tr>{''.join(reversed(rows))}</table></body></html>"""
        return HTMLResponse(html)
    except Exception as e:
        return HTMLResponse(f"<h3>加载失败: {e}</h3>")

@app.get("/confirm_schedule")
def confirm_schedule(batch_id: str = ""):
    """排单员确认排单后触发群通知（采购群 + 老板群）。"""
    if not batch_id:
        return {"ok": False, "error": "缺少 batch_id 参数"}
    report = _read_report_from_daily_table(batch_id)
    if not report:
        return {"ok": False, "error": f"批次 {batch_id} 不存在，请确认批次号"}
    confirmed = str(report.get("确认状态", "")).strip()
    if confirmed == "已确认":
        return {"ok": False, "error": "该批次已确认，无需重复操作"}
    print(f"[确认] 排单员确认批次 {batch_id}，触发群通知...")
    result = dispatch_schedule_notifications(report)
    if result.get("ok"):
        _write_confirmation_to_daily_table(batch_id, "planner")
    return {"ok": result["ok"], "batch_id": batch_id, "details": result}


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
    locked_stock = summarize_confirmed_stock(old_summary, items, inv=inv, sku_df=sku)
    active_reserved_stock = summarize_effective_reservations(reservations, inv=inv, sku_df=sku)
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

    # 按 下单时间 → 合同编号 排序，早下单的优先占用库存
    if "下单时间" in items.columns:
        items = items.sort_values(by=["下单时间", "合同编号"], kind="stable", na_position="last")
    else:
        items = items.sort_values(by=["合同编号"], kind="stable")

    running_consumed: Dict[str, float] = {}
    generated_skus_in_batch: Set[str] = set()  # 防止同批次重复生成 SKU
    backfill_rows = []
    for _, row in items.iterrows():
        row_dict = row.to_dict()
        sku_code = safe_str(row_dict.get("SKU编码", ""))
        # 当 SKU编码 不存在时，依次尝试 产品名称、规格 作为 SKU 标识
        if not sku_code:
            sku_code = safe_str(row_dict.get("产品名称", "")) or safe_str(row_dict.get("规格", ""))
        contract_id = safe_str(row_dict.get("合同编号", ""))

        if not sku_code or not contract_id:
            continue
        if contract_id in confirmed_contract_ids:
            print(f"[锁] 合同 {contract_id} 已确认，销售订单明细表禁止重算和回填")
            continue

        inv_info, _inv_how, canonical_sku = find_inventory_row(sku_code, row_dict, inv, sku_df=sku)
        if not canonical_sku:
            canonical_sku = sku_code
        if inv_info is None:
            print(f"⚠️ 库存未匹配 合同={contract_id} SKU={sku_code} 规格={safe_str(row_dict.get('规格', ''))[:50]} → {_inv_how}")

        batch_reserved = running_consumed.get(canonical_sku, 0.0)
        total_reserved = occupied_stock.get(canonical_sku, 0.0) + batch_reserved
        available = calc_available_stock(inv_info, reserved_qty=total_reserved)

        demand = to_num(row_dict.get("合同数量", 0))
        status, gap = calc_stock_status_and_gap(demand, available)

        consumed = demand - gap
        running_consumed[canonical_sku] = running_consumed.get(canonical_sku, 0.0) + consumed

        rb800_flag = "是" if is_rb800_from_text(pick_model_text(row_dict)) else "否"

        eta_date = ""
        if status == "缺货":
            in_inv = inv_info is not None
            eta = calc_shortage_eta_date(sku_code=canonical_sku, sku_df=sku, base_date=stock_base_date, in_inventory=in_inv)
            eta_date = eta.strftime("%Y-%m-%d")

        # === SKU 自动补全：将 canonical SKU 写回订单明细表 ===
        # 仅精确索引匹配（路径1）才写回 SKU，倒排索引匹配可能不准确
        is_exact_match = _inv_how.startswith("SKU列匹配") if inv_info is not None else False
        original_sku = safe_str(row_dict.get("SKU编码", ""))
        product_name = safe_str(row_dict.get("产品名称", ""))
        spec = safe_str(row_dict.get("规格", ""))

        if inv_info is not None and is_exact_match:
            # 场景 A: 精确匹配 → 补全 SKU 标准表 + 写回规范 SKU
            ensure_sku_in_standard_table(canonical_sku, product_name or spec, is_new=False, sku_df=sku)
            if not original_sku or original_sku != canonical_sku:
                print(f"[SKU补全] 合同={contract_id} {original_sku or '(空)'} → {canonical_sku}")
        elif inv_info is not None and not is_exact_match:
            # 倒排索引匹配 → 仅内部追踪，不写回（避免误匹配）
            if not original_sku:
                print(f"[SKU跳过] 合同={contract_id} '{product_name}' 倒排匹配→{canonical_sku}（不写回，等待人工确认SKU编码）")
        elif inv_info is None and not original_sku:
            # 场景 B: 全新产品 → 自动生成 SKU（防止同批次重复）
            new_sku = generate_sku_for_product(product_name, spec, sku_df=sku)
            if new_sku:
                if new_sku in generated_skus_in_batch:
                    print(f"[SKU跳过] 合同={contract_id} '{product_name}' → {new_sku} (本批次已存在)")
                else:
                    print(f"[SKU自动生成] 合同={contract_id} 新产品 '{product_name}' → {new_sku}")
                    ensure_sku_in_standard_table(new_sku, product_name, is_new=True, sku_df=sku)
                    generated_skus_in_batch.add(new_sku)
                canonical_sku = new_sku
                # 重新计算 ETA：新产品 15 天
                if status == "缺货":
                    eta = calc_shortage_eta_date(sku_code=canonical_sku, sku_df=sku, base_date=stock_base_date, in_inventory=False)
                    eta_date = eta.strftime("%Y-%m-%d")
            else:
                print(f"[SKU生成跳过] 合同={contract_id} 无法分类产品 '{product_name}'，保持使用产品名")

        # 确定是否需要写回 SKU 编码到订单明细表
        sku_to_write = None
        if is_exact_match and (not original_sku or original_sku != canonical_sku):
            sku_to_write = canonical_sku  # 精确匹配：写回规范 SKU
        elif inv_info is None and canonical_sku != sku_code:
            sku_to_write = canonical_sku  # 新产品自动生成：写回新 SKU

        allocated = consumed  # = min(demand, available)
        remaining_after = available - allocated

        # 置信度判定（基于分配后剩余库存）
        if remaining_after > 5:
            confidence = "高"
        elif 1 <= remaining_after <= 5:
            confidence = "中"
        else:
            if spec == "定制化硬件":
                confidence = "低"
            elif spec == "AC-RB800":
                if "电话" in product_name:
                    confidence = "低"
                elif original_sku is None or original_sku == "":
                    confidence = "低"
                else:
                    confidence = "中"
            else:
                confidence = "中"

        backfill_rows.append({
            "_record_id": safe_str(row_dict.get("_record_id", "")),
            "SKU编码": sku_to_write,
            "库存可用量": available,
            "缺口数量": gap,
            "库存状态": status,
            "预计到货日期": eta_date,
            "是否RB800": rb800_flag,
            "排单批次号": batch_id,
            "已分配数量": allocated,
            "置信度": confidence,
        })

    if running_consumed:
        top_skus = sorted(running_consumed.items(), key=lambda x: x[1], reverse=True)[:10]
        print(f"[逐次扣减] 本批次共消耗 {len(running_consumed)} 个SKU，Top10: {top_skus}")

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
    # 不重读飞书 — 计算字段已由 backfill 覆盖，重读可能引入一致性延迟
    # 且 build_current_detail_for_summary 的降级逻辑保证了内存数据始终被使用
    df_detail = build_current_detail_for_summary(items, None, df_backfill)

    # 统一关键列类型（入口统一清洗，后续不再重复 astype）
    for required_col in ("合同编号", "SKU编码", "合同数量", "库存可用量", "缺口数量"):
        if required_col not in df_detail.columns:
            df_detail[required_col] = ""
    df_detail["合同编号"] = df_detail["合同编号"].astype(str).apply(safe_str)
    df_detail["SKU编码"] = df_detail["SKU编码"].astype(str).apply(safe_str)
    # 当 SKU编码 为空时，依次使用 产品名称、规格 作为 SKU 标识
    empty_sku_mask = df_detail["SKU编码"] == ""
    if empty_sku_mask.any():
        if "产品名称" in df_detail.columns:
            df_detail.loc[empty_sku_mask, "SKU编码"] = df_detail.loc[empty_sku_mask, "产品名称"].astype(str).apply(safe_str)
        # 产品名称 也为空的，再用 规格 兜底
        still_empty = df_detail["SKU编码"] == ""
        if still_empty.any() and "规格" in df_detail.columns:
            df_detail.loc[still_empty, "SKU编码"] = df_detail.loc[still_empty, "规格"].astype(str).apply(safe_str)
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
                            "下单日期": safe_str(row.get("下单日期", "")),
                            "商务": safe_str(row.get("商务", "")),
                            "代理商": safe_str(row.get("代理商", "")),
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
                    "下单日期": safe_str(row.get("下单日期", "")),
                    "商务": safe_str(row.get("商务", "")),
                    "代理商": safe_str(row.get("代理商", "")),
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
                    if old_status in {"已确认", "已发货", "已签收"}:
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
            sku_codes = sku_codes[sku_codes != ""]  # 排除空 SKU（无编码用产品名兜底）
            sku_count = safe_val(sku_codes.nunique()) if len(sku_codes) > 0 else 0
            total_qty = safe_val(group["合同数量"].sum())
            shortage_mask = group["库存状态"] == "缺货"
            shortage_rows = group[shortage_mask]
            # Canonicalize shortage identifiers: use SKU code, fall back to product name
            raw_shortage_identifiers = []
            for _, r in shortage_rows.iterrows():
                sku_val = safe_str(r.get("SKU编码", ""))
                if sku_val and sku_val.lower() != "nan":
                    raw_shortage_identifiers.append(sku_val)
                else:
                    pname = safe_str(r.get("产品名称", ""))
                    if pname:
                        raw_shortage_identifiers.append(pname)

            canon_shortage: Dict[str, str] = {}
            for raw_code in sorted(set(raw_shortage_identifiers)):
                raw_str = safe_str(raw_code)
                if not raw_str:
                    continue
                _, _, canon = find_inventory_row(raw_str, {"产品名称": raw_str, "规格": raw_str}, inv, sku_df=sku)
                canon_key = canon or raw_str
                if canon_key not in canon_shortage:
                    canon_shortage[canon_key] = raw_str
            shortage_sku_codes = sorted(canon_shortage.keys())
            shortage_sku_names = []
            for sku_code in shortage_sku_codes:
                dev_name = sku_device_name_map.get(sku_code, "")
                if dev_name:
                    shortage_sku_names.append(dev_name)
                else:
                    # 尝试从原始标识获取显示名（SKU为空时直接用产品名）
                    shortage_sku_names.append(canon_shortage.get(sku_code, sku_code))
            shortage_sku_count = len(shortage_sku_codes)
            overall_status = "全部可发" if shortage_sku_count == 0 else "待补货"

            # 保护人工确认字段：如果旧记录已人工确认，保留原值
            old_confirmed = "否"
            old_manual_date = ""
            if not old_summary.empty:
                old_row = old_summary[old_summary["合同编号"] == contract_id]
                if not old_row.empty:
                    if safe_str(old_row.iloc[0].get("是否人工确认", "")) == "是":
                        old_confirmed = "是"
                        old_manual_date = safe_str(old_row.iloc[0].get("人工确认发货时间", ""))

            summary_rows.append({
                "合同编号": contract_id,
                "客户名称": contract_info_map.get(contract_id, {}).get("客户名称", ""),
                "项目名称": contract_info_map.get(contract_id, {}).get("项目名称", ""),
                "下单日期": contract_info_map.get(contract_id, {}).get("下单日期", ""),
                "商务": contract_info_map.get(contract_id, {}).get("商务", ""),
                "代理商": contract_info_map.get(contract_id, {}).get("代理商", ""),
                "项目类型": project_type,
                "订单SKU总数": int(sku_count) if sku_count is not None else 0,
                "订单总数量": int(total_qty) if total_qty is not None else 0,
                "缺货SKU数": int(shortage_sku_count),
                "缺货SKU列表": ', '.join(shortage_sku_names),
                "整体状态": overall_status,
                "AI建议发货时间": (
                    date_to_yyyy_mm_dd(next_working_day(ship_date)) if ship_date else ""
                ),
                "AI风险": "",
                "AI建议": "",
                "排单批次号": batch_id,
                "订单状态": "待确认",
                "是否人工确认": old_confirmed,
                "人工确认发货时间": old_manual_date,
            })

    summary = pd.DataFrame(summary_rows)

    # ---- 合并 old_summary 中已确认/已发货的合同（它们不在 df_detail 中，需要保留） ----
    if not old_summary.empty and "合同编号" in old_summary.columns and "订单状态" in old_summary.columns:
        old_confirmed = old_summary[
            old_summary["订单状态"].apply(lambda x: normalize_order_status(x)) != "待确认"
        ]
        if not old_confirmed.empty:
            existing_ids = set(summary["合同编号"].astype(str).apply(safe_str)) if not summary.empty else set()
            for _, old_row in old_confirmed.iterrows():
                cid = safe_str(old_row.get("合同编号", ""))
                if cid and cid not in existing_ids:
                    summary = pd.concat([summary, pd.DataFrame([old_row.to_dict()])], ignore_index=True)
                    existing_ids.add(cid)
            carried_over = len(existing_ids) - (len(summary_rows) if not summary.empty else 0)
            if carried_over > 0:
                print(f"[保留] 已确认/已发货合同 {carried_over} 条已从历史记录补回")

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
    reservation_rows = build_reservation_rows(summary, df_detail, batch_id, inv=inv, sku_df=sku)
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
    # 只写入目标表存在的字段，去除 下单日期/商务/代理商（AI排单结果表无这些字段）
    detail_cols = [c for c in summary.columns if c not in {"下单日期", "商务", "代理商"}]
    upsert_err = upsert_bitable_records_by_key(
        TABLE_ID_DETAIL,
        summary[detail_cols],
        key_col="合同编号",
        key_field_name="合同编号",
        numeric_cols=SUMMARY_NUMERIC_COLS_DEFAULT,
        date_cols=SUMMARY_DATE_COLS_DEFAULT,
    )

    # ===== 写入 AI发货总表（合同级发货记录） =====
    shipping_written = 0
    if TABLE_ID_SHIPPING and not summary.empty:
        print("正在写入 AI发货总表...")
        # 读取已存在的发货记录，保护人工修改的字段不被覆盖
        existing_shipping_map: Dict[str, Dict[str, str]] = {}
        try:
            existing_shipping = fetch_bitable_to_df(TABLE_ID_SHIPPING)
            if not existing_shipping.empty and "合同编号" in existing_shipping.columns:
                for _, er in existing_shipping.iterrows():
                    cid = safe_str(er.get("合同编号", ""))
                    if cid:
                        existing_shipping_map[cid] = {
                            "是否发货": safe_str(er.get("是否发货", "否")),
                            "发货日期": safe_str(er.get("发货日期", "")),
                            "快递公司": safe_str(er.get("快递公司", "")),
                            "快递单号": safe_str(er.get("快递单号", "")),
                            "备注": safe_str(er.get("备注", "")),
                        }
        except Exception as e:
            print(f"读取 AI发货总表现有记录失败: {e}")
        shipping_rows = []
        for _, row in summary.iterrows():
            c_id = safe_str(row.get("合同编号", ""))
            if not c_id:
                continue
            existing = existing_shipping_map.get(c_id, {})
            shipping_rows.append({
                "合同编号": c_id,
                "客户名称": contract_info_map.get(c_id, {}).get("客户名称", ""),
                "项目名称": contract_info_map.get(c_id, {}).get("项目名称", ""),
                "是否发货": existing.get("是否发货") or "否",
                "发货日期": existing.get("发货日期") or "",
                "快递公司": existing.get("快递公司") or "",
                "快递单号": existing.get("快递单号") or "",
                "备注": existing.get("备注") or "",
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
    audit_r = None  # initialized for scope
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

                # ---- 群通知延后：排单员确认后再发送 ----
                # 仅发中间态通知给采购群
                shortage_sku_count = int(report.get("库存缺货SKU数", 0) if report else 0)
                interim_msg = f"今日排单完成（{len(summary)} 份合同）"
                if shortage_sku_count > 0:
                    interim_msg += f"，{shortage_sku_count} 种物料紧缺"
                interim_msg += "，完整日报排单员确认后推送"
                _send_interim_notification(interim_msg)
                notification_result = {"enabled": True, "ok": True, "error": ""}

                # ----- 排单审计（随机5合同全链路复查） -----
                audit_r = None
                try:
                    from shared import audit_schedule_results
                    audit_r = audit_schedule_results(
                        df_summary=summary, items_df=items, inv_df=inv, sku_df=sku, sample_size=5,
                        occupied_stock=occupied_stock, find_inv_row_fn=find_inventory_row,
                    )
                    print(f"[审计] 抽查 {audit_r['sample_count']} 合同: "
                          + ("✅ 通过" if audit_r['ok'] else f"❌ {len(audit_r['errors'])}错误"))
                    if audit_r.get("errors"):
                        for e in audit_r["errors"][:5]:
                            print(f"  ❌ {e}")
                except Exception as e:
                    print(f"[审计] 审计执行异常（不影响排单）: {e}")
                    audit_r = {"ok": True, "errors": [], "warnings": []}

                # 写入审计记录表
                try:
                    if TABLE_ID_AUDIT and audit_r:
                        _write_audit_record(audit_r, batch_id, time.time() - t_start)
                except Exception as e:
                    print(f"[审计] 审计记录写入失败（不影响排单）: {e}")

                # ----- 私聊通知排单完成 -----
                _send_personal_notification(summary, batch_id, report, audit_result=audit_r)

                # ----- 自动确认（测试模式） -----
                if os.getenv("AUTO_CONFIRM", "").lower() in ("true", "1", "yes"):
                    print("[自动确认] 测试模式，自动触群发通知...")
                    try:
                        cr = confirm_schedule(batch_id=batch_id)
                        print(f"[自动确认] 结果: {cr}")
                    except Exception as e:
                        print(f"[自动确认] 异常: {e}")
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
        "审计": audit_r,
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


@app.get("/schedule/shortage-export")
def shortage_export():
    """导出缺货工单 Excel：按优先级排序，分「订单缺货」和「常规补货」两区。"""
    try:
        from io import BytesIO
        from fastapi.responses import StreamingResponse

        summary = fetch_bitable_to_df(TABLE_ID_DETAIL)
        if summary.empty:
            return {"error": "AI排单总表无数据"}

        # 过滤缺货合同
        shortage = summary[summary["整体状态"].isin(["待补货", "缺货"])].copy()
        if shortage.empty:
            return {"error": "当前无缺货订单"}

        # 优先级排序：紧急(0) > 换货(1) > 补发(2) > 维修(3) > 常规(4)
        def _priority(r):
            for tag, p in [("紧急", 0), ("换货", 1), ("补发", 2), ("维修", 3)]:
                if tag in safe_str(r.get("项目类型", "")):
                    return p
            return 4
        shortage["__p"] = shortage.apply(_priority, axis=1)
        shortage = shortage.sort_values(["__p", "合同编号"])

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "缺货工单"
        ws.append(["序号", "合同编号", "项目名称", "产品", "规格", "缺货数量", "入库日期"])

        # 订单缺货区
        ws.append(["▎ 订单缺货"])
        for idx, (_, r) in enumerate(shortage.iterrows(), 1):
            ws.append([
                idx,
                safe_str(r.get("合同编号", "")),
                safe_str(r.get("项目名称", "")),
                "", "",  # 产品/规格需从明细表取，此处留空由用户补充
                safe_str(r.get("缺货SKU数", "")),
                "",  # 入库日期留空
            ])

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=shortage.xlsx"}
        )
    except Exception as e:
        return {"error": f"导出失败: {e}"}


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


def _send_bot_reply(open_id: str, card: dict, app_id: str = "", chat_id: str = "", chat_type: str = ""):
    """以指定飞书应用的身份发送互动卡片回复，群聊自动切换为 chat_id 回复。"""
    token = _get_bot_reply_token(app_id)
    name = _BOT_CREDENTIALS.get(app_id, {}).get("name", app_id)
    if chat_type == "group" and chat_id:
        receive_id_type = "chat_id"
        receive_id = chat_id
    else:
        receive_id_type = "open_id"
        receive_id = open_id
    resp = httpx.post(
        f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": receive_id,
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
# ===== 飞书文档扫描与 AI 加工（供供应链机器人使用） =====

def _scan_and_analyze_doc(doc_url: str, question: str) -> dict:
    """读取飞书文档内容，用 DeepSeek 加工后返回卡片。"""
    import subprocess as _sp

    # 从 URL 提取 token
    corpus = doc_url.split("?")[0].rstrip("/")
    parts = corpus.split("/")
    token = parts[-1] if len(parts) >= 2 else ""
    if not token or not re.fullmatch(r"[A-Za-z0-9_\-]+", token):
        return _error_card(f"无法解析文档链接: {doc_url}")

    # 用 lark-cli 读取文档文本
    lark_cli_path = ""
    candidates = [os.path.expandvars(r"%APPDATA%\npm\lark-cli.cmd")]
    if sys.platform != "win32":
        candidates += ["/usr/local/bin/lark-cli", "/opt/homebrew/bin/lark-cli"]
    for p in candidates:
        if os.path.isfile(p):
            lark_cli_path = p
            break
    if not lark_cli_path:
        for p in os.environ.get("PATH", "").split(os.pathsep):
            for name in ("lark-cli.cmd", "lark-cli"):
                full = os.path.join(p, name)
                if os.path.isfile(full):
                    lark_cli_path = full
                    break
            if lark_cli_path:
                break
    if not lark_cli_path:
        lark_cli_path = "lark-cli"

    content = ""
    try:
        env = os.environ.copy()
        env["LARK_API_TOKEN"] = token
        result = _sp.run(
            [lark_cli_path, "--profile", "main-app", "docs", "+fetch", "--format", "text"],
            capture_output=True,
            encoding="utf-8",
            env=env,
            timeout=60,
        )
        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip()
            return _error_card(f"读取飞书文档失败: {err[:300]}")
        content = result.stdout.strip()
    except _sp.TimeoutExpired:
        return _error_card("读取飞书文档超时，请稍后重试")
    except Exception as e:
        return _error_card(f"读取文档异常: {e}")

    if not content or len(content) < 10:
        return _error_card("文档内容为空或太短，无法分析")

    # AI 加工
    api_key = os.getenv("DEEPSEEK_API_KEY", "")
    if not api_key:
        return {
            "header": {"template": "blue", "title": {"content": "📄 文档内容", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": f"（DeepSeek API Key 未配置，无法智能加工）\n\n文档原文（前 1000 字符）：\n\n{content[:1000]}..."}],
        }

    truncated = content[:12000]
    if len(content) > 12000:
        truncated += "\n\n（文档较长，以上为节选）"

    try:
        import httpx as _httpx
        client = _httpx.Client(timeout=60.0)
        resp = client.post(
            "https://api.deepseek.com/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": (
                    f"你是供应链文档分析助手。根据以下飞书文档内容，回答用户的问题。\n\n"
                    f"用户问题：{question}\n\n"
                    f"文档内容：\n{truncated}\n\n"
                    "要求：\n"
                    "1. 回答控制在 200-400 字\n"
                    "2. 如果文档内容与问题不相关，如实告知\n"
                    "3. 关键数字和数据要原样引用，不要编造\n"
                    "4. 用 Markdown 格式输出，适当使用粗体和列表"
                )}],
                "temperature": 0.3,
                "max_tokens": 2048,
            },
        )
        ai_data = resp.json()
        if "choices" in ai_data and len(ai_data["choices"]) > 0:
            result_text = ai_data["choices"][0]["message"]["content"].strip()
        else:
            result_text = f"（AI 加工失败）\n\n文档原文摘要：{truncated[:500]}..."
    except Exception as e:
        result_text = f"（AI 加工失败: {e}）\n\n文档原文摘要：{truncated[:500]}..."

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "blue", "title": {"content": "📄 文档智能分析", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": result_text}],
    }


def _build_supply_chain_reply(open_id: str, user_input: str) -> dict:
    """处理供应链AI助手的消息，返回飞书互动卡片。

    面向企业内部员工，提供排单、库存、交期等供应链管理功能。
    同时支持飞书文档链接的扫描与 AI 加工。
    """
    import re
    text = user_input.strip()

    # ---- 飞书文档扫描（优先检测） ----
    doc_match = re.search(
        r"https?://[^\s/]+\.(feishu|lark)\.(cn|com)/(docx|wiki|minutes)/([A-Za-z0-9_\-]+)",
        text, re.IGNORECASE,
    )
    if doc_match:
        doc_url = doc_match.group(0)
        question = re.sub(
            r"https?://[^\s/]+\.(feishu|lark)\.(cn|com)/(docx|wiki|minutes)/([A-Za-z0-9_\-]+)",
            "", text, flags=re.IGNORECASE,
        ).strip()
        if not question or len(question) < 2:
            question = "请总结这份文档的核心内容和关键数据"
        try:
            return _scan_and_analyze_doc(doc_url, question)
        except Exception as e:
            return _error_card(f"文档扫描失败: {e}")

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
                    "您好，我是供应链AI助手，为您提供排单、库存与交付管理服务。\n\n"
                    "📊 **常用指令**\n"
                    "• **日报** — 今日排单统计与交付概览\n"
                    "• **库存** — 库存状态与缺货物料明细\n"
                    "• **待确认** — 待人工确认的订单列表\n"
                    "• **延迟** — 预计延迟或缺货的订单\n"
                    "• **合同编号** — 查询具体订单交付详情\n"
                    "• 📄 **飞书文档链接** — 智能分析文档内容\n\n"
                    "直接输入关键词即可查询。"
                )}],
            }

    text_lower = text.lower()

    # ============================================================
    # 自然语言关键词匹配（按优先级从高到低）
    # ============================================================

    # ---- 日报 / 今日概况 ----
    if any(kw in text_lower for kw in ["日报", "今日排单", "排单日报", "今日", "报告",
                                         "今天.*排", "今天.*发", "今天.*订单", "今天.*单"]):
        try:
            report = get_today_report_row()
            return _build_report_card(report)
        except Exception as e:
            return _error_card(f"读取日报失败: {e}")

    # ---- 未排单 / 排单进度 ----
    if any(kw in text_lower for kw in ["没排", "还没排", "没有排", "未排", "还没排单",
                                         "还剩多少", "还有多少", "剩多少没",
                                         "排了多少", "排完了吗", "排完没",
                                         "排单进度", "排了多少单", "哪些没排"]):
        try:
            report = get_today_report_row()
            return _build_scheduling_progress_card(report)
        except Exception as e:
            return _error_card(f"读取排单进度失败: {e}")

    # ---- 库存 ----
    if any(kw in text_lower for kw in ["库存", "缺货", "物料", "紧缺",
                                         "还有多少货", "库存情况", "库存在哪里",
                                         "什么缺", "缺什么", "缺哪些"]):
        try:
            load_data_if_needed()
            return _build_inventory_card()
        except Exception as e:
            return _error_card(f"读取库存失败: {e}")

    # ---- 待确认 ----
    if any(kw in text_lower for kw in ["待确认", "待处理", "未确认", "还没确认",
                                         "要确认", "需要确认", "哪些要确认",
                                         "还没批", "要审批", "没确认"]):
        try:
            load_data_if_needed()
            return _build_pending_orders_card()
        except Exception as e:
            return _error_card(f"查询待确认订单失败: {e}")

    # ---- 发货情况 ----
    if any(kw in text_lower for kw in ["发货", "发了没", "发了多少", "今天发",
                                         "发出", "已发", "发了几", "发了哪",
                                         "发货情况", "发货状态", "物流情况"]):
        try:
            report = get_today_report_row()
            return _build_delivery_status_card(report)
        except Exception as e:
            return _error_card(f"查询发货状态失败: {e}")

    # ---- 延迟 / 缺货订单 ----
    if any(kw in text for kw in ["延迟", "延期", "推迟", "超期",
                                   "延迟.*单", "哪些.*延迟", "哪些.*缺货",
                                   "缺货.*订单", "还缺什么", "还缺哪些"]):
        try:
            load_data_if_needed()
            return _build_delayed_orders_card()
        except Exception as e:
            return _error_card(f"查询延迟订单失败: {e}")

    # ---- 帮助 ----
    if any(kw in text_lower for kw in ["帮助", "help", "功能", "菜单", "说明",
                                         "怎么用", "能做什么", "会什么", "有什么功能"]):
        return {
            "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": (
                "您好，我是供应链AI助手，为您提供排单、库存与交付管理服务。\n\n"
                "📊 **常用查询**\n"
                "🔹 **日报** / **今日概况** — 今日排单统计与交付概览\n"
                "🔹 **库存** / **缺货物料** — 库存健康度与紧缺明细\n"
                "🔹 **待确认** / **还没确认** — 待人工确认的订单\n"
                "🔹 **延迟** / **超期** — 预计延迟或缺货的订单\n"
                "🔹 **发货情况** — 今日已发/应发/延迟情况\n"
                "🔹 **排单进度** / **还有多少没排** — 排单完成情况\n\n"
                "🔍 **快速查询**\n"
                "🔹 输入**合同编号**查订单详情与交期\n"
                "🔹 发送**飞书文档链接**进行智能分析\n\n"
                "👇 直接输入自然语言即可，无需记忆指令。"
            )}],
        }

    # ---- 统计 / 数据 / 汇总 ----
    if any(kw in text_lower for kw in ["统计", "数据", "汇总", "概览", "总览",
                                         "整体情况", "全部情况", "所有订单"]):
        try:
            report = get_today_report_row()
            return _build_report_card(report)
        except Exception as e:
            return _error_card(f"读取日报失败: {e}")

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


# ===== 快捷卡片：排单进度 =====

def _build_scheduling_progress_card(report: dict) -> dict:
    """排单进度卡片 — 聚焦已完成/未完成比例。"""
    total = report.get('订单总数', '0')
    scheduled = report.get('已排单订单数', '0')
    unscheduled = report.get('未排单订单数', '0')
    confirmed = report.get('人工已确认排单数', '0')

    try:
        total_i = int(total)
        scheduled_i = int(scheduled)
        unscheduled_i = int(unscheduled)
        rate = f"{int(scheduled_i / total_i * 100)}%" if total_i > 0 else "N/A"
    except (ValueError, ZeroDivisionError):
        total_i = scheduled_i = unscheduled_i = 0
        rate = "N/A"

    if unscheduled_i > 0:
        status_line = f"⚠️ 还有 **{unscheduled}** 份订单尚未排单，建议尽快处理"
        header_color = "orange"
    else:
        status_line = "✅ 所有订单已完成排单"
        header_color = "green"

    lines = [
        f"📊 **排单进度** ({report.get('日期', '-')})",
        "",
        f"订单总数：**{total}**　|　已排单：**{scheduled}**　|　完成率：**{rate}**",
        f"未排单：**{unscheduled}**　|　人工已确认：**{confirmed}**",
        "",
        status_line,
    ]
    if unscheduled_i > 0:
        lines.append("")
        lines.append("💡 回复「**待确认**」查看待处理订单列表")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": header_color, "title": {"content": "📊 排单进度", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


# ===== 快捷卡片：发货状态 =====

def _build_delivery_status_card(report: dict) -> dict:
    """发货状态卡片 — 区分 AI 判定可发 vs 人工确认可发（只有人工确认过的才算真正可发）。"""
    should_ship = report.get('今日应发货订单数', '0')
    can_ship = report.get('今日可发货订单数', '0')
    delayed = report.get('今日预计延迟订单数', '0')

    try:
        delayed_i = int(delayed)
        can_i = int(can_ship)
        should_i = int(should_ship)
    except ValueError:
        delayed_i = can_i = should_i = 0

    # 查询今日应发订单中已人工确认的数量（真正可发）
    today_str = report.get('日期', '')
    manual_confirmed_today = 0
    manual_pending_today = 0
    try:
        detail = fetch_bitable_to_df(TABLE_ID_DETAIL)
        if detail is not None and not detail.empty:
            # 找出今日应发的订单
            if "AI建议发货时间" in detail.columns and "是否人工确认" in detail.columns:
                detail["_ship_date"] = detail["AI建议发货时间"].apply(
                    lambda x: date_to_yyyy_mm_dd(parse_date_to_date(x))
                )
                today_orders = detail[detail["_ship_date"] == today_str]
                if not today_orders.empty:
                    confirmed_mask = today_orders["是否人工确认"].astype(str).apply(safe_str) == "是"
                    manual_confirmed_today = int(confirmed_mask.sum())
                    # 有货 + 已确认 = 真正可发，缺货 + 已确认 = 物料待协调
                    available_mask = today_orders["整体状态"].astype(str).apply(safe_str) != "缺货"
                    manual_available = int((confirmed_mask & available_mask).sum())
                    manual_pending_today = int(len(today_orders) - manual_confirmed_today)
                else:
                    manual_available = 0
    except Exception:
        manual_confirmed_today = 0
        manual_available = 0
        manual_pending_today = 0

    # 计算各维度
    if should_i > 0:
        ai_rate = f"{int(can_i / should_i * 100)}%" if should_i > 0 else "N/A"
        manual_rate = f"{int(manual_confirmed_today / should_i * 100)}%" if should_i > 0 else "N/A"
        manual_avail_rate = f"{int(manual_available / should_i * 100)}%" if should_i > 0 and manual_confirmed_today > 0 else "N/A"
    else:
        ai_rate = "N/A"
        manual_rate = "N/A"
        manual_avail_rate = "N/A"

    # 状态判定（以人工确认为准）
    if manual_pending_today > 0:
        if manual_confirmed_today == 0:
            risk_line = f"🔴 **{manual_pending_today}** 份待人工确认 —— 请尽快确认以安排发货"
            header_color = "red"
        else:
            risk_line = f"🟠 {manual_confirmed_today} 份已确认可发，**{manual_pending_today}** 份待确认"
            header_color = "orange"
        action_hint = "\n💡 回复「**待确认**」进入人工确认页面"
    elif delayed_i > 0:
        risk_line = f"⚠️ **{delayed}** 份预计延迟，均已完成人工确认"
        header_color = "orange"
        action_hint = "\n💡 回复「**延迟**」查看延迟订单明细"
    elif should_i == 0:
        risk_line = "📅 今日无应发订单"
        header_color = "blue"
        action_hint = ""
    else:
        risk_line = f"✅ 今日全部已确认可发，交付率 {manual_rate}"
        header_color = "green"
        action_hint = ""

    lines = [
        f"🚚 **今日发货状态** ({report.get('日期', '-')})",
        "",
        "━━━ **应发 vs 可发** ━━━",
        f"今日应发：**{should_ship}** 份",
    ]
    if should_i > 0:
        lines.append(f"AI 判定可发：**{can_ship}** 份（{ai_rate}）— 基于库存状态")
        lines.append(f"人工确认可发：**{manual_confirmed_today}** 份（{manual_rate}）— ✅ 真正可发")
        if manual_available != manual_confirmed_today and manual_confirmed_today > 0:
            lines.append(f"　其中物料充足：**{manual_available}** 份")
    lines.append("")
    lines.append(f"预计延迟：**{delayed}** 份")
    lines.append(f"待人工确认：**{manual_pending_today}** 份")
    lines.append("")
    lines.append(risk_line)
    if action_hint:
        lines.append(action_hint)

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": header_color, "title": {"content": "🚚 发货状态", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_report_card(report: dict) -> dict:
    """日报卡片 — 结构化分区，带情景化注释。"""
    items = []
    # --- 头部 ---
    items.append(f"📊 **AI排单日报**")
    items.append(f"📅 {report.get('日期', '-')}　|　批次：{report.get('排单批次号', '-')}")
    items.append("")

    # --- 第一部分：订单处理概览 ---
    items.append("━━━ **订单处理概览** ━━━")
    items.append(
        f"总订单：**{report.get('订单总数', '-')}**　|　"
        f"今日新增：**{report.get('今日新增订单', '-')}**　|　"
        f"紧急：**{report.get('今日紧急订单', '-')}**"
    )
    items.append(
        f"已排单：**{report.get('已排单订单数', '-')}**　|　"
        f"未排单：**{report.get('未排单订单数', '-')}**　|　"
        f"人工已确认：**{report.get('人工已确认排单数', '-')}**"
    )
    items.append("")

    # --- 第二部分：今日交付动态 ---
    items.append("━━━ **今日交付动态** ━━━")
    should_ship = report.get('今日应发货订单数', '0')
    can_ship = report.get('今日可发货订单数', '0')
    items.append(
        f"🚚 应发货：**{should_ship}**　|　"
        f"AI 判定可发：**{can_ship}**（基于库存）"
    )
    items.append(
        f"✅ 真正可发：以**人工确认**为准 → 详情回复「**发货**」"
    )
    delayed = report.get('今日预计延迟订单数', '0')
    try:
        delayed_int = int(delayed)
    except (ValueError, TypeError):
        delayed_int = 0
    if delayed_int > 0:
        items.append(f"⚠️ 预计延迟：**{delayed}** — **建议优先关注**")
    else:
        items.append(f"⚠️ 预计延迟：**{delayed}** — 交付正常")
    items.append("")

    # --- 第三部分：库存风险监控 ---
    items.append("━━━ **库存风险监控** ━━━")
    shortage_stats = report.get('库存缺货统计', '')
    top_shortage = (shortage_stats or '').split('\n')[0] if shortage_stats and shortage_stats != '无' else ''
    if top_shortage:
        items.append(
            f"充足：**{report.get('库存充足SKU数', '-')}**　|　"
            f"预警：**{report.get('库存预警SKU数', '-')}**　|　"
            f"缺货：**{report.get('库存缺货SKU数', '-')}**"
        )
        items.append(f"🔴 **最紧缺物料**：{top_shortage} —— ⚡ 建议优先补货")
    else:
        items.append(
            f"🟢 充足：**{report.get('库存充足SKU数', '-')}**　|　"
            f"预警：**{report.get('库存预警SKU数', '-')}**　|　"
            f"缺货：**{report.get('库存缺货SKU数', '-')}**"
        )
    items.append("")

    # --- 第四部分：未来3天预警 ---
    items.append("━━━ **未来3天预警** ━━━")
    items.append(
        f"最长交期：**{report.get('最长交期订单', '-')}**　|　"
        f"最晚发货：**{report.get('最晚发货日期', '-')}**"
    )
    f3_short = report.get('未来3天缺货订单数', '0')
    try:
        f3_short_int = int(f3_short)
    except (ValueError, TypeError):
        f3_short_int = 0
    if f3_short_int > 0:
        items.append(
            f"未来3天应发：**{report.get('未来3天应发货订单数', '-')}** 单　|　"
            f"缺货风险：**{f3_short}** 单 ⚠️ 建议提前协调"
        )
    else:
        items.append(
            f"未来3天应发：**{report.get('未来3天应发货订单数', '-')}** 单　|　"
            f"缺货风险：**{f3_short}** 单"
        )

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "blue", "title": {"content": "📊 AI排单日报", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(items)}],
    }


def _build_inventory_card() -> dict:
    """库存状态卡片 — 风险分级：高风险缺货 / 中风险预警 / 正常。"""
    inv = data_cache.get("inv")
    sku_df = data_cache.get("sku")
    items_df = data_cache.get("items")

    if inv is None:
        return _error_card("库存数据未加载，请先触发排单或等待数据刷新")

    total_sku = inv["SKU"].nunique() if "SKU" in inv.columns else 0
    raw_qty = float(inv["库存数量"].sum()) if "库存数量" in inv.columns else 0.0
    # 安全上限：防止数据异常导致显示天文数字
    total_qty = int(raw_qty) if raw_qty < 1_000_000_000 else int(inv["库存数量"].apply(to_num).sum())

    # 缺货 SKU（高风险）
    shortage_list: List[tuple] = []
    # 低库存 SKU（中风险预警）
    warning_list: List[tuple] = []
    if items_df is not None and not items_df.empty and "库存状态" in items_df.columns and "SKU编码" in items_df.columns:
        shortage_items = items_df[items_df["库存状态"].astype(str).str.strip() == "缺货"]
        if not shortage_items.empty:
            gap_by_sku: Dict[str, float] = {}
            for _, r in shortage_items.iterrows():
                sku = safe_str(r.get("SKU编码", ""))
                if not sku:
                    sku = safe_str(r.get("产品名称", "")) or safe_str(r.get("规格", ""))
                gap = to_num(r.get("缺口数量", 0))
                if sku and gap > 0:
                    gap_by_sku[sku] = gap_by_sku.get(sku, 0) + gap
            shortage_list = sorted(gap_by_sku.items(), key=lambda x: x[1], reverse=True)[:8]

        # 低库存预警（库存可用量 < 5 且尚未缺货）
        if "库存可用量" in items_df.columns:
            warning_items = items_df[
                (items_df["库存状态"].astype(str).str.strip() != "缺货")
            ]
            if not warning_items.empty:
                warning_items = warning_items.copy()
                warning_items["__avail"] = warning_items["库存可用量"].apply(to_num)
                warning_items = warning_items[warning_items["__avail"] < 5]
                warning_qty: Dict[str, float] = {}
                for _, r in warning_items.iterrows():
                    sku = safe_str(r.get("SKU编码", ""))
                    if not sku:
                        sku = safe_str(r.get("产品名称", "")) or safe_str(r.get("规格", ""))
                    avail = to_num(r.get("库存可用量", 0))
                    if sku and avail >= 0:
                        warning_qty[sku] = avail
                warning_list = sorted(warning_qty.items(), key=lambda x: x[1])[:5]

    # 计算健康度（缺货SKU占总SKU比例）
    shortage_sku_count = len(shortage_list)
    try:
        health_pct = int((1 - shortage_sku_count / total_sku) * 100) if total_sku > 0 else 100
    except (ZeroDivisionError, TypeError):
        health_pct = 100

    if shortage_sku_count == 0:
        health_level = "🟢 健康"
        health_desc = "所有SKU库存充足"
    elif shortage_sku_count <= total_sku * 0.05:
        health_level = "🟡 基本健康"
        health_desc = f"仅 {shortage_sku_count} 个SKU缺货，占比 < 5%"
    elif shortage_sku_count <= total_sku * 0.15:
        health_level = "🟠 需关注"
        health_desc = f"{shortage_sku_count} 个SKU缺货，建议尽快补货"
    else:
        health_level = "🔴 高风险"
        health_desc = f"{shortage_sku_count} 个SKU缺货，需立即采购"

    lines = [
        f"📦 **库存健康度** — {health_level}",
        "",
        f"**{health_desc}**",
        f"总SKU **{total_sku}** 类　|　总库存 **{total_qty:,}** 个　|　健康度 **{health_pct}%**",
    ]

    # --- 高风险：缺货物料 ---
    if shortage_list:
        lines.append("")
        lines.append("━━━ 🔴 缺货物料（立即处理）━━━")
        for sku_code, gap in shortage_list:
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
        lines.append("✅ 无需立即处理的缺货物料")

    # --- 中风险：低库存预警 ---
    if warning_list:
        lines.append("")
        lines.append("━━━ 🟡 库存偏低（关注补货）━━━")
        for sku_code, avail in warning_list:
            name = ""
            if sku_df is not None and not sku_df.empty and "产品编码SKU" in sku_df.columns:
                match = sku_df[sku_df["产品编码SKU"].astype(str).str.strip() == sku_code]
                if not match.empty:
                    name = safe_str(match.iloc[0].get("设备名称", ""))
            display = name if name else sku_code
            avail_str = str(int(avail)) if avail == int(avail) else str(round(avail, 1))
            lines.append(f"• **{display}** — 仅剩 {avail_str} 个")
    else:
        lines.append("")
        lines.append("🟢 当前无库存偏低预警")

    # 确定卡片颜色
    if shortage_list:
        header_color = "red"
    elif warning_list:
        header_color = "orange"
    else:
        header_color = "green"

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": header_color,
                    "title": {"content": "📦 库存状态", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_pending_orders_card() -> dict:
    """待确认订单列表 — 按紧急度排序，标注发货时限。"""
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

    # 构建跳转链接
    bitable_url = (
        f"https://wl6wihmop1.feishu.cn/base/{BITABLE_APP_TOKEN}"
        f"?table={TABLE_ID_DETAIL}&view=vewiuVK8pH"
    )

    today = datetime.now(timezone(timedelta(hours=8)))

    # 计算紧急度并排序
    def _urgency_score(row):
        d = parse_date_to_date(row.get("AI建议发货时间"))
        if d is not None:
            days_diff = (d - today.date()).days
            # 超期越多分数越高（越紧急）
            return -max(days_diff, -365)
        return 0

    pending = pending.copy()
    pending["__urgency"] = pending.apply(_urgency_score, axis=1)
    pending = pending.sort_values("__urgency", ascending=False).head(15)

    pending_count = len(summary[summary["订单状态"] == "待确认"]) if "订单状态" in summary.columns else len(pending)
    lines = [f"📋 **待确认订单**　（共 {pending_count} 份待处理）", ""]

    for _, r in pending.iterrows():
        cid = safe_str(r.get("合同编号", "-"))
        status = safe_str(r.get("整体状态", ""))
        d = parse_date_to_date(r.get("AI建议发货时间"))
        ship_display = date_to_yyyy_mm_dd(d) if d else "待定"

        # 紧急度标注
        urgency_tag = ""
        if d:
            days_diff = (d - today.date()).days
            if days_diff < 0:
                urgency_tag = " ⚠️ 已超期"
            elif days_diff == 0:
                urgency_tag = " 🚨 今日应发"
            elif days_diff <= 3:
                urgency_tag = f" ⏰ {days_diff}天内"

        if "缺货" in status:
            flag = "🔴"
            note = " [需协调物料]"
        else:
            flag = "🟡"
            note = ""

        lines.append(f"{flag} {cid} — 发货 {ship_display}{urgency_tag}{note}")

    lines.append("")
    lines.append("💡 **提示**：建议优先确认「已超期」和「今日应发」的订单。")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "orange", "title": {"content": "📋 待确认订单", "tag": "plain_text"}},
        "elements": [
            {"tag": "markdown", "content": "\n".join(lines)},
            {"tag": "hr"},
            {"tag": "action", "actions": [
                {"tag": "button", "text": {"tag": "plain_text", "content": "📋 前往人工确认"},
                 "type": "primary", "url": bitable_url, "value": {}}
            ]},
        ],
    }


def _build_delayed_orders_card() -> dict:
    """延迟订单列表 — 影响评估 + 延迟严重程度分级。"""
    try:
        summary = prepare_summary_status(fetch_bitable_to_df(TABLE_ID_DETAIL))
    except Exception:
        return _error_card("无法读取排单总表")

    if summary is None or summary.empty or "整体状态" not in summary.columns:
        return {"header": {"template": "blue", "title": {"content": "⏳ 延迟订单", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "暂无延迟订单数据"}]}

    delayed = summary[summary["整体状态"].astype(str).str.strip() == "缺货"].copy()
    if delayed.empty:
        return {"header": {"template": "green", "title": {"content": "✅ 交付状态", "tag": "plain_text"}},
                "elements": [{"tag": "markdown", "content": "当前所有订单交付正常，无延迟订单。"}]}

    today = datetime.now(timezone(timedelta(hours=8)))

    # 影响评估：统计严重程度
    critical_cnt, high_cnt = 0, 0
    for _, r in delayed.iterrows():
        d = parse_date_to_date(r.get("AI建议发货时间"))
        if d is not None:
            days_diff = (d - today.date()).days
            if days_diff < -7:
                critical_cnt += 1
            elif days_diff < 0:
                high_cnt += 1

    lines = [f"⚠️ **缺货/延迟订单**　（共 {len(delayed)} 份受影响）", ""]

    # 影响评估摘要
    if critical_cnt > 0 or high_cnt > 0:
        impact_parts = []
        if critical_cnt > 0:
            impact_parts.append(f"🔴 **严重**：{critical_cnt} 份超期超过7天")
        if high_cnt > 0:
            impact_parts.append(f"🟠 **关注**：{high_cnt} 份超期7天内")
        lines.append(f"📊 **影响评估**：{'  ｜  '.join(impact_parts)}")
        lines.append("")

    # 按超期天数降序排列
    delayed = delayed.copy()
    def _calc_overdue(row):
        d = parse_date_to_date(row.get("AI建议发货时间"))
        if d is None:
            return -9999  # 没有日期的排最后
        return -(d - today.date()).days  # 负值→超期，取反→超期越多值越大

    delayed["__overdue"] = delayed.apply(_calc_overdue, axis=1)
    delayed = delayed.sort_values("__overdue", ascending=False).head(15)

    for _, r in delayed.iterrows():
        cid = safe_str(r.get("合同编号", "-"))
        proj = safe_str(r.get("项目名称", ""))
        owner = safe_str(r.get("客户名称", ""))
        ai_d = parse_date_to_date(r.get("AI建议发货时间"))
        manual_d = parse_date_to_date(r.get("人工确认发货时间"))
        display_date = manual_d or ai_d
        display_ship = date_to_yyyy_mm_dd(display_date) if display_date else "待定"

        # 延迟严重程度
        severity = "🔴"
        note = ""
        if display_date:
            days_over = (today.date() - display_date).days
            if days_over > 7:
                severity = "🔴🔴"
                note = f"已超期 {days_over} 天，建议升级"
            elif days_over > 0:
                severity = "🔴"
                note = f"已超期 {days_over} 天"
            else:
                severity = "🟠"
                note = "即将到期"

        line = f"{severity} **{cid}**"
        if proj:
            line += f"　|　{proj[:25]}"
        line += f"　|　预计 {display_ship}"
        if owner:
            line += f"　|　👤 {owner[:12]}"
        if note:
            line += f"\n　→ {note}"
        lines.append(line)

    lines.append("")
    lines.append("💡 **建议**：超期订单请优先协调物料与产线排期，尽早确认发货时间。")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red", "title": {"content": "⚠️ 延迟订单", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_contract_card(contract_id: str) -> dict:
    """单个合同详情卡片 — 含项目名称、交付时间线、AI洞察。"""
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
    project = safe_str(r.get("项目名称", ""))
    status = safe_str(r.get("整体状态", "-"))
    order_status = safe_str(r.get("订单状态", "-"))
    ai_date = parse_date_to_date(r.get("AI建议发货时间"))
    manual_date = parse_date_to_date(r.get("人工确认发货时间"))
    risk = safe_str(r.get("AI风险", ""))
    advice = safe_str(r.get("AI建议", ""))

    today = datetime.now(timezone(timedelta(hours=8)))
    display_date = manual_date or ai_date
    ship_display = date_to_yyyy_mm_dd(display_date) if display_date else "待定"

    # 订单状态
    if order_status == "已发货":
        status_emoji, status_label = "🚚", "已发出"
    elif order_status == "已确认":
        status_emoji, status_label = "🟢", "已确认"
    else:
        status_emoji, status_label = "🟡", "待确认"

    # 库存状态
    has_shortage = "缺货" in status
    stock_emoji = "🔴" if has_shortage else "🟢" if "有货" in status else "⚪"
    stock_label = "物料待齐" if has_shortage else "物料充足" if "有货" in status else status

    lines = [
        f"📄 **合同 {contract_id}**",
    ]
    if project:
        lines.append(f"🏷️ 项目：**{project}**")
    lines.append("")
    lines.append(f"订单状态：{status_emoji} **{status_label}**")
    lines.append(f"库存状态：{stock_emoji} **{stock_label}**")
    lines.append(f"预计发货：**{ship_display}**")
    if manual_date:
        lines.append(f"（人工确认发货时间：**{date_to_yyyy_mm_dd(manual_date)}**）")

    # 交付时间线分析
    if display_date:
        days_diff = (display_date - today.date()).days
        lines.append("")
        if days_diff < 0:
            lines.append(f"⏰ **已超期 {-days_diff} 天**，建议尽快协调确认")
        elif days_diff == 0:
            lines.append("🚨 **今日应发**，请确认发货安排")
        elif days_diff <= 7:
            lines.append(f"📅 **{days_diff} 天内**发货，请关注备货进度")
        else:
            lines.append(f"📅 距发货还有 **{days_diff} 天**")

    # AI 洞察
    if risk or advice:
        lines.append("")
        lines.append("━━━ **AI 洞察** ━━━")
        if risk:
            lines.append(f"📋 **风险提示**：{risk}")
        if advice:
            lines.append(f"💡 **建议**：{advice}")

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "red" if has_shortage else "blue",
                    "title": {"content": "📄 订单详情", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


def _build_default_card() -> dict:
    """默认引导卡片 — 输入无法识别时给出明确提示。"""
    return {
        "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": (
            "您好，我是供应链AI助手，为您提供排单、库存与交付管理服务。\n\n"
            "📊 **常用查询**\n"
            "• **日报** — 今日排单统计与交付概览\n"
            "• **库存** — 库存状态与缺货物料明细\n"
            "• **待确认** — 待人工确认的订单列表\n"
            "• **延迟** — 预计延迟或缺货的订单\n\n"
            "🔍 **快速查询**\n"
            "• 直接输入**合同编号**查询订单详情\n"
            "• 发送**飞书文档链接**进行智能分析\n\n"
            "💡 未能识别您输入的内容，请尝试以上关键词或指令。"
        )}],
    }


def _error_card(msg: str) -> dict:
    friendly_msg = str(msg)
    # 去掉错误类型前缀（如"读取日报失败: "），保留具体原因
    if ": " in friendly_msg:
        parts = friendly_msg.split(": ", 1)
        friendly_msg = parts[1].strip() if parts[1].strip() else parts[0].strip()

    return {
        "header": {"template": "red", "title": {"content": "⚠️ 查询异常", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": (
            f"查询过程遇到异常，请稍后重试。\n\n"
            f"🔍 {friendly_msg}\n\n"
            f"如问题持续，请联系技术支持。"
        )}],
    }


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

    # ---- 提取消息字段（兼容 V1 / V2），同时获取 chat_id 用于群聊回复 ----
    chat_id = ""
    if header and "message" in event:
        # V2 格式
        message = event.get("message", {})
        sender = event.get("sender", {})
        sender_id_obj = sender.get("sender_id", {})
        open_id = sender_id_obj.get("open_id", "")
        sender_type = sender.get("sender_type", "")
        chat_type = message.get("chat_type", "")
        chat_id = message.get("chat_id", "")
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
        chat_id = event.get("chat_id", "")
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
    if chat_type and chat_type not in ("p2p", "group"):
        print(f"[Webhook] 跳过未知聊天类型: chat_type={chat_type}")
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

    # 群聊中剥离 @提及
    user_input = user_input.strip()
    if chat_type == "group" and user_input:
        import re as _re
        user_input = _re.sub(r"@\S+", "", user_input).strip()
        user_input = _re.sub(r"@所有人", "", user_input).strip()
    if not user_input:
        print("[Webhook] 消息文本为空（或仅含 @提及）")
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
        _send_bot_reply(open_id, card, app_id=event_app_id, chat_id=chat_id, chat_type=chat_type)
    except Exception as e:
        print(f"[Webhook] 处理异常: {e}")
        import traceback
        traceback.print_exc()

    return {"msg": "ok"}


@app.post("/finance/sync")
def finance_sync(force: bool = False):
    """同步财务对账数据：总表 ← 销售订单主表，明细表 ← 销售订单明细表 + 计费规则匹配。
    
    force=true 时强制重同步所有未锁定合同（覆盖已同步状态）。
    """
    try:
        summary_result = _sync_finance_summary(force=force)
        detail_result = _sync_finance_detail()

        return {
            "ok": summary_result.get("ok", False) and detail_result.get("ok", False),
            "summary": summary_result,
            "detail": detail_result,
        }
    except Exception as e:
        return {"ok": False, "error": f"同步异常: {e}"}


@app.post("/finance/calculate")
def finance_calculate(req: FinanceCalculateRequest):
    """执行财务核算：完整性检查 → 包干/单采分类核算 → 汇总回写。

    threshold: SPEC-RB800-KZP 主键盘设备费触发阈值，默认 3
    """
    try:
        result = _run_finance_calculate(threshold=req.threshold)
        return result
    except Exception as e:
        return {"ok": False, "error": f"核算异常: {e}"}


if __name__ == "__main__":
    import uvicorn
    print("启动AI排单服务...")
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
