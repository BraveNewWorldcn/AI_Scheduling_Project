from fastapi import FastAPI, Request
import os
import re
import time
from pydantic import BaseModel
import pandas as pd
from datetime import datetime, timedelta, date, timezone
import json 
from typing import Any, Dict, List, Optional, Tuple
# =========================
# 本地缓存目录
# =========================
CACHE_DIR = "cache"
os.makedirs(CACHE_DIR, exist_ok=True)
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
from lark_oapi.api.bitable.v1.model import AppTableRecord, CreateAppTableRecordRequest
import httpx

# =========================
# 飞书字段清洗函数
# =========================
def parse_feishu_field(value):
    """清洗飞书字段：优先提取富文本包装中的纯文本，去空格，空值转空字符串"""
    # 飞书单行/多行文本实际格式：[{"text":"xxx","type":"text"}]
    if isinstance(value, list):
        texts = []
        for item in value:
            if isinstance(item, dict) and "text" in item:
                texts.append(item["text"].strip())
            else:
                texts.append(str(item).strip())
        # 单值直接返回文本，多值用逗号拼接
        if len(texts) == 1:
            return texts[0]
        return ", ".join(texts)
    elif isinstance(value, dict):
        for k in ("text", "number", "phone", "email", "url"):
            if k in value:
                return str(value[k]).strip()
        return str(value).strip()
    elif isinstance(value, (int, float)):
        # 保留数字原始类型，不转字符串
        return value
    elif isinstance(value, str):
        return value.strip()
    elif value is None:
        return ""
    else:
        return str(value).strip()

# ========== 飞书配置（请替换为你自己的信息） ==========
FEISHU_APP_ID = "cli_a96c5d017d3a1cbb"           # TODO: 替换
FEISHU_APP_SECRET = "Bk7RzLFMmeVfsERXIoazcbyXKzRm7fE5"  # TODO: 替换
BITABLE_APP_TOKEN = "C5JzbAfnia0nT3sRvjucXgUGnDc"            # TODO: 替换为多维表格的 app_token（base/ 后面那串）

# ===== 数据表 ID 配置 =====
# 【飞书表结构说明】
# 1. 销售订单总表 - 合同级别的聚合数据（暂未使用，可选增强）
#    字段：合同编号、下单日期、订单来源、是否紧急订单、是否换货订单、是否补发订单、是否维修订单、当前状态、缺货、排单批次等
# 
# 2. 销售订单明细表 ⭐ 【主表，必需】
#    字段：合同编号、SKU编码、合同数量、项目类型、产品名称、是否RB800
#
# 3. SKU标准表 ⭐ 【主表，必需】  
#    字段：产品编码SKU、是否新产品、是否外采、是否自研（用于发货周期判断）
#
# 4. 库存快照表 ⭐ 【主表，必需】
#    字段：SKU、库存数量、待采购出库、在途数量、库存日期（用于新鲜度检查）
#

TABLE_ID_ITEMS = "tblJn5iP6imjzE8h"                   # 销售订单明细表 ID（必需）
TABLE_ID_SKU   = "tblVAGWeGHvmbFgJ"                   # SKU标准表 ID（必需）
TABLE_ID_INV   = "tblFZNdEwW50izjh"                   # 库存快照表 ID（必需）
TABLE_ID_SUMMARY = "tblgMzZPyBBU5GLX"                 # 合同缺货总览表 ID（输出）
TABLE_ID_DETAIL = "tbl09Z6C7wCGh3mW"

# 环境变量优先于上方默认值（部署时可覆盖，避免硬编码写死）
if os.getenv("TABLE_ID_ITEMS"):
    TABLE_ID_ITEMS = os.getenv("TABLE_ID_ITEMS").strip()
if os.getenv("TABLE_ID_SKU"):
    TABLE_ID_SKU = os.getenv("TABLE_ID_SKU").strip()
if os.getenv("TABLE_ID_INV"):
    TABLE_ID_INV = os.getenv("TABLE_ID_INV").strip()
if os.getenv("TABLE_ID_SUMMARY"):
    TABLE_ID_SUMMARY = os.getenv("TABLE_ID_SUMMARY").strip()
if os.getenv("TABLE_ID_DETAIL"):
    TABLE_ID_DETAIL = os.getenv("TABLE_ID_DETAIL").strip()

# ===== 输出表字段映射 =====
# 当飞书表中的字段名与代码中的字段名不一致时，在这里进行映射
# 格式：{"代码中的字段名": "飞书表中的实际字段名"}
OUTPUT_DETAIL_FIELDS_MAP = {
    "排单ID": "排单ID",
    "合同编号": "合同编号",
    "SKU编码": "SKU编码",
    "合同数量": "合同数量",
    "库存可用量": "库存可用量",
    "缺口数量": "缺口数量",
    "库存状态": "库存状态",
    "预计到货日期": "预计到货日期",
    "是否RB800": "是否RB800",
    "排单批次号": "排单批次号",
}

OUTPUT_SUMMARY_FIELDS_MAP = {
    "合同编号": "合同编号",
    "项目类型": "项目类型",
    "订单SKU总数": "订单SKU总数",
    "订单总数量": "订单总数量",
    "缺货SKU数": "缺货SKU数",
    "缺货SKU列表": "缺货SKU列表",
    "整体状态": "整体状态",
    "AI建议发货时间": "AI建议发货时间",
    "排单批次号": "排单批次号",
    "是否人工确认": "是否人工确认",
    "人工确认发货时间": "人工确认发货时间"
}
# ================================================
SUMMARY_NUMERIC_COLS_DEFAULT = {"订单SKU总数", "订单总数量", "缺货SKU数"}
SUMMARY_DATE_COLS_DEFAULT = {"AI建议发货时间", "人工确认发货时间"}
app = FastAPI()

# =========================
# Pydantic 模型定义
# =========================
class ScheduleRequest(BaseModel):
    trigger: str = "test"

# =========================
# 飞书客户端（全局）
# =========================
feishu_client = lark.Client.builder() \
    .app_id(FEISHU_APP_ID) \
    .app_secret(FEISHU_APP_SECRET) \
    .build()

def fetch_bitable_to_df(table_id: str):
    """从飞书多维表格读取数据，解析富文本包装，清洗后返回 DataFrame。

    额外保留飞书 record_id 到列 `_record_id`，用于后续“回填更新”而不是新增记录。
    """
    # 1. 获取 tenant_access_token
    token_url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    token_resp = httpx.post(token_url, json={
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    }, timeout=30.0)   # ← 添加超时时间
    token_data = token_resp.json()
    if token_data.get("code") != 0:
        raise Exception(f"获取飞书token失败: {token_data}")

    access_token = token_data["tenant_access_token"]
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    # 2. 分页读取所有记录
    all_records = []
    page_token = None
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{table_id}/records"

    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token

        resp = httpx.get(url, headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"飞书读取表格失败: {data}")

        items = data.get("data", {}).get("items", [])
        for item in items:
            row = {"_record_id": item.get("record_id", "")}
            for field_name, value in item.get("fields", {}).items():
                # 先解析飞书富文本包装，再清洗
                cleaned = parse_feishu_field(value)
                row[field_name] = cleaned
            all_records.append(row)

        if not data.get("data", {}).get("has_more", False):
            break
        page_token = data["data"].get("page_token")
        if not page_token:
            break

    # 3. 构造 DataFrame 并清洗列名
    df = pd.DataFrame(all_records)

    if not df.empty:
        # 第一步：清洗所有列名（去掉首尾空格）
        df.columns = df.columns.str.strip()
        
        # 第二步：对所有列值进行字符串清洗（去空格）
        for col in df.columns:
            if df[col].dtype == 'object':
                try:
                    df[col] = df[col].astype(str).str.strip()
                except Exception:
                    pass  # 某些列无法转换，忽略
    
    # 调试：打印实际的列名，方便排查
    print(f"[DEBUG] 读取表 {table_id} 的列名: {list(df.columns)}")
    
    return df

# =========================
# ⭐新增：真正从飞书加载三张表（给缓存函数调用）
# =========================
def load_feishu_data():
    print("=" * 60)
    print("正在从飞书加载三张核心表...")
    print("=" * 60)

    # 读取三张核心表
    orders_df = fetch_bitable_to_df(TABLE_ID_ITEMS)
    sku_df    = fetch_bitable_to_df(TABLE_ID_SKU)
    stock_df  = fetch_bitable_to_df(TABLE_ID_INV)

    print("飞书数据加载完成")
    print(f"销售订单：{len(orders_df)} 条")
    print(f"SKU数据：{len(sku_df)} 条")
    print(f"库存数据：{len(stock_df)} 条")
    return orders_df, sku_df, stock_df

# 写入飞书时，字段类型必须匹配，否则会出现：
# - DatetimeFieldConvFail：日期字段写入了字符串/不支持的格式
# - TextFieldConvFail：文本字段写入了数字/时间戳
#
# 这里按“写入目标表”分别控制：不要用一个全局集合硬套所有表。
DETAIL_DATE_COLS_DEFAULT = {"预计到货日期"}   # 销售订单明细表：预计到货日期通常是日期字段
DETAIL_NUMERIC_COLS_DEFAULT = {"库存可用量", "缺口数量"}  # 明细回填：数值字段


def to_feishu_date_millis(val: Any) -> Optional[int]:
    """将日期转为飞书多维表格日期字段可用的毫秒时间戳（北京时间当日 0 点）。"""
    d = parse_date_to_date(val)
    if d is None:
        return None

    cn = timezone(timedelta(hours=8))
    dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=cn)
    return int(dt.timestamp() * 1000)


def write_df_to_bitable(
    table_id,
    df,
    fields_map=None,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
):
    """
    将 DataFrame 写入飞书多维表格（逐条写入，简单可靠）
    
    Args:
        table_id: 飞书表ID
        df: 要写入的 DataFrame
        fields_map: 字段映射字典，将代码字段名映射到飞书表中的实际字段名
                   格式: {"代码字段名": "飞书字段名"}
                   如果为None，使用原始字段名
    """
    if fields_map is None:
        fields_map = {}
    
    numeric_fields = set(numeric_cols or set())
    date_fields = set(date_cols or set())

    def feishu_value(col_name: str, val):
        if col_name in date_fields:
            ms = to_feishu_date_millis(val)
            if ms is None:
                return None
            return ms
        if isinstance(val, pd.Timestamp):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, date):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, list):
            return ", ".join([str(v) for v in val])

        # 数字字段：尽量写入 float/int，而不是字符串
        if col_name in numeric_fields:
            n = to_num(val)
            # 飞书 number 支持小数；如你字段是整数也不影响
            return n

        # 其他字段：写入字符串
        return str(val)

    for idx, row in df.iterrows():
        fields = {}
        for col in df.columns:
            if col == "_record_id" or (isinstance(col, str) and col.startswith("_")):
                continue
            val = row[col]
            
            # 跳过空值
            if pd.isna(val) or val is None or val == "":
                continue
            
            # 确定飞书表中的实际字段名
            feishu_field_name = fields_map.get(col, col)
            
            fv = feishu_value(col, val)
            if fv is None:
                continue
            fields[feishu_field_name] = fv

        try:
            record_body = AppTableRecord.builder().fields(fields).build()
            req = CreateAppTableRecordRequest.builder() \
                .app_token(BITABLE_APP_TOKEN) \
                .table_id(table_id) \
                .request_body(record_body) \
                .build()
            resp = feishu_client.bitable.v1.app_table_record.create(req)
            
            if not resp.success():
                error_msg = resp.msg if hasattr(resp, 'msg') else str(resp)
                print(f"❌ 写入记录失败 (行 {idx}): {error_msg}")
                print(f"   尝试写入的字段: {list(fields.keys())}")
                
        except Exception as e:
            print(f"❌ 写入记录异常 (行 {idx}): {str(e)}")
            print(f"   字段名: {list(fields.keys())}")


def update_bitable_records(
    table_id: str,
    df: pd.DataFrame,
    record_id_col: str = "_record_id",
    fields_map=None,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
):
    """按 record_id 更新飞书多维表格的原记录（用于“订单明细回填”）。

    说明：df 必须包含 record_id_col 列；fields_map 用于将 df 列名映射到飞书字段名。
    """
    if fields_map is None:
        fields_map = {}

    if df is None or df.empty:
        return

    if record_id_col not in df.columns:
        raise ValueError(f"df 缺少 {record_id_col} 列，无法更新回填")

    # 字段类型由调用方按目标表显式传入，避免 TextFieldConvFail / DatetimeFieldConvFail
    numeric_fields = set(numeric_cols or set())
    date_fields = set(date_cols or set())

    def feishu_value(col_name: str, val):
        if col_name in date_fields:
            ms = to_feishu_date_millis(val)
            if ms is None:
                return None
            return ms
        if isinstance(val, pd.Timestamp):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, date):
            return val.strftime("%Y-%m-%d")
        if col_name in numeric_fields:
            return to_num(val)
        return str(val)

    for idx, row in df.iterrows():
        record_id = safe_str(row.get(record_id_col, ""))
        if not record_id:
            continue

        fields = {}
        for col in df.columns:
            if col == record_id_col:
                continue

            val = row[col]
            if pd.isna(val) or val is None or val == "":
                continue

            feishu_field_name = fields_map.get(col, col)

            fv = feishu_value(col, val)
            if fv is None:
                continue
            fields[feishu_field_name] = fv

        if not fields:
            continue

        try:
            record_body = AppTableRecord.builder().fields(fields).build()
            req = UpdateAppTableRecordRequest.builder() \
                .app_token(BITABLE_APP_TOKEN) \
                .table_id(table_id) \
                .record_id(record_id) \
                .request_body(record_body) \
                .build()
            resp = feishu_client.bitable.v1.app_table_record.update(req)

            if not resp.success():
                error_msg = resp.msg if hasattr(resp, 'msg') else str(resp)
                print(f"❌ 回填更新失败 (行 {idx}): {error_msg}")
                print(f"   record_id={record_id}, 字段: {list(fields.keys())}")
        except Exception as e:
            print(f"❌ 回填更新异常 (行 {idx}): {str(e)}")
            print(f"   record_id={record_id}, 字段: {list(fields.keys())}")


def upsert_bitable_records_by_key(
    table_id: str,
    df: pd.DataFrame,
    key_col: str,
    key_field_name: str,
    *,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
) -> Optional[str]:
    """按“业务唯一键”做 upsert：已有记录则更新，没有则新增。

    用于 AI排单总表：避免每次运行都追加重复合同。

    返回：成功为 None；失败为可读错误说明（不抛异常，避免整次排单 500）。
    """
    if df is None or df.empty:
        return None

    try:
        existing = fetch_bitable_to_df(table_id)
    except Exception as e:
        err = str(e)
        if "1254004" in err or "WrongTableId" in err:
            return (
                "读取 AI排单总表失败：飞书错误 WrongTableId(1254004)，说明 TABLE_ID_DETAIL 不是有效的多维表格 table_id。"
                "常见原因：把 table_id 误写成 ttbl 开头（应为 tbl）；或在多维表格里复制了错误的 id。"
                f" 当前 TABLE_ID_DETAIL={table_id!r}。详情：{err}"
            )
        return f"读取 AI排单总表失败：{err}"
    key_to_record: Dict[str, str] = {}
    if existing is not None and not existing.empty and "_record_id" in existing.columns and key_field_name in existing.columns:
        for _, r in existing.iterrows():
            k = safe_str(r.get(key_field_name, ""))
            rid = safe_str(r.get("_record_id", ""))
            if k and rid and k not in key_to_record:
                key_to_record[k] = rid

    to_update = []
    to_create = []

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
        update_bitable_records(table_id, pd.DataFrame(to_update), record_id_col="_record_id", numeric_cols=numeric_cols, date_cols=date_cols)

    if to_create:
        write_df_to_bitable(table_id, pd.DataFrame(to_create), fields_map=None, numeric_cols=numeric_cols, date_cols=date_cols)

    return None

# =========================
# 数据缓存（避免每次请求重复读取飞书，减轻API压力）
# =========================
data_cache = {
    "items": None,
    "sku": None,
    "inv": None
}

def load_data():
    """从飞书一次性加载数据到内存，并验证关键字段"""
    print("=" * 60)
    print("正在从飞书加载数据...")
    print("=" * 60)
    
    try:
        # 加载三张核心表
        items_df = fetch_bitable_to_df(TABLE_ID_ITEMS)
        sku_df   = fetch_bitable_to_df(TABLE_ID_SKU)
        inv_df   = fetch_bitable_to_df(TABLE_ID_INV)
        
        data_cache["items"] = items_df
        data_cache["sku"]   = sku_df
        data_cache["inv"]   = inv_df
        
        # ===== 验证表字段完整性 =====
        print("\n【表字段验证】")
        
        # 字段校验分两类：
        # - 必填：影响核心计算的字段，缺失就无法运行
        # - 可选：用于回填展示/辅助信息，缺失不应阻断运行（否则会出现你看到的“误报缺字段”）
        required_items_cols = ["合同编号", "合同数量", "SKU编码"]
        optional_items_cols = ["产品名称", "规格", "库存可用量", "缺口数量", "库存状态", "预计到货日期", "是否RB800", "排单批次号",
                               "是否紧急订单", "是否换货订单", "是否补发订单", "是否维修订单"]

        required_sku_cols = ["产品编码SKU", "标准生产周期"]
        # “是否新产品/自研/外采”在你表里可能是合并列，也可能拆成三列；都当可选
        optional_sku_cols = ["设备名称", "设备型号", "是否新产品是否自研是否外采", "是否新产品", "是否自研", "是否外采"]

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
                print(f"⚠️  {table_name} 缺少【可选】字段：{missing_optional}（不影响运行）")

            print(f"✅ {table_name} 必填字段齐全")
            return True
        
        all_ok = True
        all_ok &= check_columns(items_df, "销售订单明细表", required_items_cols, optional_items_cols)
        all_ok &= check_columns(sku_df, "SKU标准表", required_sku_cols, optional_sku_cols)
        all_ok &= check_columns(inv_df, "库存快照表", required_inv_cols, optional_inv_cols)
        
        # 必要字段里已包含 `库存日期`，此处不再做可选提示
        
        print("\n【数据统计】")
        print(f"  销售订单明细数：{len(items_df)} 条")
        print(f"  SKU标准数：{len(sku_df)} 条")
        print(f"  库存快照数：{len(inv_df)} 条")
        
        if all_ok:
            print("\n✅ 所有必填字段检验完毕，数据可以使用！")
        else:
            print("\n❌ 存在必填字段缺失，无法继续运行。请检查飞书表结构/列名！")
        print("=" * 60)
    except Exception as e:
        print(f"\n❌ 飞书数据加载失败: {str(e)}")
        print("详细信息：", str(e))
        raise  # 重新抛出异常供调用者处理

# =========================
# 计算可用库存（带库存日期检查）
# =========================
def get_available_stock(inv_row):
    """兼容旧逻辑：保留函数名，但可用库存计算改由 `calc_available_stock` 统一处理。

    这里返回库存快照表的基础“库存数量”，不含锁单扣减。
    """
    return to_num(inv_row.get("库存数量", 0))

# =========================
# 辅助函数：安全获取字符串字段
# =========================
def safe_str(value):
    """安全转换为字符串并清洗"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()

# =========================
# 规则：RB800 判断（按型号包含 RB800）
# =========================
def is_rb800_from_text(text: str) -> bool:
    t = safe_str(text).upper().replace(" ", "")
    return "RB800" in t


def pick_model_text(order_row: Dict[str, Any]) -> str:
    """从订单明细行中尽量取到“型号/规格”文本用于 RB800 判断。

    你当前明细表有 `规格`/`产品名称`/`SKU编码`。
    """
    for k in ("规格", "产品名称", "SKU编码"):
        v = safe_str(order_row.get(k, ""))
        if v:
            return v
    return ""


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


def normalize_shortage_sku_list(value: Any) -> List[str]:
    """把缺货 SKU 列表规整为去重后的 SKU 数组，供数量和写入共用。"""
    if isinstance(value, list):
        raw_items = value
    else:
        s = safe_str(value)
        if not s or s.lower() in {"nan", "none", "null"}:
            return []
        raw_items = re.split(r"[,，]", s)

    seen = set()
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


def has_valid_manual_confirm_date(row: Dict[str, Any]) -> bool:
    return parse_date_to_date(row.get("人工确认发货时间")) is not None


def is_effective_manual_confirmed(row: Dict[str, Any]) -> bool:
    """只有明确人工确认且有有效人工确认发货时间，才按锁单处理。"""
    return safe_str(row.get("是否人工确认", "")) == "是" and has_valid_manual_confirm_date(row)


def calc_available_stock(inv_row: Optional[Dict[str, Any]], reserved_qty: float = 0.0) -> float:
    """可用库存 = 库存快照表的库存数量 - 已占用排单库存(仅锁单保留)。

    简单模式：库存快照每次全量覆盖；不做非锁单的占用追踪。
    """
    if not inv_row:
        base = 0.0
    else:
        # 按你表头：库存快照表字段 `库存数量`（默认只用库存数量）
        base = to_num(inv_row.get("库存数量", 0))
    return max(base - to_num(reserved_qty), 0.0)


def _norm_match_key(s: str) -> str:
    """用于 SKU 等关键字段的宽松比较：去空白、全角括号统一。"""
    t = safe_str(s).replace("\u3000", " ").replace(" ", "")
    t = t.replace("（", "(").replace("）", ")")
    return t.lower()


def _is_generic_spec_text(t: str) -> bool:
    """规格里常见「定制/外购」等无具体型号信息，无法与库存做有效匹配。"""
    s = safe_str(t).strip().replace(" ", "").replace("\u3000", "")
    if len(s) < 6:
        return True
    generic_tokens = ("定制/外购", "定制外购", "定制", "外购", "外协")
    if s in generic_tokens:
        return True
    if all(ch in "定制外购外协/\\.-" for ch in s) and len(s) <= 12:
        return True
    return False


def _inventory_best_contains_hit(inv: pd.DataFrame, text: str) -> Tuple[Optional[Dict[str, Any]], str]:
    """在库存表中用包含匹配找一行（取库存数量最大）。"""
    if len(text) < 2:
        return None, ""
    try:
        pat = re.escape(text)
    except re.error:
        return None, ""
    for col in ("国网设备型号", "国网设备名称", "SKU"):
        if col not in inv.columns:
            continue
        try:
            mask = inv[col].astype(str).str.contains(pat, case=False, na=False, regex=True)
        except Exception:
            continue
        hit = inv[mask]
        if hit.empty:
            continue
        hit = hit.copy()
        if "库存数量" in hit.columns:
            hit["_sort_q"] = hit["库存数量"].apply(to_num)
            hit = hit.sort_values("_sort_q", ascending=False)
        return hit.iloc[0].to_dict(), f"{col}包含匹配({text[:40]})"
    return None, ""


def find_inventory_row(
    sku_code: str,
    row_dict: Dict[str, Any],
    inv: pd.DataFrame,
    sku_df: Optional[pd.DataFrame] = None,
) -> Tuple[Optional[Dict[str, Any]], str]:
    """在库存快照中查找对应行。

    1) 优先按 `SKU` 列与明细 `SKU编码` 精确/规范化匹配（避免「表里有货但明细对不上 key」）。
    2) 通过 SKU标准表：用 `产品编码SKU` 找到 `设备型号`/`设备名称`，再去库存表做包含匹配（解决明细规格写成「定制/外购」等情况）。
    3) 再按明细 `规格`/`产品名称`（过滤无效规格）与库存表做包含匹配。
    多条命中时取 **库存数量最大** 的一条。
    """
    if inv is None or inv.empty:
        return None, "库存表为空"

    code = safe_str(sku_code)

    if "SKU" in inv.columns:
        sku_col = inv["SKU"].astype(str).str.strip()
        sub = inv[sku_col == code]
        if sub.empty and code:
            sub = inv[sku_col.apply(_norm_match_key) == _norm_match_key(code)]
        if not sub.empty:
            return sub.iloc[0].to_dict(), "SKU列匹配"

    # 经 SKU 主数据拿到型号/名称再匹配库存
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
                hit, how = _inventory_best_contains_hit(inv, t)
                if hit is not None:
                    return hit, f"SKU标准表.{key}→{how}"

    texts: List[str] = []
    for k in ("规格", "产品名称"):
        t = safe_str(row_dict.get(k, ""))
        if len(t) >= 2 and not _is_generic_spec_text(t):
            texts.append(t)
    if code and all(t != code for t in texts):
        texts.insert(0, code)

    for text in texts:
        hit, how = _inventory_best_contains_hit(inv, text)
        if hit is not None:
            return hit, how

    return None, f"未匹配(SKU编码={code!r}, 规格={safe_str(row_dict.get('规格', ''))[:40]!r})"


def calc_stock_status_and_gap(demand: float, available: float) -> Tuple[str, float]:
    gap = max(to_num(demand) - to_num(available), 0.0)
    return ("有货" if gap <= 0 else "缺货"), (0.0 if gap <= 0 else gap)


def calc_shortage_eta_date(
    sku_code: str,
    sku_df: pd.DataFrame,
    base_date: date,
) -> date:
    """仅缺货时使用：新产品(查不到SKU)→25天；有标准周期→按 `标准生产周期`(天)。

    注意：按你的要求，“预计到货日期”以库存快照表的 `库存日期` 作为基准日。
    """
    match = sku_df[sku_df["产品编码SKU"] == sku_code] if (sku_df is not None and not sku_df.empty) else pd.DataFrame()
    if match.empty:
        cycle = 25
        print("新产品，返回基准日+25天")
    else:
        sku_info = match.iloc[0].to_dict()
        cycle = to_num(sku_info.get("标准生产周期", 0))
        if cycle > 0:
            print(f"有标准周期，返回基准日+{int(cycle)}天")
        else:
            cycle = 25
            print("无标准周期，按新产品处理，返回基准日+25天")
    
    final_date = add_calendar_days_then_working_day(base_date, int(cycle))
    print(f'物料 {sku_code}, 基准日 {base_date}, 标准周期 {cycle}天, 计算结果 {final_date}')
    return final_date


def calc_daily_capacity(sku_kinds: int, total_qty: float, base: int = 5) -> int:
    """简单产能函数：标准日排单量 ≤ 5 单。

    若当日订单的 SKU种类 < 10 或 总数量 < 15，则可增加 1-2 单处理能力。
    这里用一个简单可解释的阶梯规则：
    - 特别轻量（sku_kinds < 5 或 total_qty < 8）→ +2
    - 轻量（sku_kinds < 10 或 total_qty < 15）→ +1
    - 否则 +0
    """
    sku_kinds = int(to_num(sku_kinds))
    total_qty = to_num(total_qty)

    bonus = 0
    if sku_kinds < 5 or total_qty < 8:
        bonus = 2
    elif sku_kinds < 10 or total_qty < 15:
        bonus = 1

    return int(base + bonus)


def apply_capacity_scheduling(summary: pd.DataFrame, today: date) -> Tuple[pd.DataFrame, int]:
    """对 AI排单总表应用每日产能限制。

    - 锁单（是否人工确认=是）：完全不动
    - 特殊订单（项目类型=特殊订单）：不计入每日5单限制，也不动
    - 其他订单：若某日超产能，将 `AI建议发货时间` 顺延到后续日期

    返回：(调整后的 summary, 被顺延的订单数量)
    """
    if summary is None or summary.empty or "AI建议发货时间" not in summary.columns:
        return summary, 0

    df = summary.copy()
    df["__ship_date"] = df["AI建议发货时间"].apply(lambda x: next_working_day(parse_date_to_date(x) or today))

    df["__locked"] = df.apply(lambda r: is_effective_manual_confirmed(r.to_dict()), axis=1)
    df["__special"] = df.get("项目类型", "").astype(str) == "特殊订单"

    # 只对“非锁定 + 非特殊”应用产能
    normal = df[~df["__locked"] & ~df["__special"]].copy()
    others = df[df["__locked"] | df["__special"]].copy()

    normal = normal.sort_values(by=["__ship_date", "合同编号"], kind="stable")

    day_count: Dict[date, int] = {}
    day_sku_kinds: Dict[date, int] = {}
    day_total_qty: Dict[date, float] = {}

    delayed = 0
    assigned_dates = {}

    for _, row in normal.iterrows():
        base_date = next_working_day(row["__ship_date"])
        order_sku_kinds = int(to_num(row.get("订单SKU总数", 0)))
        order_qty = to_num(row.get("订单总数量", 0))

        d = base_date
        while True:
            c = day_count.get(d, 0)
            sk = day_sku_kinds.get(d, 0)
            tq = day_total_qty.get(d, 0.0)

            cap = calc_daily_capacity(sk, tq, base=5)
            if c < cap:
                # 放入当天
                day_count[d] = c + 1
                day_sku_kinds[d] = sk + order_sku_kinds   # 简单模式：按订单SKU数累加
                day_total_qty[d] = tq + order_qty
                assigned_dates[safe_str(row.get("合同编号", ""))] = d
                if d != base_date:
                    delayed += 1
                break

            d = next_working_day(d + timedelta(days=1))

    if assigned_dates:
        normal["__ship_date_new"] = normal["合同编号"].astype(str).map(lambda x: assigned_dates.get(safe_str(x), next_working_day(today)))
        normal["AI建议发货时间"] = normal["__ship_date_new"].apply(date_to_yyyy_mm_dd)
        normal = normal.drop(columns=["__ship_date_new"])

    out = pd.concat([normal, others], ignore_index=True).drop(columns=["__ship_date", "__locked", "__special"], errors="ignore")
    if "AI建议发货时间" in out.columns:
        out["AI建议发货时间"] = out["AI建议发货时间"].apply(lambda x: date_to_yyyy_mm_dd(next_working_day(parse_date_to_date(x) or today)))
    # 输出表更友好：按新发货时间排序
    if "AI建议发货时间" in out.columns:
        out = out.sort_values(by=["AI建议发货时间", "合同编号"], kind="stable")

    return out, delayed

# =========================
# 计算发货周期（核心规则）
# =========================
def sku_supply_type_from_sku_row(sku_row: Dict[str, Any]) -> str:
    """从SKU标准表字段 `是否新产品是否自研是否外采` 粗略解析供给类型。

    返回：'外采' / '自研' / '新产品' / ''
    """
    # 兼容两种表结构：
    # 1) 合并列：是否新产品是否自研是否外采
    # 2) 拆分列：是否新产品 / 是否自研 / 是否外采
    raw = safe_str(sku_row.get("是否新产品是否自研是否外采", ""))
    if "外采" in raw:
        return "外采"
    if "自研" in raw:
        return "自研"
    if "新" in raw:
        return "新产品"

    if safe_str(sku_row.get("是否外采", "")) == "是":
        return "外采"
    if safe_str(sku_row.get("是否自研", "")) == "是":
        return "自研"
    if safe_str(sku_row.get("是否新产品", "")) == "是":
        return "新产品"
    return ""


def next_working_day(start_date: date) -> date:
    """从 start_date 开始，找到下一个工作日（跳过周末）。"""
    start_date = parse_date_to_date(start_date) or datetime.today().date()
    if start_date.weekday() < 5:  # 周一到周五
        return start_date
    # 周六或周日，跳到下周一
    days_to_add = 7 - start_date.weekday()
    return start_date + timedelta(days=days_to_add)


def add_calendar_days_then_working_day(start_date: date, days: int) -> date:
    """先加自然日，再把结果顺延到工作日。"""
    start_date = parse_date_to_date(start_date) or datetime.today().date()
    return next_working_day(start_date + timedelta(days=int(days)))


def calc_order_ship_date_for_group(
    group: pd.DataFrame,
    sku_df: pd.DataFrame,
    today: date,
) -> Tuple[Optional[date], Optional[date]]:
    """计算订单级别的预计发货时间：返回 (latest_eta, ship_date)"""
    shortage_mask = group["库存状态"] == "缺货"
    shortage_rows = group[shortage_mask]
    
    if shortage_rows.empty:
        # 无缺货，直接排到最近工作日
        return None, next_working_day(today)
    
    valid_dates = [d for d in shortage_rows['预计到货日期'].apply(parse_date_to_date).tolist() if d is not None]
    if valid_dates:
        latest_eta = max(valid_dates)
        ship_date = next_working_day(latest_eta)
        return latest_eta, ship_date

    # 缺货但暂时没有有效 ETA 时也必须进入 AI排单总表，避免新增订单被漏排。
    fallback_eta = add_calendar_days_then_working_day(today, 25)
    return fallback_eta, fallback_eta

# =========================
# 主排单API
# =========================


@app.post("/schedule")
def run_scheduler(payload: ScheduleRequest = ScheduleRequest()):
    """AI排单主函数"""

    # =========================
    # 从飞书直接加载最新数据
    # =========================
    print("🔄 从飞书加载最新数据")
    load_data()

    items = data_cache["items"]
    sku   = data_cache["sku"]
    inv   = data_cache["inv"]
    # ===== 生成排单ID & 排单批次号 =====
    schedule_id = datetime.now().strftime("SCH%Y%m%d%H%M%S")
    batch_id = datetime.now().strftime("BATCH%Y%m%d")
    
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
    # 读取历史排单（用于人工锁单保护 + 锁单库存占用）
    # =========================
    try:
        old_summary = pd.read_csv(f"{CACHE_DIR}/ai_schedule_summary.csv")
        print("已加载历史排单缓存")
    except:
        old_summary = pd.DataFrame()

    # 兼容旧缓存列名：历史版本可能用 AI订单ID 表示合同编号
    if not old_summary.empty:
        if "合同编号" not in old_summary.columns and "AI订单ID" in old_summary.columns:
            old_summary = old_summary.rename(columns={"AI订单ID": "合同编号"})
        if "合同编号" in old_summary.columns:
            old_summary["合同编号"] = old_summary["合同编号"].astype(str).apply(safe_str)
        if "是否人工确认" not in old_summary.columns:
            old_summary["是否人工确认"] = "否"
        if "人工确认发货时间" not in old_summary.columns:
            old_summary["人工确认发货时间"] = ""
        old_summary["是否人工确认"] = old_summary.apply(
            lambda r: "是" if is_effective_manual_confirmed(r.to_dict()) else "否",
            axis=1,
        )
        if "AI建议发货时间" in old_summary.columns:
            old_summary["AI建议发货时间"] = old_summary["AI建议发货时间"].apply(date_to_yyyy_mm_dd)
        old_summary["人工确认发货时间"] = old_summary["人工确认发货时间"].apply(date_to_yyyy_mm_dd)
    # =========================
    # 计算已锁单库存占用（关键AI逻辑）
    # =========================
    locked_stock: Dict[str, float] = {}

    if not old_summary.empty:
        locked_orders = old_summary[old_summary.apply(lambda r: is_effective_manual_confirmed(r.to_dict()), axis=1)]

        if not locked_orders.empty:
            try:
                old_detail = pd.read_csv(f"{CACHE_DIR}/ai_schedule_detail.csv")
            except:
                old_detail = pd.DataFrame()

            if not old_detail.empty:
                # old_detail 使用“销售订单明细表”的合同编号
                locked_detail = old_detail[old_detail["合同编号"].isin(locked_orders["合同编号"])]

                for _, r in locked_detail.iterrows():
                    sku_code = safe_str(r.get("SKU编码", ""))
                    qty = to_num(r.get("合同数量", 0))

                    locked_stock[sku_code] = locked_stock.get(sku_code, 0) + qty
        print("🔒 已冻结库存：", locked_stock)

    result_rows = []

    # =========================
    # 第一步：回填“销售订单明细表”字段（库存可用量/缺口/库存状态/预计到货/是否RB800/排单批次号）
    # =========================
    today = datetime.today().date()

    # 库存快照表的“全表最新库存日期”（当某SKU未匹配到库存行时，用它做 ETA 基准日）
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

        inv_info, _inv_how = find_inventory_row(sku_code, row_dict, inv, sku_df=sku)
        if inv_info is None:
            print(f"⚠️ 库存未匹配 合同={contract_id} SKU={sku_code} 规格={safe_str(row_dict.get('规格', ''))[:50]} → {_inv_how}")

        # 以库存快照最新日期或今天作为 ETA 基准日

        reserved = locked_stock.get(sku_code, 0.0)
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
    print("正在回填销售订单明细表...")
    update_bitable_records(
        TABLE_ID_ITEMS,
        df_backfill,
        record_id_col="_record_id",
        numeric_cols=DETAIL_NUMERIC_COLS_DEFAULT,
        date_cols=DETAIL_DATE_COLS_DEFAULT,
    )

    # =========================
    # 第二步开始前：重新读取“已回填”的销售订单明细表
    # =========================
    print("正在重新读取回填后的销售订单明细表...")
    df_detail_reread = fetch_bitable_to_df(TABLE_ID_ITEMS)
    df_detail = build_current_detail_for_summary(items, df_detail_reread, df_backfill)
    # 统一关键列类型
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

    # 本地缓存明细（用于下次锁单占用）
    try:
        df_detail.to_csv(f"{CACHE_DIR}/ai_schedule_detail.csv", index=False)
    except Exception as e:
        print("⚠️ 明细缓存保存失败：", str(e))

    # ===== 第二步：生成“AI排单总表”（按合同汇总，一单一行） =====
    print('明细表列名:', df_detail.columns.tolist())
    print(df_detail[['合同编号', 'SKU编码', '库存状态', '预计到货日期']].head(3))
    summary_rows = []
    if not df_detail.empty:
        def safe_val(x):
            if isinstance(x, float) and pd.isna(x):
                return None
            if isinstance(x, pd.Timestamp) and pd.isna(x):
                return None
            return x

        for contract_id, group in df_detail.groupby("合同编号"):
            # =========================
            # 人工锁单检查
            # =========================
            if not old_summary.empty:
                old_row = old_summary[old_summary["合同编号"] == contract_id]

                if not old_row.empty:
                    if is_effective_manual_confirmed(old_row.iloc[0].to_dict()):
                        print(f"🔒 合同 {contract_id} 已人工锁单，跳过AI重算")
                        summary_rows.append(old_row.iloc[0].to_dict())
                        continue

            latest_eta, ship_date = calc_order_ship_date_for_group(group=group, sku_df=sku, today=today)
            if ship_date is None:
                ship_date = next_working_day(today)
                print(f'⚠️ 订单 {contract_id} 暂无有效发货日期，已兜底排到最近工作日 {ship_date}')
            print(f'订单 {contract_id}: latest_eta={latest_eta}, ship_date={ship_date}')

            # 订单类型（用于输出字段 `项目类型`）
            # 与规则1一致：远控 / 特殊 / 常规
            any_rb800 = False
            for _, r in group.iterrows():
                if is_rb800_from_text(pick_model_text(r.to_dict())):
                    any_rb800 = True
                    break
            is_urgent = (group.get("是否紧急订单", pd.Series(dtype=str)).astype(str) == "是").any()
            is_exchange = (group.get("是否换货订单", pd.Series(dtype=str)).astype(str) == "是").any()
            is_resend = (group.get("是否补发订单", pd.Series(dtype=str)).astype(str) == "是").any()
            is_repair = (group.get("是否维修订单", pd.Series(dtype=str)).astype(str) == "是").any()

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
            shortage_skus_unique = sorted(normalize_shortage_sku_list(shortage_rows['SKU编码'].dropna().unique().tolist()))
            shortage_sku_count = len(shortage_skus_unique)
            overall_status = "缺货" if shortage_sku_count else "有货"

            summary_rows.append({
                "合同编号": contract_id,
                "项目类型": project_type,
                "订单SKU总数": sku_count if sku_count is not None else 0,
                "订单总数量": total_qty if total_qty is not None else 0,
                "缺货SKU数": shortage_sku_count,
                "缺货SKU列表": ', '.join(shortage_skus_unique),
                "整体状态": overall_status,
                "AI建议发货时间": date_to_yyyy_mm_dd(next_working_day(ship_date)) if ship_date else "",
                "排单批次号": batch_id,
                "是否人工确认": "否",
                "人工确认发货时间": ""
            })
    summary = pd.DataFrame(summary_rows)
    # 统一合同编号，避免同一合同因空格/格式差异被重复写入
    if not summary.empty:
        summary["合同编号"] = summary["合同编号"].astype(str).apply(safe_str)
        summary = summary.drop_duplicates(subset=["合同编号"], keep="last")
    print("summary 列名：", list(summary.columns))

    # ===== 第三步：应用每日产能限制（仅非锁单/非特殊订单）=====
    summary, delayed_cnt = apply_capacity_scheduling(summary, today=today)
    if delayed_cnt:
        print(f"📦 产能限制生效：{delayed_cnt} 单被顺延到后续日期")

    # ===== 强制类型转换并打印 =====
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
            elif key == "是否人工确认":
                converted[key] = "是" if safe_str(value) == "是" else "否"
            elif key == "人工确认发货时间":
                converted[key] = date_to_yyyy_mm_dd(value)
            else:
                converted[key] = value
        converted["缺货SKU列表"] = ", ".join(normalize_shortage_sku_list(converted.get("缺货SKU列表", "")))
        converted["缺货SKU数"] = shortage_sku_count_from_list(converted.get("缺货SKU列表", ""))
        converted["是否人工确认"] = "是" if is_effective_manual_confirmed(converted) else "否"
        return converted

    converted_summary_rows = []
    for _, row in summary.iterrows():
        converted_row = convert_summary_row(row.to_dict())
        if converted_row.get("合同编号") == "GW20260410-3":
            print(f"GW20260410-3 转换后数据: {converted_row}")
        converted_summary_rows.append(converted_row)

    summary = pd.DataFrame(converted_summary_rows)

    # ===== 写入飞书结果表 =====
    # 先清空结果表（可选，避免数据重复）—— 此处简单起见，直接追加写入
    # 如果想要覆盖，需要先删除原记录，这里暂不处理，你可手动清空结果表
    
    print("\n【写入结果到飞书】")

    # 明细回填已通过 update_bitable_records 写回销售订单明细表；这里不再新增“明细输出表”。
    # =========================
    # 保存本地缓存（供明天锁单用）
    # =========================
    summary.to_csv(f"{CACHE_DIR}/ai_schedule_summary.csv", index=False)
    print("排单缓存已保存")
    
    summary = summary.where(pd.notnull(summary), None)
    df_detail = df_detail.where(pd.notnull(df_detail), None)

    # 汇总数据 → 写入 AI排单总表
    print("正在写入 AI排单总表...")
    upsert_err = upsert_bitable_records_by_key(
        TABLE_ID_DETAIL,
        summary,
        key_col="合同编号",
        key_field_name="合同编号",
        numeric_cols=SUMMARY_NUMERIC_COLS_DEFAULT,
        date_cols=SUMMARY_DATE_COLS_DEFAULT,
    )

    if upsert_err:
        print(f"❌ {upsert_err}")
        return {
            "msg": "明细已回填，但 AI排单总表写入失败",
            "error": upsert_err,
            "AI排单总表": summary.to_dict(orient="records") if not summary.empty else [],
            "回填明细行数": 0 if df_backfill is None else int(len(df_backfill)),
            "顺延订单数": int(delayed_cnt),
            "TABLE_ID_DETAIL": TABLE_ID_DETAIL,
            "notes": "请核对多维表格中「AI排单总表」的真实 table_id（应以 tbl 开头），或通过环境变量 TABLE_ID_DETAIL 覆盖。",
        }

    return {
        "msg": "AI排单完成",
        "AI排单总表": summary.to_dict(orient="records") if not summary.empty else [],
        "回填明细行数": 0 if df_backfill is None else int(len(df_backfill)),
        "顺延订单数": int(delayed_cnt),
        "notes": "回填已更新销售订单明细表原记录；总表按合同编号更新/新增；特殊订单不计入5单限制；如写入失败请检查飞书字段名是否与代码一致"
    }

# =========================
# 飞书 Webhook 触发接口
# =========================
@app.post("/webhook/feishu")
async def feishu_webhook(request: Request):
    body = await request.json()

    # ① 飞书URL验证（第一次配置时用）
    if "challenge" in body:
        print("飞书URL验证成功")
        return {"challenge": body["challenge"]}

    print("========== 收到飞书事件 ==========")

    # ② 重新加载飞书数据
    try:
        print("开始重新加载飞书数据...")
        load_data()
        print("飞书数据加载成功，当前数据量：", len(data_cache["items"]))
    except Exception as e:
        print("飞书数据加载失败：", str(e))
        return {"error": f"飞书数据加载失败: {str(e)}"}

if __name__ == "__main__":
    import uvicorn

    print("启动AI排单服务...")
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
