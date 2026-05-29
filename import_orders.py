#!/usr/bin/env python3
"""
OA 订单数据 & 库存快照 导入工具

读取 Excel → 校验 → 自动写入飞书多维表格。

用法:
    python import_orders.py 文件.xlsx                # 导入（含订单主表+明细表+库存快照）
    python import_orders.py 文件.xlsx --dry-run       # 仅校验不写入
    python import_orders.py 文件.xlsx --auto-fix      # 自动修复已知问题

Excel 结构要求:
    Sheet1（订单表）: 订单编号, 客户名称, 项目名称, 商务, 下单时间, 国网设备名称, 设备型号, 数量
    Sheet2（库存表）: 国网设备名称, 型号, 库存数量   （可选，存在则清除后重新导入）
"""

import os
import sys
import argparse
import re
from datetime import datetime, date, timezone, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


# =========================
# 加载 .env
# =========================
def _load_dotenv():
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

# 飞书配置
FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
BITABLE_APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_ID_ITEMS = os.getenv("TABLE_ID_ITEMS", "")       # 销售订单明细表
TABLE_ID_MAIN = os.getenv("TABLE_ID_MAIN", "")          # 销售订单主表
TABLE_ID_INV = os.getenv("TABLE_ID_INV", "")            # 库存快照表

FEISHU_BITABLE_BASE = "https://open.feishu.cn/open-apis/bitable/v1/apps"

# =========================
# 列名映射
# =========================

# OA 导出列名 → 标准列名（仅用于订单表 Sheet）
OA_COLUMN_MAP = {
    "订单编号": "合同编号",
    "数量": "合同数量",
    "设备型号": "规格",
    "设备名称": "产品名称",
    "国网设备名称": "产品名称",
    "下单日期": "下单时间",
    "订单时间": "下单时间",
    "下单时间": "下单时间",      # 保留原列名（部分 Excel 已使用此列名）
    "申请时间": "下单时间",      # 部分 Excel 用"申请时间"作为下单时间
    "申请人": "商务",           # 部分 Excel 用"申请人"作为商务
}

# 订单主表：标准列名 → 飞书字段名（精确匹配）
MAIN_FIELD_MAP = {
    "合同编号": "合同编号",
    "下单时间": "下单日期",      # MAIN 表的日期字段叫"下单日期"
    "客户名称": "客户名称",
    "项目名称": "项目名称",
    "商务": "商务",
}

# 订单明细表：标准列名 → 飞书字段名（精确匹配）
ITEMS_FIELD_MAP = {
    "合同编号": "合同编号",
    "下单时间": "下单时间",
    "产品名称": "产品名称",
    "规格": "规格",
    "合同数量": "合同数量",
}

# 库存快照表：Excel 原始列名 → 飞书字段名（精确匹配）
INV_COLUMN_MAP = {
    "国网设备名称": "国网设备名称",
    "型号": "国网设备型号",
    "库存数量": "库存数量",
}

# 主表去重列
MAIN_COLUMNS = ["合同编号", "下单时间", "客户名称", "项目名称", "商务"]
# 明细表列
ITEMS_COLUMNS = ["合同编号", "下单时间", "产品名称", "规格", "合同数量"]
# 库存表列
INV_COLUMNS = ["国网设备名称", "国网设备型号", "库存数量"]


# =========================
# 工具函数
# =========================
def safe_str(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float):
        try:
            if pd.isna(val):
                return ""
        except Exception:
            pass
    if isinstance(val, float) and (val != val):
        return ""
    return str(val).strip()


def to_num(val: Any) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def to_date_millis(val: Any) -> Optional[int]:
    """将日期值转为飞书 DateTime 字段可用的毫秒时间戳（北京时间 0 点）。

    支持：datetime、date、字符串、Excel 数字日期（如 46156 = 2026-05-29）。
    """
    if val is None:
        return None
    if isinstance(val, pd.Timestamp):
        d = val.date()
    elif isinstance(val, datetime):
        d = val.date()
    elif isinstance(val, date):
        d = val
    elif isinstance(val, (int, float)):
        if pd.isna(val) or val <= 0:
            return None
        # Excel 数字日期（Windows 1900 日期系统，epoch = 1899-12-30）
        excel_epoch = date(1899, 12, 30)
        d = excel_epoch + timedelta(days=int(val))
    elif isinstance(val, str):
        val = val.strip()
        if not val:
            return None
        # 先尝试作为数字字符串解析
        try:
            num = float(val)
            if num > 0:
                excel_epoch = date(1899, 12, 30)
                d = excel_epoch + timedelta(days=int(num))
                cn = timezone(timedelta(hours=8))
                dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=cn)
                return int(dt.timestamp() * 1000)
        except ValueError:
            pass
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"]:
            try:
                d = datetime.strptime(val, fmt).date()
                break
            except ValueError:
                continue
        else:
            return None
    else:
        return None
    cn = timezone(timedelta(hours=8))
    dt = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=cn)
    return int(dt.timestamp() * 1000)


# =========================
# 飞书 API
# =========================
_token_cache: Dict[str, Any] = {}


def get_access_token() -> str:
    import httpx
    now = datetime.now().timestamp()
    if _token_cache.get("expires_at", 0) > now + 300:
        return _token_cache["token"]

    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
        timeout=15.0,
    )
    data = resp.json()
    _token_cache["token"] = data["tenant_access_token"]
    _token_cache["expires_at"] = now + data.get("expire", 7200)
    return _token_cache["token"]


def _feishu_headers():
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json; charset=utf-8",
    }


def _safe_http_json(resp, label: str = "") -> dict:
    try:
        return resp.json()
    except Exception:
        text = resp.text[:500] if hasattr(resp, "text") else str(resp)[:500]
        raise Exception(f"{label} 响应非 JSON: {text}")


def delete_all_records(table_id: str) -> int:
    """删除指定飞书多维表格的全部记录，返回删除条数。"""
    import httpx
    headers = _feishu_headers()

    # 先查出所有 record_id
    all_ids = []
    page_token = None
    while True:
        url = f"{FEISHU_BITABLE_BASE}/{BITABLE_APP_TOKEN}/tables/{table_id}/records/search"
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

    # 分批删除（每批最多 500 条）
    delete_url = f"{FEISHU_BITABLE_BASE}/{BITABLE_APP_TOKEN}/tables/{table_id}/records/batch_delete"
    deleted = 0
    for i in range(0, len(all_ids), 500):
        batch = all_ids[i:i + 500]
        resp = httpx.post(delete_url, headers=headers, json={"records": batch}, timeout=60.0)
        data = _safe_http_json(resp, "删除记录")
        if data.get("code") != 0:
            raise Exception(f"删除库存记录失败: code={data.get('code')}, msg={data.get('msg')}")
        deleted += len(batch)
    return deleted


def write_df_to_bitable(
    table_id: str,
    df: pd.DataFrame,
    field_map: Optional[Dict[str, str]] = None,
    numeric_cols: Optional[set] = None,
    date_cols: Optional[set] = None,
):
    """批量写入飞书多维表格（新增记录）。

    df 的列名是"标准列名"，field_map 将其映射到飞书实际字段名。
    只有 field_map 中列出的列才会被写入。
    """
    import httpx
    field_map = field_map or {}
    numeric_cols = numeric_cols or set()
    date_cols = date_cols or set()

    headers = _feishu_headers()
    url = f"{FEISHU_BITABLE_BASE}/{BITABLE_APP_TOKEN}/tables/{table_id}/records/batch_create"
    records_batch: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        fields: Dict[str, Any] = {}
        for std_col, feishu_col in field_map.items():
            if std_col not in df.columns:
                continue
            val = row[std_col]
            if pd.isna(val) or val is None or val == "":
                continue
            if std_col in date_cols:
                ms = to_date_millis(val)
                if ms is None:
                    continue
                val = ms
            elif std_col in numeric_cols:
                val = to_num(val)
            else:
                val = str(val).strip()
                if not val:
                    continue
            fields[feishu_col] = val

        if not fields:
            continue
        records_batch.append({"fields": fields})

        if len(records_batch) >= 500:
            resp = httpx.post(url, headers=headers, json={"records": records_batch}, timeout=60.0)
            data = _safe_http_json(resp, "写入飞书")
            if data.get("code") != 0:
                raise Exception(f"写入飞书失败: code={data.get('code')}, msg={data.get('msg')}")
            records_batch.clear()

    if records_batch:
        resp = httpx.post(url, headers=headers, json={"records": records_batch}, timeout=60.0)
        data = _safe_http_json(resp, "写入飞书")
        if data.get("code") != 0:
            raise Exception(f"写入飞书失败: code={data.get('code')}, msg={data.get('msg')}")


# =========================
# Sheet 类型检测
# =========================
def detect_sheet_type(df: pd.DataFrame) -> str:
    """检测 Sheet 是订单表还是库存表。"""
    cols = set(df.columns)
    # 订单表特征：有订单编号/合同编号，且有数量相关列
    has_order_id = "订单编号" in cols or "合同编号" in cols
    has_quantity = bool({"数量", "合同数量"} & cols)
    has_device = bool({"国网设备名称", "设备名称", "产品名称", "设备型号", "规格"} & cols)
    if has_order_id and (has_quantity or has_device):
        return "orders"
    # 库存表特征：有库存数量
    if "库存数量" in cols:
        return "inventory"
    # 兜底：按 sheet 名称判断
    return "unknown"


# =========================
# 清洗订单表
# =========================
def clean_orders_df(df: pd.DataFrame) -> pd.DataFrame:
    """清洗订单表：应用 OA 列名映射，删除无名列，清洗字符串。"""
    df = df.copy()

    # 删除无名列
    df = df.loc[:, ~df.columns.astype(str).str.match(r'^Unnamed')]

    # 应用 OA 列名映射
    df = df.rename(columns=lambda c: OA_COLUMN_MAP.get(c, c))
    # 去重列名（OA_COLUMN_MAP 可能产生重复，如 设备名称 和 国网设备名称 都映射到 产品名称）
    df = df.loc[:, ~df.columns.duplicated()]

    # 清洗字符串列
    str_cols = {"合同编号", "产品名称", "规格", "客户名称", "项目名称", "商务", "SKU编码"}
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].apply(safe_str)

    return df


# =========================
# 校验
# =========================
class ValidationError(Exception):
    pass


def validate_orders(df: pd.DataFrame, auto_fix: bool = False) -> Tuple[pd.DataFrame, List[str]]:
    """校验订单表。"""
    errors = []
    warnings = []
    df = df.copy()

    # 检查 合同编号
    if "合同编号" not in df.columns:
        raise ValidationError("订单表缺少「订单编号」列")

    empty_contract = df["合同编号"] == ""
    if empty_contract.any():
        rows = df[empty_contract].index.tolist()
        raise ValidationError(f"合同编号为空，共 {len(rows)} 行 (行号: {rows[:10]})")

    # 检查 合同数量
    if "合同数量" in df.columns:
        df["合同数量"] = df["合同数量"].apply(to_num)
        zero_qty = df["合同数量"] <= 0
        if zero_qty.any():
            errors.append(f"合同数量 <= 0，共 {zero_qty.sum()} 行")

        huge_qty = df["合同数量"] > 10000
        if huge_qty.any():
            for _, r in df[huge_qty].iterrows():
                warnings.append(f"合同数量异常大: {r['合同编号']} 数量={r['合同数量']}")

    # 检查 下单时间
    if "下单时间" in df.columns:
        # 检测是否为 Excel 数字日期（如 46156 = 2026-05-29）
        sample = df["下单时间"].dropna()
        if len(sample) > 0 and pd.api.types.is_numeric_dtype(sample):
            # Excel 数字日期 → datetime
            excel_epoch = pd.Timestamp("1899-12-30")
            df["下单时间"] = df["下单时间"].apply(
                lambda x: excel_epoch + pd.Timedelta(days=int(x)) if pd.notna(x) and x > 0 else pd.NaT
            )
        else:
            df["下单时间"] = pd.to_datetime(df["下单时间"], errors="coerce")
        no_date = df["下单时间"].isna()
        if no_date.any():
            if auto_fix:
                df.loc[no_date, "下单时间"] = pd.Timestamp.now()
                warnings.append(f"下单时间为空 {no_date.sum()} 行，已自动填入当天日期")
            else:
                warnings.append(f"下单时间为空 {no_date.sum()} 行")
    else:
        if auto_fix:
            df["下单时间"] = pd.Timestamp.now()
            warnings.append("缺少下单时间列，已自动填入当天日期")

    # 检查 产品名称
    if "产品名称" in df.columns:
        empty_name = df["产品名称"] == ""
        if empty_name.any():
            errors.append(f"产品名称为空，共 {empty_name.sum()} 行（请检查是否有「国网设备名称」或「设备名称」列）")
    else:
        errors.append("订单表缺少产品名称列（需有「国网设备名称」或「设备名称」列）")

    # 检查 SKU（非阻塞）
    if "SKU编码" in df.columns:
        empty_sku = df["SKU编码"] == ""
        if empty_sku.any():
            if auto_fix:
                warnings.append(f"SKU编码为空 {empty_sku.sum()} 行，将在排单时自动匹配")
            else:
                warnings.append(f"SKU编码为空 {empty_sku.sum()} 行（排单时会自动匹配）")

    if errors:
        raise ValidationError("数据校验失败:\n" + "\n".join(f"  - {e}" for e in errors))

    return df, warnings


def validate_inventory(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """校验库存表（接收已做列名映射后的 DataFrame）。"""
    errors = []
    warnings = []
    df = df.copy()

    if "库存数量" not in df.columns:
        raise ValidationError("库存表缺少「库存数量」列")

    if "国网设备名称" not in df.columns:
        raise ValidationError("库存表缺少「国网设备名称」列")

    if "国网设备型号" not in df.columns:
        raise ValidationError("库存表缺少「型号」列（将映射为「国网设备型号」）")

    df["库存数量"] = df["库存数量"].apply(to_num)
    negative_qty = df["库存数量"] < 0
    if negative_qty.any():
        warnings.append(f"库存数量为负，共 {negative_qty.sum()} 行，将按实际值写入")

    empty_name = df["国网设备名称"].apply(safe_str) == ""
    if empty_name.any():
        errors.append(f"国网设备名称为空，共 {empty_name.sum()} 行")

    empty_model = df["国网设备型号"].apply(safe_str) == ""
    if empty_model.any():
        errors.append(f"国网设备型号为空，共 {empty_model.sum()} 行")

    if errors:
        raise ValidationError("数据校验失败:\n" + "\n".join(f"  - {e}" for e in errors))

    return df, warnings


# =========================
# 主流程
# =========================
def import_excel(
    filepath: str,
    dry_run: bool = False,
    auto_fix: bool = False,
):
    """主入口：读取 Excel → 按 Sheet 类型分发 → 校验 → 写入飞书。"""
    print(f"读取文件: {filepath}")
    xl = pd.ExcelFile(filepath)
    print(f"发现 {len(xl.sheet_names)} 个 Sheet: {xl.sheet_names}\n")

    orders_done = False
    inventory_done = False

    for sheet_name in xl.sheet_names:
        print(f"{'=' * 60}")
        print(f"处理 Sheet: {sheet_name}")
        print(f"{'=' * 60}")

        df_raw = pd.read_excel(filepath, sheet_name=sheet_name)
        print(f"读取 {len(df_raw)} 行, {len(df_raw.columns)} 列")
        print(f"原始列名: {list(df_raw.columns)}")

        sheet_type = detect_sheet_type(df_raw)
        print(f"检测类型: {sheet_type}")

        # ============================================================
        # 订单表 → 主表 + 明细表
        # ============================================================
        if sheet_type == "orders":
            if orders_done:
                print("  已有订单表已处理，跳过。")
                continue

            # --- 清洗 ---
            try:
                df_orders, warnings = validate_orders(clean_orders_df(df_raw), auto_fix=auto_fix)
            except ValidationError as e:
                print(f"\n  校验失败! {e}")
                continue

            if warnings:
                print(f"\n  校验警告 ({len(warnings)} 条):")
                for w in warnings[:20]:
                    print(f"    - {w}")

            # --- 1) 写入订单主表（去重） ---
            print(f"\n  [1/2] 订单主表 ({TABLE_ID_MAIN})")
            # 只取主表需要的列
            main_avail_cols = [c for c in MAIN_COLUMNS if c in df_orders.columns]
            df_main = df_orders[main_avail_cols].drop_duplicates(subset=["合同编号"], keep="first")
            print(f"    去重后: {len(df_main)} 行 (原始 {len(df_orders)} 行)")

            if dry_run:
                print(f"    [DRY RUN] 将导入 {len(df_main)} 行")
                print(f"    字段映射: { {c: MAIN_FIELD_MAP.get(c, c) for c in main_avail_cols} }")
                print(f"    前 3 行:")
                print(df_main.head(3).to_string(index=False))
            else:
                write_df_to_bitable(
                    TABLE_ID_MAIN, df_main,
                    field_map=MAIN_FIELD_MAP,
                    date_cols={"下单时间"},
                )
                print(f"    写入成功! {len(df_main)} 行已导入。")

            # --- 2) 写入订单明细表 ---
            print(f"\n  [2/2] 订单明细表 ({TABLE_ID_ITEMS})")
            items_avail_cols = [c for c in ITEMS_COLUMNS if c in df_orders.columns]
            df_items = df_orders[items_avail_cols]

            # 过滤掉产品名称包含"费"的明细（费用类不需要出库）
            if "产品名称" in df_items.columns:
                fee_mask = df_items["产品名称"].str.contains("费", na=False)
                fee_count = fee_mask.sum()
                if fee_count > 0:
                    df_items = df_items[~fee_mask]
                    print(f"    已过滤 {fee_count} 行含'费'的明细（费用类无需出库）")

            print(f"    共 {len(df_items)} 行")

            if dry_run:
                print(f"    [DRY RUN] 将导入 {len(df_items)} 行")
                print(f"    字段映射: { {c: ITEMS_FIELD_MAP.get(c, c) for c in items_avail_cols} }")
                print(f"    前 3 行:")
                print(df_items.head(3).to_string(index=False))
            else:
                write_df_to_bitable(
                    TABLE_ID_ITEMS, df_items,
                    field_map=ITEMS_FIELD_MAP,
                    numeric_cols={"合同数量"},
                    date_cols={"下单时间"},
                )
                print(f"    写入成功! {len(df_items)} 行已导入。")

            orders_done = True

        # ============================================================
        # 库存表 → 库存快照表（清除后导入）
        # ============================================================
        elif sheet_type == "inventory":
            if inventory_done:
                print("  已有库存表已处理，跳过。")
                continue

            # --- 清洗 ---
            # 库存表不应用 OA_COLUMN_MAP，使用专用映射
            df_inv = df_raw.copy()
            df_inv = df_inv.loc[:, ~df_inv.columns.astype(str).str.match(r'^Unnamed')]
            df_inv = df_inv.rename(columns=lambda c: INV_COLUMN_MAP.get(c, c))

            # 清洗字符串
            for col in ["国网设备名称", "国网设备型号"]:
                if col in df_inv.columns:
                    df_inv[col] = df_inv[col].apply(safe_str)

            try:
                df_inv, warnings = validate_inventory(df_inv)
            except ValidationError as e:
                print(f"\n  校验失败! {e}")
                continue

            if warnings:
                print(f"\n  校验警告 ({len(warnings)} 条):")
                for w in warnings[:20]:
                    print(f"    - {w}")

            print(f"\n  [库存快照表] ({TABLE_ID_INV})")
            print(f"    共 {len(df_inv)} 行")

            # 只取库存表需要的列
            inv_avail_cols = [c for c in INV_COLUMNS if c in df_inv.columns]
            df_inv_write = df_inv[inv_avail_cols]

            # 构建 field map
            inv_field_map = {c: c for c in inv_avail_cols}

            if dry_run:
                print(f"    [DRY RUN] 将先清除全部旧数据，再导入 {len(df_inv_write)} 行")
                print(f"    字段映射: {inv_field_map}")
                print(f"    前 3 行:")
                print(df_inv_write.head(3).to_string(index=False))
            else:
                # 先清除旧库存
                deleted = delete_all_records(TABLE_ID_INV)
                print(f"    已清除旧库存 {deleted} 条记录")

                # 写入新库存
                write_df_to_bitable(
                    TABLE_ID_INV, df_inv_write,
                    field_map=inv_field_map,
                    numeric_cols={"库存数量"},
                )
                print(f"    写入成功! {len(df_inv_write)} 行已导入。")

            inventory_done = True

        else:
            print(f"  无法识别 Sheet 类型，跳过。")
            print(f"  提示: 订单表需包含「订单编号」列，库存表需包含「库存数量」列。")
            continue

    print(f"\n{'=' * 60}")
    print("导入完成。")
    if dry_run:
        print("（DRY RUN 模式，未实际写入飞书）")


def main():
    parser = argparse.ArgumentParser(
        description="OA 订单数据 & 库存快照 导入工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    python import_orders.py 数据.xlsx
    python import_orders.py 数据.xlsx --dry-run
    python import_orders.py 数据.xlsx --auto-fix

Excel 格式:
    Sheet1（订单表）: 订单编号, 客户名称, 项目名称, 商务, 下单时间, 国网设备名称, 设备型号, 数量
    Sheet2（库存表）: 国网设备名称, 型号, 库存数量
        """,
    )
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--dry-run", action="store_true", help="仅校验不写入（预览模式）")
    parser.add_argument("--auto-fix", action="store_true", help="自动修复已知问题（如填入默认日期、允许SKU为空等）")
    args = parser.parse_args()

    if not os.path.isfile(args.file):
        print(f"文件不存在: {args.file}")
        sys.exit(1)

    if not FEISHU_APP_SECRET:
        print("未配置 FEISHU_APP_SECRET，请在 .env 文件中设置")
        sys.exit(1)

    try:
        import_excel(args.file, dry_run=args.dry_run, auto_fix=args.auto_fix)
    except Exception as e:
        print(f"\n导入失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
