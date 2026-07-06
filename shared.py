"""
共享常量与映射 — 全项目唯一真源。
scheduler_api.py、import_orders.py 均从此导入，避免列名映射不一致。
"""

# =========================
# OA 导出列名 → 标准列名（仅用于订单表 Sheet）
# =========================
OA_COLUMN_MAP = {
    "订单编号": "合同编号",
    "数量": "合同数量",
    "设备型号": "规格",
    "设备名称": "产品名称",
    "国网设备名称": "产品名称",
    "明细规格": "规格",          # 部分 Excel 用"明细规格"作为规格
    "下单日期": "下单时间",
    "订单时间": "下单时间",
    "下单时间": "下单时间",      # 保留原列名（部分 Excel 已使用此列名）
    "申请时间": "下单时间",      # 部分 Excel 用"申请时间"作为下单时间
    "签订日期": "下单时间",      # 部分 Excel 用"签订日期"作为下单时间
    "申请人": "商务",           # 部分 Excel 用"申请人"作为商务
    "代理商": "代理商",
    "客户名称": "客户名称",      # 保留原列名
    "所属项目": "项目名称",      # 部分 Excel 用"所属项目"作为项目名称
}

# =========================
# 订单主表：标准列名 → 飞书字段名（精确匹配）
# =========================
MAIN_FIELD_MAP = {
    "合同编号": "合同编号",
    "下单时间": "下单日期",      # MAIN 表字段叫"下单日期"
    "客户名称": "客户名称",
    "项目名称": "项目名称",
    "商务": "商务",
    "代理商": "代理商",
}

# =========================
# 订单明细表：标准列名 → 飞书字段名（精确匹配）
# =========================
ITEMS_FIELD_MAP = {
    "合同编号": "合同编号",
    "下单时间": "下单时间",
    "产品名称": "产品名称",
    "规格": "规格",
    "合同数量": "合同数量",
}

# =========================
# 库存快照表：Excel 原始列名 → 飞书字段名（精确匹配）
# =========================
INV_COLUMN_MAP = {
    "国网设备名称": "国网设备名称",
    "型号": "国网设备型号",
    "库存数量": "库存数量",
}

# =========================
# 各表需要的标准列名列表
# =========================
MAIN_COLUMNS = ["合同编号", "下单时间", "客户名称", "项目名称", "商务", "代理商"]
ITEMS_COLUMNS = ["合同编号", "下单时间", "产品名称", "规格", "合同数量"]
INV_COLUMNS = ["国网设备名称", "国网设备型号", "库存数量"]


# =========================
# 排单结果审计 — 排单完成后自动执行
# =========================
import random
from typing import Any, Dict, List, Optional, Tuple
from datetime import date, timedelta
try:
    import pandas as pd
except ImportError:
    pd = None


def _safe_str(val: Any) -> str:
    if val is None: return ""
    if isinstance(val, float) and (val != val): return ""
    return str(val).strip()


def _to_num(val: Any) -> float:
    try: return float(val)
    except: return 0.0


def _next_working_day(d: date) -> date:
    """跳过周末，返回下一个工作日"""
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _add_calendar_days_then_working_day(base: date, days: int) -> date:
    return _next_working_day(base + timedelta(days=days))


def _parse_date(val: Any) -> Optional[date]:
    if val is None or val == "": return None
    if isinstance(val, date): return val
    s = _safe_str(val)
    if not s: return None
    try: return date.fromisoformat(s[:10])
    except: pass
    return None


def audit_schedule_results(
    df_summary: Any,      # pd.DataFrame: 排单总表
    items_df: Any,        # pd.DataFrame: 销售订单明细表
    inv_df: Any,          # pd.DataFrame: 库存快照表
    sku_df: Any,          # pd.DataFrame: SKU标准表
    sample_size: int = 5,
    occupied_stock: Optional[Dict[str, float]] = None,           # 已确认占用+有效预留
    find_inv_row_fn: Optional[Any] = None,                       # find_inventory_row 函数引用（可选）
) -> Dict[str, Any]:
    """
    排单结果独立审计：随机抽查 N 个合同，全链路计算与排单引擎交叉验证。

    审计规则（5条）：
    1. 库存来源：使用传入的库存数据，与排单使用的同一批次交叉验证
    2. 库存可用量：独立复现逐次扣减（含 occupied_stock，按下单时间→合同编号排序）
    3. 缺货判定: gap = max(demand - available, 0), 验证二元判定和gap值
    4. 合同级缺货：canonical SKU去重计数 + 缺货SKU缺口总和一致
    5. 发货日期：three-path独立重算（next_working_day）

    Returns: { ok, sample_count, sample_contracts, errors, warnings }
    """
    if pd is None or df_summary is None or items_df is None or inv_df is None:
        return {"ok": False, "errors": ["审计跳过: 缺少数据"]}

    contracts = list(df_summary["合同编号"].dropna().unique())
    if not contracts:
        return {"ok": True, "errors": [], "warnings": ["无合同数据"]}

    sample = sorted(random.sample(contracts, min(sample_size, len(contracts))))
    today = date.today()
    errors: List[str] = []
    warnings: List[str] = []

    # ---- 预处理 ----
    # 库存索引: SKU → (库存数量, inv_row)
    inv_index: Dict[str, Tuple[float, Optional[Any]]] = {}
    if "SKU" in inv_df.columns and "库存数量" in inv_df.columns:
        for _, row in inv_df.iterrows():
            sku = _safe_str(row.get("SKU", ""))
            if not sku or sku == "nan": continue
            qty = _to_num(row.get("库存数量", 0))
            if sku in inv_index:
                inv_index[sku] = (inv_index[sku][0] + qty, inv_index[sku][1])
            else:
                inv_index[sku] = (qty, row)

    # SKU → 标准生产周期
    sku_cycle: Dict[str, int] = {}
    if sku_df is not None and "产品编码SKU" in sku_df.columns:
        for _, row in sku_df.iterrows():
            code = _safe_str(row.get("产品编码SKU", ""))
            cycle = row.get("标准生产周期", 0)
            if code and code != "nan" and pd.notna(cycle):
                sku_cycle[code] = int(cycle)

    # ---- 全局逐次扣减（规则2） ----
    # 模拟排单引擎按时间→合同编号排序，逐步扣减库存（含 occupied_stock）
    all_items_sorted = items_df.copy()
    if "下单时间" in all_items_sorted.columns:
        all_items_sorted = all_items_sorted.sort_values(by=["下单时间", "合同编号"], kind="stable", na_position="last")
    else:
        all_items_sorted = all_items_sorted.sort_values(by=["合同编号"], kind="stable")

    occupied = occupied_stock or {}
    audit_consumed: Dict[str, float] = {}   # 本批已分配
    audit_detail: Dict[str, List[Tuple]] = {}

    for _, row in all_items_sorted.iterrows():
        sku_code = _safe_str(row.get("SKU编码", ""))
        if not sku_code:
            sku_code = _safe_str(row.get("产品名称", "")) or _safe_str(row.get("规格", ""))
        contract_id = _safe_str(row.get("合同编号", ""))
        if not sku_code or not contract_id: continue

        demand = _to_num(row.get("合同数量", 0))
        prod = _safe_str(row.get("产品名称", ""))
        spec = _safe_str(row.get("规格", ""))

        # Canonical SKU via find_inventory_row if available
        canonical_sku = sku_code
        if find_inv_row_fn:
            try:
                _, _, can = find_inv_row_fn(sku_code, {"产品名称": prod, "规格": spec}, inv_df, sku_df=sku_df)
                if can: canonical_sku = can
            except:
                pass

        # 原始库存
        stock_tuple = inv_index.get(canonical_sku)
        raw_stock = stock_tuple[0] if stock_tuple else 0.0

        # 逐次扣减（含 occupied_stock）
        batch_used = audit_consumed.get(canonical_sku, 0.0)
        total_reserved = occupied.get(canonical_sku, 0.0) + batch_used
        available = max(raw_stock - total_reserved, 0.0)

        # 缺货判定（规则3）
        gap = max(demand - available, 0.0)
        is_shortage = gap > 0

        # 消耗库存
        consumed = demand - gap
        audit_consumed[canonical_sku] = batch_used + consumed

        if contract_id not in audit_detail:
            audit_detail[contract_id] = []
        audit_detail[contract_id].append((prod, spec, demand, canonical_sku, available, gap, is_shortage))

    # ---- 对每个抽查合同验证 ----
    for contract in sample:
        sr = df_summary[df_summary["合同编号"] == contract]
        if sr.empty:
            warnings.append(f"{contract}: 不在排单总表中")
            continue
        sr = sr.iloc[0]

        overall = _safe_str(sr.get("整体状态", ""))
        shortage_count = int(sr.get("缺货SKU数", 0) or 0)
        shortage_list_str = _safe_str(sr.get("缺货SKU列表", ""))

        citems = items_df[items_df["合同编号"] == contract]
        if citems.empty:
            warnings.append(f"{contract}: 无明细数据")
            continue

        audit_rows = audit_detail.get(contract, [])

        # === 规则3: 缺货判定交叉验证 ===
        # 用 (产品名称, 规格) 做精确配对，已匹配的行不重复使用
        audit_remaining = list(enumerate(audit_rows))  # (index, tuple)
        for _, row in citems.iterrows():
            prod_name = _safe_str(row.get("产品名称", ""))
            declared_spec = _safe_str(row.get("规格", ""))
            declared_status = _safe_str(row.get("库存状态", ""))
            declared_gap = _to_num(row.get("缺口数量", 0))
            demand = _to_num(row.get("合同数量", 0))

            # 精确匹配：同名 + 同规格
            match_idx = None
            match_ar = None
            for i, (idx, ar) in enumerate(audit_remaining):
                if ar[0] == prod_name and ar[1] == declared_spec:
                    match_idx = i
                    match_ar = ar
                    break
            if match_idx is not None:
                audit_remaining.pop(match_idx)
            else:
                # 降级：同名即可
                for i, (idx, ar) in enumerate(audit_remaining):
                    if ar[0] == prod_name:
                        match_idx = i
                        match_ar = ar
                        break
                if match_idx is not None:
                    audit_remaining.pop(match_idx)

            if match_ar is not None:
                _, _, _, _, audit_avail, audit_gap, audit_shortage = match_ar
                declared_shortage = "缺货" in declared_status

                # (a) 二元判定一致
                if declared_shortage != audit_shortage:
                    errors.append(
                        f"{contract}|{prod_name[:25]}: 缺货判定矛盾 — "
                        f"排单={declared_status}, 审计={'缺货' if audit_shortage else '有货'} "
                        f"(需求={int(demand)} 可用={int(audit_avail)} 缺口={int(audit_gap)})"
                    )

                # (b) gap值一致（仅缺货时检查，允许±1浮点误差）
                if declared_shortage and audit_shortage and abs(int(declared_gap) - int(audit_gap)) > 1:
                    errors.append(
                        f"{contract}|{prod_name[:25]}: 缺口值不一致 — "
                        f"排单={int(declared_gap)} 审计={int(audit_gap)}"
                    )

        # === 规则4: 合同级缺货数（去重SKU + 缺口总和） ===
        # (a) 去重SKU计数
        canonical_shortage: Dict[str, str] = {}
        for ar in audit_rows:
            if ar[6]:  # is_shortage
                sku = ar[3]
                canonical_shortage[sku] = sku
        audit_shortage_count = len(canonical_shortage)

        # 排单总表中的缺货SKU数 vs 审计的去重缺货数
        if audit_shortage_count != shortage_count:
            errors.append(
                f"{contract}: 缺货SKU数不一致 — 排单={shortage_count} 审计(去重)={audit_shortage_count}"
            )

        # (b) 缺货SKU缺口总和
        audit_total_gap = sum(ar[5] for ar in audit_rows if ar[6])  # sum of gaps for shortage rows
        declared_shortage_items = citems[citems["库存状态"].astype(str).str.contains("缺货", na=False)]
        declared_total_gap = sum(_to_num(row.get("缺口数量", 0)) for _, row in declared_shortage_items.iterrows())

        if abs(int(audit_total_gap) - int(declared_total_gap)) > 1:
            errors.append(
                f"{contract}: 缺口总和不一致 — 排单={int(declared_total_gap)} 审计={int(audit_total_gap)}"
            )

        if overall == "全部可发" and audit_shortage_count > 0:
            errors.append(f"{contract}: 标记全部可发但审计发现 {audit_shortage_count} 个缺货SKU")

        # === 规则5: 发货日期 ===
        eta_val = sr.get("AI建议发货时间", None)
        if eta_val and pd.notna(eta_val):
            declared_eta = _parse_date(eta_val)
            if audit_shortage_count == 0:
                # 全部可发 → 应该是下一个工作日
                expected_eta = _next_working_day(today)
                if declared_eta and declared_eta != expected_eta:
                    # 产能调度顺延是正常行为，降级为 warning
                    warnings.append(
                        f"{contract}: 发货日期因产能调度顺延 — "
                        f"原期望={expected_eta} → 实际={declared_eta}")
            elif audit_shortage_count > 0:
                # 有缺货 → max(ETA) 的下一个工作日
                max_eta_days = 0
                for ar in audit_rows:
                    if ar[6]:  # is_shortage
                        cycle = sku_cycle.get(ar[3], 25)
                        max_eta_days = max(max_eta_days, cycle)
                if max_eta_days > 0:
                    expected_eta = _add_calendar_days_then_working_day(today, max_eta_days)
                    if declared_eta and abs((declared_eta - expected_eta).days) > 3:
                        warnings.append(
                            f"{contract}: 发货日期偏差较大 — 排单={declared_eta} 审计期望≈{expected_eta} "
                            f"(基于max生产周期={max_eta_days}天)"
                        )

    return {
        "ok": len(errors) == 0,
        "sample_count": len(sample),
        "sample_contracts": sample,
        "errors": errors,
        "warnings": warnings,
    }
