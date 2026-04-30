from fastapi import FastAPI
from pydantic import BaseModel          
import pandas as pd
from datetime import datetime, timedelta

app = FastAPI()

class ScheduleRequest(BaseModel):
    trigger: str = "test"   

app = FastAPI()

# =========================
# 计算可用库存
# =========================
def get_available_stock(inv_row):
    return inv_row["库存数量"] - inv_row["待采购出库"] + inv_row["在途数量"]

# =========================
# 计算发货周期（核心规则）
# =========================
def calc_days(order, sku_info, stock_ok):

    project = order["项目类型"]

    # 优先级最高规则
    if order["是否紧急订单"] == "是":
        return 2
    if order["是否换货订单"] == "是":
        return 2
    if order["是否补发订单"] == "是":
        return 2
    if order["是否维修订单"] == "是":
        return 10

    # =====================
    # 远程项目规则
    # =====================
    if project == "远程控制项目" or order["是否RB800"] == "是":

        if stock_ok:
            return 7

        if "电话" in order["产品名称"]:
            return 15

        if sku_info["是否新产品"] == "是":
            return 25

        return 10

    # =====================
    # 常规项目规则
    # =====================
    else:
        if stock_ok:
            return 3

        if sku_info["是否外采"] == "是":
            return 7

        if sku_info["是否自研"] == "是":
            return 5

    return 7

# =========================
# 主排单API
# =========================
@app.post("/schedule")
def run_scheduler(payload: ScheduleRequest):

    orders = pd.read_excel("F:/AI/AI_Scheduling_Project/data/sales_orders.xlsx")
    items = pd.read_excel("F:/AI/AI_Scheduling_Project/data/sales_order_items.xlsx")
    sku = pd.read_excel("F:/AI/AI_Scheduling_Project/data/sku_master.xlsx")
    inv = pd.read_excel("F:/AI/AI_Scheduling_Project/data/inventory.xlsx")

 # ===== 新增：清洗列名和SKU数据 =====
    items.columns = items.columns.str.strip()
    sku.columns = sku.columns.str.strip()
    inv.columns = inv.columns.str.strip()

    result_rows = []

    for _, row in items.iterrows():

        sku_code = row["SKU编码"]
# ===== 改动：安全查找 SKU 主数据 =====
        sku_match = sku[sku["产品编码SKU"] == sku_code]
        if sku_match.empty:
            print(f"警告：SKU主数据中未找到 {sku_code}，跳过订单 {row['合同编号']}")
            continue
        sku_info = sku_match.iloc[0]
        # ===================================

        # ===== 改动：安全查找库存数据 =====
        inv_match = inv[inv["SKU"] == sku_code]
        if inv_match.empty:
            print(f"警告：库存表中未找到 {sku_code}，跳过订单 {row['合同编号']}")
            continue
        inv_info = inv_match.iloc[0]

        # ===== 新增：把 NaN 安全转成 0 =====
        stock_qty = inv_info["库存数量"]
        out_qty = inv_info["待采购出库"]
        transit_qty = inv_info["在途数量"]

        if pd.isna(stock_qty):
            stock_qty = 0
        if pd.isna(out_qty):
            out_qty = 0
        if pd.isna(transit_qty):
            transit_qty = 0

        available_stock = stock_qty - out_qty + transit_qty
        # ===================================
        gap = max(row["合同数量"] - available_stock, 0)
        stock_ok = gap == 0

        days = calc_days(row, sku_info, stock_ok)

        today = datetime.today().date()
        ship_date = today + timedelta(days=days)

        result_rows.append({
            "AI订单ID": row["合同编号"],
            "项目类型": row["项目类型"],
            "产品SKU": sku_code,
            "订单数量": row["合同数量"],
            "当前库存": available_stock,
            "库存状态": "有货" if stock_ok else "缺货",
            "缺口数量": gap,
            "计算发货周期(天)": days,
            "建议排产日期": today,
            "预计发货日期": ship_date
        })

    df_detail = pd.DataFrame(result_rows)
    df_detail = pd.DataFrame(result_rows)

    # ===== 按合同汇总（循环构建，稳定不出错） =====
      # 安全值转换：把 NaN/NaT 变成 None，字符串 "nan" 变成空
    def safe_value(val):
        if isinstance(val, float) and (pd.isna(val) or pd.isnull(val)):
            return None
        if isinstance(val, pd.Timestamp) and pd.isna(val):
            return None
        return val

    summary_rows = []
    if not df_detail.empty:
        for contract_id, group in df_detail.groupby("AI订单ID"):
            project_type = safe_value(group["项目类型"].iloc[0])
            sku_count = safe_value(group["产品SKU"].nunique())
            total_qty = safe_value(group["订单数量"].sum())
            shortage_mask = group["库存状态"] == "缺货"
            shortage_sku_count = safe_value(shortage_mask.sum())
            shortage_sku_list = group.loc[shortage_mask, "产品SKU"].unique().tolist()
            overall_status = "缺货" if shortage_sku_count else "有货"  # shortage_sku_count 可能为 0 或 None，None 时为假
            earliest_ship = safe_value(group["预计发货日期"].min())
            summary_rows.append({
                "AI订单ID": safe_value(contract_id),
                "项目类型": project_type if project_type is not None else "",
                "订单SKU总数": sku_count if sku_count is not None else 0,
                "订单总数量": total_qty if total_qty is not None else 0,
                "缺货SKU数": shortage_sku_count if shortage_sku_count is not None else 0,
                "缺货SKU列表": shortage_sku_list,
                "整体状态": overall_status,
                "最早预计发货": earliest_ship
            })