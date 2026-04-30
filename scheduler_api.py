from fastapi import FastAPI, Request
from pydantic import BaseModel
import pandas as pd
from datetime import datetime, timedelta
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
def parse_feishu_field(value):
    """清洗飞书字段：提取纯文本、去空格、空值转空字符串、列表拼接"""
    # 处理飞书文本包装 {"text": "...", "type": "text"}
    if isinstance(value, dict):
        # 取第一个有效值（通常 text, number 等）
        for k in ("text", "number", "phone", "email", "url"):
            if k in value:
                return str(value[k]).strip()
        # 如果都没命中，转字符串（兜底）
        return str(value).strip()
    elif isinstance(value, str):
        return value.strip()
    elif isinstance(value, list):
        return ", ".join([str(v).strip() for v in value])
    elif value is None:
        return ""
    else:
        return value

# ========== 飞书配置（请替换为你自己的信息） ==========
FEISHU_APP_ID = "cli_a96c5d017d3a1cbb"           # TODO: 替换
FEISHU_APP_SECRET = "Bk7RzLFMmeVfsERXIoazcbyXKzRm7fE5"  # TODO: 替换
BITABLE_APP_TOKEN = "C5JzbAfnia0nT3sRvjucXgUGnDc"            # TODO: 替换为多维表格的 app_token（base/ 后面那串）
# 四个子表 ID
TABLE_ID_ITEMS = "tblJn5iP6imjzE8h"                   # TODO: 销售订单明细表 ID
TABLE_ID_SKU   = "tblVAGWeGHvmbFgJ"                   # TODO: SKU标准表 ID
TABLE_ID_INV   = "tblFZNdEwW50izjh"                   # TODO: 库存快照表 ID
# 结果表（需要先在飞书多维表格里创建两个子表：合同缺货总览 和 明细）
TABLE_ID_SUMMARY = "tblgMzZPyBBU5GLX"                 # TODO: 合同缺货总览表 ID
TABLE_ID_DETAIL  = "tblVRYTmbOsImzW2"                 # TODO: 明细表 ID
# ================================================

app = FastAPI()

class ScheduleRequest(BaseModel):
    trigger: str = "test"

# =========================
# 飞书客户端（全局）
# =========================
feishu_client = lark.Client.builder() \
    .app_id(FEISHU_APP_ID) \
    .app_secret(FEISHU_APP_SECRET) \
    .build()

import httpx

def fetch_bitable_to_df(table_id):
    """使用 HTTP 直接调用飞书 API，稳定读取多维表格（无 SDK 分页问题）"""
    # 1. 获取 tenant_access_token
    token_url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    token_resp = httpx.post(token_url, json={
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    })
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

        resp = httpx.get(url, headers=headers, params=params)
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"飞书读取表格失败: {data}")

        items = data.get("data", {}).get("items", [])
        for item in items:
            row = {}
            for field_name, value in item.get("fields", {}).items():
                value = parse_feishu_field(value)   # ⭐关键：先脱壳
                if isinstance(value, list):
                    row[field_name] = ", ".join([str(v) for v in value])
                else:
                    row[field_name] = value
            all_records.append(row)

        if not data.get("data", {}).get("has_more", False):
            break
        page_token = data["data"].get("page_token")
        if not page_token:
            break

    # 3. 构造 DataFrame 并清洗列名
    df = pd.DataFrame(all_records)
    if not df.empty:
        df.columns = df.columns.str.strip()
        # 清洗 SKU 相关列
        for col in ["SKU编码", "产品编码SKU", "SKU"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
    return df

def write_df_to_bitable(table_id, df):
    """将 DataFrame 写入飞书多维表格（逐条写入，简单可靠）"""
    for _, row in df.iterrows():
        fields = {}
        for col in df.columns:
            val = row[col]
            if pd.isna(val):
                continue          # 跳过空值
            elif isinstance(val, pd.Timestamp):
                fields[col] = val.strftime("%Y-%m-%d")
            elif isinstance(val, list):
                fields[col] = ", ".join([str(v) for v in val])
            else:
                fields[col] = str(val)

        record_body = AppTableRecord.builder().fields(fields).build()
        req = CreateAppTableRecordRequest.builder() \
            .app_token(BITABLE_APP_TOKEN) \
            .table_id(table_id) \
            .request_body(record_body) \
            .build()
        resp = feishu_client.bitable.v1.app_table_record.create(req)
        if not resp.success():
            print(f"写入记录失败: {resp.msg}")
            # 这里可以添加记录失败的行号，方便排查

# =========================
# 数据缓存（避免每次请求重复读取飞书，减轻API压力）
# =========================
data_cache = {
    "items": None,
    "sku": None,
    "inv": None
}

def load_data():
    """从飞书一次性加载数据到内存"""
    print("正在从飞书加载数据...")
    data_cache["items"] = fetch_bitable_to_df(TABLE_ID_ITEMS)
    data_cache["sku"]   = fetch_bitable_to_df(TABLE_ID_SKU)
    data_cache["inv"]   = fetch_bitable_to_df(TABLE_ID_INV)
    print("数据加载完毕。")

# =========================
# 计算可用库存
# =========================
def get_available_stock(inv_row):
    stock_qty = inv_row.get("库存数量", 0)
    out_qty = inv_row.get("待采购出库", 0)
    transit_qty = inv_row.get("在途数量", 0)
    # 安全处理 NaN
    if pd.isna(stock_qty): stock_qty = 0
    if pd.isna(out_qty): out_qty = 0
    if pd.isna(transit_qty): transit_qty = 0
    return stock_qty - out_qty + transit_qty

# =========================
# 计算发货周期（核心规则）
# =========================
def calc_days(order, sku_info, stock_ok):
    project = order["项目类型"]

    if order["是否紧急订单"] == "是":
        return 2
    if order["是否换货订单"] == "是":
        return 2
    if order["是否补发订单"] == "是":
        return 2
    if order["是否维修订单"] == "是":
        return 10

    if project == "远程控制项目" or order["是否RB800"] == "是":
        if stock_ok:
            return 7
        if "电话" in order["产品名称"]:
            return 15
        if sku_info["是否新产品"] == "是":
            return 25
        return 10
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
def run_scheduler(payload: ScheduleRequest = None):
    # 默认参数
    if payload is None:
        payload = ScheduleRequest()

    # 如果缓存为空，尝试加载飞书数据
    if data_cache["items"] is None:
        try:
            load_data()
        except Exception as e:
            return {"error": f"飞书数据加载失败: {str(e)}"}

    items = data_cache["items"]
    sku   = data_cache["sku"]
    inv   = data_cache["inv"]

    result_rows = []

    for _, row in items.iterrows():
        if "SKU编码" not in row:
            print("错误：销售订单明细表缺少“SKU编码”列")
            continue

        sku_code = row["SKU编码"]
        if not sku_code or sku_code == "nan":
            print(f"警告：订单 {row.get('合同编号', '未知')} 的SKU编码为空，跳过")
            continue

        # 安全查找 SKU 主数据
        sku_match = sku[sku["产品编码SKU"] == sku_code]
        if sku_match.empty:
            print(f"警告：SKU主数据中未找到 {sku_code}，跳过订单 {row['合同编号']}")
            continue
        sku_info = sku_match.iloc[0]

        # 安全查找库存数据
        inv_match = inv[inv["SKU"] == sku_code]
        if inv_match.empty:
            print(f"警告：库存表中未找到 {sku_code}，跳过订单 {row['合同编号']}")
            continue
        inv_info = inv_match.iloc[0]

        available_stock = get_available_stock(inv_info)
        order_qty = row["合同数量"]
        gap = max(order_qty - available_stock, 0)
        stock_ok = gap == 0

        days = calc_days(row, sku_info, stock_ok)

        today = datetime.today().date()
        ship_date = today + timedelta(days=days)

        result_rows.append({
            "AI订单ID": row["合同编号"],
            "项目类型": row["项目类型"],
            "产品SKU": sku_code,
            "订单数量": order_qty,
            "当前库存": available_stock,
            "库存状态": "有货" if stock_ok else "缺货",
            "缺口数量": gap,
            "计算发货周期(天)": days,
            "建议排产日期": today,
            "预计发货日期": ship_date
        })

    df_detail = pd.DataFrame(result_rows)

    # ===== 按合同汇总 =====
    summary_rows = []
    if not df_detail.empty:
        def safe_val(x):
            if isinstance(x, float) and pd.isna(x):
                return None
            if isinstance(x, pd.Timestamp) and pd.isna(x):
                return None
            return x

        for contract_id, group in df_detail.groupby("AI订单ID"):
            project_type = safe_val(group["项目类型"].iloc[0])
            sku_count = safe_val(group["产品SKU"].nunique())
            total_qty = safe_val(group["订单数量"].sum())
            shortage_mask = group["库存状态"] == "缺货"
            shortage_sku_count = safe_val(shortage_mask.sum())
            shortage_sku_list = group.loc[shortage_mask, "产品SKU"].unique().tolist()
            overall_status = "缺货" if shortage_sku_count else "有货"
            earliest_ship = safe_val(group["预计发货日期"].min())

            summary_rows.append({
                "AI订单ID": safe_val(contract_id),
                "项目类型": project_type if project_type is not None else "",
                "订单SKU总数": sku_count if sku_count is not None else 0,
                "订单总数量": total_qty if total_qty is not None else 0,
                "缺货SKU数": shortage_sku_count if shortage_sku_count is not None else 0,
                "缺货SKU列表": ", ".join(shortage_sku_list) if shortage_sku_list else "",
                "整体状态": overall_status,
                "最早预计发货": earliest_ship if earliest_ship is not None else ""
            })
    summary = pd.DataFrame(summary_rows)

    # ===== 写入飞书结果表 =====
    # 先清空结果表（可选，避免数据重复）—— 此处简单起见，直接追加写入
    # 如果想要覆盖，需要先删除原记录，这里暂不处理，你可手动清空结果表
    print("正在写入合同缺货总览...")
    write_df_to_bitable(TABLE_ID_SUMMARY, summary)
    print("正在写入明细细...")
    write_df_to_bitable(TABLE_ID_DETAIL, df_detail)
    print("飞书写入完成。")

    return {
        "msg": "AI排单完成",
        "合同总览": summary.to_dict(orient="records") if not summary.empty else [],
        "明细": df_detail.to_dict(orient="records")
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

    # ③ 自动执行排单
    try:
        print("开始执行自动排单...")
        result = run_scheduler(ScheduleRequest())
        print("排单完成")
        return {"msg": "排单完成", "detail": result}
    except Exception as e:
        print("排单失败：", str(e))
        return {"error": f"排单失败: {str(e)}"}