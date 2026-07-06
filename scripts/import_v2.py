"""
导入 v2 Demo 数据到飞书多维表格
1. 清空所有旧记录
2. 写入主表（80合同）→ 写入明细表（241订单行）→ 写入库存表（587条）
"""
import httpx, time, json, openpyxl, os

# ===== Config =====
APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_MAIN = "tbl06oxGEdMNTEB8"     # 销售订单主表
TABLE_ITEMS = "tblJn5iP6imjzE8h"    # 销售订单明细表
TABLE_INV = "tblFZNdEwW50izjh"      # 库存快照表

# ===== Auth =====
r = httpx.post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=15)
TOKEN = r.json()["tenant_access_token"]
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
BASE = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}"


def list_all_records(table_id):
    """获取表中全部记录的 ID"""
    ids = []
    page_token = None
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        r = httpx.get(f"{BASE}/tables/{table_id}/records", headers=HEADERS,
                      params=params, timeout=60)
        d = r.json()
        if d.get("code") != 0:
            print(f"  list error: {d}")
            break
        for it in d["data"]["items"]:
            ids.append(it["record_id"])
        if not d["data"].get("has_more"):
            break
        page_token = d["data"]["page_token"]
    return ids


def delete_all_records(table_id, table_name):
    """全量删除表中记录"""
    ids = list_all_records(table_id)
    if not ids:
        print(f"  [{table_name}] 没有记录，跳过")
        return
    print(f"  [{table_name}] 删除 {len(ids)} 条...")
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        r = httpx.post(f"{BASE}/tables/{table_id}/records/batch_delete",
                       headers=HEADERS, json={"records": batch}, timeout=60)
        jd = r.json()
        if jd.get("code") != 0:
            print(f"    删除失败: {jd.get('code')} {jd.get('msg')}")
            return
        print(f"    已删 {min(i+200, len(ids))}/{len(ids)}")
        time.sleep(0.3)
    print(f"  [{table_name}] 删除完成")


def batch_write(table_id, table_name, records, batch_size=100):
    """批量写入记录"""
    total = len(records)
    written = 0
    for i in range(0, total, batch_size):
        batch = records[i:i+batch_size]
        payload = {"records": [{"fields": r} for r in batch]}
        r = httpx.post(f"{BASE}/tables/{table_id}/records/batch_create",
                       headers=HEADERS, json=payload, timeout=60)
        jd = r.json()
        if jd.get("code") != 0:
            print(f"    写入失败: {jd.get('code')} {jd.get('msg')}")
            print(f"    response: {json.dumps(jd, ensure_ascii=False)[:500]}")
            return written
        written += len(batch)
        print(f"    [{table_name}] {written}/{total}")
        time.sleep(0.3)
    return written


# ============================================================
# 第一步: 清空旧数据 (SKIP - 已在上次运行中清空)
# ============================================================
print("=== 第一步: 跳过(表已空) ===")

# ============================================================
# 第二步: 读取 Excel 数据
# ============================================================
print("\n=== 第二步: 读取 Excel ===")
XLSX = "/Users/xiaowang/AI/AI/AI_Scheduling_Project/测试集_演示_v2.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)

# 订单表
ws_orders = wb["订单表"]
order_headers = [c.value for c in ws_orders[1]]
print(f"  订单表: {ws_orders.max_row-1} 行, 列: {order_headers}")

orders_raw = []
for row in ws_orders.iter_rows(min_row=2, values_only=True):
    orders_raw.append(dict(zip(order_headers, row)))

# 库存表
ws_inv = wb["库存表"]
inv_headers = [c.value for c in ws_inv[1]]
print(f"  库存表: {ws_inv.max_row-1} 行")

inv_raw = []
for row in ws_inv.iter_rows(min_row=2, values_only=True):
    inv_raw.append(dict(zip(inv_headers, row)))

# ============================================================
# 第三步: 写入主表（合同级，按合同编号去重）
# ============================================================
print("\n=== 第三步: 写入主表 ===")
seen_contracts = {}
main_records = []
for o in orders_raw:
    cno = str(o.get("合同编号", ""))
    if cno in seen_contracts:
        continue
    seen_contracts[cno] = o
    # 日期格式转化
    sign_date = o.get("签订日期")
    if hasattr(sign_date, "strftime"):
        sign_date = sign_date.strftime("%Y-%m-%d")
    main_records.append({
        "合同编号": cno,
        "客户名称": str(o.get("客户名称", "")),
        "代理商": str(o.get("代理商", "")),
        "项目名称": str(o.get("所属项目", "")),
        "下单日期": int(time.mktime(time.strptime(str(sign_date), "%Y-%m-%d"))) * 1000
            if sign_date else None,
    })

n_main = batch_write(TABLE_MAIN, "主表", main_records)
print(f"  主表写入 {n_main} 条合同")

# ============================================================
# 第四步: 写入明细表（订单行级）
# ============================================================
print("\n=== 第四步: 写入明细表 ===")
items_records = []
for o in orders_raw:
    cno = str(o.get("合同编号", ""))
    pname = str(o.get("产品名称", ""))
    spec = str(o.get("明细规格", "")) if o.get("明细规格") else ""
    qty = o.get("数量")
    sign_date = o.get("签订日期")
    if hasattr(sign_date, "strftime"):
        sign_date = sign_date.strftime("%Y-%m-%d")

    items_records.append({
        "合同编号": cno,
        "产品名称": pname,
        "规格": spec,
        "合同数量": float(qty) if qty else 0,
        "下单时间": int(time.mktime(time.strptime(str(sign_date), "%Y-%m-%d"))) * 1000
            if sign_date else None,
    })

n_items = batch_write(TABLE_ITEMS, "明细表", items_records, batch_size=100)
print(f"  明细表写入 {n_items} 条订单行")

# ============================================================
# 第五步: 写入库存表
# ============================================================
print("\n=== 第五步: 写入库存表 ===")
inv_records = []
for inv in inv_raw:
    brand = inv.get("销售品牌", "")
    inv_records.append({
        "国网设备名称": str(inv.get("国网设备名称", "")),
        "国网设备型号": str(inv.get("国网设备型号", "")) if inv.get("国网设备型号") else "",
        "库存数量": int(inv.get("库存数量", 0)) if inv.get("库存数量") else 0,
        "待采购出库": int(inv.get("待采购出库", 0)) if inv.get("待采购出库") else 0,
        "数据来源": str(brand) if brand else "智造科技",
        "在途数量": str(int(inv.get("待采购出库", 0) or 0) * 2 // 3) if inv.get("待采购出库") else "0",
    })

n_inv = batch_write(TABLE_INV, "库存表", inv_records, batch_size=100)
print(f"  库存表写入 {n_inv} 条")

print(f"\n=== 导入完成! ===")
print(f"  主表: {n_main} 合同")
print(f"  明细表: {n_items} 订单行")
print(f"  库存表: {n_inv} 条")
