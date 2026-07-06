"""
一键导入+同步财务 v3 — 先清表，再写入主表/明细表/库存表，最后同步财务对账
"""
import httpx, time, json, openpyxl, sys, os

APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_MAIN = "tbl06oxGEdMNTEB8"
TABLE_ITEMS = "tblJn5iP6imjzE8h"
TABLE_INV = "tblFZNdEwW50izjh"

r = httpx.post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=15)
TOKEN = r.json()["tenant_access_token"]
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
BASE = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}"

def list_all_records(table_id, table_name):
    ids = []
    page_token = None
    while True:
        params = {"page_size": 500}
        if page_token: params["page_token"] = page_token
        r = httpx.get(f"{BASE}/tables/{table_id}/records", headers=HEADERS, params=params, timeout=60)
        d = r.json()
        if d.get("code") != 0:
            print(f"  [{table_name}] list error: {d.get('code')} {d.get('msg')}")
            return ids
        items = d.get("data", {}).get("items", [])
        if not items: break
        ids.extend(it["record_id"] for it in items)
        if not d["data"].get("has_more"): break
        page_token = d["data"]["page_token"]
    return ids

def delete_all(table_id, table_name):
    ids = list_all_records(table_id, table_name)
    if not ids:
        print(f"  [{table_name}] 无记录")
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
        print(f"    {min(i+200,len(ids))}/{len(ids)}")
        time.sleep(0.3)
    print(f"  [{table_name}] 完成")

def batch_create(table_id, table_name, records):
    total = len(records)
    written = 0
    for i in range(0, total, 100):
        batch = records[i:i+100]
        payload = {"records": [{"fields": r} for r in batch]}
        r = httpx.post(f"{BASE}/tables/{table_id}/records/batch_create",
                       headers=HEADERS, json=payload, timeout=60)
        jd = r.json()
        if jd.get("code") != 0:
            print(f"    写入失败: {jd.get('code')} {jd.get('msg')}")
            return written
        written += len(batch)
        print(f"    [{table_name}] {written}/{total}")
        time.sleep(0.3)
    return written

# ===== 第一步：清空三表 =====
print("=== 1/4 清空旧数据 ===")
for tid, tn in [(TABLE_MAIN, "主表"), (TABLE_ITEMS, "明细表"), (TABLE_INV, "库存表")]:
    delete_all(tid, tn)
time.sleep(2)

# ===== 第二步：读 Excel =====
print("\n=== 2/4 读取Excel ===")
XLSX = sys.argv[1] if len(sys.argv) > 1 else "/Users/xiaowang/AI/AI/AI_Scheduling_Project/测试集_演示_v2.xlsx"
wb = openpyxl.load_workbook(XLSX, data_only=True)

ws1 = wb["订单表"]
headers1 = [c.value for c in ws1[1]]
orders = []
for row in ws1.iter_rows(min_row=2, values_only=True):
    orders.append(dict(zip(headers1, row)))
print(f"  订单表: {len(orders)} 行")

ws2 = wb["库存表"]
headers2 = [c.value for c in ws2[1]]
inventory = []
for row in ws2.iter_rows(min_row=2, values_only=True):
    inventory.append(dict(zip(headers2, row)))
print(f"  库存表: {len(inventory)} 行")

# ===== 第三步：写入主表和明细表 =====
print("\n=== 3/4 写入订单数据 ===")

# 主表：按合同编号去重
seen = {}
main_records = []
for o in orders:
    cno = str(o.get("合同编号", ""))
    if cno in seen: continue
    seen[cno] = o
    sd = o.get("签订日期")
    if hasattr(sd, "strftime"): sd = sd.strftime("%Y-%m-%d")
    ts = int(time.mktime(time.strptime(str(sd), "%Y-%m-%d"))) * 1000 if sd else None
    main_records.append({
        "合同编号": cno,
        "客户名称": str(o.get("客户名称", "")),
        "代理商": str(o.get("代理商", "")),
        "项目名称": str(o.get("所属项目", "")),
        "下单日期": ts,
    })

n_main = batch_create(TABLE_MAIN, "主表", main_records)

# 明细表
items_records = []
for o in orders:
    sd = o.get("签订日期")
    if hasattr(sd, "strftime"): sd = sd.strftime("%Y-%m-%d")
    ts = int(time.mktime(time.strptime(str(sd), "%Y-%m-%d"))) * 1000 if sd else None
    items_records.append({
        "合同编号": str(o.get("合同编号", "")),
        "产品名称": str(o.get("产品名称", "")),
        "规格": str(o.get("明细规格", "")) if o.get("明细规格") else "",
        "合同数量": float(o.get("数量", 0)) if o.get("数量") else 0,
        "下单时间": ts,
    })

n_items = batch_create(TABLE_ITEMS, "明细表", items_records)

# 库存表
inv_records = []
import random
random.seed(42)
for inv in inventory:
    brand = inv.get("销售品牌", "")
    purchase = inv.get("待采购出库", 0) or 0
    inv_records.append({
        "国网设备名称": str(inv.get("国网设备名称", "")),
        "国网设备型号": str(inv.get("国网设备型号", "")) if inv.get("国网设备型号") else "",
        "库存数量": int(inv.get("库存数量", 0)) if inv.get("库存数量") else 0,
        "待采购出库": int(purchase),
        "数据来源": str(brand) if brand else "智造科技",
        "在途数量": str(int(purchase) * 2 // 3),
    })

n_inv = batch_create(TABLE_INV, "库存表", inv_records)

# ===== 第四步：同步财务对账 =====
print("\n=== 4/4 同步财务对账 ===")

# 4a：主表 → 对账总表
# 直接通过 API 写入，不依赖 scheduler_api.py 的函数
FINANCE_SUMMARY = "tbl6DaJN26jZFrz6"
FINANCE_DETAIL = "tblR5gHW5tOgGylp"

# 先清对账表旧数据
delete_all(FINANCE_SUMMARY, "对账总表")
delete_all(FINANCE_DETAIL, "对账明细表")

# 写入对账总表 (基本字段)
summary_records = []
for o in orders:
    cno = str(o.get("合同编号", ""))
    if hasattr(o.get("签订日期"), "strftime"):
        sd = o.get("签订日期").strftime("%Y-%m-%d")
    else:
        sd = str(o.get("签订日期", ""))
    ts = int(time.mktime(time.strptime(sd, "%Y-%m-%d"))) * 1000 if sd else None
    summary_records.append({
        "合同编号": cno,
        "下单日期": ts,
        "客户名称": str(o.get("客户名称", "")),
        "项目名称": str(o.get("所属项目", "")),
        "代理商": str(o.get("代理商", "")),
    })

# 去重写入
seen_summary = {}
uniq_summary = []
for r in summary_records:
    if r["合同编号"] not in seen_summary:
        seen_summary[r["合同编号"]] = True
        uniq_summary.append(r)
n_sum = batch_create(FINANCE_SUMMARY, "对账总表", uniq_summary)
print(f"  对账总表: {n_sum} 条")

# 写入对账明细表
detail_records = []
for o in orders:
    sd = o.get("签订日期")
    if hasattr(sd, "strftime"): sd = sd.strftime("%Y-%m-%d")
    ts = int(time.mktime(time.strptime(str(sd), "%Y-%m-%d"))) * 1000 if sd else None
    detail_records.append({
        "合同编号": str(o.get("合同编号", "")),
        "产品名称": str(o.get("产品名称", "")),
        "规格": str(o.get("明细规格", "")) if o.get("明细规格") else "",
        "合同数量": float(o.get("数量", 0)) if o.get("数量") else 0,
        "项目名称": str(o.get("所属项目", "")),
    })

n_det = batch_create(FINANCE_DETAIL, "对账明细表", detail_records)
print(f"  对账明细表: {n_det} 条")

print(f"\n=== 全部完成! ===")
print(f"  主表: {n_main} 合同")
print(f"  明细表: {n_items} 订单行")
print(f"  库存表: {n_inv} 条")
print(f"  对账总表: {n_sum} 条")
print(f"  对账明细表: {n_det} 条")
