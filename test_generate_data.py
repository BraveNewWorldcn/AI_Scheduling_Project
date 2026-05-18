"""生成测试数据并写入飞书——只新增，不修改/删除任何已有记录。"""
import httpx, json, time, sys, random, os
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8')

APP_ID = os.getenv("FEISHU_APP_ID", "cli_a96c5d017d3a1cbb")
APP_SECRET = os.getenv("FEISHU_APP_SECRET")
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "C5JzbAfnia0nT3sRvjucXgUGnDc")

if not APP_SECRET:
    raise ValueError("请先设置环境变量 FEISHU_APP_SECRET")

# ---------- 获取 token ----------
resp = httpx.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
    json={'app_id': APP_ID, 'app_secret': APP_SECRET}, timeout=30)
TOKEN = resp.json()['tenant_access_token']
HEADERS = {'Authorization': f'Bearer {TOKEN}', 'Content-Type': 'application/json; charset=utf-8'}

def write_records(table_id, records, label):
    """批量写入记录，每批 500 条。"""
    url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{table_id}/records/batch_create'
    batch = []
    ok, fail = 0, 0
    for rec in records:
        batch.append({'fields': rec})
        if len(batch) >= 500:
            r = httpx.post(url, headers=HEADERS, json={'records': batch}, timeout=60)
            d = r.json()
            if d.get('code') == 0: ok += len(batch)
            else: fail += len(batch); print(f'  ❌ {label} batch fail: {d.get("msg",d)}')
            batch.clear()
    if batch:
        r = httpx.post(url, headers=HEADERS, json={'records': batch}, timeout=60)
        d = r.json()
        if d.get('code') == 0: ok += len(batch)
        else: fail += len(batch); print(f'  ❌ {label} batch fail: {d.get("msg",d)}')
    print(f'  ✅ {label}: {ok} ok, {fail} fail')
    return ok, fail

# ====== 1. SKU 标准表 (30 种 TEST_ SKU) ======
print('\n' + '=' * 60)
print('1. 生成 SKU 标准表测试数据 (30 条)')
print('=' * 60)

skus = []
for i in range(1, 31):
    # 不同类型：普通、RB800、高安全库存、无安全库存、0/空生产周期
    category = '成品'
    if i <= 5:
        dev_prefix, sku_type = '消控室值班助手', 'RB800设备'
    elif i <= 12:
        dev_prefix, sku_type = '无线压力监测终端', '传感器'
    elif i <= 18:
        dev_prefix, sku_type = '智能网关', '通信设备'
    elif i <= 24:
        dev_prefix, sku_type = 'NFC标签', '配件'
    else:
        dev_prefix, sku_type = '电源模块', '电源'

    sku_code = f'TEST_SKU-{i:03d}'
    dev_name = f'{dev_prefix}-{sku_type}-型号{i}'

    # 安全库存: 部分有值, 部分0, 部分空
    if i <= 10:
        safety_stock = str(random.choice([3, 5, 8, 10]))
    elif i <= 20:
        safety_stock = '0'
    else:
        safety_stock = ''

    # 标准生产周期: 部分正常, 部分0, 部分空
    if i <= 15:
        cycle = str(random.choice([7, 14, 21, 25, 30]))
    elif i <= 25:
        cycle = '0'
    else:
        cycle = ''

    # 是否RB800
    is_rb800 = 'RB800' in sku_code or i <= 5

    skus.append({
        '产品编码SKU': sku_code,
        '设备名称': dev_name,
        '设备型号': f'MODEL-{i:03d}',
        '设备单位': random.choice(['台', '个', '套']),
        '分类': category,
        '安全库存': safety_stock,
        '标准生产周期': cycle,
        '是否新产品': '否' if i > 3 else '是',
        '是否自研': '是' if i % 3 != 0 else '否',
        '是否外采': '是' if i % 3 == 0 else '否',
    })

write_records('tblVAGWeGHvmbFgJ', skus, 'SKU标准表')
time.sleep(0.5)

# ====== 2. 销售订单主表 (15 个 TEST_ 合同) ======
print('\n' + '=' * 60)
print('2. 生成销售订单主表测试数据 (15 条)')
print('=' * 60)

customers = [
    ('TEST_云盾智慧物联科技有限公司', 'TEST_城市商业综合体消控室远程值守一期'),
    ('TEST_鹏程安全技术工程有限公司', 'TEST_产业园区消控室物联网集中监控'),
    ('TEST_智安消防服务集团', 'TEST_集团8处厂区消控室远程联网托管'),
    ('TEST_恒达物业资产管理有限公司', 'TEST_高端写字楼消控室远程控制预警'),
    ('TEST_盛安科技消防安全有限公司', 'TEST_化工园区消控室远程遥控应急'),
    ('TEST_联创安全科技股份公司', 'TEST_地铁隧道消控室集中监控'),
    ('TEST_安泰消防技术服务公司', 'TEST_大型商场消控室联网改造'),
    ('TEST_瑞安智慧城市科技公司', 'TEST_医院消控室物联网平台建设'),
    ('TEST_华宇安全系统有限公司', 'TEST_学校消控室远程值守项目'),
    ('TEST_东方安防科技股份公司', 'TEST_数据中心消控室智能监控'),
    ('TEST_博创消防安全有限公司', 'TEST_机场航站楼消控室联网'),
    ('TEST_天卫安防技术有限公司', 'TEST_仓储物流消控室集中管理'),
    ('TEST_海安智慧消防科技公司', 'TEST_港口码头消控室远程控制'),
    ('TEST_中安消防工程有限公司', 'TEST_政府大楼消控室改造升级'),
    ('TEST_金盾安全系统集成公司', 'TEST_体育馆消控室物联网接入'),
]

main_orders = []
today = datetime(2026, 5, 15)
for i, (cust, proj) in enumerate(customers):
    cid = f'TEST_HT-{i+1:03d}'
    is_urgent = '是' if i in [0, 4, 8, 12] else '否'  # 4 个紧急
    is_exchange = '是' if i == 2 else '否'
    is_resend = '是' if i == 6 else '否'
    is_repair = '是' if i == 10 else '否'
    # 下单日期: 今天或前几天
    days_ago = random.choice([0, 0, 0, 1, 2, 3, 5])  # 更多今天的
    order_date = today - timedelta(days=days_ago)

    main_orders.append({
        '合同编号': cid,
        '客户名称': cust,
        '项目名称': proj,
        '下单日期': int(order_date.timestamp() * 1000),
        '订单来源': 'OA下单',
        '是否紧急订单': is_urgent,
        '是否换货订单': is_exchange,
        '是否补发订单': is_resend,
        '是否维修订单': is_repair,
        '优先级': '高' if is_urgent == '是' else '普通',
    })

write_records('tbl06oxGEdMNTEB8', main_orders, '销售订单主表')
time.sleep(0.5)

# ====== 3. 库存快照表 (对 30 个 TEST_ SKU, 3 天库存) ======
print('\n' + '=' * 60)
print('3. 生成库存快照表测试数据')
print('=' * 60)

inv_records = []
# 库存日期: 2026-05-13, 14, 15
inv_dates = [datetime(2026, 5, d) for d in [13, 14, 15]]
stock_scenarios = [
    # (库存数量, 待采购出库, 在途数量) -> 可用库存
    # scenarios 0-14: 充足
    (50, 5, 0), (100, 10, 5), (200, 20, 10), (80, 0, 0), (150, 30, 20),
    (60, 5, 5), (120, 10, 0), (90, 0, 10), (70, 5, 5), (110, 20, 10),
    (40, 0, 5), (130, 10, 15), (85, 5, 0), (95, 15, 10), (75, 0, 0),
    # scenarios 15-24: 预警 (可用库存 < 安全库存, 但 > 0)
    (5, 0, 0), (3, 1, 0), (8, 5, 0), (6, 0, 0), (4, 2, 0),
    (7, 3, 0), (2, 0, 0), (9, 7, 0), (1, 0, 0), (5, 2, 0),
    # scenarios 25-29: 缺货
    (0, 0, 0), (0, 5, 0), (2, 10, 0), (0, 0, 0), (1, 8, 0),
]

for sku_idx, sku in enumerate(skus):
    sku_code = sku['产品编码SKU']
    scenario = stock_scenarios[sku_idx % 30]
    stock_qty, pending, transit = scenario

    for inv_date in inv_dates:
        # 添加一些日期间的变化（旧日期稍少，新日期稍多）
        day_factor = 1.0 if inv_date == inv_dates[-1] else (0.95 if inv_date == inv_dates[-2] else 0.9)
        q = max(0, int(stock_qty * day_factor + random.randint(-2, 2)))
        p = max(0, int(pending * day_factor))
        t = max(0, int(transit * day_factor))

        inv_records.append({
            '库存日期': int(inv_date.timestamp() * 1000),
            'SKU': sku_code,
            '库存数量': str(q),
            '待采购出库': str(p),
            '在途数量': str(t),
            '国网设备名称': sku['设备名称'],
            '国网设备型号': sku['设备型号'],
            '数据来源': '手动录入',
            '导入批次号': int(inv_date.timestamp() * 1000),
        })

write_records('tblFZNdEwW50izjh', inv_records, '库存快照表')
time.sleep(0.5)

# ====== 4. 销售订单明细表 (约 60 条订单行) ======
print('\n' + '=' * 60)
print('4. 生成销售订单明细表测试数据 (~60 条)')
print('=' * 60)

items = []
# 为每个合同分配 SKU
for i, (cust, proj) in enumerate(customers):
    cid = f'TEST_HT-{i+1:03d}'
    # 每个合同 2-6 个 SKU
    num_skus = random.choice([2, 2, 3, 3, 4, 4, 5, 6])
    # 给合同分配 SKU index
    assigned_skus = []
    for j in range(num_skus):
        # 分配不同类型的 SKU
        if i in [0, 1]:  # RB800 远程控制项目
            sku_idx = random.choice(range(0, 5))  # SKU 1-5 是 RB800
        elif i in [2, 6, 10]:  # 特殊订单 (换货/补发/维修)
            sku_idx = random.choice(range(0, 20))  # 混合
        elif i in [4, 8, 12]:  # 紧急订单
            sku_idx = random.choice(range(0, 25))
        else:
            sku_idx = random.choice(range(0, 30))
        if sku_idx not in assigned_skus:
            assigned_skus.append(sku_idx)

    for seq, sku_idx in enumerate(assigned_skus):
        sku = skus[sku_idx]
        qty = random.choice([1, 2, 3, 5, 8, 10, 15, 20])
        spec = f'TEST-{sku["设备型号"]}-定制' if sku_idx % 5 == 0 else f'TEST-{sku["设备型号"]}'

        items.append({
            '合同编号': cid,
            'SKU编码': sku['产品编码SKU'],
            '合同数量': qty,
            '产品名称': sku['设备名称'],
            '规格': spec,
            '下单时间': int((today - timedelta(days=random.choice([0, 0, 0, 1, 2, 3, 5]))).timestamp() * 1000),
            '是否RB800': '是' if sku_idx < 5 else '否',
        })

# 添加一个已确认的测试合同 (模拟已人工确认)
confirmed_cid = 'TEST_HT-CONFIRMED'
main_orders.append({
    '合同编号': confirmed_cid,
    '客户名称': 'TEST_已确认测试客户',
    '项目名称': 'TEST_已确认合同项目',
    '下单日期': int(today.timestamp() * 1000),
    '订单来源': 'OA下单',
    '是否紧急订单': '否',
    '是否换货订单': '否',
    '是否补发订单': '否',
    '是否维修订单': '否',
    '优先级': '普通',
})
# 该合同只有 2 个 SKU
for j in range(2):
    sku = skus[j]
    items.append({
        '合同编号': confirmed_cid,
        'SKU编码': sku['产品编码SKU'],
        '合同数量': random.choice([3, 5]),
        '产品名称': sku['设备名称'],
        '规格': sku['设备型号'],
        '下单时间': int(today.timestamp() * 1000),
        '是否RB800': '否',
    })

write_records('tbl06oxGEdMNTEB8', [main_orders[-1]], '销售订单主表(已确认合同)')
write_records('tblJn5iP6imjzE8h', items, '销售订单明细表')
time.sleep(0.5)

print(f'\n{"=" * 60}')
print(f'测试数据生成完毕')
print(f'  SKU标准表: {len(skus)} 条')
print(f'  销售订单主表: {len(main_orders)} 条')
print(f'  库存快照表: {len(inv_records)} 条')
print(f'  销售订单明细表: {len(items)} 条')
print(f'  总计: {len(skus)+len(main_orders)+len(inv_records)+len(items)} 条')
print(f'  合同数: {len(main_orders)}')
print(f'  覆盖场景: 普通/紧急/换货/补发/维修/RB800')
print(f'  库存场景: 充足/预警/缺货')
print(f'{"=" * 60}')
