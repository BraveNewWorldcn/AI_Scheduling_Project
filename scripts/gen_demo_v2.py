"""
生成面试Demo数据 v2 — 多产品合同版本
设计目标：
  80合同 × 平均3产品 = ~240条订单行
  15类库存产品，战略分布（健康/预警/临界）
  看板数据好看、通知有缺货场景、财务有梯度
"""
import openpyxl
import random
from datetime import datetime, timedelta

random.seed(42)

# ============================================================
# 1. 基础数据
# ============================================================
COMPANIES = [
    "苏州智造安防科技有限公司", "江苏恒通智能设备有限公司",
    "昆山明泰消防工程有限公司", "无锡瑞联科技发展有限公司",
    "常州中安智能科技有限公司", "南京博创物联技术有限公司",
    "苏州工业园区艾科智能有限公司", "上海锦程安全技术有限公司",
    "杭州启明智联科技有限公司", "南通华远智慧消防有限公司",
    "苏州高新智联安防有限公司", "太仓瑞驰科技有限公司",
]

AGENTS = ["智造安防代理", "恒通智能代理", "瑞联科技代理", "中安智能代理"]

PROJECTS = [
    "太湖新城智慧社区一期", "苏州中心智能化改造", "昆山花桥产业园消防升级",
    "常熟高新区人才公寓", "张家港保税区物流中心", "吴江太湖科创园",
    "工业园区生物医药基地", "相城高铁新城商务中心", "姑苏区历史文化街区安防",
    "吴中区智能制造产业园", "新区科技城研发大楼", "太仓港综合保税区",
    "苏州湾文化中心", "阳澄湖数字文创园", "独墅湖高教创新区",
]

# 来源/商务/结算方式
SOURCES = ["代理商报备", "直销", "招标", "框架协议", "战略合作"]
BIZ = ["张三", "李四", "王五", "赵六", "钱七"]
SETTLE = ["标准品硬件", "标准品硬件", "标准品硬件", "Iot产品", "Iot产品", "技术服务"]

# ============================================================
# 2. 产品目录 — 8个核心产品 + 规格 + SKU + 库存基线
# ============================================================
PRODUCTS = [
    # name, spec, sku, inventory_qty, unit_price, popularity(1-10), stock_level
    ("消控室值班助手-主键盘",  "A116-MAIN",   "SKU-001-MAIN",  85,  4800, 10, "高"),
    ("消控室值班助手-分机",    "A116-SUB",    "SKU-001-SUB",   72,  2800, 8,  "高"),
    ("智能烟感探测器-GS210",   "GS-NB-IoT",   "SKU-002-GS210", 210, 580,  9,  "高"),
    ("智慧消防网关-GT300",     "GT-4G-V2",    "SKU-003-GT300", 45,  3200, 7,  "中"),
    ("消防水压监测终端-WP400", "WP-LoRa",     "SKU-004-WP400", 55,  1800, 5,  "中"),
    ("电气火灾监控器-EF100",   "EF-485-V3",   "SKU-005-EF100", 30,  2200, 4,  "中"),
    ("智能疏散指示系统-EL500", "EL-BLE-V2",   "SKU-006-EL500", 120, 950,  3,  "高"),
    ("4G通信模块",             "4G-MOD-V3",   "SKU-007-4GMOD", 320, 280,  6,  "高"),
]

# 产品亲和组合: 网关经常搭配烟感和模块，消控主键搭配分机
AFFINITY_GROUPS = [
    [0, 2, 7],  # 主键盘+烟感+4G模块 (消控室方案)
    [0, 1],     # 主键盘+分机
    [2, 3, 7],  # 烟感+网关+4G模块 (烟感联网)
    [3, 4],     # 网关+水压监测 (智慧消防)
    [4, 5],     # 水压+电气火灾监控 (消防水+电)
    [6, 7],     # 疏散指示+4G模块
    [0, 2, 3, 7], # 全套消控方案
]

# ============================================================
# 3. 生成合同 — 80个合同，不同产品数
# ============================================================
# 合同产品数分布: 1品/2品/3品/4品/5品
CONTRACT_SIZES = [1]*10 + [2]*18 + [3]*26 + [4]*16 + [5]*7 + [6]*3
random.shuffle(CONTRACT_SIZES)

BASE_DATE = datetime(2026, 6, 21)

contracts = []  # list of {合同编号, 日期, 客户, 代理商, 项目, 来源, 商务, 是否紧急}
order_lines = []  # list of {合同编号, 产品名, 规格, SKU, 数量, 结算方式}

contract_id = 1
for size in CONTRACT_SIZES:
    cno = f"DEMO-{7000+contract_id}-{random.randint(1,99):02d}"
    days_ago = random.choices([0,1,2,3,5,7,10,14,20,28],
                              weights=[3,4,5,5,4,3,3,2,1,1])[0]
    sign_date = (BASE_DATE - timedelta(days=days_ago)).strftime("%Y-%m-%d")
    company = random.choice(COMPANIES)
    agent = random.choice(AGENTS)
    project = random.choice(PROJECTS)
    source = random.choice(SOURCES)
    biz_person = random.choice(BIZ)
    is_urgent = "是" if random.random() < 0.12 else "否"  # ~12%紧急

    contract = {
        "合同编号": cno, "签订日期": sign_date, "客户名称": company,
        "代理商": agent, "所属项目": project, "订单来源": source,
        "商务": biz_person, "是否紧急订单": is_urgent,
    }
    contracts.append(contract)

    # 选产品: 优先从亲和组里挑
    picked = set()
    if size >= 2 and random.random() < 0.7:
        group = random.choice(AFFINITY_GROUPS)
        n_from_group = min(size, len(group))
        picked = set(random.sample(group, n_from_group))
    # 补足到size
    while len(picked) < size:
        candidates = [i for i in range(len(PRODUCTS)) if i not in picked]
        if not candidates:
            break
        # 权重=热度
        weights = [PRODUCTS[i][5] for i in candidates]
        picked.add(random.choices(candidates, weights=weights, k=1)[0])

    for pi in picked:
        p = PRODUCTS[pi]
        inv_qty = p[3]
        # 数量分布: 75%合理(库存覆盖), 20%紧张, 5%超库存(真实缺货)
        roll = random.random()
        if roll < 0.75:
            qty = random.randint(1, max(1, inv_qty // 4))         # 小批量，轻松满足
        elif roll < 0.95:
            qty = random.randint(inv_qty // 4, inv_qty // 2)       # 中等批量
        else:
            qty = random.randint(inv_qty, inv_qty * 2)             # 超库存 → 真实缺货

        order_lines.append({
            "合同编号": cno,
            "产品名称": p[0],
            "规格": p[1],
            "SKU编码": p[2],
            "数量": qty,
            "结算方式": random.choice(SETTLE),
            "签订日期": sign_date,
            "所属项目": project,
        })

    contract_id += 1

print(f"生成了 {len(contracts)} 个合同, {len(order_lines)} 条订单行")
print(f"平均每合同 {len(order_lines)/len(contracts):.1f} 个产品")

# ============================================================
# 4. 生成库存数据 — 21种唯一产品
# ============================================================
INVENTORY_BASE = [
    # 名称, 型号, 品牌, 类型, 库存数, 累计销售, 待采购, 锁定库存, 方案, 预警, 备货周期
    ("NFC卡", "待定", "建和智能", "外购复销类", 800, 24605, 200, 150, "门禁方案", 100, 5),
    ("物联网卡（电信4G）", "100M/月/3年期/新开卡", "智造科技", "外购复销类", 1200, 1514, 200, 180, "报警主机联网方案", 50, 10),
    ("无线压力监测终端", "AC-SYL211-4G-16", "奥创", "Iot产品", 200, 4516, 60, 80, "水系统监测方案", 30, 5),
    ("定制化硬件", "定制化硬件", "智造科技", "外购复销类", 120, 1237, 20, 30, None, 15, 7),
    ("消控室值班助手-主键盘", "A116-MAIN", "高新投三江", "Iot产品", 85, 2340, 35, 40, "消控室联网方案", 20, 7),
    ("消控室值班助手-分机", "A116-SUB", "高新投三江", "Iot产品", 72, 1890, 28, 35, "消控室联网方案", 15, 7),
    ("智能烟感探测器-GS210", "GS-NB-IoT", "智造科技", "Iot产品", 210, 4520, 50, 60, "烟感监测方案", 30, 5),
    ("电气火灾监控器-EF100", "EF-485-V3", "智造科技", "Iot产品", 30, 1120, 15, 20, "电气监控方案", 10, 10),
    ("消防水压监测终端-WP400", "WP-LoRa", "奥创", "Iot产品", 55, 890, 15, 20, "水系统监测方案", 12, 5),
    ("智慧消防网关-GT300", "GT-4G-V2", "智造科技", "Iot产品", 45, 2100, 20, 25, "网关联网方案", 15, 7),
    ("智能疏散指示系统-EL500", "EL-BLE-V2", "智造科技", "Iot产品", 120, 780, 25, 35, "疏散指示方案", 25, 5),
    ("可燃气体探测器-GD600", "GD-ZigBee", "奥创", "Iot产品", 100, 1560, 20, 30, "燃气监测方案", 20, 5),
    ("4G通信模块", "4G-MOD-V3", "智造科技", "Iot产品", 320, 5680, 80, 100, "通用通信方案", 50, 5),
    ("物联网卡（移动4G）", "300M/月/3年期", "智造科技", "外购复销类", 500, 890, 60, 80, "备用通信方案", 60, 10),
    ("电源管理模块", "PM-12V-5A", "智造科技", "Iot产品", 200, 3200, 50, 60, "通用电源方案", 40, 7),
    ("智能温湿度传感器", "TH-NB-IoT", "智造科技", "传感器类", 180, 650, 30, 40, "环境监测方案", 25, 5),
    ("消防广播系统-功放", "PA-500W", "奥创", "终端类", 25, 340, 10, 15, "广播方案", 8, 10),
    ("消防电话系统-主机", "FP-64L", "智造科技", "终端类", 15, 210, 8, 12, "电话方案", 5, 14),
    ("手动报警按钮", "MCP-NB", "智造科技", "传感器类", 350, 1200, 80, 100, "报警方案", 50, 5),
    ("声光报警器", "SA-NB-IoT", "奥创", "终端类", 280, 980, 50, 60, "报警方案", 40, 5),
    ("输入输出模块", "IOM-485", "智造科技", "通讯类", 420, 2100, 60, 80, "通用方案", 60, 5),
]

inv_rows = []
for idx, t in enumerate(INVENTORY_BASE):
    name, model, brand, ptype, stock, sales, purchase, locked, sol, warn, lead = t
    remaining = max(0, stock - locked)  # 可用库存
    inv_rows.append([idx + 1, name, model, brand, ptype, stock, sales, purchase,
                     locked, remaining, sol or "", warn, lead])

# 统计
crit = sum(1 for r in inv_rows if r[5] < 15)
warn = sum(1 for r in inv_rows if 15 <= r[5] < 60)
healthy = sum(1 for r in inv_rows if r[5] >= 60)
print(f"库存: {healthy}健康 / {warn}预警 / {crit}临界 "
      f"({100*healthy/len(inv_rows):.0f}%/{100*warn/len(inv_rows):.0f}%/{100*crit/len(inv_rows):.0f}%)")
total_inv = len(inv_rows)

# ============================================================
# 5. 写入Excel
# ============================================================
wb = openpyxl.Workbook()

# Sheet 1: 订单表
ws1 = wb.active
ws1.title = "订单表"
headers1 = ["序号", "合同编号", "签订日期", "客户名称", "代理商", "所属项目",
            "产品名称", "结算方式", "明细规格", "数量", "备注"]
ws1.append(headers1)
for i, ol in enumerate(order_lines):
    ws1.append([
        i + 1, ol["合同编号"], ol["签订日期"],
        contracts[[c["合同编号"] for c in contracts].index(ol["合同编号"])]["客户名称"],
        contracts[[c["合同编号"] for c in contracts].index(ol["合同编号"])]["代理商"],
        ol["所属项目"], ol["产品名称"], ol["结算方式"],
        ol["规格"], ol["数量"], ""
    ])

# Sheet 2: 库存表
ws2 = wb.create_sheet("库存表")
headers2 = ["序号", "国网设备名称", "国网设备型号", "销售品牌", "产品类型",
            "库存数量", "累计销售", "待采购出库", "锁定总库存", "锁定剩余库存",
            "硬件方案", "预警数量", "备货周期(天)", "保质期（年）"]
ws2.append(headers2)
for r in inv_rows:
    ws2.append([r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                r[8], r[9], r[10] or "", r[11], r[12], None])

OUT = "/Users/xiaowang/AI/AI/AI_Scheduling_Project/测试集_演示_v2.xlsx"
wb.save(OUT)
print(f"\n保存: {OUT}")

# 数据亮点
unique_contracts = len(set(ol["合同编号"] for ol in order_lines))
product_counts = {}
for ol in order_lines:
    pn = ol["产品名称"]
    product_counts[pn] = product_counts.get(pn, 0) + 1
print(f"\n=== 数据亮点 ===")
print(f"合同数: {unique_contracts}")
print(f"订单行: {len(order_lines)}")
print(f"公司数: {len(set(c['客户名称'] for c in contracts))}")
print(f"紧急订单: {sum(1 for c in contracts if c['是否紧急订单']=='是')}")
print(f"\n产品覆盖:")
for pn, ct in sorted(product_counts.items(), key=lambda x: -x[1]):
    print(f"  {pn}: {ct}次")
shortage = sum(1 for ol in order_lines if ol["数量"] > 50)
print(f"\n潜在缺货行: {shortage}/{len(order_lines)} ({100*shortage/len(order_lines):.0f}%)")
