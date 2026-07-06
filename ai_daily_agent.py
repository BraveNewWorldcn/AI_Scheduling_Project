#!/usr/bin/env python3
"""
AI排单日报机器人
---------------
每日自动读取飞书《AI排单日报》多维表格最新数据，
调用 DeepSeek 生成运营日报，创建飞书文档并推送到群。

运行方式：
    python ai_daily_agent.py

环境变量（必需）：
    DEEPSEEK_API_KEY       DeepSeek API 密钥
    FEISHU_CHAT_ID         飞书群 chat_id（oc_xxx）

环境变量（可选，已有默认值）：
    BITABLE_APP_TOKEN      飞书多维表格 app token
    TABLE_ID_DAILY_REPORT  日报表 ID
    FEISHU_TASKLIST_ID     飞书任务清单 ID（不设置则自动发现）

飞书应用权限（需在开发者后台开启）：
    确保应用 cli_a96c5d017d3a1cbb 已开启以下权限：
    - docx:document 或 docx:document:create  （创建文档）
    - im:message                              （发送消息）
    - task:task:write 或 task:task:writeonly  （创建任务）
    权限配置地址：https://open.feishu.cn/app/cli_a96c5d017d3a1cbb/auth
"""

from __future__ import annotations

import subprocess
import json
import os
import sys
import re
import csv
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

# ===== 自动发现 lark-cli 路径 =====
def _find_lark_cli() -> str:
    """查找 lark-cli 可执行文件路径。"""
    # 优先检查常见的 npm 全局安装路径
    candidates = [
        os.path.expandvars(r"%APPDATA%\npm\lark-cli.cmd"),
        os.path.expandvars(r"%APPDATA%\npm\lark-cli"),
        "/usr/local/bin/lark-cli",
        "/opt/homebrew/bin/lark-cli",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    # 回退到 PATH 查找
    for p in os.environ.get("PATH", "").split(os.pathsep):
        for name in ("lark-cli.cmd", "lark-cli"):
            full = os.path.join(p, name)
            if os.path.isfile(full):
                return full
    return "lark-cli"  # 最后的尝试

LARK_CLI = _find_lark_cli()

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

# ===== 配置 =====
APP_ID = os.getenv("FEISHU_APP_ID", "")
APP_SECRET = os.getenv("FEISHU_APP_SECRET")
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_ID_DAILY_REPORT = os.getenv("TABLE_ID_DAILY_REPORT", "")
CHAT_ID = os.getenv("FEISHU_CHAT_ID", "")
TASKLIST_ID = os.getenv("FEISHU_TASKLIST_ID", "")
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

CST = timezone(timedelta(hours=8))

if not DEEPSEEK_API_KEY:
    sys.exit("[X] 请先设置环境变量 DEEPSEEK_API_KEY")
if not CHAT_ID:
    sys.exit("[X] 请先设置环境变量 FEISHU_CHAT_ID（飞书群 chat_id，以 oc_ 开头）")
if not APP_SECRET:
    sys.exit("[X] 请先设置环境变量 FEISHU_APP_SECRET")

# DeepSeek client (OpenAI-compatible)
from openai import OpenAI
deepseek = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")

# httpx for bot-identity API calls (bitable read) + webhook cards
import httpx

# 自动加载 .env（优先 python-dotenv，失败则回退到手动解析）
try:
    from dotenv import load_dotenv
    load_dotenv(override=True)
except ImportError:
    pass  # _load_dotenv() already called above


def _get_bot_token() -> str:
    """获取飞书 tenant_access_token（bot 身份），用于读表操作。"""
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=30.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取飞书Token失败: {data}")
    return data["tenant_access_token"]


# ===== 工具函数 =====

def _parse_feishu_value(value: Any) -> str:
    """将飞书 API 返回的字段值转为纯文本。兼容富文本数组、数字、字符串。"""
    if value is None:
        return ""
    if isinstance(value, list):
        texts = []
        for item in value:
            if isinstance(item, dict) and "text" in item:
                texts.append(item["text"].strip())
            else:
                texts.append(str(item).strip())
        return ", ".join(texts)
    if isinstance(value, dict):
        for k in ("text", "number", "phone", "email", "url"):
            if k in value:
                return str(value[k]).strip()
        return str(value).strip()
    if isinstance(value, (int, float)):
        return str(value)
    return str(value).strip()


def _parse_date_millis(value: Any) -> str:
    """将飞书日期字段（毫秒时间戳或字符串）转为 YYYY-MM-DD。"""
    if value is None:
        return ""
    if isinstance(value, (int, float)) and value > 0:
        try:
            if value >= 10_000_000_000:
                ts = value / 1000
            else:
                ts = value
            return datetime.fromtimestamp(ts, tz=CST).strftime("%Y-%m-%d")
        except Exception:
            pass
    s = _parse_feishu_value(value)
    return s[:10] if len(s) >= 10 else s


def run_lark(*args: str, stdin_text: Optional[str] = None) -> dict:
    """调用 lark-cli，返回解析后的 JSON（bot 身份）。"""
    cmd = [LARK_CLI] + list(args)
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            encoding="utf-8",
            timeout=120,
            input=stdin_text,
        )
    except subprocess.TimeoutExpired:
        raise Exception(f"lark-cli 执行超时: {' '.join(cmd)}")

    if result.returncode != 0:
        err = (result.stderr or "").strip() or (result.stdout or "").strip()
        raise Exception(f"lark-cli 失败 [{result.returncode}]: {err[:800]}")

    stdout = (result.stdout or "").strip()
    if not stdout:
        return {}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        raise Exception(f"lark-cli 返回非 JSON: {stdout[:500]}")


# ===== 步骤 1：读取最新记录 =====

def fetch_latest_record() -> Dict[str, str]:
    """读取 AI排单日报表 最新一行（httpx 直调 API，bot 身份）。"""
    token = _get_bot_token()
    url = (f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}"
           f"/tables/{TABLE_ID_DAILY_REPORT}/records/search")
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"page_size": 1, "sort": [{"field_name": "排单运行时间", "desc": True}]}

    resp = httpx.post(url, headers=headers, json=payload, timeout=30.0)
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"读取日报表API失败: code={data.get('code')}, msg={data.get('msg')}")

    items = data.get("data", {}).get("items", [])
    if not items:
        raise ValueError("日报表为空，没有找到任何排单日报数据。")

    fields = items[0].get("fields", {})

    # 所有字段使用安全提取 + 富文本解析
    record = {
        "日期":               _parse_date_millis(fields.get("日期")),
        "排单运行时间":       _parse_feishu_value(fields.get("排单运行时间")),
        "排单批次号":         _parse_feishu_value(fields.get("排单批次号")),
        "订单总数":           _parse_feishu_value(fields.get("订单总数")),
        "今日新增订单":       _parse_feishu_value(fields.get("今日新增订单")),
        "今日紧急订单":       _parse_feishu_value(fields.get("今日紧急订单")),
        "未排单订单数":       _parse_feishu_value(fields.get("未排单订单数")),
        "已排单订单数":       _parse_feishu_value(fields.get("已排单订单数")),
        "人工已确认排单数":   _parse_feishu_value(fields.get("人工已确认排单数")),
        "今日应发货订单数":   _parse_feishu_value(fields.get("今日应发货订单数")),
        "今日可发货订单数":   _parse_feishu_value(fields.get("今日可发货订单数")),
        "今日预计延迟订单数": _parse_feishu_value(fields.get("今日预计延迟订单数")),
        "SKU总数":            _parse_feishu_value(fields.get("SKU总数")),
        "库存充足SKU数":      _parse_feishu_value(fields.get("库存充足SKU数")),
        "库存预警SKU数":      _parse_feishu_value(fields.get("库存预警SKU数")),
        "库存缺货SKU数":      _parse_feishu_value(fields.get("库存缺货SKU数")),
        "库存缺货统计":      _parse_feishu_value(fields.get("库存缺货统计")),
        "最紧缺SKU缺口":      _parse_feishu_value(fields.get("最紧缺SKU缺口")),
        "最长交期订单":       _parse_feishu_value(fields.get("最长交期订单")),
        "最晚发货日期":       _parse_feishu_value(fields.get("最晚发货日期")),
        "排单占库存比例":     _parse_feishu_value(fields.get("排单占库存比例")),
        "未来3天应发货订单数": _parse_feishu_value(fields.get("未来3天应发货订单数")),
        "未来3天缺货订单数":  _parse_feishu_value(fields.get("未来3天缺货订单数")),
        "未来7天应发货订单数": _parse_feishu_value(fields.get("未来7天应发货订单数")),
        "缺货SKU明细":        _parse_feishu_value(fields.get("缺货SKU明细")) or "[]",
    }

    # 补默认值：飞书空单元格可能字段不存在
    for key, default in [
        ("库存缺货统计", "无"),
        ("最紧缺SKU缺口", "0"),
        ("今日预计延迟订单数", "0"),
    ]:
        if not record.get(key):
            record[key] = default

    return record


# ===== 辅助：加载 SKU 主数据（设备名称 + 设备型号 + 安全库存） =====

_sku_master_cache: Optional[Dict[str, Dict[str, str]]] = None

def _load_sku_master() -> Dict[str, Dict[str, str]]:
    """读取 SKU 标准表，返回 {sku_code: {name, model, safety}}。结果缓存。"""
    global _sku_master_cache
    if _sku_master_cache is not None:
        return _sku_master_cache

    token = _get_bot_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    base = "https://open.feishu.cn/open-apis/bitable/v1/apps"
    sku_table_id = os.getenv("TABLE_ID_SKU", "tblVAGWeGHvmbFgJ")

    sku_map: Dict[str, Dict[str, str]] = {}
    page_token = None
    while True:
        params = {"page_size": 200}
        if page_token:
            params["page_token"] = page_token
        resp = httpx.get(f"{base}/{APP_TOKEN}/tables/{sku_table_id}/records",
                         headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            break
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            code = _parse_feishu_value(f.get("产品编码SKU"))
            if not code:
                continue
            sku_map[code] = {
                "name": _parse_feishu_value(f.get("设备名称")),
                "model": _parse_feishu_value(f.get("设备型号")),
                "safety": int(_parse_feishu_value(f.get("安全库存")) or "0"),
            }
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")

    _sku_master_cache = sku_map
    return sku_map


# ===== 辅助：构建缺货明细表 =====

def _build_shortage_table(record: Dict[str, str]) -> str:
    """从销售订单明细表实时读取缺货数据，生成 Markdown 表格。"""
    token = _get_bot_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    base = "https://open.feishu.cn/open-apis/bitable/v1/apps"
    items_table_id = os.getenv("TABLE_ID_ITEMS", "tblJn5iP6imjzE8h")

    # 读取明细表中库存状态 == "缺货" 的行（用 search API 不够精确，改用全量 + filter）
    gap_by_sku: Dict[str, float] = {}
    page_token = None
    while True:
        params = {"page_size": 200}
        if page_token:
            params["page_token"] = page_token
        resp = httpx.get(f"{base}/{APP_TOKEN}/tables/{items_table_id}/records",
                         headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            break
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            status = _parse_feishu_value(f.get("库存状态"))
            if status != "缺货":
                continue
            sku = _parse_feishu_value(f.get("SKU编码"))
            gap = float(_parse_feishu_value(f.get("缺口数量")) or "0")
            if sku and gap > 0:
                gap_by_sku[sku] = gap_by_sku.get(sku, 0) + gap
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")

    if not gap_by_sku:
        return ""

    sorted_items = sorted(gap_by_sku.items(), key=lambda x: x[1], reverse=True)
    sku_map = _load_sku_master()
    table = "| 产品名称 | 设备型号 | 订单缺货数量 |\n|---------|--------|------------|\n"
    for sku_code, gap in sorted_items:
        info = sku_map.get(sku_code, {})
        name = info.get("name") or sku_code
        model = info.get("model", "-")
        table += f"| {name} | {model} | {int(gap)} |\n"
    return table


# ===== 辅助：构建常规备货建议表 =====

def _build_restock_table(record: Dict[str, str]) -> str:
    """读取库存快照表，计算 安全库存 - 库存数量 = 常规备货数。返回 Markdown 表格。"""
    sku_map = _load_sku_master()
    token = _get_bot_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    base = "https://open.feishu.cn/open-apis/bitable/v1/apps"
    inv_table_id = os.getenv("TABLE_ID_INV", "tblFZNdEwW50izjh")

    # 读取库存快照表，取最新日期的库存
    inv_stock: Dict[str, int] = {}
    latest_date = 0
    page_token = None
    while True:
        params = {"page_size": 200}
        if page_token:
            params["page_token"] = page_token
        resp = httpx.get(f"{base}/{APP_TOKEN}/tables/{inv_table_id}/records",
                         headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            break
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            inv_date_raw = f.get("库存日期")
            inv_date = 0
            if isinstance(inv_date_raw, (int, float)) and inv_date_raw > 0:
                inv_date = int(inv_date_raw)
            if inv_date < latest_date:
                continue
            if inv_date > latest_date:
                latest_date = inv_date
                inv_stock.clear()
            sku = _parse_feishu_value(f.get("SKU"))
            qty = int(_parse_feishu_value(f.get("库存数量")) or "0")
            if sku:
                inv_stock[sku] = max(inv_stock.get(sku, 0), qty)
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")

    # 计算备货需求：安全库存 - 当前库存 > 0 才需要备货
    rows = []
    for code, info in sku_map.items():
        safety = info["safety"]
        if safety <= 0:
            continue
        stock = inv_stock.get(code, 0)
        need = safety - stock
        if need > 0:
            rows.append((info["name"] or code, info.get("model", "-"), need, safety, stock))

    if not rows:
        return ""
    rows.sort(key=lambda x: x[2], reverse=True)
    table = "| 产品名称 | 设备型号 | 常规备货数量 |\n|---------|--------|------------|\n"
    for name, model, need, safety, stock in rows:
        table += f"| {name} | {model} | {need} |\n"
    return table


# ===== 步骤 2：DeepSeek 生成日报 =====

def generate_daily_report(record: Dict[str, str]) -> tuple:
    """调用 DeepSeek 生成完整日报和群消息。返回 (full_report, group_msg)。"""

    batch = record.get("排单批次号", "")

    # 预计算关键比率
    total_sku = int(record.get("SKU总数", "0") or "0")
    shortage_sku = int(record.get("库存缺货SKU数", "0") or "0")
    warning_sku = int(record.get("库存预警SKU数", "0") or "0")
    sufficient_sku = int(record.get("库存充足SKU数", "0") or "0")
    delay = int(record.get("今日预计延迟订单数", "0") or "0")
    new_orders = int(record.get("今日新增订单", "0") or "0")
    future_short = int(record.get("未来3天缺货订单数", "0") or "0")
    can_ship = int(record.get("今日可发货订单数", "0") or "0")
    should_ship = int(record.get("今日应发货订单数", "0") or "0")
    emergency = record.get("今日紧急订单", "0") or "0"

    shortage_pct = round(shortage_sku / total_sku * 100, 1) if total_sku > 0 else 0
    health_rate = round(sufficient_sku / total_sku * 100, 1) if total_sku > 0 else 0
    delivery_rate = round(can_ship / should_ship * 100, 1) if should_ship > 0 else 100

    # 注入预计算数据
    record["_库存健康率"] = f"{health_rate}%"
    record["_缺货率"] = f"{shortage_pct}%"
    record["_交付率"] = f"{delivery_rate}%"

    prompt = f"""你是一名资深的供应链运营总监（15年经验）。你的任务是生成一份给高管和业务团队的飞书运营日报。

【铁律】
1. 只使用传入数据，严禁编造数字。某项为 0 或空时，描述为"无"或"正常"。
2. 直接输出日报正文，不要"好的"、"以下是日报"等废话。
3. 语气果断、锐利，像真实高管在暴露问题。正常时说"平稳"，异常时直接警告。
4. 每个数字必须带单位（单/个/%），预计算的比率数据已提供，直接引用。

【分级规则】
🚨 红色警报（文档+群消息都要突出）：
  - 今日预计延迟订单 > 0
  - 库存缺货SKU数 ≥ 5
  - 未来3天缺货订单 > 3
⚠️ 黄色关注：
  - 库存预警SKU数 ≥ 库存充足SKU数
  - 排单占库存比例 > 0.5
✅ 正常：以上均不满足时直接说"各项指标正常"。

【输出结构 - 文档版】
📊 **供应链 AI 运营日报 ({datetime.now(CST).strftime('%m/%d')})**

🎯 **一、大盘速览**
- 总订单 {record.get('订单总数',0)} 单，今日新增 {new_orders} 单（紧急 {emergency}）。
- 给出定性判断（单量平稳/上升/激增）。

🚚 **二、交付诊断**
- 应发 {should_ship} → 可发 {can_ship}，交付率 {delivery_rate}%。
- 延迟 {delay} 单。{ "🚨 存在交付延迟，需立即介入！" if delay > 0 else "交付链路通畅。" }

📦 **三、物料风险**
- 库存健康率 {health_rate}%（充足 {sufficient_sku} / 预警 {warning_sku} / 缺货 {shortage_sku}）。
- { "当前物料充足。" if shortage_sku == 0 else f"缺货率 {shortage_pct}%，最紧缺为「{(record.get('库存缺货统计','?') or '?').split(chr(10))[0]}」。建议采购优先补货。" }

🔮 **四、未来 72h 预测**
- 未来3天应发 {record.get('未来3天应发货订单数','0')} 单，缺货风险 {future_short} 单。
- { "暂无交付风险。" if future_short == 0 else f"⚠️ {future_short} 单面临缺货风险，建议提前锁定库存。" }
- 最晚发货日期：{record.get('最晚发货日期','-')}

===原始数据===
{json.dumps(record, ensure_ascii=False, indent=2)}

在输出的末尾，另起一行写 [GROUP_MSG_START]，然后紧接着输出群消息版本。格式如下：

（此处为完整日报正文 Markdown，200-350字）
[GROUP_MSG_START]
（此处为群消息摘要，100-200字，用 \\n\\n 分段，粗体突出关键数字和异常）"""

    try:
        response = deepseek.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=4096,
            timeout=90,
        )
        text = response.choices[0].message.content.strip()
    except Exception as e:
        raise Exception(f"DeepSeek API 调用失败: {e}")

    marker = "[GROUP_MSG_START]"
    if marker in text:
        parts = text.split(marker, 1)
        full_report = parts[0].strip()
        group_msg = parts[1].strip() if len(parts) > 1 else ""
    else:
        full_report = text
        group_msg = ""

    # 容错：full 为空但 group_msg 有内容 → 交换
    if not full_report and group_msg:
        full_report = group_msg
        group_msg = ""

    # 容错：group_msg 为空 → 从 full 截取前 200 字
    if not group_msg and full_report:
        group_msg = full_report[:200]

    return full_report, group_msg


# ===== 步骤 3：创建飞书文档 =====

def create_feishu_doc(title: str, content: str,
                       shortage_table: str = "",
                       restock_table: str = "") -> str:
    """在飞书创建文档，返回文档 URL。"""
    now_str = datetime.now(CST).strftime("%Y-%m-%d %H:%M:%S")
    doc_parts = [f"""# {title}

> 生成时间：{now_str} | 数据来源：AI排单系统 | 机器人自动生成

---

{content}
"""]

    if shortage_table:
        doc_parts.append(f"""
---

## 📦 缺货明细

{shortage_table}
""")

    if restock_table:
        doc_parts.append(f"""
---

## 🛒 常规备货建议

> 计算规则：安全库存（SKU标准表） - 当前库存（最新快照） = 建议备货数

{restock_table}
""")

    doc_content = "\n".join(doc_parts)
    doc_content += "\n\n---\n\n*本文档由供应链AI助手自动生成*"

    try:
        data = run_lark(
            "docs", "+create",
            "--api-version", "v2",
            "--doc-format", "markdown",
            "--content", "-",
            stdin_text=doc_content,
        )
    except Exception as e:
        err = str(e)
        if "99991672" in err or "permission" in err.lower():
            raise Exception(
                f"创建飞书文档失败：应用缺少 docx:document 权限。"
                f"请在飞书开发者后台开启后重试。{e}"
            )
        raise Exception(f"创建飞书文档失败: {e}")

    # 尝试从响应中提取文档 URL
    doc = data.get("data", {}).get("document", {})
    doc_id = doc.get("document_id", "")
    doc_url = doc.get("url", "")

    if not doc_url:
        # 飞书文档 URL 格式：https://{domain}/docx/{doc_id}
        brand = "cccli4.feishu.cn"  # 飞书国内版域名
        doc_url = f"https://{brand}/docx/{doc_id}"

    if not doc_url or doc_url == f"https://cccli4.feishu.cn/docx/":
        # 最后的兜底：从原始输出中尝试提取 URL
        raw = json.dumps(data)
        match = re.search(r'https?://[^\s"]+/docx/[^\s"]+', raw)
        if match:
            doc_url = match.group(0)

    return doc_url


# ===== 步骤 4：发送群消息 =====

def send_group_message(chat_id: str, record: Dict[str, str], doc_url: str, group_msg: str):
    """向飞书群发送日报（lark-cli user 身份）。机器人进群后改为 bot token。"""
    full_msg = group_msg.strip()
    if doc_url:
        full_msg += f"\n\n📄 [更多请查看完整日报]({doc_url})"
    try:
        run_lark("im", "+messages-send", "--chat-id", chat_id, "--markdown", full_msg)
    except Exception as e:
        raise Exception(f"发送群消息失败: {e}")


# ===== 步骤 5：自动创建延期跟进任务 =====

def maybe_create_followup_task(record: Dict[str, str], doc_url: str) -> bool:
    """当今日预计延迟订单数 > 0 时创建飞书任务。返回是否创建。"""
    delay_str = record.get("今日预计延迟订单数", "0") or "0"
    try:
        delay = int(delay_str)
    except ValueError:
        delay = 0

    if delay <= 0:
        return False

    shortage_full = record.get("库存缺货统计", "无") or "无"
    shortage_sku = shortage_full.split('\n')[0] if '\n' in shortage_full else shortage_full

    # 截止时间：明天 12:00 CST
    tomorrow = datetime.now(CST) + timedelta(days=1)
    due_str = tomorrow.strftime("%Y-%m-%dT12:00:00+08:00")

    summary = f"【跟进】{datetime.now(CST).strftime('%m.%d')}存在延期订单需处理"

    description = (
        f"今日预计延迟订单数：{delay} 单\n"
        f"最紧缺物料：{shortage_sku}\n"
        f"日报链接：{doc_url}\n"
        f"\n请尽快跟进处理延期订单。"
    )

    args: List[str] = [
        "task", "+create",
        "--summary", summary,
        "--description", description,
        "--due", due_str,
    ]

    # 尝试自动发现任务清单
    tasklist_id = TASKLIST_ID or _discover_tasklist_id()
    if tasklist_id:
        args.extend(["--tasklist-id", tasklist_id])

    try:
        run_lark(*args)
        return True
    except Exception as e:
        err = str(e)
        if "99991672" in err or "permission" in err.lower():
            print(f"  [WARN] 创建任务失败：应用缺少 task:task:write 权限，请先在飞书开发者后台开启")
        else:
            print(f"  [WARN] 创建任务失败: {e}")
        return False


def _discover_tasklist_id() -> str:
    """尝试自动发现可用的任务清单 ID。"""
    try:
        data = run_lark("task", "tasklists", "list", "--page-size", "5")
        items = data.get("data", {}).get("items", [])
        if items:
            return items[0].get("id", "")
    except Exception:
        pass
    return ""


# ===== 三线异步流卡片模块 =====

def _safe_int(val, default=0):
    """安全转为整数，飞书字段可能是字符串/列表/数字。"""
    try:
        s = str(val).strip()
        return int(float(s)) if s else default
    except (ValueError, TypeError):
        return default


def _build_card_color(report: Dict[str, str]) -> str:
    """根据数据判定卡片主题色（战略大盘流用）。"""
    delay = _safe_int(report.get("今日预计延迟订单数", "0"))
    future_short = _safe_int(report.get("未来3天缺货订单数", "0"))
    shortage_sku = _safe_int(report.get("库存缺货SKU数", "0"))
    if delay > 0 or future_short > 0:
        return "red"
    if shortage_sku > 0:
        return "yellow"
    return "green"


def _format_pct(numerator: int, denominator: int) -> str:
    return f"{round(numerator / denominator * 100, 1)}%" if denominator > 0 else "-"


def _parse_card_date(date_text: str) -> Optional[datetime]:
    if not date_text:
        return None
    s = str(date_text).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s[:10], fmt).replace(tzinfo=CST)
        except ValueError:
            continue
    return None


def _sla_countdown_text(latest_ship_date: str) -> str:
    ship_date = _parse_card_date(latest_ship_date)
    if not ship_date:
        return "未知"
    today = datetime.now(CST).date()
    days_left = (ship_date.date() - today).days
    if days_left < 0:
        return f"已超期 {abs(days_left)} 天"
    if days_left == 0:
        return "今日到期"
    return f"剩余 {days_left} 天"


def _sla_days_left(latest_ship_date: str) -> Optional[int]:
    ship_date = _parse_card_date(latest_ship_date)
    if not ship_date:
        return None
    return (ship_date.date() - datetime.now(CST).date()).days


def _format_short_date(date_text: str) -> str:
    ship_date = _parse_card_date(date_text)
    if not ship_date:
        return date_text or "-"
    return ship_date.strftime("%-m/%-d") if os.name != "nt" else ship_date.strftime("%#m/%#d")


def _load_shortage_items(report: Dict[str, str]) -> List[Dict[str, str]]:
    """兼容 JSON/普通文本的缺货明细，供交互式卡片突出业务影响。"""
    raw = (report.get("缺货SKU明细") or "").strip()
    if not raw or raw == "[]":
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return [{"raw": raw}]

    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        return []

    items = []
    for item in data:
        if not isinstance(item, dict):
            continue
        items.append({str(k): _parse_feishu_value(v) for k, v in item.items()})
    return items


def _shortage_item_sku(item: Dict[str, str]) -> str:
    return (
        item.get("SKU")
        or item.get("SKU编码")
        or item.get("产品编码SKU")
        or item.get("断供物料")
        or item.get("物料名称")
        or item.get("sku")
        or ""
    )


def _shortage_item_gap(item: Dict[str, str]) -> str:
    return (
        item.get("缺口数量")
        or item.get("缺口")
        or item.get("最紧缺SKU缺口")
        or item.get("gap")
        or "0"
    )


def _pick_shortage_item(report: Dict[str, str], top_sku: str) -> Dict[str, str]:
    items = _load_shortage_items(report)
    if not items:
        return {}
    normalized_top = (top_sku or "").strip().lower()
    for item in items:
        sku = _shortage_item_sku(item)
        if normalized_top and sku.strip().lower() == normalized_top:
            return item
    return items[0]


def fetch_project_name_by_id(order_id: str) -> str:
    """TODO: 后续可在这里对接多维表格，将合同/订单号翻译为项目名或客户名。"""
    order_id = (order_id or "").strip()
    if not order_id:
        return ""
    cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache", "ai_schedule_summary.csv")
    try:
        with open(cache_path, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if (row.get("合同编号") or "").strip() != order_id:
                    continue
                project = (row.get("项目名称") or "").strip()
                customer = (row.get("客户名称") or "").strip()
                return project or customer or order_id
    except Exception:
        pass
    return order_id


def _shortage_business_impact(report: Dict[str, str], item: Dict[str, str]) -> str:
    impact = (
        item.get("关联订单项目")
        or item.get("订单项目")
        or item.get("项目名称")
        or item.get("客户名称")
    )
    if impact:
        return impact
    order_id = (
        item.get("合同编号")
        or item.get("订单号")
        or report.get("最长交期订单")
    )
    translated = fetch_project_name_by_id(order_id)
    return translated or "暂未识别到关联项目"


def _shortage_latest_ship_date(report: Dict[str, str], item: Dict[str, str]) -> str:
    return (
        item.get("最晚发货日期")
        or item.get("最晚发货日")
        or item.get("要求发货日期")
        or item.get("AI建议发货时间")
        or report.get("最晚发货日期")
        or "-"
    )


def _procurement_owner_mention() -> str:
    owner_open_id = os.getenv("PROCUREMENT_OWNER_OPEN_ID", "").strip()
    if owner_open_id:
        return f"<at id={owner_open_id}></at>"
    return "@采购负责人"


def _build_shortage_brief_list(report: Dict[str, str], top_sku: str) -> str:
    items = _load_shortage_items(report)
    briefs = []
    for item in items:
        sku = _shortage_item_sku(item)
        if not sku:
            continue
        if top_sku and sku.strip() == top_sku.strip():
            continue
        gap = _shortage_item_gap(item)
        briefs.append(f"{sku}(缺{gap})")
    if briefs:
        return " | ".join(briefs[:6])

    top_raw = top_sku or report.get("库存缺货统计", "")
    top = top_raw.split('\n')[0] if '\n' in (top_raw or '') else (top_raw or '')
    return top if top and top != "无" else "暂无其他断供项"


def _estimate_lead_time_days(report: Dict[str, str]) -> str:
    explicit = report.get("本周平均履约周期") or report.get("平均履约周期") or report.get("Lead Time")
    if explicit:
        return str(explicit).replace("天", "")
    latest = _parse_card_date(report.get("最晚发货日期", ""))
    if latest:
        return str(max((latest.date() - datetime.now(CST).date()).days, 0))
    return "-"


def _estimate_capacity_hunger_days(report: Dict[str, str]) -> str:
    explicit = report.get("产能池饥饿度") or report.get("积压可流转天数")
    if explicit:
        return str(explicit).replace("天", "")
    backlog = _safe_int(report.get("已排单订单数", "0")) + _safe_int(report.get("未排单订单数", "0"))
    today_should = _safe_int(report.get("今日应发货订单数", "0"))
    future_7 = _safe_int(report.get("未来7天应发货订单数", "0"))
    daily_burn = max(today_should, round(future_7 / 7) if future_7 > 0 else 0, 1)
    return str(round(backlog / daily_burn, 1))


def _boss_ai_insight(report: Dict[str, str], color: str, lead_time: str, hunger_days: str) -> str:
    should_ship = report.get("今日应发货订单数", "0")
    can_ship = report.get("今日可发货订单数", "0")
    delay = report.get("今日预计延迟订单数", "0")
    shortage_sku = report.get("库存缺货SKU数", "0")
    if color == "red":
        return (
            f"今日可发 {can_ship}/{should_ship} 单，已有 {delay} 单交付被阻断；"
            f"产能池预计支撑 {hunger_days} 天，核心断供物料正在放大库存呆滞与项目延期风险。"
        )
    if color == "yellow":
        return (
            f"今日履约基本可控，但 {shortage_sku} 个核心物料断供已形成前瞻风险；"
            f"当前积压仅够流转 {hunger_days} 天，需同步销售补充有效订单水位。"
        )
    return (
        f"今日交付链路健康，履约周期约 {lead_time} 天；"
        f"产能池可维持 {hunger_days} 天流转，暂未暴露资金或库存呆滞风险。"
    )


# ===== 三线卡片构建函数 =====

def _build_procurement_card(report: Dict[str, str]) -> dict:
    """Track 1 - 采购/物料产协群：基于业务影响施压，而不是罗列系统数字。"""
    batch = report.get("排单批次号", "-")
    top_raw = report.get("库存缺货统计", "无") or "无"
    top_sku = top_raw.split('\n')[0] if '\n' in top_raw else top_raw
    shortage_sku = _safe_int(report.get("库存缺货SKU数", "0"))
    shortage_item = _pick_shortage_item(report, top_sku)
    business_impact = _shortage_business_impact(report, shortage_item)
    latest_ship_date = _shortage_latest_ship_date(report, shortage_item)
    sla_countdown = _sla_countdown_text(latest_ship_date)
    days_left = _sla_days_left(latest_ship_date)
    sla_level = (
        f"<font color='red'>**{sla_countdown}**</font>"
        if days_left is not None and days_left < 3
        else f"<font color='orange'>**{sla_countdown}**</font>"
        if days_left is not None and days_left < 7
        else f"**{sla_countdown}**"
    )
    shortage_level = "高" if shortage_sku > 0 else "中"
    backup_list = _build_shortage_brief_list(report, top_sku)
    owner = _procurement_owner_mention()
    action_url = os.getenv("PROCUREMENT_ACTION_URL", "")

    # ---- 三格仪表盘 ----
    dash_columns = [
        {
            "tag": "column",
            "width": "weighted", "weight": 1,
            "elements": [{"tag": "markdown", "content": f"<font color='red'>**{shortage_sku}**</font>\n断供物料", "text_align": "center"}],
        },
        {
            "tag": "column",
            "width": "weighted", "weight": 1,
            "elements": [{"tag": "markdown", "content": f"<font color='red'>**{_format_short_date(latest_ship_date)}**</font>\nSLA截止", "text_align": "center"}],
        },
        {
            "tag": "column",
            "width": "weighted", "weight": 1,
            "elements": [{"tag": "markdown", "content": f"<text_tag color='red'>高风险</text_tag>\n等级", "text_align": "center"}],
        },
    ]

    # ---- 最紧急物料 ----
    spotlight = (
        f"{owner}\n\n"
        f"**<font color='red'>最紧急物料</font>**：{top_sku}\n"
        f"业务影响：{business_impact}\n"
        f"SLA倒计时：{sla_level}（最晚{_format_short_date(latest_ship_date)}）"
    )

    # ---- Top 5 其他物料 ----
    other_items = _load_shortage_items(report)
    other_lines = []
    count = 0
    for item in other_items:
        sku = _shortage_item_sku(item)
        if not sku or sku.strip() == top_sku.strip():
            continue
        gap = _shortage_item_gap(item)
        impact = _shortage_business_impact(report, item)
        other_lines.append(f"{sku}(缺{gap}) | {impact}")
        count += 1
        if count >= 5:
            break
    if len(other_items) > count + 1:
        other_lines.append(f"+{len(other_items) - count - 1} 项详见清单")

    # ---- 组装 card ----
    elements = [
        {"tag": "column_set", "flex_mode": "trisect", "columns": dash_columns},
        {"tag": "hr"},
        {"tag": "markdown", "content": spotlight},
    ]
    if other_lines:
        elements.append({"tag": "hr"})
        elements.append({"tag": "markdown", "content": "**其他高风险物料**\n" + "\n".join(other_lines)})
    elements.append({"tag": "hr"})
    elements.append({"tag": "markdown", "content": "<text_tag color='orange'>今日需完成</text_tag>\n- 更新承诺到货交期\n- 无法承诺则同步替代料/拆单方案"})

    # 按钮
    buttons = [{
        "tag": "button",
        "text": {"tag": "plain_text", "content": "录入最新到货交期"},
        "type": "primary",
        "url": action_url,
        "multi_url": {"url": action_url, "pc_url": action_url, "android_url": action_url, "ios_url": action_url}
    }] if action_url else []
    # secondary button: 查物料清单
    inv_url = os.getenv("BITABLE_INV_URL", "")
    if inv_url:
        buttons.append({
            "tag": "button",
            "text": {"tag": "plain_text", "content": "查看紧缺物料清单"},
            "type": "default",
            "url": inv_url,
            "multi_url": {"url": inv_url, "pc_url": inv_url, "android_url": inv_url, "ios_url": inv_url}
        })
    if buttons:
        elements.append({"tag": "action", "actions": buttons})

    # note 时间戳
    now_ts = datetime.now(CST).strftime("%Y-%m-%d %H:%M CST")
    elements.append({"tag": "note", "elements": [{"tag": "plain_text", "content": f"批次 {batch} | {now_ts}"}]})

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "red" if shortage_sku > 0 else "orange",
            "title": {
                "content": f"物料断供阻断预警 ({batch})",
                "tag": "plain_text"
            }
        },
        "elements": elements
    }
    return card


def _build_planner_card(report: Dict[str, str]) -> dict:
    """Track 2 - 订单确认流：必定触发，蓝色卡片。"""
    batch = report.get("排单批次号", "-")
    scheduled = report.get("已排单订单数", "0")
    unscheduled = report.get("未排单订单数", "0")
    action_url = os.getenv("PLANNER_ACTION_URL", "")

    markdown = (
        "AI 影子计算已完成，请排单员进行人工确认。\n\n"
        f"已排单：**{scheduled}** 单\n"
        f"未排单：**{unscheduled}** 单"
    )

    elements = [
        {"tag": "markdown", "content": markdown}
    ]

    if action_url:
        elements.append({
            "tag": "action",
            "actions": [{
                "tag": "button",
                "text": {"tag": "plain_text", "content": "确认排期，一键同步生产"},
                "type": "primary",
                "url": action_url,
                "multi_url": {"url": action_url, "pc_url": action_url,
                              "android_url": action_url, "ios_url": action_url}
            }]
        })

    # note 时间戳
    now_ts = datetime.now(CST).strftime("%Y-%m-%d %H:%M CST")
    elements.append({"tag": "note", "elements": [{"tag": "plain_text", "content": f"批次 {batch} | {now_ts}"}]})

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "blue",
            "title": {
                "content": f"排单确认通知 ({batch})",
                "tag": "plain_text"
            }
        },
        "elements": elements
    }
    return card


def _build_boss_card(report: Dict[str, str]) -> dict:
    """Track 3 - 老板/全局经营指挥群：报资产效率与前瞻风险。"""
    batch = report.get("排单批次号", "-")
    color = _build_card_color(report)
    total = report.get("订单总数", "0")
    should_ship = report.get("今日应发货订单数", "0")
    can_ship = report.get("今日可发货订单数", "0")
    delay = report.get("今日预计延迟订单数", "0")
    shortage_sku = report.get("库存缺货SKU数", "0")
    future_short = report.get("未来3天缺货订单数", "0")
    top_raw = report.get("库存缺货统计", "无") or "无"
    top_sku = top_raw.split('\n')[0] if '\n' in top_raw else top_raw
    lead_time = _estimate_lead_time_days(report)
    hunger_days = _estimate_capacity_hunger_days(report)
    dashboard_url = os.getenv("BITABLE_DASHBOARD_URL", "")

    ss = _safe_int(should_ship)
    cs = _safe_int(can_ship)
    delivery_rate = _format_pct(cs, ss)
    d_int = _safe_int(delay)

    # ---- MVP: 三格仪表盘 (column_set trisect) ----
    # 交付率（含目标对比）、库存健康度（含目标对比）、需关注风险项
    dr_target = float(os.getenv("KPI_DELIVERY_TARGET", "90"))
    dr_ok = cs > 0 and ss > 0 and (float(cs)/float(ss) * 100) >= dr_target

    sufficient_sku = _safe_int(report.get("库存充足SKU数", "0"))
    warning_sku = _safe_int(report.get("库存预警SKU数", "0"))
    total_sku = _safe_int(report.get("SKU总数", "0"))
    health_rate = _format_pct(sufficient_sku, total_sku)
    health_target = float(os.getenv("KPI_HEALTH_TARGET", "80"))
    health_ok = total_sku > 0 and (sufficient_sku / total_sku * 100) >= health_target

    # 需关注项数：延迟 + 缺货 + 预警
    attention_count = d_int + warning_sku + (1 if _safe_int(shortage_sku) > 0 and top_sku != "无" else 0)
    attention_ok = attention_count <= 2

    # 双通道编码：颜色 + 文字标签
    dr_pct = (float(cs) / float(ss) * 100) if ss > 0 else 0
    dr_label = "<text_tag color='green'>正常</text_tag>" if dr_ok else ("<text_tag color='orange'>关注</text_tag>" if ss > 0 and dr_pct >= 80 else "<text_tag color='red'>告警</text_tag>")
    health_pct = (float(sufficient_sku) / float(total_sku) * 100) if total_sku > 0 else 0
    health_label = "<text_tag color='green'>正常</text_tag>" if health_ok else ("<text_tag color='orange'>关注</text_tag>" if total_sku > 0 and health_pct >= 60 else "<text_tag color='red'>告警</text_tag>")
    att_label = "<text_tag color='green'>正常</text_tag>" if attention_ok else ("<text_tag color='orange'>关注</text_tag>" if attention_count <= 5 else "<text_tag color='red'>告警</text_tag>")

    dash_columns = [
        {"tag": "column", "width": "weighted", "weight": 1,
         "elements": [{"tag": "markdown", "content": f"<font color='green'>**{delivery_rate}**</font>\n交付率\n目标 {int(dr_target)}% {dr_label}", "text_align": "center"}] 
         if dr_ok else 
         [{"tag": "markdown", "content": f"<font color='orange'>**{delivery_rate}**</font>\n交付率\n目标 {int(dr_target)}% {dr_label}", "text_align": "center"}]
         if ss > 0 and dr_pct >= 80 else
         [{"tag": "markdown", "content": f"<font color='red'>**{delivery_rate}**</font>\n交付率\n目标 {int(dr_target)}% {dr_label}", "text_align": "center"}]},
        {"tag": "column", "width": "weighted", "weight": 1,
         "elements": [{"tag": "markdown", "content": f"<font color='green'>**{health_rate}**</font>\n库存健康度\n目标 {int(health_target)}% {health_label}", "text_align": "center"}] 
         if health_ok else
         [{"tag": "markdown", "content": f"<font color='orange'>**{health_rate}**</font>\n库存健康度\n目标 {int(health_target)}% {health_label}", "text_align": "center"}]
         if total_sku > 0 and health_pct >= 60 else
         [{"tag": "markdown", "content": f"<font color='red'>**{health_rate}**</font>\n库存健康度\n目标 {int(health_target)}% {health_label}", "text_align": "center"}]},
        {"tag": "column", "width": "weighted", "weight": 1,
         "elements": [{"tag": "markdown", "content": f"<font color='green'>**{attention_count}项**</font>\n需关注\n{att_label}", "text_align": "center"}] 
         if attention_ok else
         [{"tag": "markdown", "content": f"<font color='orange'>**{attention_count}项**</font>\n需关注\n{att_label}", "text_align": "center"}]
         if attention_count <= 5 else
         [{"tag": "markdown", "content": f"<font color='red'>**{attention_count}项**</font>\n需关注\n{att_label}", "text_align": "center"}]},
    ]

    # ---- 今日结论（灰底背景块） ----
    trend_text = ""
    if lead_time:
        trend_text = f" | Lead Time {lead_time}天"
    conclusion = f"交付率 {delivery_rate}（目标 {int(dr_target)}%）{trend_text}，{attention_count} 项需关注"
    alert_status = "告警" if not dr_ok and attention_count > 5 else ("关注" if not dr_ok or not attention_ok else "正常")
    status_icon = "●" if alert_status == "告警" else ("⚠" if alert_status == "关注" else "✓")
    conclusion_md = (
        f"**{status_icon} 今日{alert_status}**\n"
        f"<font color='grey'>{conclusion}</font>"
    )

    # ---- 需关注区域（最多3条 + 角色标签） ----
    attention_items = []
    if d_int > 0:
        attention_items.append(f"- [交付] {delay} 单延迟，需跟进处理")
    if _safe_int(shortage_sku) > 0 and top_sku != "无":
        attention_items.append(f"- [采购] 核心物料 {top_sku} 断供，影响 {shortage_sku} 项")
    if warning_sku > 0:
        attention_items.append(f"- [库存] {warning_sku} 个SKU预警，建议补货")
    if future_short and _safe_int(future_short) > 0:
        attention_items.append(f"- [预测] 未来3天 {future_short} 单有缺货风险")
    attention_text = "\n".join(attention_items[:3])
    if len(attention_items) > 3:
        attention_text += f"\n+{len(attention_items) - 3} 项详情"

    # ---- 灰度信息行 ----
    total_anomalies = d_int + _safe_int(shortage_sku) + warning_sku
    gray_note = f"今日共 {total_anomalies} 项异常，{min(len(attention_items), 3)} 项需关注，其余执行层闭环中 | {total} 单资产池"

    # ---- header 颜色：正常时蓝色 ----
    if not dr_ok and attention_count > 5:
        header_color = "red"
    elif not dr_ok or not attention_ok:
        header_color = "yellow"
    else:
        header_color = "blue"

    # ---- 组装 card ----
    elements = [
        {"tag": "column_set", "flex_mode": "trisect", "columns": dash_columns},
        {"tag": "hr"},
        {"tag": "markdown", "content": conclusion_md},
    ]
    if attention_text:
        elements.append({"tag": "hr"})
        elements.append({"tag": "markdown", "content": "**需关注**\n" + attention_text})

    if dashboard_url:
        elements.append({
            "tag": "action",
            "actions": [
                {
                    "tag": "button",
                    "text": {"tag": "plain_text", "content": "查看详情"},
                    "type": "primary",
                    "url": dashboard_url,
                    "multi_url": {"url": dashboard_url, "pc_url": dashboard_url, "android_url": dashboard_url, "ios_url": dashboard_url}
                },
                {
                    "tag": "button",
                    "text": {"tag": "plain_text", "content": "查看交付明细"},
                    "type": "default",
                    "url": dashboard_url,
                    "multi_url": {"url": dashboard_url, "pc_url": dashboard_url, "android_url": dashboard_url, "ios_url": dashboard_url}
                },
            ]
        })

    # note: 时间戳 + 灰度信息
    now_ts = datetime.now(CST).strftime("%Y-%m-%d %H:%M CST")
    elements.append({"tag": "note", "elements": [
        {"tag": "plain_text", "content": f"{gray_note}"},
        {"tag": "plain_text", "content": f"批次 {batch} | {now_ts} | AI排单引擎"},
    ]})

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": header_color,
            "title": {
                "content": f"供应链经营日报 ({batch})",
                "tag": "plain_text"
            }
        },
        "elements": elements
    }
    return card


# ===== 发送基础设施（保留不动） =====

def _send_card_to_webhook(webhook_url: str, card: dict, label: str):
    """发送卡片到指定 Webhook，带完整错误日志。"""
    # 校验 URL 格式：必须包含 /bot/v2/hook
    if "/bot/v2/hook" not in webhook_url:
        print(f"  [SKIP] {label}：URL 不是飞书 Bot Webhook 地址，当前值={webhook_url[:60]}...")
        return False

    payload = {"msg_type": "interactive", "card": card}
    try:
        resp = httpx.post(
            webhook_url,
            json=payload,
            timeout=15.0,
            headers={"Content-Type": "application/json; charset=utf-8"},
        )
        data = resp.json()
        code_ok = data.get("code", 0) == 0
        status_ok = data.get("StatusCode", 0) == 0
        message_ok = data.get("StatusMessage", "success") == "success"
        if not (code_ok and status_ok and message_ok):
            print(f"  [WARN] {label} 发送异常: {json.dumps(data, ensure_ascii=False)}")
            return False
        print(f"  [OK] {label} 发送成功")
        return True
    except httpx.RequestError as e:
        print(f"  [X] {label} 网络错误: {e}")
        return False
    except Exception as e:
        print(f"  [X] {label} 未知错误: {e}")
        return False


# ===== 三线分流主控入口 =====

def execute_triple_track_dispatch(report: Dict[str, str]):
    """事件驱动型三线异步流（3-Track Parallel Flow）。

    Track 1 — 物料异常流：仅 库存缺货SKU数 > 0 时触发 → PROCUREMENT_WEBHOOK_URL
    Track 2 — 排单确认流：仅发送到独立个人通知 Webhook，不复用采购群 Webhook
    Track 3 — 战略大盘流：默认发送 → BOSS_WEBHOOK_URL
    """
    shortage_sku = _safe_int(report.get("库存缺货SKU数", "0"))
    procurement_url = os.getenv("PROCUREMENT_WEBHOOK_URL", "")
    planner_url = os.getenv("PLANNER_WEBHOOK_URL", "")
    boss_url = os.getenv("BOSS_WEBHOOK_URL", "")

    any_attempted = False
    all_ok = True

    # ---- Track 1：物料异常流（条件触发） ----
    if shortage_sku > 0 and procurement_url:
        try:
            card = _build_procurement_card(report)
            all_ok = _send_card_to_webhook(procurement_url, card, "Track1-物料异常流") and all_ok
            any_attempted = True
        except Exception as e:
            print(f"  [X] Track1-物料异常流 构建或发送异常: {e}")
            all_ok = False
    elif shortage_sku > 0 and not procurement_url:
        print("  [SKIP] Track1-物料异常流：PROCUREMENT_WEBHOOK_URL 未配置")
    else:
        print("  [SKIP] Track1-物料异常流：库存缺货SKU数为0，无需触发")

    # ---- Track 2：排单确认流（个人通知，不复用采购/产协群） ----
    if planner_url and planner_url != procurement_url:
        try:
            card = _build_planner_card(report)
            all_ok = _send_card_to_webhook(planner_url, card, "Track2-订单确认流") and all_ok
            any_attempted = True
        except Exception as e:
            print(f"  [X] Track2-订单确认流 构建或发送异常: {e}")
            all_ok = False
    elif planner_url and planner_url == procurement_url:
        print("  [SKIP] Track2-订单确认流：PLANNER_WEBHOOK_URL 与采购群相同，避免把个人通知发到群里")
    else:
        print("  [SKIP] Track2-订单确认流：PLANNER_WEBHOOK_URL 未配置")

    # ---- Track 3：战略大盘流（默认发送） ----
    if boss_url:
        try:
            card = _build_boss_card(report)
            all_ok = _send_card_to_webhook(boss_url, card, "Track3-战略大盘流") and all_ok
            any_attempted = True
        except Exception as e:
            print(f"  [X] Track3-战略大盘流 构建或发送异常: {e}")
            all_ok = False
    else:
        print("  [SKIP] Track3-战略大盘流：BOSS_WEBHOOK_URL 未配置")

    if not any_attempted:
        print("  [SKIP] 三线分流均未执行：所有 Webhook URL 均未配置")
        return False

    return all_ok


# ===== 主流程 =====

def main():
    print("=" * 55)
    print("  AI排单日报机器人")
    print(f"  启动时间：{datetime.now(CST).strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    try:
        # ---- 步骤 1 ----
        print("[1/7] 读取多维表格最新数据...")
        record = fetch_latest_record()
        print(f"  [OK] 读取表格成功 — 批次号: {record.get('排单批次号')}, "
              f"日期: {record.get('日期')}, "
              f"订单总数: {record.get('订单总数')}")

        # ---- 步骤 2 ----
        print("[2/7] 调用 DeepSeek 生成日报...")
        full_report, group_msg = generate_daily_report(record)
        print(f"  [OK] 生成日报成功 — 正文 {len(full_report)} 字，群消息 {len(group_msg)} 字")

        # ---- 步骤 3 ----
        print("[3/7] 创建飞书文档...")
        print("       生成缺货明细表...")
        shortage_table = _build_shortage_table(record)
        print("       生成常规备货建议表（读取SKU+库存表）...")
        restock_table = _build_restock_table(record)
        today_str = datetime.now(CST).strftime("%Y%m%d")
        doc_url = create_feishu_doc(
            f"AI排单运营日报-{today_str}", full_report,
            shortage_table=shortage_table,
            restock_table=restock_table,
        )
        print(f"  [OK] 创建文档成功: {doc_url}")

        # ---- 步骤 4 ----
        print("[4/7] 发送群消息（机器人身份）...")
        send_group_message(CHAT_ID, record, doc_url, group_msg)
        print("  [OK] 发送消息成功")

        # ---- 步骤 5 ----
        print("[5/7] 三线异步流卡片推送（已迁移到排单员确认流程，此处跳过）...")
        # execute_triple_track_dispatch(record)  # 群通知由排单员确认后触发，避免独立流程绕过确认

        # ---- 步骤 6 ----
        print("[6/7] 判断是否需要创建跟进任务...")
        created = maybe_create_followup_task(record, doc_url)
        if created:
            print("  [OK] 已创建延期跟进任务")
        else:
            print("  [SKIP] 无延期订单，不创建任务")

        # ---- 步骤 7 ----
        print("[7/7] 全部流程执行完毕 [OK]")
        print("=" * 55)
        print(f"  文档链接: {doc_url}")
        print("=" * 55)

    except Exception as e:
        print(f"\n[X] 执行失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
