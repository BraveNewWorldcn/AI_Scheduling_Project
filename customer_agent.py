#!/usr/bin/env python3
"""
全自助客户查单中枢（飞书长连接版）
----------------------------------
通过 lark-cli 长连接接收事件，不依赖公网 URL。
四层漏斗：身份绑定 → 语义提取 → 消歧卡片 → 精准交付

运行方式：python customer_agent.py（常驻进程）
"""

from __future__ import annotations

import json
import os
import sys
import re
import time
import subprocess
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional
from difflib import SequenceMatcher

import httpx

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

# ----- 自动加载 .env -----
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

CST = timezone(timedelta(hours=8))

APP_ID = os.getenv("CUSTOMER_BOT_APP_ID", "")
APP_SECRET = os.getenv("CUSTOMER_BOT_APP_SECRET")
DATA_APP_ID = os.getenv("FEISHU_APP_ID", APP_ID)
DATA_APP_SECRET = os.getenv("FEISHU_APP_SECRET", APP_SECRET)
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_ID_MAIN = os.getenv("TABLE_ID_MAIN", "tbl06oxGEdMNTEB8")
TABLE_ID_DETAIL = os.getenv("TABLE_ID_DETAIL", "tbl09Z6C7wCGh3mW")

if not APP_SECRET:
    sys.exit("[X] 请先设置 CUSTOMER_BOT_APP_SECRET")

# ===== 会话缓存（消歧卡片选择） =====
# {open_id: [{"合同编号":..., "项目名称":...}, ...]}
_session_cache: Dict[str, List[Dict[str, str]]] = {}
_seen_event_ids: set[str] = set()
_token_cache: Dict[str, tuple[str, datetime]] = {}
_lark_cli_path: str = ""


def _find_lark_cli() -> str:
    global _lark_cli_path
    if _lark_cli_path:
        return _lark_cli_path
    candidates = [
        os.path.expandvars(r"%APPDATA%\npm\lark-cli.cmd"),
    ]
    if sys.platform != "win32":
        candidates += ["/usr/local/bin/lark-cli", "/opt/homebrew/bin/lark-cli"]
    for p in candidates:
        if os.path.isfile(p):
            _lark_cli_path = p
            return p
    for p in os.environ.get("PATH", "").split(os.pathsep):
        for name in ("lark-cli.cmd", "lark-cli"):
            full = os.path.join(p, name)
            if os.path.isfile(full):
                _lark_cli_path = full
                return full
    return "lark-cli"


# ===== 工具函数 =====

def _get_token(app_id: str, app_secret: str) -> str:
    cached = _token_cache.get(app_id)
    if cached and cached[1] > datetime.now(CST) + timedelta(minutes=5):
        return cached[0]
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": app_id, "app_secret": app_secret}, timeout=30.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取 Token 失败: {data}")
    token = data["tenant_access_token"]
    _token_cache[app_id] = (token, datetime.now(CST) + timedelta(seconds=6600))
    return token


def _get_bot_token() -> str:
    return _get_token(APP_ID, APP_SECRET)


def _get_data_token() -> str:
    return _get_token(DATA_APP_ID, DATA_APP_SECRET)


def _parse_feishu_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        texts = [item["text"].strip() if isinstance(item, dict) and "text" in item
                 else str(item).strip() for item in value]
        return ", ".join(texts)
    if isinstance(value, dict):
        for k in ("text", "number", "phone", "email", "url"):
            if k in value:
                return str(value[k]).strip()
        return str(value).strip()
    if isinstance(value, (int, float)):
        return str(value)
    return str(value).strip()


def _parse_feishu_date(value: Any) -> str:
    if isinstance(value, (int, float)) and value > 0:
        try:
            ts = value / 1000 if value >= 10_000_000_000 else value
            return datetime.fromtimestamp(ts, tz=CST).strftime("%Y-%m-%d")
        except Exception:
            pass
    text = _parse_feishu_value(value)
    if re.fullmatch(r"\d{10,13}", text):
        try:
            n = int(text)
            ts = n / 1000 if n >= 10_000_000_000 else n
            return datetime.fromtimestamp(ts, tz=CST).strftime("%Y-%m-%d")
        except Exception:
            pass
    return text[:10] if len(text) >= 10 else text


def _fuzzy_score(query: str, target: str) -> float:
    q, t = query.lower().strip(), target.lower().strip()
    if not q or not t:
        return 0.0
    if q in t:
        return 0.9
    return SequenceMatcher(None, q, t).ratio()


# ===== 第一层：身份提取 =====

_identity_cache: Dict[str, str] = {}


def _extract_company_name(open_id: str) -> Optional[str]:
    if open_id in _identity_cache:
        return _identity_cache[open_id]
    try:
        token = _get_bot_token()
        resp = httpx.get(
            f"https://open.feishu.cn/open-apis/contact/v3/users/{open_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"user_id_type": "open_id"}, timeout=10.0,
        )
        data = resp.json()
        if data.get("code") == 0:
            user = data.get("data", {}).get("user", {})
            company = user.get("company_name", "")
            if company:
                _identity_cache[open_id] = company
                return company
    except Exception:
        pass
    return None


# ===== 第二层：模糊匹配 =====

def _load_company_orders(company_name: str) -> List[Dict[str, str]]:
    token = _get_data_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    base = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}"
    orders: List[Dict[str, str]] = []
    page_token = None
    while True:
        params: Dict[str, Any] = {"page_size": 200}
        if page_token:
            params["page_token"] = page_token
        resp = httpx.get(f"{base}/tables/{TABLE_ID_MAIN}/records",
                         headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"读取销售订单主表失败: {data}")
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            cust = _parse_feishu_value(f.get("客户名称"))
            # 公司匹配：放宽阈值到 0.15，避免因飞书企业名与合同客户名细微差异导致漏单
            if company_name and cust and _fuzzy_score(company_name, cust) < 0.15:
                continue
            orders.append({
                "合同编号": _parse_feishu_value(f.get("合同编号")),
                "客户名称": cust,
                "项目名称": _parse_feishu_value(f.get("项目名称")),
            })
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return orders


def _extract_keywords(user_input: str) -> List[str]:
    """智能提取查询关键词：去除口语化噪音，保留项目/设备/交付相关实词。"""
    # 纯噪音词（完全无信息量的虚词/口语词）
    noise = ["帮我", "查一下", "那个", "这个", "什么", "怎么", "单子",
             "一下", "请问", "我想", "我要", "看看", "有没有",
             "能不能", "可以", "请", "麻烦", "帮忙", "谢谢",
             "查一查", "查询", "查找"]

    cleaned = user_input
    for w in noise:
        cleaned = cleaned.replace(w, " ")

    # 步骤1：按标点拆分
    raw_tokens = [t.strip() for t in re.split(r"[\s,，、。！？]+", cleaned) if len(t.strip()) >= 2]

    # 步骤2：对每个 token 做进一步拆分和清洗
    # 去掉常见口语后缀（发了吗→发，到了没→到，啥时候→空）
    refined: List[str] = []
    for tok in raw_tokens:
        # 去掉口语化后缀（保留核心词）
        core = re.sub(r"(了吗|了没|没有|没呢|到了吗|啥时候|什么时候|怎么样|的情况|的进度)$", "", tok)
        if len(core.strip()) >= 2:
            refined.append(core.strip())
        # 如果原始 token 包含"项目""设备""系统"等词，保留整个 token
        if core.strip() != tok and any(w in tok for w in ["项目", "设备", "系统", "工程", "园区", "大楼"]):
            refined.append(tok.strip())

    # 步骤3：对含中文的长 token（≥4字符）尝试通过常见分隔词拆分
    extra: List[str] = []
    for tok in refined:
        if len(tok) >= 4 and re.search(r"[一-鿿]", tok):
            # 在"项目""设备""系统""工程""园区"处拆分
            parts = re.split(r"(项目|设备|系统|工程|园区|大楼|中心|工厂|车站)", tok)
            for p in parts:
                p = p.strip()
                if len(p) >= 2:
                    extra.append(p)
    refined.extend(extra)

    # 去重，保留顺序，最多取 8 个
    seen: set = set()
    result = []
    for t in refined:
        if t not in seen:
            seen.add(t)
            result.append(t)
    return result[:8]


def _fuzzy_match_orders(orders: List[Dict[str, str]], user_input: str) -> List[Dict[str, Any]]:
    """多级匹配策略：精确合同号 → 关键词组合 → 整体模糊匹配。"""
    query = user_input.strip().lower()

    # Level 1: 精确合同编号匹配
    exact_contracts = [
        {**order, "_score": 1.0}
        for order in orders
        if order.get("合同编号", "").strip().lower() and order.get("合同编号", "").strip().lower() in query
    ]
    if exact_contracts:
        return exact_contracts[:1]

    keywords = _extract_keywords(user_input)

    # 同时提取更短的关键词（≥1字符，如"医院"、"发"）
    short_keywords = [t.strip() for t in re.split(r"[\s,，、。！？]+", user_input) if 1 <= len(t.strip()) <= 3]

    scored: List[tuple] = []
    for order in orders:
        proj = order.get("项目名称", "")
        contract = order.get("合同编号", "")
        target = f"{proj} {contract}".lower()
        proj_lower = proj.lower()

        # 基础分：整体模糊匹配
        score = _fuzzy_score(user_input, f"{proj} {contract}")

        # 加分项 1：关键词命中项目名称
        keyword_hits = 0
        for kw in keywords:
            if len(kw) >= 2 and kw.lower() in proj_lower:
                keyword_hits += 1
                score += 0.25
            elif len(kw) >= 2 and kw.lower() in contract.lower():
                score += 0.2

        # 加分项 2：短关键词部分匹配（如"医院"、"写字楼"）
        for sk in short_keywords:
            if sk.lower() in proj_lower:
                score += 0.15

        # 加分项 3：用户查询整体是项目名的子串
        if query in proj_lower:
            score += 0.4

        # 加分项 4：关键词较多时给予组合奖励
        if keyword_hits >= 2:
            score += 0.2
        elif keyword_hits >= 3:
            score += 0.3

        score = min(score, 1.0)
        if score > 0.12:  # 略降低阈值
            scored.append((score, {**order, "_score": round(score, 2)}))

    scored.sort(key=lambda x: x[0], reverse=True)
    return [item for _, item in scored]


# ===== 第三层：消歧卡片 =====

def _build_disambiguation_card(candidates: List[Dict[str, Any]], open_id: str) -> dict:
    # 缓存候选订单，供用户回复序号时查询
    _session_cache[open_id] = [
        {"合同编号": c["合同编号"], "项目名称": c["项目名称"]} for c in candidates[:9]
    ]

    lines = [f"🧐 帮您找到了 **{len(candidates)}** 个相关项目：\n"]
    for i, c in enumerate(candidates[:9]):
        proj = c.get("项目名称", "未知")[:50]
        lines.append(f"**{i + 1}**. {proj}")

    lines.append(f"\n⚙️ 都不对请输入 **0**")

    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "wathet",
            "title": {"content": "🔍 请选择要查询的订单", "tag": "plain_text"},
        },
        "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
    }


# ===== 第四层：交付卡片 =====

def _load_order_detail(contract_id: str) -> Optional[Dict[str, str]]:
    token = _get_data_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    base = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{APP_TOKEN}"
    page_token = None
    while True:
        params: Dict[str, Any] = {"page_size": 200}
        if page_token:
            params["page_token"] = page_token
        resp = httpx.get(f"{base}/tables/{TABLE_ID_DETAIL}/records",
                         headers=headers, params=params, timeout=30.0)
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"读取AI排单总表失败: {data}")
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            if _parse_feishu_value(f.get("合同编号")) == contract_id:
                return {
                    "合同编号": contract_id,
                    "项目名称": _parse_feishu_value(f.get("项目名称")),
                    "AI建议发货时间": _parse_feishu_date(f.get("AI建议发货时间")),
                    "人工确认发货时间": _parse_feishu_date(f.get("人工确认发货时间")),
                    "整体状态": _parse_feishu_value(f.get("整体状态")),
                    "订单状态": _parse_feishu_value(f.get("订单状态")),
                    "AI风险": _parse_feishu_value(f.get("AI风险")),
                    "AI建议": _parse_feishu_value(f.get("AI建议")),
                }
        if not data.get("data", {}).get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return None


def _build_delivery_card(order: Dict[str, str]) -> dict:
    contract = order.get("合同编号", "-")
    project = order.get("项目名称", "-")
    status = order.get("整体状态", "有货")
    order_status = order.get("订单状态", "待确认")
    ai_ship = order.get("AI建议发货时间", "")
    manual_ship = order.get("人工确认发货时间", "")

    ship_date_raw = (manual_ship or ai_ship or "").strip()
    ship_date_display = "待定"
    if ship_date_raw:
        try:
            d = datetime.strptime(ship_date_raw[:10], "%Y-%m-%d")
            ship_date_display = f"{d.month}月{d.day}日"
        except Exception:
            ship_date_display = ship_date_raw

    # ---- 文案包装（绝不可出现：缺货/断供/风险/待补/紧急采购等）----
    if order_status == "已发货":
        header = "🚚 您的设备已发出"
        color = "green"
    elif order_status == "已确认":
        header = "📋 您的交付批次已确认"
        color = "blue"
    else:
        header = "📦 项目交付进度"
        color = "blue"

    if status == "缺货":
        status_line = (
            f"关于您关注的【**{project}**】项目（合同编号：{contract}），"
            f"目前设备正在按生产计划有序推进，各核心环节均在严格把控之中。"
        )
        detail_line = "设备核心模块正在进行出厂前的全面质量验证与联调测试，确保交付后即装即用、稳定运行。"
        if ship_date_raw:
            delivery_line = (
                f"您的交付批次已锁定，预计 **{ship_date_display}** 为您准时安排发出。"
            )
        else:
            delivery_line = "您的交付批次确认后，我将第一时间为您同步预计发出时间。"
    else:
        status_line = (
            f"关于您关注的【**{project}**】项目（合同编号：{contract}），一切进展顺利。"
        )
        detail_line = "设备已完成生产，正在进行出厂前的最终调试与专业防震包装，确保运输安全与到场即用。"
        if ship_date_raw:
            delivery_line = (
                f"您的交付批次已排定，预计 **{ship_date_display}** 为您准时发出。"
            )
        else:
            delivery_line = "您的交付批次确认后，我将第一时间为您同步预计发出时间。"

    markdown = (
        f"{status_line}\n\n"
        f"{detail_line}\n\n"
        f"{delivery_line}\n\n"
        f"📮 发货后我会第一时间为您同步物流单号及追踪信息，请您放心。"
    )

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": color, "title": {"content": header, "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": markdown}],
    }


# ===== 兜底卡片 =====

def _build_simple_card(title: str, text: str, color: str = "red") -> dict:
    return {
        "header": {"template": color, "title": {"content": title, "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": text}],
    }


# ===== 意图分类 =====

# 订单相关关键词（匹配任一即视为订单查询意图）
_ORDER_KEYWORDS = [
    r"[A-Za-z0-9\-_]{6,}",            # 合同编号格式
    r"订单|合同|项目|交期|发货|物流|送货|到货|签收|排产|排单",
    r"进度|跟踪|状态|批次|在途|运输|发出|收到|到达",
    r"什么时候.*到|何时.*发|查一下.*单|查.*进度|有没有.*发货",
    r"单号|编号|查询|查一查|查看|我的.*货|我们的.*单",
    r"项目|工程|设备|安装|调试|验收",
    r"合同号|合同.*编号|项目.*名称|发到哪里|走到哪|啥时候",
    r"帮我.*查|帮我.*看|请问.*订单|请问.*发货",
]

# 明确非订单相关的内容（闲聊、无关话题）
_OFF_TOPIC_PATTERNS = [
    r"^(hi|hello|你好|在吗|嗨|哈喽|hello)[\s!！。.]*$",
    r"^(早安|午安|晚安|早上好|下午好|晚上好|morning|evening)[\s!！。.]*$",
    r"^(谢谢|感谢|多谢|好的|ok|okay|收到|知道了|明白|懂了)[\s!！。.]*$",
    r"^(再见|拜拜|bye|see you|回头见)[\s!！。.]*$",
    r"^(你是谁|你叫什么|你是什么|你老板|谁做的|机器人)",
    r"^(今天)?天气|股票|新闻|今天.*星期|几点了|时间",
    r"^(你会|你能|你可以).*(写代码|聊天|唱歌|讲笑话|干嘛|做什么)",
    r"价格|多少钱|报价|报价单|优惠|折扣|便宜|议价|定价",
    r"招聘|面试|工资|待遇|请假|入职",
    r"^[A-Za-z]{2,10}$",             # 短英文（HAHAH, test, abc 等）
    r"^(测试|test|ceshi)$",
]

# 保留的旧闲聊模式（兼容用）
_CHITCHAT_PATTERNS = [
    r"^(hi|hello|你好|在吗|嗨|哈喽)[\s!！。.]*$",
    r"^(今天)?天气", r"^(你会|你能|你可以).*(写代码|聊天|唱歌|讲笑话)",
    r"^(测试|test)$",
    r"^(早安|午安|晚安|早上好|下午好|晚上好)$",
    r"^(谢谢|感谢|好的|ok|okay|收到|知道了)[\s!！。.]*$",
    r"^(你是谁|你叫什么|你是什么)", r"^(再见|拜拜|bye)$",
    r"^[A-Za-z]{2,8}$",
]


def _is_order_related(text: str) -> bool:
    """判断用户输入是否与订单/交付查询相关。"""
    cleaned = text.strip()
    # 消歧序号回复（0-9），直接放行
    if re.match(r"^[0-9]$", cleaned):
        return True
    # 先检查非订单相关模式
    for pat in _OFF_TOPIC_PATTERNS:
        if re.search(pat, cleaned, re.IGNORECASE):
            return False
    # 检查订单相关关键词
    for pat in _ORDER_KEYWORDS:
        if re.search(pat, cleaned, re.IGNORECASE):
            return True
    # 包含2个以上中文字符且不匹配无关模式 → 可能是项目名称/订单查询
    if len(re.findall(r"[一-鿿]", cleaned)) >= 2:
        return True
    # 短文本/单字母/纯英文 → 视为无关
    return False


def _is_chitchat(text: str) -> bool:
    for pat in _CHITCHAT_PATTERNS:
        if re.match(pat, text.strip(), re.IGNORECASE):
            return True
    return len(text.strip()) < 2


# ===== 主处理逻辑 =====

def process_message(open_id: str, user_input: str) -> dict:
    """处理单条用户消息，返回飞书互动卡片 dict。"""

    # 第一关：意图分类 — 是否与订单/交付相关
    if not _is_order_related(user_input):
        # 礼貌用语（谢谢/好的/收到/再见等）
        polite_patterns = [
            r"^(谢谢|感谢|多谢|好的|ok|okay|收到|知道了|明白|懂了)[\s!！。.]*$",
            r"^(再见|拜拜|bye)[\s!！。.]*$",
        ]
        is_polite = any(re.match(p, user_input.strip(), re.IGNORECASE) for p in polite_patterns)
        if is_polite:
            return _build_simple_card(
                "🤖 订单查询助手",
                "不客气！如需查询订单交付进度或物流信息，请随时告诉我您的**合同编号**或**项目名称**。",
                "blue",
            )
        # 测试/调试
        if re.match(r"^(测试|test|ceshi)$", user_input.strip(), re.IGNORECASE):
            return _build_simple_card(
                "🤖 订单查询助手",
                "您好，我是订单查询助手，系统运行正常。请发送**合同编号**或**项目名称**查询订单交付进度。",
                "blue",
            )
        # 其他无关内容 → 专业引导
        return _build_simple_card(
            "🤖 订单查询助手",
            "您好，我是订单查询助手，专注于为您提供订单交付进度与物流信息查询服务。\n\n"
            "请直接告诉我以下任一信息，我将为您查询最新的交付状态：\n"
            "• **合同编号**（如 HT-2024-001）\n"
            "• **项目名称**（如 城市商业综合体消控室项目）\n"
            "• 输入 **「查一下我的订单进度」**\n\n"
            "📌 如需价格咨询、技术方案等其他服务，请联系您的专属销售顾问。",
            "blue",
        )

    # 检查是否在消歧会话中（用户回复了序号）
    if open_id in _session_cache:
        candidates = _session_cache.pop(open_id)
        try:
            idx = int(user_input.strip()) - 1
            if 0 <= idx < len(candidates):
                contract_id = candidates[idx]["合同编号"]
                detail = _load_order_detail(contract_id)
                if detail:
                    return _build_delivery_card(detail)
                return _build_simple_card(
                    "📋 信息确认中",
                    "该订单的交付计划正在确认中，待批次确认后将第一时间为您同步预计发出时间。如需加急请联系您的专属销售顾问。",
                    "blue",
                )
            else:
                return _build_simple_card(
                    "👋 已取消",
                    "如需进一步查询订单信息，请随时发送项目名称或合同编号。",
                    "blue",
                )
        except ValueError:
            pass  # 不是数字，继续正常查询流程

    # 第一层：身份提取
    company = _extract_company_name(open_id)

    # 加载订单子集。外部联系人可能取不到 company_name，允许用合同号/项目关键词兜底查询。
    orders = _load_company_orders(company or "")
    if not orders:
        # 如果有公司身份但公司过滤后为空 → 可能是飞书企业名与合同客户名不一致 → 兜底全量查询
        if company:
            orders = _load_company_orders("")
        if not orders:
            if not company:
                return _build_simple_card(
                    "🔐 身份确认中",
                    "您好，系统暂未识别到您的企业信息。\n\n"
                    "请直接发送以下任一信息，我将立即为您查询：\n"
                    "• **合同编号**（如 HT-2024-001）\n"
                    "• **项目名称**\n\n"
                    "收到后我会第一时间为您反馈交付进度。",
                    "blue",
                )
            return _build_simple_card(
                "📭 暂无进行中的订单",
                f"已识别贵司「{company}」，当前系统中暂无进行中的订单。\n\n"
                "如您已签订新合同但查询不到信息，请联系您的专属销售顾问协助确认。",
                "blue",
            )

    # 第二层：模糊匹配
    candidates = _fuzzy_match_orders(orders, user_input)
    if not candidates:
        scope_text = f"贵司「{company}」的订单" if company else "当前订单库"
        return _build_simple_card(
            "🔍 未找到匹配项目",
            f"在{scope_text}中未找到与您提供的信息相匹配的订单。\n\n"
            "建议您尝试：\n"
            "• 输入完整的**合同编号**（如 HT-2024-001）\n"
            "• 输入准确的**项目名称**关键词\n\n"
            "如需其他帮助，请联系您的专属销售顾问。",
            "blue",
        )

    # 第三层：多条消歧
    if len(candidates) >= 2:
        return _build_disambiguation_card(candidates, open_id)

    # 第四层：精准交付
    contract_id = candidates[0]["合同编号"]
    detail = _load_order_detail(contract_id)
    if not detail:
        proj_name = candidates[0].get("项目名称", "")
        return _build_simple_card(
            "📋 订单信息确认中",
            f"已找到项目「{proj_name}」，目前该订单的交付计划正在确认中。\n\n"
            "待交付批次确认后，我将第一时间为您同步预计发出时间。\n\n"
            "如需加急处理，请联系您的专属销售顾问。",
            "blue",
        )
    return _build_delivery_card(detail)


# ===== 发送回复 =====

def _send_reply(open_id: str, card: dict, chat_id: str = "", chat_type: str = ""):
    """发送回复，群聊自动切换为 chat_id 回复。"""
    token = _get_bot_token()
    if chat_type == "group" and chat_id:
        receive_id_type = "chat_id"
        receive_id = chat_id
    else:
        receive_id_type = "open_id"
        receive_id = open_id
    resp = httpx.post(
        f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": receive_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        },
        timeout=15.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"发送回复失败: {data}")


def _extract_message_from_event(event: dict) -> tuple:
    """兼容 lark-cli 扁平事件与飞书 V2 envelope，返回 (open_id, chat_id, chat_type, message_type, text)。"""
    event_id = str(event.get("event_id") or event.get("header", {}).get("event_id") or "")
    if event_id:
        if event_id in _seen_event_ids:
            return "", "", "", "", ""
        _seen_event_ids.add(event_id)

    # 公共字段
    chat_id = str(event.get("chat_id") or "")
    chat_type = str(event.get("chat_type") or "")

    # lark-cli event schema im.message.receive_v1 当前输出为扁平结构：
    # {sender_id, message_type, content, chat_id, chat_type, ...}
    if "sender_id" in event or "message_type" in event:
        open_id = str(event.get("sender_id") or "")
        msg_type = str(event.get("message_type") or "text")
        content = event.get("content") or ""
        return open_id, chat_id, chat_type, msg_type, str(content)

    payload = event.get("event", {})
    message = payload.get("message", {})
    sender = payload.get("sender", {}).get("sender_id", {})
    open_id = sender.get("open_id", "")
    msg_type = message.get("msg_type", "text")
    content_str = message.get("content", "{}")
    chat_id_v2 = message.get("chat_id", "")
    if not chat_id and chat_id_v2:
        chat_id = chat_id_v2

    try:
        content = json.loads(content_str)
    except json.JSONDecodeError:
        content = {}

    user_input = ""
    if msg_type == "text":
        user_input = content.get("text", "")
    elif msg_type == "post":
        for block in content.get("content", []):
            for elem in block:
                if elem.get("tag") == "text":
                    user_input += elem.get("text", "")
    return open_id, chat_id, chat_type, msg_type, user_input


def _strip_at_mention(text: str) -> str:
    """移除群聊中的 @ 提及（@所有人、@机器人等），返回纯净用户输入。"""
    import re as _re
    t = _re.sub(r"@\S+", "", text).strip()
    t = _re.sub(r"@所有人", "", t).strip()
    return t


# ===== 双 App 长连接事件主循环 =====

# 主 App（供应链AI助手）的回复配置
MAIN_APP_ID = os.getenv("FEISHU_APP_ID", "")
MAIN_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
MAIN_APP_TOKEN_CACHE: Dict[str, tuple] = {}


def _get_main_app_token() -> str:
    """获取主 App（供应链AI助手）的 token。"""
    if not MAIN_APP_SECRET:
        raise Exception("缺少 FEISHU_APP_SECRET")
    cached = MAIN_APP_TOKEN_CACHE.get("token")
    if cached and cached[1] > datetime.now(CST) + timedelta(minutes=5):
        return cached[0]
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": MAIN_APP_ID, "app_secret": MAIN_APP_SECRET}, timeout=30.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取主App Token 失败: {data}")
    token = data["tenant_access_token"]
    MAIN_APP_TOKEN_CACHE["token"] = (token, datetime.now(CST) + timedelta(seconds=6600))
    return token


def _send_main_app_reply(open_id: str, card: dict, chat_id: str = "", chat_type: str = ""):
    """以供应链AI助手身份发送回复，群聊自动切换为 chat_id 回复。"""
    token = _get_main_app_token()
    if chat_type == "group" and chat_id:
        receive_id_type = "chat_id"
        receive_id = chat_id
    else:
        receive_id_type = "open_id"
        receive_id = open_id
    resp = httpx.post(
        f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": receive_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        },
        timeout=15.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        print(f"  [供应链助手回复] 失败: {data}")
        return False
    return True


def _run_event_loop(lark_cli_path: str, profile_args: List[str], label: str,
                    handler_func, reply_func):
    """单个 App 的事件消费循环（在独立线程中运行），同时支持 p2p 和群聊。"""
    print(f"[{label}] 启动长连接...")

    while True:
        try:
            cmd = [lark_cli_path] + profile_args + ["event", "consume", "im.message.receive_v1", "--as", "bot"]
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                bufsize=1,
            )

            ready_event = threading.Event()

            def _read_stderr():
                try:
                    for line in iter(proc.stderr.readline, ""):
                        stripped = line.rstrip("\n").rstrip("\r")
                        if stripped:
                            print(f"  [{label}] {stripped}")
                        if "[event] ready" in stripped:
                            ready_event.set()
                except Exception:
                    pass

            t = threading.Thread(target=_read_stderr, daemon=True)
            t.start()

            if not ready_event.wait(timeout=30):
                print(f"[{label}] 启动超时，5秒后重试...")
                proc.kill()
                proc.wait()
                time.sleep(5)
                continue

            if proc.poll() is not None:
                print(f"[{label}] 异常退出 (code={proc.returncode})，5秒后重试...")
                time.sleep(5)
                continue

            print(f"[{label}] 就绪，接收事件中（支持 p2p + 群聊）")

            _msg_count = 0
            for line in iter(proc.stdout.readline, ""):
                stripped = line.strip()
                if not stripped:
                    continue
                try:
                    event = json.loads(stripped)
                except json.JSONDecodeError:
                    continue

                chat_type = event.get("chat_type", "")
                # 支持 p2p 和 group 两种聊天类型
                if chat_type not in ("p2p", "group"):
                    continue

                open_id, chat_id, _ct, msg_type, user_input = _extract_message_from_event(event)
                if not open_id or not user_input.strip():
                    continue
                if msg_type != "text":
                    continue

                # 群聊中剥离 @提及
                if chat_type == "group":
                    user_input = _strip_at_mention(user_input)
                    if not user_input.strip():
                        continue

                _msg_count += 1
                now_str = datetime.now(CST).strftime("%H:%M:%S")
                chat_label = "群聊" if chat_type == "group" else "私聊"
                print(f"[{label}] {now_str} [{chat_label}] #{_msg_count} {open_id[:12]}...: {user_input[:60]}")

                try:
                    card = handler_func(open_id, user_input)
                    reply_func(open_id, card, chat_id=chat_id, chat_type=chat_type)
                    print(f"[{label}]  -> 已回复")
                except Exception as e:
                    print(f"[{label}]  [X] 处理异常: {e}")

            ret = proc.wait()
            print(f"[{label}] 进程退出 (code={ret})，本会话收到 {_msg_count} 条消息，5秒后重连...")
            time.sleep(5)

        except Exception as e:
            print(f"[{label}] 异常: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(5)


# ===== 飞书文档扫描与 AI 加工 =====

_FEISHU_DOC_URL_RE = re.compile(
    r"https?://[^\s/]+\.(feishu|lark)\.(cn|com)/(docx|wiki|minutes)/([A-Za-z0-9_\-]+)",
    re.IGNORECASE,
)

_DEEPSEEK_CLIENT_CACHE: Any = None


def _get_deepseek_client():
    global _DEEPSEEK_CLIENT_CACHE
    if _DEEPSEEK_CLIENT_CACHE is None:
        api_key = os.getenv("DEEPSEEK_API_KEY", "")
        if api_key:
            from openai import OpenAI
            _DEEPSEEK_CLIENT_CACHE = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    return _DEEPSEEK_CLIENT_CACHE


def _scan_feishu_doc(doc_url: str) -> Optional[str]:
    """通过 lark-cli 读取飞书文档的完整文本内容。支持 docx/wiki/minutes。"""
    # 拆分 URL 获取 token
    corpus = doc_url.split("?")[0]
    parts = corpus.rstrip("/").split("/")
    if len(parts) < 2:
        return None
    token = parts[-1]
    if not re.fullmatch(r"[A-Za-z0-9_\-]+", token):
        return None

    lark_cli = _find_lark_cli()

    # 读取文档纯文本
    try:
        result = subprocess.run(
            [lark_cli, "--profile", "main-app", "docs", "+fetch", "--token", token, "--format", "text"],
            capture_output=True,
            encoding="utf-8",
            timeout=60,
        )
        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip()
            print(f"  [文档扫描] lark-cli 读取失败: {err[:200]}")
            return None
        content = result.stdout.strip()
        if content and len(content) >= 10:
            return content
    except subprocess.TimeoutExpired:
        print(f"  [文档扫描] lark-cli 读取超时")
    except Exception as e:
        print(f"  [文档扫描] 异常: {e}")
    return None


def _ai_process_doc(content: str, user_question: str) -> str:
    """用 DeepSeek 对文档内容进行加工，回答用户问题或生成摘要。"""
    client = _get_deepseek_client()
    if client is None:
        return "（文档已读取，但 DeepSeek API Key 未配置，无法智能加工）"

    # 截断超长文档，保留前 12000 字符
    truncated = content[:12000]
    if len(content) > 12000:
        truncated += "\n\n（文档较长，以上为节选）"

    prompt = f"""你是供应链文档分析助手。根据以下飞书文档内容，回答用户的问题。

用户问题：{user_question}

文档内容：
{truncated}

要求：
1. 回答控制在 200-400 字
2. 如果文档内容与问题不相关，如实告知
3. 关键数字和数据要原样引用，不要编造
4. 用 Markdown 格式输出，适当使用粗体和列表"""

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=2048,
            timeout=60,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"（AI 加工失败: {e}）\n\n文档原文摘要：{truncated[:500]}..."


def _handle_doc_scan(user_input: str) -> Optional[dict]:
    """检测消息中是否包含飞书文档链接，如是则读取并加工后返回卡片。"""
    match = _FEISHU_DOC_URL_RE.search(user_input)
    if not match:
        return None

    doc_url = match.group(0)
    print(f"  [文档扫描] 检测到飞书文档: {doc_url}")

    # 提取用户的具体问题（去掉 URL 后的剩余文本）
    question = _FEISHU_DOC_URL_RE.sub("", user_input).strip()
    if not question or len(question) < 2:
        question = "请总结这份文档的核心内容和关键数据"

    # 读取文档
    content = _scan_feishu_doc(doc_url)
    if content is None:
        return {
            "header": {"template": "red", "title": {"content": "📄 文档读取失败", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": (
                f"无法读取该飞书文档：{doc_url}\n\n"
                "可能原因：\n"
                "• 文档权限限制（需确保供应链AI助手有访问权限）\n"
                "• 链接格式不正确\n"
                "• 文档类型暂不支持\n\n"
                "请确认文档链接或尝试发送文档内容截图。"
            )}],
        }

    print(f"  [文档扫描] 文档长度: {len(content)} 字符")

    # AI 加工
    result = _ai_process_doc(content, question)

    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": "blue", "title": {"content": "📄 文档智能分析", "tag": "plain_text"}},
        "elements": [{"tag": "markdown", "content": result}],
    }


def _supply_chain_handler(open_id: str, user_input: str) -> dict:
    """供应链AI助手消息处理（延迟导入 scheduler_api 模块以复用逻辑）。

    优先检测飞书文档链接，其次交给 scheduler_api 做供应链关键词匹配。
    """
    # 飞书文档扫描优先
    doc_card = _handle_doc_scan(user_input)
    if doc_card is not None:
        return doc_card

    try:
        from scheduler_api import _build_supply_chain_reply
        return _build_supply_chain_reply(open_id, user_input)
    except ImportError:
        # 回退：简单的引导卡片
        return {
            "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": (
                "我是供应链AI助手。\n\n"
                "📊 发送 **日报** / **库存** / **待确认** / **延迟** 查询相关信息\n"
                "🔍 发送**合同编号**查询订单详情\n"
                "📄 发送**飞书文档链接**让我帮你分析和总结文档内容"
            )}],
        }


def main():
    print("=" * 55)
    print("  双机器人长连接中枢")
    print(f"  启动时间：{datetime.now(CST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  订单查询助手 App：{APP_ID}")
    print(f"  供应链AI助手 App：{MAIN_APP_ID}")
    print("=" * 55)

    lark_cli = _find_lark_cli()
    print(f"  lark-cli: {lark_cli}")
    print(f"  支持消息类型: 私聊(p2p) + 群聊(group)")
    print(f"  飞书文档扫描: 已启用（docx/wiki/minutes）")

    # 事件订阅诊断
    try:
        result = subprocess.run(
            [lark_cli, "event", "status"], capture_output=True, encoding="utf-8", timeout=15,
        )
        if result.returncode == 0 and result.stdout.strip():
            print(f"\n  [诊断] 事件总线状态:")
            for line in result.stdout.strip().split("\n")[:8]:
                if line.strip():
                    print(f"    {line.strip()}")
            print()
    except Exception:
        pass

    print("[OK] 启动双路长连接，等待消息...\n")

    # 订单查询助手（默认 profile = 客服机器人）
    t1 = threading.Thread(
        target=_run_event_loop,
        args=(lark_cli, [], "订单查询助手", process_message, _send_reply),
        daemon=True,
    )
    t1.start()

    # 供应链AI助手（--profile main-app）
    t2 = threading.Thread(
        target=_run_event_loop,
        args=(lark_cli, ["--profile", "main-app"], "供应链AI助手",
              _supply_chain_handler, _send_main_app_reply),
        daemon=True,
    )
    t2.start()

    print("[OK] 两个机器人都已启动\n")

    try:
        while t1.is_alive() or t2.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[OK] 正常退出")


if __name__ == "__main__":
    main()
