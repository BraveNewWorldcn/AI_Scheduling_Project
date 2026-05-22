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
            if company_name and _fuzzy_score(company_name, cust) < 0.3:
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
    noise = ["帮我", "查一下", "那个", "这个", "什么", "怎么", "进度", "单子",
             "订单", "一下", "请问", "我想", "我要", "看看", "有没有", "项目"]
    cleaned = user_input
    for w in noise:
        cleaned = cleaned.replace(w, " ")
    tokens = [t.strip() for t in re.split(r"[\s,，、]+", cleaned) if len(t.strip()) >= 2]
    return tokens[:5]


def _fuzzy_match_orders(orders: List[Dict[str, str]], user_input: str) -> List[Dict[str, Any]]:
    query = user_input.strip().lower()
    exact_contracts = [
        {**order, "_score": 1.0}
        for order in orders
        if order.get("合同编号", "").strip().lower() and order.get("合同编号", "").strip().lower() in query
    ]
    if exact_contracts:
        return exact_contracts[:1]

    keywords = _extract_keywords(user_input)
    scored: List[tuple] = []
    for order in orders:
        proj = order.get("项目名称", "")
        contract = order.get("合同编号", "")
        score = _fuzzy_score(user_input, f"{proj} {contract}")
        for kw in keywords:
            if kw.lower() in proj.lower():
                score += 0.3
            if kw.lower() in contract.lower():
                score += 0.2
        if score > 0.15:
            scored.append((min(score, 1.0), {**order, "_score": round(min(score, 1.0), 2)}))
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
            f"目前设备正在按计划为您有序排产备货中。"
        )
        detail_line = "核心模块正在进行严密的出厂前质量验证与软硬件联调测试，确保设备到场后稳定运行。"
        if ship_date_raw:
            delivery_line = (
                f"您的专属交付批次已锁定，我们将确保在 **{ship_date_display}** 为您准时安排发出。"
            )
        else:
            delivery_line = "您的专属交付批次确认后，我会第一时间为您同步预计发出时间。"
    else:
        status_line = (
            f"关于您关注的【**{project}**】项目（合同编号：{contract}），一切进展顺利。"
        )
        detail_line = "设备正在完成出厂前的最终调试与防震包装，确保运输安全与到场即用。"
        if ship_date_raw:
            delivery_line = (
                f"您的专属交付批次已排定，预计 **{ship_date_display}** 为您准时发出。"
            )
        else:
            delivery_line = "您的专属交付批次确认后，我会第一时间为您同步预计发出时间。"

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


# ===== 闲聊检测 =====

_CHITCHAT_PATTERNS = [
    r"^(hi|hello|你好|在吗|嗨|哈喽)[\s!！。.]*$",
    r"^(今天)?天气", r"^(你会|你能|你可以).*(写代码|聊天|唱歌|讲笑话)",
    r"^(测试|test)$",
]


def _is_chitchat(text: str) -> bool:
    for pat in _CHITCHAT_PATTERNS:
        if re.match(pat, text.strip(), re.IGNORECASE):
            return True
    return len(text.strip()) < 2


# ===== 主处理逻辑 =====

def process_message(open_id: str, user_input: str) -> dict:
    """处理单条用户消息，返回飞书互动卡片 dict。"""

    # 闲聊拦截
    if _is_chitchat(user_input):
        return _build_simple_card(
            "🤖 供应链物流管家",
            "我是您的供应链物流管家，只能帮您查询项目订单交期与物流。请问您的项目名称是什么？",
            "red",
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
                return _build_simple_card("📦 暂无数据", "该订单暂无履约数据，请联系销售。", "orange")
            else:
                return _build_simple_card("👋 已取消", "如需帮助请联系您的专属销售。", "blue")
        except ValueError:
            pass  # 不是数字，继续正常查询流程

    # 第一层：身份提取
    company = _extract_company_name(open_id)

    # 加载订单子集。外部联系人可能取不到 company_name，允许用合同号/项目关键词兜底查询。
    orders = _load_company_orders(company or "")
    if not orders:
        if not company:
            return _build_simple_card(
                "🔐 身份未识别",
                "抱歉，未识别到您的企业身份。请直接发送合同编号或项目关键词，我会继续帮您查找。",
                "orange",
            )
        return _build_simple_card(
            "📭 暂无在途订单",
            f"已识别「{company}」，但当前系统中暂无在途订单。如需帮助请联系销售。",
            "blue",
        )

    # 第二层：模糊匹配
    candidates = _fuzzy_match_orders(orders, user_input)
    if not candidates:
        scope_text = f"「{company}」的订单" if company else "当前订单库"
        return _build_simple_card(
            "🔍 未找到匹配项目",
            f"在{scope_text}中未找到与「{user_input}」匹配的项目。请尝试输入合同编号或更完整的项目关键词。",
            "orange",
        )

    # 第三层：多条消歧
    if len(candidates) >= 2:
        return _build_disambiguation_card(candidates, open_id)

    # 第四层：精准交付
    contract_id = candidates[0]["合同编号"]
    detail = _load_order_detail(contract_id)
    if not detail:
        return _build_simple_card(
            "📦 暂无履约数据",
            f"找到「{candidates[0].get('项目名称','?')}」，但暂无履约数据。",
            "orange",
        )
    return _build_delivery_card(detail)


# ===== 发送回复 =====

def _send_reply(open_id: str, card: dict):
    token = _get_bot_token()
    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id"
    resp = httpx.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": open_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        },
        timeout=15.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"发送回复失败: {data}")


def _extract_message_from_event(event: dict) -> tuple[str, str, str]:
    """兼容 lark-cli 扁平事件与飞书 V2 envelope，返回 open_id/message_type/text。"""
    event_id = str(event.get("event_id") or event.get("header", {}).get("event_id") or "")
    if event_id:
        if event_id in _seen_event_ids:
            return "", "", ""
        _seen_event_ids.add(event_id)

    # lark-cli event schema im.message.receive_v1 当前输出为扁平结构：
    # {sender_id, message_type, content, ...}，其中 content 对 text 已预渲染为纯文本。
    if "sender_id" in event or "message_type" in event:
        open_id = str(event.get("sender_id") or "")
        msg_type = str(event.get("message_type") or "text")
        content = event.get("content") or ""
        return open_id, msg_type, str(content)

    payload = event.get("event", {})
    message = payload.get("message", {})
    sender = payload.get("sender", {}).get("sender_id", {})
    open_id = sender.get("open_id", "")
    msg_type = message.get("msg_type", "text")
    content_str = message.get("content", "{}")

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
    return open_id, msg_type, user_input


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


def _send_main_app_reply(open_id: str, card: dict):
    """以供应链AI助手身份发送回复。"""
    token = _get_main_app_token()
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": open_id,
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
    """单个 App 的事件消费循环（在独立线程中运行）。"""
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

            print(f"[{label}] 就绪，接收事件中")

            for line in iter(proc.stdout.readline, ""):
                stripped = line.strip()
                if not stripped:
                    continue
                try:
                    event = json.loads(stripped)
                except json.JSONDecodeError:
                    continue

                chat_type = event.get("chat_type", "")
                if chat_type != "p2p":
                    continue

                open_id, msg_type, user_input = _extract_message_from_event(event)
                if not open_id or not user_input.strip():
                    continue
                if msg_type != "text":
                    continue

                now_str = datetime.now(CST).strftime("%H:%M:%S")
                print(f"[{label}] {now_str} {open_id[:12]}...: {user_input[:60]}")

                try:
                    card = handler_func(open_id, user_input)
                    reply_func(open_id, card)
                    print(f"[{label}]  -> 已回复")
                except Exception as e:
                    print(f"[{label}]  [X] 处理异常: {e}")

            ret = proc.wait()
            print(f"[{label}] 进程退出 (code={ret})，5秒后重连...")
            time.sleep(5)

        except Exception as e:
            print(f"[{label}] 异常: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(5)


def _supply_chain_handler(open_id: str, user_input: str) -> dict:
    """供应链AI助手消息处理（延迟导入 scheduler_api 模块以复用逻辑）。"""
    try:
        from scheduler_api import _build_supply_chain_reply
        return _build_supply_chain_reply(open_id, user_input)
    except ImportError:
        # 回退：简单的引导卡片
        return {
            "header": {"template": "blue", "title": {"content": "🏭 供应链AI助手", "tag": "plain_text"}},
            "elements": [{"tag": "markdown", "content": (
                "我是供应链AI助手。\n\n📊 发送 **日报** / **库存** / **待确认** / **延迟** 查询相关信息\n🔍 发送**合同编号**查询订单详情"
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
