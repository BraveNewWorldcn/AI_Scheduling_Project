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
import subprocess
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional
from difflib import SequenceMatcher

import httpx

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
APP_TOKEN = os.getenv("BITABLE_APP_TOKEN", "")
TABLE_ID_MAIN = os.getenv("TABLE_ID_MAIN", "tbl06oxGEdMNTEB8")
TABLE_ID_DETAIL = os.getenv("TABLE_ID_DETAIL", "tbl09Z6C7wCGh3mW")

LARK_CLI = os.path.expandvars(r"%APPDATA%\npm\lark-cli.cmd")
if not os.path.isfile(LARK_CLI):
    LARK_CLI = "lark-cli"

if not APP_SECRET:
    sys.exit("[X] 请先设置 FEISHU_APP_SECRET")

# ===== 会话缓存（消歧卡片选择） =====
# {open_id: [{"合同编号":..., "项目名称":...}, ...]}
_session_cache: Dict[str, List[Dict[str, str]]] = {}


# ===== 工具函数 =====

def _get_bot_token() -> str:
    resp = httpx.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=30.0,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取 Token 失败: {data}")
    return data["tenant_access_token"]


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
    token = _get_bot_token()
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
            break
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            cust = _parse_feishu_value(f.get("客户名称"))
            if _fuzzy_score(company_name, cust) < 0.3:
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
    token = _get_bot_token()
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
            break
        for item in data.get("data", {}).get("items", []):
            f = item.get("fields", {})
            if _parse_feishu_value(f.get("合同编号")) == contract_id:
                return {
                    "合同编号": contract_id,
                    "项目名称": _parse_feishu_value(f.get("项目名称")),
                    "AI建议发货时间": _parse_feishu_value(f.get("AI建议发货时间")),
                    "人工确认发货时间": _parse_feishu_value(f.get("人工确认发货时间")),
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
    status = order.get("整体状态", "未知")
    order_status = order.get("订单状态", "待确认")
    ai_ship = order.get("AI建议发货时间", "-")
    manual_ship = order.get("人工确认发货时间", "")
    ai_risk = order.get("AI风险", "")
    ai_advice = order.get("AI建议", "")

    if status == "缺货":
        color, status_text = "red", "🔴 缺货待补"
    elif order_status == "已发货":
        color, status_text = "green", "🟢 已发货"
    elif order_status == "已确认":
        color, status_text = "blue", "🔵 已确认待发货"
    else:
        color, status_text = "orange", "🟠 AI 排单中"

    ship_date = manual_ship or ai_ship or "待定"
    markdown = (
        f"**合同编号**：{contract}\n"
        f"**项目名称**：{project}\n"
        f"**当前状态**：{status_text}\n"
        f"**预计发货**：{ship_date}\n"
    )
    if ai_risk and ai_risk != "无明显风险":
        markdown += f"**风险提示**：{ai_risk}\n"
    if ai_advice:
        markdown += f"**AI 建议**：{ai_advice}\n"

    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": color,
            "title": {"content": "📦 订单履约状态", "tag": "plain_text"},
        },
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
    if not company:
        return _build_simple_card(
            "🔐 身份未识别",
            "抱歉，未识别到您的企业身份。请联系您的专属销售进行信息绑定。",
            "red",
        )

    # 加载公司订单子集
    orders = _load_company_orders(company)
    if not orders:
        return _build_simple_card(
            "📭 暂无在途订单",
            f"已识别「{company}」，但当前系统中暂无在途订单。如需帮助请联系销售。",
            "blue",
        )

    # 第二层：模糊匹配
    candidates = _fuzzy_match_orders(orders, user_input)
    if not candidates:
        return _build_simple_card(
            "🔍 未找到匹配项目",
            f"在「{company}」的订单中未找到与「{user_input}」匹配的项目。请尝试输入项目关键词。",
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
    httpx.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "receive_id": open_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        },
        timeout=15.0,
    )


# ===== 长连接事件主循环 =====

def main():
    print("=" * 55)
    print("  客户查单中枢 (长连接模式)")
    print(f"  启动时间：{datetime.now(CST).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  应用 ID：{APP_ID}")
    print("  监听事件：im.message.receive_v1")
    print("=" * 55)

    cmd = [LARK_CLI, "event", "consume", "im.message.receive_v1", "--as", "bot"]

    env = os.environ.copy()
    env["LARK_APP_ID"] = APP_ID
    env["LARK_APP_SECRET"] = APP_SECRET

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        encoding="utf-8",
        bufsize=1,
        env=env,
    )

    print("[OK] 长连接已建立，等待客户消息...\n")

    try:
        assert proc.stdout is not None
        for line in proc.stdout:
            line = line.strip()
            if not line:
                continue

            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue

            # 提取消息内容
            message = event.get("event", {}).get("message", {})
            sender = event.get("event", {}).get("sender", {}).get("sender_id", {})
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

            if not user_input.strip() or not open_id:
                continue

            now = datetime.now(CST).strftime("%H:%M:%S")
            print(f"[{now}] {open_id[:12]}...: {user_input[:60]}")

            try:
                card = process_message(open_id, user_input)
                _send_reply(open_id, card)
                print(f"  -> 已回复")
            except Exception as e:
                print(f"  [X] 处理异常: {e}")

    except KeyboardInterrupt:
        print("\n[OK] 正常退出")
    finally:
        proc.terminate()
        proc.wait()


if __name__ == "__main__":
    main()
