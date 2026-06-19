"""
顺丰物流模块 — 丰桥开放平台 API 对接

功能：
  1. 运单查询（EXP_RECE_QUERY_SFWAYBILL）— 按单号查物流轨迹
  2. /shipping/* API 路由（挂载到主 FastAPI app）
  3. 物流状态定时轮询（每 2 小时，8:00-20:00）
  4. 飞书发货总表自动回写

凭证从 .env 读取：
  SF_PARTNER_ID      顾客编码
  SF_CHECKWORD       校验码
  SF_MONTHLY_CARD    月结卡号
  SF_SERVER_URL      API地址（沙箱: sfapi-sbox.sf-express.com / 生产: sfapi.sf-express.com）

项目地址格式：收件人，电话，详细地址【项目：项目名称】
"""

from __future__ import annotations

import hashlib
import base64
import json
import logging
import os
import re
import sqlite3
import threading
import time
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional

import httpx
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler

logger = logging.getLogger("sf_shipping")

# =========================
# 环境变量
# =========================
SF_PARTNER_ID   = os.getenv("SF_PARTNER_ID", "")
SF_CHECKWORD    = os.getenv("SF_CHECKWORD", "")
SF_MONTHLY_CARD = os.getenv("SF_MONTHLY_CARD", "")
SF_SERVER_URL   = os.getenv("SF_SERVER_URL", "https://sfapi-sbox.sf-express.com/std/service")


# =========================
# 请求模型
# =========================
class ManualOrderRequest(BaseModel):
    """手动录入单号"""
    tracking_numbers: List[str]  # 可一次性填多个单号
    project_name: str = ""


class TrackingQueryRequest(BaseModel):
    """按条件查询物流"""
    tracking_numbers: Optional[List[str]] = None  # 指定单号列表
    status_filter: str = "all"  # all / active / signed


# =========================
# 物流状态映射（opCode → 展示状态）
# =========================
OPCODE_STATUS_MAP = {
    "50": "已揽收",   # 已收件
    "30": "运输中",   # 运输中
    "31": "运输中",   # 到达中转
    "36": "运输中",   # 离开中转
    "44": "派送中",   # 派送中
    "204": "派送中",   # 派送中（新版）
    "80": "已签收",   # 已签收
    "99": "异常",     # 异常
    "92": "异常",     # 退件
}

# firstStatusName 后备映射
STATUS_NAME_MAP = {
    "已揽收": "已揽收",
    "运送中": "运输中",
    "派送中": "派送中",
    "已签收": "已签收",
}


def _map_status(op_code: str, first_status_name: str = "") -> str:
    """将 opCode + firstStatusName 映射为展示状态"""
    if op_code in OPCODE_STATUS_MAP:
        return OPCODE_STATUS_MAP[op_code]
    if first_status_name in STATUS_NAME_MAP:
        return STATUS_NAME_MAP[first_status_name]
    return "运输中"


def extract_project_from_address(address: str) -> str:
    """从地址中提取【项目：xxx】中的项目名"""
    m = re.search(r"【项目[：:](.*?)】", address)
    if m:
        return m.group(1).strip()
    m = re.search(r"【(.*?)】", address)
    if m:
        return m.group(1).strip()
    return ""


# =========================
# SF API Client
# =========================
class SFClient:
    """顺丰丰桥 API 客户端"""

    def __init__(self, partner_id: str, checkword: str, monthly_card: str, server_url: str):
        self.partner_id = partner_id
        self.checkword = checkword
        self.monthly_card = monthly_card
        self.server_url = server_url

    def _sign(self, msg_data: str, timestamp: str) -> str:
        """MD5 签名：Base64(MD5(msgData + timestamp + checkWord))  不 URL 编码"""
        raw = msg_data + timestamp + self.checkword
        md5_digest = hashlib.md5(raw.encode("utf-8")).digest()
        return base64.b64encode(md5_digest).decode("utf-8")

    def _call(self, service_code: str, msg_data: dict) -> dict:
        """通用 API 调用"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")[:-3]
        msg_data_str = json.dumps(msg_data, ensure_ascii=False)
        msg_digest = self._sign(msg_data_str, timestamp)

        headers = {
            "appCode": self.partner_id,
            "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
            "timestamp": timestamp,
        }
        form = {
            "partnerID": self.partner_id,
            "requestID": uuid.uuid4().hex,
            "serviceCode": service_code,
            "timestamp": timestamp,
            "msgData": msg_data_str,
            "msgDigest": msg_digest,
        }

        try:
            resp = httpx.post(self.server_url, headers=headers, data=form, timeout=30.0)
            result = resp.json()
            return result
        except httpx.TimeoutException:
            logger.error(f"SF API [{service_code}] timeout")
            raise HTTPException(status_code=504, detail="顺丰接口超时，请稍后重试")
        except Exception as e:
            logger.error(f"SF API [{service_code}] error: {e}")
            raise HTTPException(status_code=502, detail=f"顺丰接口异常: {e}")

    # ---- SFWAYBILL 运单查询（替代路由查询） ----
    def query_sfwaybill(self, search_no: str, search_type: str = "1") -> dict:
        """
        清单运费查询 / 运单轨迹查询
        :param search_no:   运单号 或 订单号
        :param search_type:  1=运单号  2=订单号
        :return: API 响应（apiResultData 包含 JSON 字符串）
        """
        msg_data = {
            "language": "0",
            "searchType": search_type,
            "searchNo": search_no,
        }
        return self._call("EXP_RECE_QUERY_SFWAYBILL", msg_data)


def parse_sfwaybill_response(api_response: dict) -> dict:
    """
    解析 QUERY_SFWAYBILL 返回，提取物流轨迹
    返回格式兼容两种查询模式：
      - 按运单号：apiResultData.msgData.routeResps[].routes[]
      - 按月结卡号：apiResultData.msgData.waybillInfo（仅基本信息）
    """
    default = {
        "status": "未知",
        "detail": "",
        "update_time": "",
        "project_name": "",
        "routes": [],
    }

    if api_response.get("apiResultCode") != "A1000":
        return default

    # apiResultData 是 JSON 字符串，需要二次解析
    result_data_str = api_response.get("apiResultData", "{}")
    try:
        result_data = json.loads(result_data_str) if isinstance(result_data_str, str) else result_data_str
    except (json.JSONDecodeError, TypeError):
        return default

    if not result_data.get("success"):
        error_msg = result_data.get("errorMsg", "")
        if "找不到" in str(error_msg):
            return {"status": "未知", "detail": f"未查到运单信息", "update_time": "", "project_name": "", "routes": []}
        return default

    msg_data = result_data.get("msgData", {})

    # 尝试解析 routeResps（运单号查询）
    route_resps = msg_data.get("routeResps", [])
    if route_resps and isinstance(route_resps, list) and len(route_resps) > 0:
        first = route_resps[0]
        mail_no = first.get("mailNo", "")
        routes = first.get("routes", [])
        if routes:
            latest = routes[-1]
            op_code = str(latest.get("opCode", ""))
            first_status = latest.get("firstStatusName", "")
            status = _map_status(op_code, first_status)
            detail = latest.get("remark", "")
            update_time = latest.get("acceptTime", "")

            return {
                "status": status,
                "detail": detail,
                "update_time": update_time,
                "project_name": "",
                "mail_no": mail_no,
                "routes": [
                    {
                        "time": r.get("acceptTime", ""),
                        "remark": r.get("remark", ""),
                        "status": _map_status(str(r.get("opCode", "")), r.get("firstStatusName", "")),
                        "address": r.get("acceptAddress", ""),
                    }
                    for r in routes
                ],
            }

    # 按月结卡号查询 — 返回运单基本信息（无路由轨迹）
    waybill_info = msg_data.get("waybillInfo")
    if waybill_info:
        address = waybill_info.get("addresseeAddr", "")
        project_name = extract_project_from_address(address)
        return {
            "status": "已揽收",
            "detail": "",
            "update_time": "",
            "project_name": project_name,
            "recipient": waybill_info.get("addresseeContName", ""),
            "recipient_phone": waybill_info.get("addresseeMobile", ""),
            "address": address,
            "sender": waybill_info.get("consignorContName", ""),
            "sender_address": waybill_info.get("consignorAddr", ""),
            "routes": [],
            "source": "monthly_card",
        }

    return default


# =========================
# SQLite 发货记录缓存
# =========================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shipment_tracking.db")

_db_lock = threading.Lock()


def _get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def db_init():
    """初始化表"""
    with _db_lock:
        conn = _get_db()
        conn.execute("""
            CREATE TABLE IF NOT EXISTS tracking_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tracking_number TEXT NOT NULL UNIQUE,
                project_name TEXT DEFAULT '',
                status TEXT DEFAULT '未知',
                detail TEXT DEFAULT '',
                update_time TEXT DEFAULT '',
                last_checked TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.commit()
        conn.close()


def db_upsert_tracking(tracking_number: str, project_name: str = "",
                       status: str = "未知", detail: str = "", update_time: str = ""):
    """插入或更新物流记录"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _db_lock:
        conn = _get_db()
        conn.execute("""
            INSERT INTO tracking_log (tracking_number, project_name, status, detail, update_time, last_checked)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(tracking_number) DO UPDATE SET
                status=excluded.status, detail=excluded.detail,
                update_time=excluded.update_time, last_checked=excluded.last_checked
        """, (tracking_number, project_name, status, detail, update_time, now))
        conn.commit()
        conn.close()


def db_get_active_trackings() -> List[dict]:
    """获取未签收的记录"""
    with _db_lock:
        conn = _get_db()
        rows = conn.execute("""
            SELECT tracking_number, project_name, status, detail
            FROM tracking_log
            WHERE status != '已签收' AND tracking_number != ''
        """).fetchall()
        result = [dict(r) for r in rows]
        conn.close()
        return result


def db_get_all_trackings() -> List[dict]:
    """获取全部记录"""
    with _db_lock:
        conn = _get_db()
        rows = conn.execute("""
            SELECT * FROM tracking_log ORDER BY created_at DESC
        """).fetchall()
        result = [dict(r) for r in rows]
        conn.close()
        return result


def db_get_stats() -> dict:
    """统计"""
    with _db_lock:
        conn = _get_db()
        rows = conn.execute("SELECT status, COUNT(*) as cnt FROM tracking_log GROUP BY status").fetchall()
        out = {r["status"]: r["cnt"] for r in rows}
        conn.close()
        return out


# =========================
# 飞书适配器
# =========================
_feishu_adapter = None


def set_feishu_adapter(adapter):
    global _feishu_adapter
    _feishu_adapter = adapter


def _feishu_update_shipping_table(project_name: str, tracking_number: str, info: dict):
    """按项目名更新飞书发货总表的物流字段"""
    if not _feishu_adapter or not project_name:
        return

    TABLE_ID = os.getenv("TABLE_ID_SHIPPING", "")
    bitable_token = os.getenv("BITABLE_APP_TOKEN", "")
    if not TABLE_ID or not bitable_token:
        return

    try:
        token = _feishu_adapter.get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}

        # 1. 按项目名搜索
        search_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{bitable_token}/tables/{TABLE_ID}/records/search"
        resp = httpx.post(search_url, headers=headers, json={
            "filter": {
                "conjunction": "and",
                "conditions": [{"field_name": "项目名称", "operator": "contains", "value": [project_name]}]
            },
            "page_size": 5,
        }, timeout=15.0)

        data = resp.json()
        items = data.get("data", {}).get("items", [])
        if not items:
            logger.debug(f"飞书未找到项目: {project_name}")
            return

        # 2. 对每个匹配的记录，检查快递单号是否为空，有空的就写入
        for item in items:
            fields = item.get("fields", {})
            existing_no = _parse_field(fields.get("快递单号", ""))
            if existing_no:
                continue  # 已有单号，跳过

            record_id = item["record_id"]
            update_fields = {
                "快递公司": "顺丰速运",
                "快递单号": tracking_number,
                "物流状态": info.get("status", ""),
            }
            if info.get("update_time"):
                update_fields["物流更新时间"] = info["update_time"]
            if info.get("detail"):
                update_fields["物流详情"] = info["detail"]

            update_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{bitable_token}/tables/{TABLE_ID}/records/{record_id}"
            httpx.put(update_url, headers=headers, json={"fields": update_fields}, timeout=15.0)
            logger.info(f"飞书发货表更新: 项目={project_name} 单号={tracking_number} 状态={info.get('status')}")
            break  # 只写第一个匹配的

    except Exception as e:
        logger.error(f"飞书更新失败: {e}")


def _feishu_search_by_tracking(tracking_number: str) -> List[dict]:
    """按快递单号在飞书发货总表中搜索"""
    if not _feishu_adapter:
        return []

    TABLE_ID = os.getenv("TABLE_ID_SHIPPING", "")
    bitable_token = os.getenv("BITABLE_APP_TOKEN", "")
    if not TABLE_ID or not bitable_token:
        return []

    try:
        token = _feishu_adapter.get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}
        search_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{bitable_token}/tables/{TABLE_ID}/records/search"

        resp = httpx.post(search_url, headers=headers, json={
            "filter": {
                "conjunction": "and",
                "conditions": [{"field_name": "快递单号", "operator": "is", "value": [tracking_number]}]
            },
            "page_size": 10,
        }, timeout=15.0)

        return resp.json().get("data", {}).get("items", [])
    except Exception:
        return []


def _parse_field(value) -> str:
    if isinstance(value, list):
        return "".join(item.get("text", "") for item in value if isinstance(item, dict))
    return str(value).strip() if value else ""


# =========================
# FastAPI Router
# =========================
router = APIRouter(prefix="/shipping", tags=["顺丰物流"])


def _get_sf_client() -> Optional[SFClient]:
    if not SF_PARTNER_ID or not SF_CHECKWORD:
        return None
    return SFClient(SF_PARTNER_ID, SF_CHECKWORD, SF_MONTHLY_CARD, SF_SERVER_URL)


@router.post("/tracking/import")
async def tracking_import(req: ManualOrderRequest):
    """
    录入快递单号并立即查询物流状态
    请求体: {"tracking_numbers": ["SF123", "SF456"], "project_name": "XX项目"}
    """
    if not req.tracking_numbers:
        raise HTTPException(status_code=400, detail="快递单号不能为空")

    sf = _get_sf_client()
    if not sf:
        raise HTTPException(status_code=400, detail="顺丰凭证未配置")

    results = []
    for tn in req.tracking_numbers:
        tn = tn.strip()
        if not tn:
            continue
        try:
            resp = sf.query_sfwaybill(tn)
            info = parse_sfwaybill_response(resp)

            # 优先用返回数据中的地址提取项目名；没有则用请求带的
            project = info.get("project_name", "") or req.project_name

            db_upsert_tracking(tn, project, info["status"], info["detail"], info["update_time"])

            # 有项目名才写飞书
            if project:
                _feishu_update_shipping_table(project, tn, info)

            results.append({
                "tracking_number": tn,
                "status": info["status"],
                "detail": info["detail"],
                "project_name": project,
            })
            time.sleep(0.2)  # 限频
        except HTTPException:
            raise
        except Exception as e:
            results.append({"tracking_number": tn, "status": "error", "detail": str(e)})

    return {"results": results, "total": len(results)}


@router.post("/tracking/refresh")
async def tracking_refresh():
    """手动刷新所有活跃运单的物流状态"""
    result = refresh_all_active()
    return {"status": "ok", "processed": len(result), "details": result}


@router.post("/tracking/refresh/one")
async def tracking_refresh_one(tracking_number: str = ""):
    """刷新单个运单"""
    if not tracking_number:
        raise HTTPException(status_code=400, detail="缺少 tracking_number")
    sf = _get_sf_client()
    if not sf:
        raise HTTPException(status_code=400, detail="顺丰凭证未配置")

    resp = sf.query_sfwaybill(tracking_number)
    info = parse_sfwaybill_response(resp)
    db_upsert_tracking(tracking_number, info.get("project_name", ""),
                       info["status"], info["detail"], info["update_time"])
    return {"tracking_number": tracking_number, **info}


@router.get("/tracking/status")
async def tracking_status():
    """物流状态统计"""
    stats = db_get_stats()
    return {
        "total": sum(stats.values()),
        "stats": stats,
        "active": sum(v for k, v in stats.items() if k != "已签收"),
        "signed": stats.get("已签收", 0),
    }


@router.get("/tracking/list")
async def tracking_list():
    """所有运单列表"""
    return {"items": db_get_all_trackings()}


# =========================
# 物流状态刷新任务
# =========================
def refresh_all_active() -> List[dict]:
    """刷新所有未签收运单"""
    sf = _get_sf_client()
    if not sf:
        logger.warning("顺丰凭证未配置，跳过刷新")
        return []

    active = db_get_active_trackings()
    if not active:
        return []

    results = []
    newly_signed = []

    for row in active:
        tn = row["tracking_number"]
        try:
            resp = sf.query_sfwaybill(tn)
            info = parse_sfwaybill_response(resp)

            if info["status"] == "未知":
                continue

            db_upsert_tracking(tn, row.get("project_name", ""),
                               info["status"], info["detail"], info["update_time"])

            # 飞书回写
            project_name = row.get("project_name", "") or info.get("project_name", "")
            if project_name:
                _feishu_update_shipping_table(project_name, tn, info)
            else:
                # 无项目名，尝试按单号在飞书中搜索并更新
                _update_feishu_by_tracking(tn, info)

            if info["status"] != row.get("status", ""):
                if info["status"] == "已签收":
                    newly_signed.append({"tracking_number": tn, "project_name": project_name})

            results.append({"tracking_number": tn, "status": info["status"], "detail": info["detail"]})
            time.sleep(0.2)

        except Exception as e:
            logger.error(f"查询 {tn} 失败: {e}")

    if newly_signed:
        _notify_signed(newly_signed)

    return results


def _update_feishu_by_tracking(tracking_number: str, info: dict):
    """按单号在飞书中搜索并更新"""
    items = _feishu_search_by_tracking(tracking_number)
    if not items or not _feishu_adapter:
        return
    try:
        TABLE_ID = os.getenv("TABLE_ID_SHIPPING", "")
        bitable_token = os.getenv("BITABLE_APP_TOKEN", "")
        token = _feishu_adapter.get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}

        for item in items:
            record_id = item["record_id"]
            fields = item.get("fields", {})
            exist_status = _parse_field(fields.get("物流状态", ""))
            if exist_status == info["status"]:
                continue  # 状态没变，不重复更新
            update_url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{bitable_token}/tables/{TABLE_ID}/records/{record_id}"
            httpx.put(update_url, headers=headers, json={"fields": {
                "物流状态": info["status"],
                "物流详情": info["detail"],
                "物流更新时间": info["update_time"],
            }}, timeout=15.0)
    except Exception as e:
        logger.error(f"按单号飞书更新失败: {e}")


def _notify_signed(rows: List[dict]):
    """飞书群签收通知"""
    chat_id = os.getenv("FEISHU_CHAT_ID", "")
    if not chat_id or not _feishu_adapter:
        return

    lines = ["📦 **快递签收通知**"]
    for r in rows[:10]:
        proj = r.get("project_name") or "—"
        lines.append(f"· {proj} — {r['tracking_number']} 已签收")

    try:
        token = _feishu_adapter.get_access_token()
        httpx.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"},
            json={
                "receive_id": chat_id,
                "msg_type": "interactive",
                "content": json.dumps({
                    "elements": [{"tag": "markdown", "content": "\n".join(lines)}],
                }),
            },
            timeout=10.0,
        )
    except Exception as e:
        logger.error(f"签收通知失败: {e}")


# =========================
# 定时任务
# =========================
_scheduler = BackgroundScheduler(timezone="Asia/Shanghai")


def start_scheduler():
    """启动定时任务（每 2 小时，8-20 点）"""
    if not SF_PARTNER_ID or not SF_CHECKWORD:
        logger.info("顺丰凭证未配置，跳过定时任务")
        return

    db_init()
    _scheduler.add_job(
        refresh_all_active,
        "cron",
        hour="8-20/2",
        id="sf_tracking_refresh",
        name="顺丰物流状态刷新",
        replace_existing=True,
    )
    _scheduler.start()
    logger.info("物流刷新定时任务已启动（每2小时, 8:00-20:00）")


def stop_scheduler():
    if _scheduler.running:
        _scheduler.shutdown(wait=False)
        logger.info("物流刷新定时任务已停止")


# =========================
# 导出
# =========================
__all__ = [
    "router",
    "start_scheduler",
    "stop_scheduler",
    "set_feishu_adapter",
    "db_init",
    "SFClient",
    "parse_sfwaybill_response",
    "extract_project_from_address",
]
