#!/usr/bin/env python3
"""
MIoT 服务层核心逻辑
- 查询服务列表
- 创建服务
- 修正 siid
- 导出服务/属性
"""

import json
import os
import time

import requests

__all__ = [
    "get_services", "sync_services",
    "read_service_config_excel", "read_service_list_excel",
    "parse_service_str", "check_product_status",
    "modify_iid", "modify_with_retry",
]

from miot_common import (
    BASE,
    DEFAULT_HEADERS as SERVICE_HEADERS,
    build_cookies as _build_cookies,
    build_params as _build_params,
    safe_request as _safe_request,
    parse_json_response,
    is_success_response,
    response_message,
)

SERVICE_HEADERS = dict(SERVICE_HEADERS)  # 向后兼容：保留模块级名称

GET_SERVICES_API  = f"{BASE}/cgi-std/api/v1/functionDefine/getInstanceServices"
ADD_SERVICE_API   = f"{BASE}/cgi-std/post/api/v1/functionDefine/addInstanceService"
MODIFY_SIID_API   = f"{BASE}/cgi-op/api/v1/speccenter/specV2/instanceProperty/modifyPropertyIid"
PRODUCT_LIST_API  = f"{BASE}/cgi-op/api/v1/product/list/get"


# ─── 产品状态检查 ─────────────────────────────────────────────

# 产品状态映射
PRODUCT_STATUS_MAP = {
    0: "测试中",
    1: "开发中",
    2: "审核中",
    3: "已发布",
    4: "已下架",
}

def check_product_status(config: dict) -> tuple:
    """
    检查目标产品的状态，只有 status=0（测试中）才允许创建。
    返回 (is_ok: bool, status: int, status_name: str, message: str)
    调用方应在创建操作前优先调用此函数，非测试中状态直接拒绝创建。
    """
    pd_id = config.get("pdId", "")
    if not pd_id:
        return (False, -1, "未知", f"未指定产品ID (pdId)，无法检查产品状态")

    params = {
        "userId": str(config.get("userId", "")),
        "xiaomiiot_ph": str(config.get("xiaomiiot_ph", "")),
        "searchWords": str(config.get("model", "")),  # 用 model 作为搜索关键词
        "region": -1,
        "productTypeId": -1,
        "connectType": -1,
    }
    # 如果有 groupId 则传入，提高查找效率
    if config.get("groupId"):
        params["groupId"] = str(config["groupId"])

    # cgi-op 域名需要 /fe-op/productCenter 作为 referer
    check_headers = dict(SERVICE_HEADERS)
    check_headers["referer"] = f"{BASE}/fe-op/productCenter"

    try:
        resp = _safe_request("GET", PRODUCT_LIST_API, params=params,
                             cookies=_cookies(config), headers=check_headers)
        data = parse_json_response(resp, "查询产品状态 API")
    except Exception as e:
        return (False, -1, "查询失败", f"查询产品状态失败: {e}")

    if not is_success_response(data):
        return (False, -1, "查询失败", f"查询产品状态失败: {response_message(data, str(data))}")

    products = data.get("result") or []
    target = None
    for p in products:
        if str(p.get("pdId")) == str(pd_id):
            target = p
            break

    if not target:
        return (False, -1, "未找到", f"未在产品列表中找到 pdId={pd_id} 的产品，请确认产品ID是否正确")

    status = target.get("status", -1)
    status_name = PRODUCT_STATUS_MAP.get(status, f"未知({status})")

    if status == 0:
        return (True, status, status_name, f"产品 {target.get('name', pd_id)} (pdId={pd_id}) 状态: {status_name}，允许创建")
    else:
        return (False, status, status_name,
                f"产品 {target.get('name', pd_id)} (pdId={pd_id}) 状态: {status_name}（status={status}），仅测试中（status=0）才允许创建！")


# ─── 低层 HTTP ────────────────────────────────────────────────

def _cookies(config: dict) -> dict:
    return _build_cookies(config)

def _params(config: dict) -> dict:
    return _build_params(config)

def _headers(pd_id: int | str | None = None) -> dict:
    h = dict(SERVICE_HEADERS)
    h["referer"] = (
        f"{BASE}/fe-op/productCenter/config/function?productId={pd_id or ''}"
    )
    return h


# ─── 查询服务 ─────────────────────────────────────────────────

def get_services(config: dict) -> list[dict]:
    """查询产品下的所有服务列表"""
    params = {
        **_params(config),
        "model": config.get("model", ""),
        "language": config.get("language", "zh_cn"),
        "version": config.get("version", "1"),
    }
    # pdId 可选
    if config.get("pdId"):
        params["pdId"] = str(config["pdId"])

    resp = _safe_request("GET", GET_SERVICES_API, params=params,
                         cookies=_cookies(config), headers=_headers())
    data = parse_json_response(resp, "查询服务列表 API")
    return data.get("result") or data.get("data") or []


# ─── 创建服务 ─────────────────────────────────────────────────

def create_service(config: dict, name: str, description: str = "",
                   norm_desc: str = "", standard: bool = False) -> dict:
    """
    创建单个服务
    返回原始 API 响应，同时尝试注入 .siid 字段方便调用方使用
    """
    pd_id = config.get("pdId", "")
    payload = {
        "version": "1",
        "model": config.get("model", ""),
        "language": config.get("language", "zh_cn"),
        "standard": standard,
        "name": name,
        "normalizationDesc": norm_desc or name,
        "description": description or " ",
        "pdId": int(pd_id) if pd_id else 0,
    }
    resp = _safe_request("POST", ADD_SERVICE_API,
                         params=_params(config),
                         cookies=_cookies(config),
                         headers=_headers(pd_id),
                         json=payload)
    result = parse_json_response(resp, "创建服务 API")

    # 注入 .siid 方便上层读取
    if result.get("code") == 0:
        data = result.get("data")
        result["siid"] = data.get("siid") if isinstance(data, dict) else data
    elif "result" in result:
        r = result.get("result")
        result["siid"] = r.get("siid") if isinstance(r, dict) else r

    return result


# ─── 修正 siid ────────────────────────────────────────────────

def modify_siid(config: dict, service_id: int | str, old_siid: int, new_siid: int) -> dict:
    """修正服务的 siid"""
    pd_id = config.get("pdId", "")
    payload = {
        "model": config.get("model", ""),
        "pdId": int(pd_id) if pd_id else 0,
        "version": "1",
        "serviceId": service_id,
        "oldIid": old_siid,
        "whichIid": "SIID",
        "newIid": new_siid,
    }
    resp = _safe_request("POST", MODIFY_SIID_API,
                         params=_params(config),
                         cookies=_cookies(config),
                         headers=_headers(pd_id),
                         json=payload)
    return parse_json_response(resp, "修正 siid API")


# ─── 通用 IID 修正（PIID / AIID / EIID）────────────────────────

def modify_iid(config: dict, service_id: int | str, old_iid: int, new_iid: int, which_iid: str) -> dict:
    """修正属性/方法/事件的 IID
    which_iid: "PIID" / "AIID" / "EIID"
    复用与 modify_siid 相同的 API 端点，仅 whichIid 字段不同
    """
    pd_id = config.get("pdId", "")
    payload = {
        "model": config.get("model", ""),
        "pdId": int(pd_id) if pd_id else 0,
        "version": "1",
        "serviceId": service_id,
        "oldIid": old_iid,
        "whichIid": which_iid,
        "newIid": new_iid,
    }
    resp = _safe_request("POST", MODIFY_SIID_API,
                         params=_params(config),
                         cookies=_cookies(config),
                         headers=_headers(pd_id),
                         json=payload)
    return parse_json_response(resp, f"修正 {which_iid} API")


# ─── 带重试的 IID 修正 ────────────────────────────────────────

def modify_with_retry(modify_fn, *args, log_fn=None, **kwargs):
    """
    包装 modify_siid / modify_iid，失败后自动重试：
      第1次失败 → 等5秒重试
      第2次失败 → 等10秒重试
      第3次失败 → 返回 (False, last_response)

    返回 (success: bool, response: dict)
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    last_resp = None
    retry_delays = [5, 10]  # 第一次失败等5s，第二次失败等10s

    for attempt in range(1 + len(retry_delays)):  # 最多3次尝试
        resp = modify_fn(*args, **kwargs)
        if is_success_response(resp):
            return True, resp
        last_resp = resp
        if attempt < len(retry_delays):
            delay = retry_delays[attempt]
            msg_resp = response_message(resp, str(resp))
            log(f"    ⚠️ 修正失败({msg_resp})，等待{delay}秒后重试...")
            time.sleep(delay)

    return False, last_resp


# ─── 解析服务的 serviceStr ────────────────────────────────────

def parse_service_str(svc: dict) -> dict:
    """
    解析 svc["serviceStr"] JSON，返回
    { "properties": [...], "events": [...], "actions": [...] }
    """
    svc_str = svc.get("serviceStr", "")
    if not svc_str:
        return {"properties": [], "events": [], "actions": []}
    try:
        sd = json.loads(svc_str)
    except Exception:
        return {"properties": [], "events": [], "actions": []}
    return {
        "properties": sd.get("required-properties", []),
        "events":     sd.get("required-events", []),
        "actions":    sd.get("required-actions", []),
    }


# ─── 读取服务 Excel 配置 ──────────────────────────────────────

def read_service_config_excel(path: str) -> dict:
    """
    读取服务 Excel 的「产品配置」Sheet（列为 参数名/值）
    返回 config dict
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    config = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) >= 2 and row[0] is not None:
            key = str(row[0]).strip()
            val = str(row[1]).strip() if row[1] is not None else ""
            config[key] = val
    wb.close()
    return config


def read_service_list_excel(path: str) -> list[dict]:
    """
    读取服务 Excel 的「服务列表」Sheet
    返回 list of dict（列：服务ID / 服务名称 / 服务描述 / 标准化描述 / 是否标准服务）
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[1] if len(wb.worksheets) > 1 else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    result = []
    for row in rows[1:]:
        record = {}
        for i, h in enumerate(headers):
            val = row[i] if i < len(row) else None
            record[h] = str(val) if val is not None else ""
        # 跳过 服务名称 为空的行
        if record.get("服务名称", "").strip():
            result.append(record)
    return result


# ─── 批量同步服务（创建 + 修正 siid）────────────────────────────

def sync_services(
    config: dict,
    service_rows: list[dict],
    dry_run: bool = False,
    delay: float = 0.0,
    log_fn=None,
    cancelled_fn=None,
) -> dict:
    """
    批量同步服务（创建 / 跳过 / 修正 siid）

    service_rows 每行字段：
        服务名称, 服务描述(可选), 服务ID(期望siid), 标准化描述(可选), 是否标准服务(可选)

    log_fn(msg: str) 若提供则实时回调日志
    cancelled_fn() -> bool 若提供且返回 True 则中止

    返回 { created, skipped, fixed, errors, results: list }
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    def is_cancelled():
        return cancelled_fn() if cancelled_fn else False

    # ─── 优先检查产品状态（仅非 dry-run 时）────────────────────
    if not dry_run:
        log("🔍 正在检查产品状态...")
        is_ok, status, status_name, msg = check_product_status(config)
        if is_ok:
            log(f"✅ {msg}")
        else:
            log(f"❌ {msg}")
            return {
                "created": 0, "skipped": 0,
                "fixed": 0, "errors": len(service_rows),
                "results": [{"action": "blocked", "error": msg} for _ in service_rows],
            }

    # 获取已有服务
    log(f"正在获取 {config.get('model')} 的已有服务...")
    all_services = get_services(config)
    existing = {}
    for svc in all_services:
        key = (
            svc.get("name", ""),
            svc.get("description", "") or svc.get("normalizationDesc", ""),
        )
        existing[key] = svc
    log(f"平台已有 {len(all_services)} 个服务")

    total = len(service_rows)
    log(f"Excel 中有 {total} 个服务待处理\n")

    results = []
    created = skipped = fixed = errors = 0

    for row_num, row in enumerate(service_rows, 1):
        if is_cancelled():
            log("\n⚠️ 已取消")
            break

        name         = str(row.get("服务名称", "")).strip()
        desc         = str(row.get("服务描述", "")).strip() if row.get("服务描述") else ""
        norm_desc    = str(row.get("标准化描述", "")).strip() if row.get("标准化描述") else ""
        standard     = str(row.get("是否标准服务", "false")).lower() == "true"
        expected_siid_raw = row.get("服务ID")
        try:
            expected_siid = int(expected_siid_raw) if expected_siid_raw not in (None, "", float("nan")) else None
        except (ValueError, TypeError):
            expected_siid = None

        key = (name, desc)

        if key in existing:
            actual_siid = int(existing[key].get("siid", 0))
            if expected_siid and actual_siid != expected_siid:
                log(f"[{row_num}/{total}] 🔧 {name} siid={actual_siid}→{expected_siid} 修正中...")
                if not dry_run:
                    svc_id = existing[key].get("serviceId", actual_siid)
                    ok, r = modify_with_retry(modify_siid, config, svc_id, actual_siid, expected_siid, log_fn=log)
                    if ok:
                        log(f"    ✅ 修正成功")
                        fixed += 1
                        results.append({"name": name, "action": "fix", "siid": expected_siid})
                        if delay > 0:
                            time.sleep(delay)
                    else:
                        log(f"    ❌ 修正失败(已重试3次): {r}")
                        errors += 1
                        results.append({"name": name, "action": "fix_fail", "siid": actual_siid, "error": str(r)})
                        log("\n⛔ 修正失败，停止创建")
                        break
                else:
                    log(f"    [干跑] 需要修正 siid {actual_siid} → {expected_siid}")
                    fixed += 1
            else:
                log(f"[{row_num}/{total}] ⏭️ {name} siid={actual_siid} 已存在，跳过")
                skipped += 1
                results.append({"name": name, "action": "skip", "siid": actual_siid})
        else:
            log(f"[{row_num}/{total}] 🆕 {name} 创建中...")
            if not dry_run:
                r = create_service(config, name, desc, norm_desc, standard)
                new_siid = r.get("siid")
                if new_siid:
                    if expected_siid and new_siid != expected_siid:
                        log(f"    siid={new_siid}，期望={expected_siid}，修正中...")
                        svc_data = r.get("data") or r.get("result")
                        svc_id = svc_data.get("serviceId") if isinstance(svc_data, dict) else new_siid
                        ok_f, fr = modify_with_retry(modify_siid, config, svc_id or new_siid, new_siid, expected_siid, log_fn=log)
                        if ok_f:
                            log(f"    ✅ 创建成功 siid={expected_siid} (修正自{new_siid})")
                            created += 1
                            results.append({"name": name, "action": "create_fix", "siid": expected_siid, "original_siid": new_siid})
                        else:
                            log(f"    ⚠️ 创建成功 siid={new_siid}，修正到{expected_siid}失败(已重试3次)")
                            errors += 1
                            results.append({"name": name, "action": "create_fix_fail", "siid": new_siid, "expected_siid": expected_siid})
                            log("\n⛔ siid修正失败，停止创建")
                            break
                    else:
                        log(f"    ✅ 创建成功 siid={new_siid}")
                        created += 1
                        results.append({"name": name, "action": "create", "siid": new_siid})
                    if delay > 0:
                        time.sleep(delay)
                else:
                    log(f"    ❌ 创建失败: {r}")
                    errors += 1
                    results.append({"name": name, "action": "fail", "error": str(r)})
            else:
                log(f"    [干跑] 将创建")
                created += 1

    summary = f"\n{'='*40}\n完成！创建: {created} | 跳过: {skipped} | 修正: {fixed} | 错误: {errors}"
    log(summary)

    return {
        "created": created, "skipped": skipped,
        "fixed": fixed,    "errors": errors,
        "results": results,
    }
