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

BASE = "https://iot.mi.com"

SERVICE_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "origin": BASE,
    "user-agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
}

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
        data = resp.json()
    except Exception as e:
        return (False, -1, "查询失败", f"查询产品状态失败: {e}")

    if data.get("status") != 200:
        return (False, -1, "查询失败", f"查询产品状态失败: {data.get('message', data)}")

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
    return {
        "userId": str(config.get("userId", "")),
        "xiaomiiot_ph": str(config.get("xiaomiiot_ph", "")),
        "serviceToken": str(config.get("serviceToken", "")),
    }

def _params(config: dict) -> dict:
    return {
        "userId": str(config.get("userId", "")),
        "xiaomiiot_ph": str(config.get("xiaomiiot_ph", "")),
    }

def _headers(pd_id=None) -> dict:
    h = dict(SERVICE_HEADERS)
    h["referer"] = (
        f"{BASE}/fe-op/productCenter/config/function?productId={pd_id or ''}"
    )
    return h


def _safe_request(method: str, url: str, *, max_retries: int = 3,
                  retry_delay: float = 1.0, log_fn=None, **kwargs) -> requests.Response:
    """带重试的 HTTP 请求，网络错误自动重试"""
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.request(method, url, timeout=30, **kwargs)
            return resp
        except (ConnectionResetError, ConnectionError, OSError) as e:
            last_exc = e
            if attempt < max_retries:
                if log_fn:
                    log_fn(f"  ⚠️ 请求失败(第{attempt}次)，{retry_delay}s后重试: {e}")
                else:
                    print(f"  ⚠️ 请求失败(第{attempt}次)，{retry_delay}s后重试: {e}")
                time.sleep(retry_delay)
    raise last_exc  # type: ignore


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
    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"API 返回非 JSON (HTTP {resp.status_code}): {resp.text[:300] or '(空响应)'}。常见原因：Cookie 过期")
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
    try:
        result = resp.json()
    except Exception:
        raise RuntimeError(f"创建服务 API 返回非 JSON (HTTP {resp.status_code}): {resp.text[:300] or '(空响应)'}")

    # 注入 .siid 方便上层读取
    if result.get("code") == 0:
        data = result.get("data")
        result["siid"] = data.get("siid") if isinstance(data, dict) else data
    elif "result" in result:
        r = result.get("result")
        result["siid"] = r.get("siid") if isinstance(r, dict) else r

    return result


# ─── 修正 siid ────────────────────────────────────────────────

def modify_siid(config: dict, service_id, old_siid: int, new_siid: int) -> dict:
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
    try:
        return resp.json()
    except Exception:
        raise RuntimeError(f"修正 siid API 返回非 JSON (HTTP {resp.status_code}): {resp.text[:300] or '(空响应)'}")


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
    读取服务 Excel 的「产品配置」Sheet（pandas 格式，列为 参数名/值）
    返回 config dict
    """
    import pandas as pd
    df = pd.read_excel(path, sheet_name=0)
    config = dict(zip(df["参数名"], df["值"]))
    return {k: str(v) for k, v in config.items() if v is not None}


def read_service_list_excel(path: str) -> list[dict]:
    """
    读取服务 Excel 的「服务列表」Sheet
    返回 list of dict（列：服务ID / 服务名称 / 服务描述 / 标准化描述 / 是否标准服务）
    """
    import pandas as pd
    df = pd.read_excel(path, sheet_name=1)
    df = df.dropna(subset=["服务名称"])
    df["服务名称"] = df["服务名称"].astype(str)
    return df.to_dict("records")


# ─── 批量同步服务（创建 + 修正 siid）────────────────────────────

def sync_services(
    config: dict,
    service_rows: list[dict],
    dry_run: bool = False,
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
                    r = modify_siid(config, svc_id, actual_siid, expected_siid)
                    if r.get("code") == 0 or r.get("status") == 200:
                        log(f"    ✅ 修正成功")
                        fixed += 1
                        results.append({"name": name, "action": "fix", "siid": expected_siid})
                    else:
                        log(f"    ❌ 修正失败: {r}")
                        errors += 1
                        results.append({"name": name, "action": "fix_fail", "siid": actual_siid, "error": str(r)})
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
                        fr = modify_siid(config, svc_id or new_siid, new_siid, expected_siid)
                        if fr.get("code") == 0 or fr.get("status") == 200:
                            log(f"    ✅ 创建成功 siid={expected_siid} (修正自{new_siid})")
                            created += 1
                            results.append({"name": name, "action": "create_fix", "siid": expected_siid, "original_siid": new_siid})
                        else:
                            log(f"    ⚠️ 创建成功 siid={new_siid}，修正到{expected_siid}失败")
                            created += 1
                            results.append({"name": name, "action": "create_fix_fail", "siid": new_siid, "expected_siid": expected_siid})
                    else:
                        log(f"    ✅ 创建成功 siid={new_siid}")
                        created += 1
                        results.append({"name": name, "action": "create", "siid": new_siid})
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
