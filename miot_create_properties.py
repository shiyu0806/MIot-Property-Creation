#!/usr/bin/env python3
"""
MIoT 属性批量创建工具（通用版）
- 自动查询产品服务列表
- 从 Excel 读取属性定义
- 自动识别属性类型（bool/数值/枚举）
- 批量调用 API 创建属性
"""

import argparse
import json
import os
import sys
import time

import requests
from openpyxl import load_workbook

# ─── API ──────────────────────────────────────────────────────
BASE = "https://iot.mi.com"
CREATE_API = f"{BASE}/cgi-std/post/api/v1/functionDefine/addInstanceProperty"
QUERY_PROPS_API = f"{BASE}/cgi-std/api/v1/functionDefine/getInstanceProperties"
QUERY_SERVICES_API = f"{BASE}/cgi-std/api/v1/functionDefine/getInstanceServices"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/146.0.0.0 Safari/537.36",
    "origin": BASE,
    "referer": f"{BASE}/",
}


# ─── Excel 读取 ──────────────────────────────────────────────

def read_config(ws) -> dict:
    """读取公共配置 Sheet"""
    cfg = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        key, val = row[0], row[1]
        if key and val is not None:
            cfg[str(key).strip()] = str(val).strip() if val else ""
    return cfg


def read_properties(ws) -> list[dict]:
    """读取属性定义 Sheet，返回属性列表"""
    # 第一行是表头，第二行是说明，数据从第三行开始
    headers = None
    props = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            headers = [str(c).strip() if c else "" for c in row]
            continue
        if i == 2:
            # 说明行，跳过
            continue
        # 空行跳过 - 用 name 列（第3列，index=2）判断
        name_val = row[2] if len(row) > 2 else None
        if not name_val:
            continue
        d = {}
        for j, h in enumerate(headers):
            val = row[j] if j < len(row) else None
            d[h] = val
        props.append(d)
    return props


# ─── API 请求 ─────────────────────────────────────────────────

def build_cookies(config: dict) -> dict:
    """构建请求 Cookie"""
    return {
        "serviceToken": config.get("serviceToken", ""),
        "userId": config.get("userId", ""),
        "xiaomiiot_ph": config.get("xiaomiiot_ph", ""),
    }


def build_query_params(config: dict, **extra) -> dict:
    """构建 Query 参数"""
    params = {
        "userId": config.get("userId", ""),
        "xiaomiiot_ph": config.get("xiaomiiot_ph", ""),
        "pdId": config.get("pdId", ""),
    }
    params.update(extra)
    return params


def query_services(config: dict) -> list[dict]:
    """查询产品的所有服务列表"""
    params = build_query_params(
        config,
        model=config.get("model", ""),
        connectType=config.get("connectType", "16"),
        language=config.get("language", "zh_cn"),
        version=config.get("version", "1"),
        status=config.get("status", "0"),
    )
    resp = requests.get(QUERY_SERVICES_API, params=params,
                        headers=HEADERS, cookies=build_cookies(config), timeout=15)
    data = resp.json()
    if data.get("status") != 200:
        print(f"❌ 查询服务列表失败: {data}")
        return []
    return data.get("result", [])


def query_properties(siid: int, service_type: str, config: dict) -> list[dict]:
    """查询某个服务下的已有属性"""
    params = build_query_params(config)
    params.update({
        "version": config.get("version", "1"),
        "status": config.get("status", "0"),
        "siid": str(siid),
        "serviceType": service_type,
        "model": config.get("model", ""),
        "connectType": config.get("connectType", "16"),
        "language": config.get("language", "zh_cn"),
    })
    resp = requests.get(QUERY_PROPS_API, params=params,
                        headers=HEADERS, cookies=build_cookies(config), timeout=15)
    data = resp.json()
    if data.get("status") != 200:
        return []
    return data.get("result", [])


def create_property(body: dict, config: dict) -> dict:
    """调用创建属性接口"""
    params = build_query_params(config)
    resp = requests.post(CREATE_API, params=params, json=body,
                         headers=HEADERS, cookies=build_cookies(config), timeout=15)
    return resp.json()


# ─── 属性类型识别 & 请求体构造 ───────────────────────────────

def detect_value_type(fmt: str, prop: dict) -> str:
    """自动识别值类型: bool_range / number / enum / string"""
    vt = str(prop.get("value_type") or "").strip().lower()
    if vt in ("bool_range", "number", "enum", "string"):
        return vt

    # 自动判断
    if fmt == "bool":
        return "bool_range"
    if fmt == "string":
        return "string"

    # 有枚举值定义 → enum
    vl = str(prop.get("value_list") or "").strip()
    if vl:
        return "enum"

    # 有数值范围 → number
    vr_min = prop.get("value_range_min")
    if vr_min is not None and str(vr_min).strip():
        return "number"

    # 默认
    return "number"


def parse_value_list(raw: str) -> list:
    """解析枚举值: '0:关闭,1:开启,2:待机' → [{value:0,description:'关闭'}, ...]"""
    if not raw or not str(raw).strip():
        return []
    result = []
    for item in str(raw).strip().split(","):
        item = item.strip()
        if ":" in item:
            val, desc = item.split(":", 1)
            result.append({"value": int(val.strip()), "description": desc.strip()})
        elif item.isdigit():
            result.append({"value": int(item), "description": f"值{item}"})
    return result


def parse_access(raw: str) -> list:
    """解析权限: 'read,write,notify' → ['read','write','notify']"""
    if not raw:
        return ["read", "write", "notify"]
    return [a.strip() for a in str(raw).split(",") if a.strip()]


def parse_bool(raw) -> bool:
    """解析布尔值"""
    if isinstance(raw, bool):
        return raw
    s = str(raw).strip().lower()
    return s in ("true", "1", "yes")


def build_request_body(prop: dict, config: dict, service_info: dict = None) -> dict:
    """构造创建属性的请求体"""

    def safe_int(val, default=0):
        """安全转换为 int，空字符串返回 default"""
        if val is None or str(val).strip() == "":
            return default
        try:
            return int(val)
        except (ValueError, TypeError):
            return default

    fmt = str(prop.get("format", "bool")).strip()
    vtype = detect_value_type(fmt, prop)

    # valueList & valueRange
    if vtype == "bool_range":
        value_list = []
        value_range = []
    elif vtype == "enum":
        value_list = parse_value_list(prop.get("value_list", ""))
        value_range = []
    elif vtype == "string":
        value_list = []
        value_range = []
    else:  # number
        value_list = []
        vr_min = prop.get("value_range_min")
        vr_max = prop.get("value_range_max")
        vr_step = prop.get("value_range_step")
        value_range = [
            int(vr_min) if vr_min is not None and str(vr_min).strip() else 0,
            int(vr_max) if vr_max is not None and str(vr_max).strip() else 65535,
            int(vr_step) if vr_step is not None and str(vr_step).strip() else 1,
        ]

    # 服务信息 - 优先用属性行内值，其次用查到的服务信息，最后用公共配置
    siid = prop.get("siid")
    if siid is None or str(siid).strip() == "":
        if service_info:
            siid = service_info.get("siid")
        else:
            siid = 0
    # 确保 siid 是 int（服务返回的 siid 可能是字符串）
    siid = safe_int(siid, 0)

    service_type = prop.get("service_type") or prop.get("serviceType", "")
    if not str(service_type).strip() and service_info:
        service_type = service_info.get("type", "")

    body = {
        "language": config.get("language", "zh_cn") or "zh_cn",
        "description": str(prop.get("description", "")),
        "format": fmt,
        "valueList": value_list,
        "valueRange": value_range,
        "access": parse_access(prop.get("access") or config.get("access", "read,write,notify")),
        "source": safe_int(prop.get("source") or config.get("source"), 4),
        "name": str(prop.get("name", "")),
        "gattAccess": parse_access(prop.get("access") or config.get("access", "read,write,notify")),
        "version": safe_int(config.get("version"), 1),
        "status": safe_int(config.get("status"), 0),
        "siid": safe_int(siid, 0),
        "serviceType": str(service_type),
        "model": str(prop.get("model") or config.get("model", "")),
        "connectType": safe_int(prop.get("connectType") or config.get("connectType"), 16),
        "standard": parse_bool(prop.get("standard") or config.get("standard", False)),
        "specCategory": None,
        "pdId": safe_int(prop.get("pdId") or config.get("pdId"), 0),
    }
    return body


# ─── 服务匹配 ────────────────────────────────────────────────

def match_service(prop: dict, services: list[dict]) -> dict | None:
    """
    根据属性行的 service_name / service_desc / siid 匹配服务
    优先级:
      1. service_name + service_desc 同时精确匹配（区分同名服务）
      2. service_desc 精确匹配
      3. service_name + service_desc 模糊匹配
      4. service_name 精确匹配
      5. service_desc 包含匹配
      6. siid 直接匹配（仅当以上都无法匹配时兜底）
    注意：siid 优先级最低，因为换产品后 siid 会变化
    """
    prop_siid = prop.get("siid")
    prop_sname = str(prop.get("service_name") or "").strip()
    prop_sdesc = str(prop.get("service_desc") or "").strip()

    # 1. service_name + service_desc 同时精确匹配
    if prop_sname and prop_sdesc:
        for svc in services:
            if svc.get("name") == prop_sname and svc.get("description") == prop_sdesc:
                return svc

    # 2. service_desc 精确匹配
    if prop_sdesc:
        for svc in services:
            if svc.get("description") == prop_sdesc:
                return svc

    # 3. service_name + service_desc 模糊匹配
    if prop_sname and prop_sdesc:
        for svc in services:
            svc_name = svc.get("name", "")
            svc_desc = svc.get("description", "")
            if prop_sname in svc_name and prop_sdesc in svc_desc:
                return svc

    # 4. service_name 精确匹配（仅当 desc 为空时）
    if prop_sname:
        for svc in services:
            if svc.get("name") == prop_sname:
                return svc

    # 5. service_desc 包含匹配
    if prop_sdesc:
        for svc in services:
            if prop_sdesc in svc.get("description", ""):
                return svc

    # 6. siid 直接匹配（兜底，仅在名称/描述都无法匹配时使用）
    if prop_siid is not None and str(prop_siid).strip():
        try:
            target_siid = int(prop_siid)
            for svc in services:
                svc_siid = svc.get("siid")
                if svc_siid is not None and int(svc_siid) == target_siid:
                    return svc
        except (ValueError, TypeError):
            pass

    return None


# ─── 主流程 ──────────────────────────────────────────────────

def print_services(services: list[dict]):
    """打印服务列表"""
    print(f"\n📋 产品服务列表（共 {len(services)} 个）:")
    print("-" * 80)
    print(f"{'siid':>4} | {'name':<22} | {'description':<20} | type")
    print("-" * 80)
    for svc in services:
        siid = svc.get("siid", "?")
        name = svc.get("name", "?")
        desc = svc.get("description", "")
        stype = svc.get("type", "")
        print(f"{siid:>4} | {name:<22} | {desc:<20} | {stype}")


def main():
    parser = argparse.ArgumentParser(description="MIoT 属性批量创建工具")
    parser.add_argument("--excel", default="MIoT_属性创建模板.xlsx",
                        help="Excel 文件路径")
    parser.add_argument("--dry-run", action="store_true",
                        help="只解析不创建")
    parser.add_argument("--yes", "-y", action="store_true",
                        help="跳过确认直接创建")
    parser.add_argument("--skip-verify", action="store_true",
                        help="跳过验证步骤（推荐）")
    parser.add_argument("--delay", type=float, default=0.5,
                        help="每条创建之间的间隔秒数（默认0.5）")
    parser.add_argument("--list-services", action="store_true",
                        help="只列出产品服务，不创建")
    parser.add_argument("--siid", type=int, default=None,
                        help="只创建指定 siid 下的属性")
    args = parser.parse_args()

    # 读取 Excel
    if not os.path.exists(args.excel):
        print(f"❌ 文件不存在: {args.excel}")
        sys.exit(1)

    wb = load_workbook(args.excel)
    config = read_config(wb["公共配置"])
    props = read_properties(wb["属性定义"])

    # 检查必填配置
    missing = []
    for key in ["serviceToken", "xiaomiiot_ph", "userId", "pdId", "model"]:
        if not config.get(key):
            missing.append(key)
    if missing:
        print(f"❌ 公共配置缺少必填项: {', '.join(missing)}")
        print("   请在 Excel「公共配置」Sheet 中填写")
        sys.exit(1)

    if not props:
        print("❌ 属性定义为空，请在 Excel 中填写属性行")
        sys.exit(1)

    # 查询服务列表
    print("🔍 查询产品服务列表...")
    services = query_services(config)
    if not services:
        print("❌ 未查到服务列表，请检查 Cookie 和产品信息")
        sys.exit(1)

    print_services(services)

    if args.list_services:
        return

    # 匹配属性行与服务
    print(f"\n📝 解析 {len(props)} 条属性定义...")
    tasks = []
    for i, prop in enumerate(props):
        name = prop.get("name", f"行{i+1}")
        fmt = prop.get("format", "?")
        vtype = detect_value_type(fmt, prop)

        # 匹配服务
        svc = match_service(prop, services)
        if svc:
            siid = svc["siid"]
            sname = svc.get("description", svc.get("name", ""))
        else:
            siid = prop.get("siid", "?")
            sname = "未匹配"

        # 过滤指定 siid
        if args.siid is not None and int(siid) != args.siid:
            continue

        body = build_request_body(prop, config, svc)

        tasks.append({
            "index": i + 1,
            "name": name,
            "desc": prop.get("description", ""),
            "format": fmt,
            "value_type": vtype,
            "siid": siid,
            "service_name": sname,
            "body": body,
        })

    if not tasks:
        print("❌ 没有匹配的属性任务")
        sys.exit(1)

    # 打印任务列表
    print(f"\n{'#':>3} | {'name':<20} | {'desc':<12} | {'format':<8} | {'vtype':<10} | siid | 服务")
    print("-" * 90)
    for t in tasks:
        print(f"{t['index']:>3} | {t['name']:<20} | {t['desc']:<12} | "
              f"{t['format']:<8} | {t['value_type']:<10} | {t['siid']:<4} | {t['service_name']}")

    if args.dry_run:
        print(f"\n🏁 干跑模式，共 {len(tasks)} 条属性待创建（未实际执行）")
        # 打印第一条的请求体示例
        if tasks:
            print(f"\n📄 第1条请求体示例:")
            print(json.dumps(tasks[0]["body"], ensure_ascii=False, indent=2))
        return

    # 确认
    if not args.yes:
        ans = input(f"\n⚠️  即将创建 {len(tasks)} 条属性，是否继续？[y/N] ").strip().lower()
        if ans not in ("y", "yes"):
            print("已取消")
            return

    # 批量创建
    print(f"\n🚀 开始创建 {len(tasks)} 条属性...\n")
    success = 0
    failed = 0
    results = []

    for t in tasks:
        print(f"  [{t['index']}] {t['name']} ({t['desc']}) → siid={t['siid']} ... ", end="", flush=True)
        try:
            resp = create_property(t["body"], config)
            status = resp.get("status")
            result_val = resp.get("result")
            if status == 200:
                piid = result_val  # result 是新属性的 piid
                print(f"✅ 成功 (piid={piid})")
                success += 1
                results.append({"name": t["name"], "status": "success", "piid": piid, "siid": t["siid"]})
            else:
                msg = resp.get("message", resp.get("msg", json.dumps(resp, ensure_ascii=False)))
                print(f"❌ 失败 ({msg})")
                failed += 1
                results.append({"name": t["name"], "status": "failed", "error": msg, "siid": t["siid"]})
        except Exception as e:
            print(f"❌ 异常 ({e})")
            failed += 1
            results.append({"name": t["name"], "status": "error", "error": str(e), "siid": t["siid"]})

        time.sleep(args.delay)

    # 汇总
    print(f"\n{'='*50}")
    print(f"📊 创建完成: 成功 {success} / 失败 {failed} / 共 {len(tasks)}")
    if failed > 0:
        print("\n❌ 失败列表:")
        for r in results:
            if r["status"] != "success":
                print(f"  - {r['name']} (siid={r['siid']}): {r.get('error', '未知错误')}")

    # 保存结果
    result_file = "miot_create_result.json"
    with open(result_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\n💾 结果已保存到 {result_file}")


if __name__ == "__main__":
    main()
