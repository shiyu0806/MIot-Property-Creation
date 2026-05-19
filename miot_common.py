#!/usr/bin/env python3
"""
MIoT 公共工具模块
统一管理 BASE、HEADERS、Cookie/参数构建、带重试的 HTTP 请求
避免在各核心模块中重复定义相同逻辑
"""

import time
import requests

__all__ = [
    "BASE", "DEFAULT_HEADERS",
    "TEMPLATE_VERSION",
    "build_cookies", "build_params", "build_headers", "safe_request",
    "parse_json_response", "is_success_response", "response_message",
    "safe_int", "ApiAuthError",
    "PROPERTY_COLUMNS", "ACTION_COLUMNS", "EVENT_COLUMNS",
]

# ─── 公共常量 ─────────────────────────────────────────────────

BASE = "https://iot.mi.com"

TEMPLATE_VERSION = "2.0"

DEFAULT_HEADERS = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "user-agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "origin": BASE,
    "referer": f"{BASE}/",
}


# ─── Cookie / 参数构建 ───────────────────────────────────────

def build_cookies(config: dict) -> dict:
    """从 config 构建请求 Cookie（统一实现）"""
    return {
        "serviceToken": str(config.get("serviceToken", "")),
        "userId": str(config.get("userId", "")),
        "xiaomiiot_ph": str(config.get("xiaomiiot_ph", "")),
    }


def build_params(config: dict) -> dict:
    """从 config 构建公共查询参数（userId + xiaomiiot_ph）"""
    return {
        "userId": str(config.get("userId", "")),
        "xiaomiiot_ph": str(config.get("xiaomiiot_ph", "")),
    }


def build_headers(referer: str = None) -> dict:
    """构建请求 headers，可选覆盖 referer"""
    h = dict(DEFAULT_HEADERS)
    if referer:
        h["referer"] = referer
    return h


# ─── 带重试的 HTTP 请求 ──────────────────────────────────────

def safe_request(
    method: str,
    url: str,
    *,
    max_retries: int = 3,
    retry_delay: float = 1.0,
    timeout: int = 30,
    log_fn=None,
    **kwargs,
) -> requests.Response:
    """
    带重试的 HTTP 请求，网络错误自动重试，业务错误直接返回。
    各核心模块的 _safe_request 统一使用此实现。
    """
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.request(method, url, timeout=timeout, **kwargs)
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


# ─── API 响应解析 ─────────────────────────────────────────────

class ApiAuthError(RuntimeError):
    """Raised when MIoT rejects a request because login credentials are invalid."""


def parse_json_response(resp: requests.Response, context: str = "API") -> dict:
    """解析 JSON 响应；非 JSON 时给出更可读的登录/响应异常提示。"""
    if resp.status_code in (401, 403):
        raise ApiAuthError(
            f"{context} 鉴权失败 (HTTP {resp.status_code})：登录态失效或 Cookie 不正确，请重新登录后再试"
        )

    try:
        data = resp.json()
    except Exception:
        text = resp.text[:300] if getattr(resp, "text", "") else "(空响应)"
        lower = text.lower()
        hint = ""
        if "<html" in lower or "login" in lower or "passport" in lower:
            hint = "。常见原因：Cookie 过期或登录态失效，请重新登录"
            raise ApiAuthError(f"{context} 返回非 JSON (HTTP {resp.status_code}): {text}{hint}")
        raise RuntimeError(f"{context} 返回非 JSON (HTTP {resp.status_code}): {text}{hint}")

    if not isinstance(data, dict):
        raise RuntimeError(f"{context} 返回结构异常，应为 JSON 对象，实际为 {type(data).__name__}")
    return data


def is_success_response(data: dict) -> bool:
    """兼容 MIoT 平台常见成功标记：status=200 或 code=0。"""
    return data.get("status") == 200 or data.get("code") == 0


def response_message(data: dict, default: str = "未知错误") -> str:
    """从 API 响应中提取用户可读错误信息。"""
    return str(data.get("message") or data.get("msg") or data.get("desc") or default)


# ─── 安全类型转换 ─────────────────────────────────────────────

def safe_int(val, default: int = 0) -> int:  # val: Any
    """安全转换为 int，空值或无效值返回 default"""
    if val is None or str(val).strip() == "":
        return default
    try:
        return int(val)
    except (ValueError, TypeError):
        return default


# ─── Excel 列定义（属性 / 方法 / 事件） ─────────────────────
# 统一定义，供 create_template.py 和 miot_export_template.py 共用
# 格式: (key, header_name, width, description, required)

PROPERTY_COLUMNS = [
    ("name",              "name",              20, "属性英文名\n如: on, mode, delay-time",          True),
    ("description",       "description",       25, "属性中文描述\n如: 开关, 模式, 延时时间",          True),
    ("format",            "format",            12, "数据格式\nbool/uint8/uint16/uint32\n/int8/int16/int32/float/string", True),
    ("service_desc",      "service_desc",      22, "服务中文描述（推荐）\n如: 开关一键、按键1点动毫秒数", True),
    ("value_list",        "value_list",        35, "枚举值（仅enum类型）\n格式: 0:关闭,1:开启,2:待机",  False),
    ("value_range_min",   "value_range_min",   14, "数值最小值\n（仅number类型）",                    False),
    ("value_range_max",   "value_range_max",   14, "数值最大值\n（仅number类型）",                    False),
    ("value_range_step",  "value_range_step",  14, "数值步长\n（仅number类型）",                     False),
    ("service_name",      "service_name",      20, "服务英文名\n如: switch, jog-delay-time",         False),
    ("siid",              "siid",               8, "服务ID（备选）\n直接指定siid，填了则忽略service匹配", False),
    ("access",            "access",            20, "访问权限\n默认: read,write,notify\n（gattAccess自动等同于access）", False),
    ("piid",              "piid",               8, "属性ID\n（导出时自动填入，创建后可校验修正）", False),
]

ACTION_COLUMNS = [
    ("name",              "name",              20, "方法英文名\n如: toggle, play, reset",           True),
    ("description",       "description",       25, "方法中文描述\n如: 切换, 播放, 重置",             True),
    ("normalizationDesc", "normalizationDesc", 20, "规范描述（通常=英文名）",                        False),
    ("service_desc",      "service_desc",      22, "服务中文描述（推荐）\n如: 开关一键、按键1点动毫秒数", True),
    ("service_name",      "service_name",      20, "服务英文名\n如: switch, jog-delay-time",         False),
    ("siid",              "siid",               8, "服务ID（备选）\n直接指定siid，填了则忽略service匹配", False),
    ("aiid",              "aiid",               8, "方法ID\n（导出时自动填入，创建后可校验修正）", False),
]

EVENT_COLUMNS = [
    ("name",              "name",              20, "事件英文名\n如: click, timeout, alarm",         True),
    ("description",       "description",       25, "事件中文描述\n如: 点击, 超时, 告警",             True),
    ("normalizationDesc", "normalizationDesc", 20, "规范描述（通常=英文名）",                        False),
    ("service_desc",      "service_desc",      22, "服务中文描述（推荐）\n如: 开关一键、按键1点动毫秒数", True),
    ("service_name",      "service_name",      20, "服务英文名\n如: switch, jog-delay-time",         False),
    ("siid",              "siid",               8, "服务ID（备选）\n直接指定siid，填了则忽略service匹配", False),
    ("eiid",              "eiid",               8, "事件ID\n（导出时自动填入，创建后可校验修正）", False),
]
