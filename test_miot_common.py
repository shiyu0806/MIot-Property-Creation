#!/usr/bin/env python3
"""
miot_common.py + miot_create_properties.py 核心函数单元测试

运行: python3 -m pytest test_miot_common.py -v
或:   python3 test_miot_common.py
"""

import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from miot_common import (
    build_cookies, build_params, build_headers, safe_int, safe_request,
    BASE, DEFAULT_HEADERS,
    PROPERTY_COLUMNS, ACTION_COLUMNS, EVENT_COLUMNS,
)
from miot_create_properties import (
    match_service, detect_value_type, parse_value_list,
    build_request_body, build_action_request_body, build_event_request_body,
    _build_action_event_base, parse_bool, parse_access,
)


# ═══════════════════════════════════════════════════════════════════
# miot_common 测试
# ═══════════════════════════════════════════════════════════════════

class TestSafeInt(unittest.TestCase):
    def test_normal_int(self):
        assert safe_int("123") == 123

    def test_none(self):
        assert safe_int(None) == 0

    def test_empty_string(self):
        assert safe_int("") == 0

    def test_whitespace(self):
        assert safe_int("   ") == 0

    def test_invalid(self):
        assert safe_int("abc") == 0

    def test_custom_default(self):
        assert safe_int(None, default=-1) == -1

    def test_float_string(self):
        assert safe_int("3.14") == 0

    def test_negative(self):
        assert safe_int("-5") == -5


class TestBuildCookies(unittest.TestCase):
    def test_normal(self):
        cfg = {"serviceToken": "abc", "userId": "123", "xiaomiiot_ph": "xyz"}
        cookies = build_cookies(cfg)
        assert cookies["serviceToken"] == "abc"
        assert cookies["userId"] == "123"
        assert cookies["xiaomiiot_ph"] == "xyz"

    def test_missing_keys(self):
        cookies = build_cookies({})
        assert cookies["serviceToken"] == ""
        assert cookies["userId"] == ""

    def test_numeric_values(self):
        cookies = build_cookies({"userId": 12345})
        assert cookies["userId"] == "12345"


class TestBuildParams(unittest.TestCase):
    def test_normal(self):
        cfg = {"userId": "u1", "xiaomiiot_ph": "p1", "extra": "ignored"}
        params = build_params(cfg)
        assert params == {"userId": "u1", "xiaomiiot_ph": "p1"}


class TestBuildHeaders(unittest.TestCase):
    def test_default(self):
        h = build_headers()
        assert h["referer"] == f"{BASE}/"
        assert "user-agent" in h

    def test_custom_referer(self):
        h = build_headers("https://example.com")
        assert h["referer"] == "https://example.com"


class TestConstants(unittest.TestCase):
    def test_base_url(self):
        assert BASE == "https://iot.mi.com"

    def test_column_definitions(self):
        for col_def in PROPERTY_COLUMNS + ACTION_COLUMNS + EVENT_COLUMNS:
            assert len(col_def) == 5, f"列定义应为 5 元组: {col_def}"
            assert isinstance(col_def[2], int), f"宽度应为 int: {col_def}"
            assert isinstance(col_def[4], bool), f"required 应为 bool: {col_def}"


# ═══════════════════════════════════════════════════════════════════
# match_service 测试
# ═══════════════════════════════════════════════════════════════════

class TestMatchService(unittest.TestCase):
    def setUp(self):
        self.services = [
            {"siid": 2, "name": "light", "type": "urn:miot-spec-v2:device:light:1", "description": "Light"},
            {"siid": 3, "name": "switch", "type": "urn:miot-spec-v2:service:switch:1", "description": "Switch"},
            {"siid": 4, "name": "mode", "type": "urn:miot-spec-v2:service:mode:1", "description": "Mode"},
            {"siid": 10, "name": "dimmer", "type": "urn:miot-spec-v2:service:dimmer:1", "description": "Dimmer"},
            {"siid": 20, "name": "countdown", "type": "urn:miot-spec-v2:service:countdown:1", "description": "Countdown"},
        ]

    def test_exact_desc_and_name(self):
        prop = {"service_desc": "Light", "name": "on", "service_name": "light"}
        result = match_service(prop, self.services)
        assert result is not None
        assert result["siid"] == 2

    def test_exact_desc_only(self):
        prop = {"service_desc": "Switch", "service_name": "xxx"}
        result = match_service(prop, self.services)
        assert result is not None
        assert result["siid"] == 3

    def test_fuzzy_desc_includes(self):
        """优先级 5: service_desc 在 description 中"""
        prop = {"service_desc": "ight", "service_name": "xxx"}
        result = match_service(prop, self.services)
        assert result is not None  # "ight" 是 "Light" 的子串
        assert result["siid"] == 2

    def test_name_matching(self):
        prop = {"service_desc": "nonexistent", "service_name": "countdown"}
        result = match_service(prop, self.services)
        assert result is not None
        assert result["siid"] == 20

    def test_siid_fallback(self):
        prop = {"siid": 4, "service_desc": "nonexistent", "service_name": "xxx"}
        result = match_service(prop, self.services)
        assert result is not None
        assert result["siid"] == 4

    def test_no_match(self):
        prop = {"service_desc": "nonexistent", "service_name": "xxx"}
        result = match_service(prop, self.services)
        assert result is None

    def test_empty_services(self):
        prop = {"service_desc": "Light"}
        result = match_service(prop, [])
        assert result is None

    def test_exact_desc_priority_over_siid(self):
        prop = {"siid": 99, "service_desc": "Switch", "service_name": "switch"}
        result = match_service(prop, self.services)
        assert result is not None
        assert result["siid"] == 3


# ═══════════════════════════════════════════════════════════════════
# detect_value_type / parse_value_list 测试
# ═══════════════════════════════════════════════════════════════════

class TestDetectValueType(unittest.TestCase):
    def test_enum(self):
        prop = {"format": "uint8", "value_list": "0:关闭,1:开启,2:待机", "value_range_min": "", "value_range_max": ""}
        assert detect_value_type("uint8", prop) == "enum"

    def test_number(self):
        prop = {"format": "uint16", "value_list": "", "value_range_min": "0", "value_range_max": "100"}
        assert detect_value_type("uint16", prop) == "number"

    def test_bool_type(self):
        prop = {"format": "bool", "value_list": "", "value_range_min": "", "value_range_max": ""}
        assert detect_value_type("bool", prop) == "bool_range"

    def test_string_type(self):
        prop = {"format": "string", "value_list": "", "value_range_min": "", "value_range_max": ""}
        assert detect_value_type("string", prop) == "string"


class TestParseValueList(unittest.TestCase):
    def test_standard(self):
        result = parse_value_list("0:关闭,1:开启,2:待机")
        assert result == [
            {"description": "关闭", "value": 0},
            {"description": "开启", "value": 1},
            {"description": "待机", "value": 2},
        ]

    def test_whitespace(self):
        result = parse_value_list(" 0 : 关闭 , 1 : 开启 ")
        assert len(result) == 2
        assert result[0]["value"] == 0

    def test_empty(self):
        assert parse_value_list("") == []
        assert parse_value_list(None) == []

    def test_single(self):
        result = parse_value_list("1:on")
        assert result == [{"description": "on", "value": 1}]


# ═══════════════════════════════════════════════════════════════════
# parse_bool / parse_access 测试
# ═══════════════════════════════════════════════════════════════════

class TestParseBool(unittest.TestCase):
    def test_true_values(self):
        for v in ["1", "true", "True", "TRUE", "yes", 1, True]:
            assert parse_bool(v) is True, f"应为 True: {v!r}"

    def test_false_values(self):
        for v in ["0", "false", "no", "off", "", None, 0]:
            assert parse_bool(v) is False, f"应为 False: {v!r}"


class TestParseAccess(unittest.TestCase):
    def test_standard(self):
        assert parse_access("read,write,notify") == ["read", "write", "notify"]

    def test_single(self):
        assert parse_access("read") == ["read"]

    def test_empty(self):
        assert parse_access("") == ["read", "write", "notify"]

    def test_whitespace(self):
        assert parse_access(" read , write ") == ["read", "write"]


# ═══════════════════════════════════════════════════════════════════
# build_request_body 测试
# ═══════════════════════════════════════════════════════════════════

class TestBuildRequestBody(unittest.TestCase):
    def setUp(self):
        self.config = {"model": "test.device", "pdId": "123", "version": "1", "language": "zh_cn"}
        self.service_info = {"siid": 3, "type": "urn:test:switch:1"}

    def test_basic_property(self):
        prop = {
            "name": "on", "description": "开关", "format": "bool",
            "access": "read,write,notify",
            "value_list": "", "value_range_min": "", "value_range_max": "",
            "value_range_step": "", "standard": "", "valueType": "", "unit": "", "piid": "",
        }
        body = build_request_body(prop, self.config, self.service_info)
        assert body["siid"] == 3
        assert body["description"] == "开关"
        assert body["format"] == "bool"
        assert body["access"] == ["read", "write", "notify"]

    def test_enum_property(self):
        prop = {
            "name": "mode", "description": "模式", "format": "uint8",
            "access": "read,write", "value_list": "0:自动,1:手动",
            "value_range_min": "", "value_range_max": "", "value_range_step": "",
            "standard": "", "valueType": "", "unit": "", "piid": "",
        }
        body = build_request_body(prop, self.config, self.service_info)
        assert body["format"] == "uint8"
        assert body["access"] == ["read", "write"]

    def test_number_property(self):
        prop = {
            "name": "brightness", "description": "亮度", "format": "uint8",
            "access": "read,write", "value_list": "",
            "value_range_min": "0", "value_range_max": "100", "value_range_step": "1",
            "standard": "", "valueType": "", "unit": "", "piid": "",
        }
        body = build_request_body(prop, self.config, self.service_info)
        assert body["format"] == "uint8"
        assert body["valueRange"] == [0, 100, 1]


# ═══════════════════════════════════════════════════════════════════
# build_action / build_event 测试
# ═══════════════════════════════════════════════════════════════════

class TestBuildActionEventBody(unittest.TestCase):
    def setUp(self):
        self.config = {"model": "test.device", "pdId": "123", "version": "1", "language": "zh_cn"}
        self.service_info = {"siid": 3, "type": "urn:test:switch:1"}

    def test_action_body(self):
        item = {"name": "toggle", "description": "切换", "normalizationDesc": "toggle"}
        body = build_action_request_body(item, self.config, self.service_info)
        assert body["siid"] == 3
        assert body["name"] == "toggle"
        assert body["description"] == "切换"

    def test_event_body(self):
        item = {"name": "alarm", "description": "告警", "normalizationDesc": "alarm"}
        body = build_event_request_body(item, self.config, self.service_info)
        assert body["siid"] == 3
        assert body["name"] == "alarm"

    def test_action_defaults(self):
        item = {"name": "reset", "description": "重置", "normalizationDesc": ""}
        body = build_action_request_body(item, self.config, self.service_info)
        assert body["normalizationDesc"] == "reset"


# ═══════════════════════════════════════════════════════════════════
# safe_request 测试
# ═══════════════════════════════════════════════════════════════════

class TestSafeRequest(unittest.TestCase):
    def test_success(self):
        resp = safe_request("GET", "https://httpbin.org/get", timeout=5, max_retries=1)
        assert resp.status_code == 200

    def test_retry_on_connection_error(self):
        try:
            safe_request("GET", "http://localhost:1", timeout=1, max_retries=2)
        except Exception:
            pass


if __name__ == "__main__":
    print("运行 miot_common + miot_create_properties 核心函数单元测试...\n")
    unittest.main(verbosity=2)
