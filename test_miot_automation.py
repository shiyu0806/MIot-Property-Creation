#!/usr/bin/env python3
"""
miot_automation_core.py 纯函数单元测试

运行: python3 -m pytest test_miot_automation.py -v
或:   python3 test_miot_automation.py
"""

import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from miot_automation_core import (
    _replace_source_model,
    _parse_spec_type,
    _parse_spec_relate,
    _fix_item_model,
    read_automation_excel,
)


# ═══════════════════════════════════════════════════════════════════
# _replace_source_model 测试
# ═══════════════════════════════════════════════════════════════════

class TestReplaceSourceModel(unittest.TestCase):
    def test_basic_replace(self):
        result = _replace_source_model(
            "xhuan.switch.4prz03.set_properties",
            "xhuan.switch.4prz03",
            "gudi.switch.swy007"
        )
        assert result == "gudi.switch.swy007.set_properties"

    def test_no_replace_needed(self):
        """相同 model 不替换"""
        result = _replace_source_model(
            "xhuan.switch.4prz03.set_properties",
            "xhuan.switch.4prz03",
            "xhuan.switch.4prz03"
        )
        assert result == "xhuan.switch.4prz03.set_properties"

    def test_empty_value(self):
        assert _replace_source_model("", "a", "b") == ""

    def test_empty_source(self):
        assert _replace_source_model("hello", "", "b") == "hello"

    def test_none_value(self):
        assert _replace_source_model(None, "a", "b") is None

    def test_multiple_occurrences(self):
        result = _replace_source_model(
            "xhuan.switch.4prz03.action.xhuan.switch.4prz03",
            "xhuan.switch.4prz03",
            "gudi.switch.swy007"
        )
        assert result == "gudi.switch.swy007.action.gudi.switch.swy007"


# ═══════════════════════════════════════════════════════════════════
# _parse_spec_type 测试
# ═══════════════════════════════════════════════════════════════════

class TestParseSpecType(unittest.TestCase):
    def test_property(self):
        assert _parse_spec_type("property.25.1") == "prop"

    def test_event(self):
        assert _parse_spec_type("event.6.2") == "event"

    def test_empty(self):
        assert _parse_spec_type("") == "event"

    def test_none(self):
        assert _parse_spec_type(None) == "event"

    def test_unknown_prefix(self):
        assert _parse_spec_type("action.1.2") == "action"


# ═══════════════════════════════════════════════════════════════════
# _parse_spec_relate 测试
# ═══════════════════════════════════════════════════════════════════

class TestParseSpecRelate(unittest.TestCase):
    def test_property(self):
        siid, sub = _parse_spec_relate("property.25.1")
        assert siid == "25"
        assert sub == "1"

    def test_event(self):
        siid, sub = _parse_spec_relate("event.6.2")
        assert siid == "6"
        assert sub == "2"

    def test_empty(self):
        assert _parse_spec_relate("") == ("", "")

    def test_none(self):
        assert _parse_spec_relate(None) == ("", "")

    def test_two_parts(self):
        """只有两段时应返回空"""
        assert _parse_spec_relate("property.25") == ("", "")


# ═══════════════════════════════════════════════════════════════════
# _fix_item_model 测试
# ═══════════════════════════════════════════════════════════════════

class TestFixItemModel(unittest.TestCase):
    def test_replace_from_command(self):
        """从 command 提取源 model 并替换"""
        config = {"model": "gudi.switch.swy007"}
        item = {
            "command": "xhuan.switch.4prz03.set_properties",
            "key": "4.1.2",
            "model": "xhuan.switch.4prz03",
        }
        _fix_item_model(config, item)
        assert item["command"] == "gudi.switch.swy007.set_properties"
        assert item["model"] == "gudi.switch.swy007"

    def test_replace_from_key(self):
        """从 key 提取源 model 并替换"""
        config = {"model": "gudi.switch.swy007"}
        item = {
            "command": "",
            "key": "4.xhuan.switch.4prz03.3.1",
            "model": "",
        }
        _fix_item_model(config, item)
        assert "gudi.switch.swy007" in item["key"]

    def test_no_replace_when_same_model(self):
        """目标 model 和源 model 相同则不替换"""
        config = {"model": "xhuan.switch.4prz03"}
        item = {"command": "xhuan.switch.4prz03.set_properties", "model": "xhuan.switch.4prz03"}
        _fix_item_model(config, item)
        assert item["command"] == "xhuan.switch.4prz03.set_properties"

    def test_empty_model(self):
        """config 没有 model，不处理"""
        config = {"model": ""}
        item = {"command": "xhuan.switch.4prz03.set_properties"}
        _fix_item_model(config, item)
        assert item["command"] == "xhuan.switch.4prz03.set_properties"

    def test_replace_group_scene_dto(self):
        """替换 groupSceneDto 中的字段"""
        config = {"model": "gudi.switch.swy007"}
        item = {
            "command": "",
            "key": "4.xhuan.switch.4prz03.3.1",
            "model": "",
            "groupSceneDto": {
                "command": "xhuan.switch.4prz03.set_properties",
                "key": "4.xhuan.switch.4prz03.3.1",
            },
        }
        _fix_item_model(config, item)
        assert "gudi.switch.swy007" in item["groupSceneDto"]["command"]
        assert "gudi.switch.swy007" in item["groupSceneDto"]["key"]


# ═══════════════════════════════════════════════════════════════════
# 运行入口
# ═══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("运行 miot_automation_core 纯函数单元测试...\n")
    unittest.main(verbosity=2)
