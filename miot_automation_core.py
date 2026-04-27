#!/usr/bin/env python3
"""
MIoT 自定义自动化核心模块
API:
  - 查询自动化列表: GET /cgi-op/api/v1/productcenter/automation/list
  - 检查标准自动化: POST /cgi-op/api/v1/productcenter/automation/check/standard/automation (multipart)
  - 保存 then 普通动作: POST /cgi-op/api/v1/productcenter/automation/action/save (JSON)
  - 保存 then 组合动作: POST /cgi-op/api/v1/productcenter/automation/group/action/save (JSON, 扁平payload含actionList)
  - 保存 if（触发条件）: POST /cgi-op/api/v1/productcenter/automation/launch/save (JSON)

自动化分两种类型:
  - then（执行动作）: trId=201, specType=prop, appValueStyle=4, autoType=1
    有 command、actionList，value=null，actionList 内含实际执行值
  - if（触发条件）: trId=101, specType=event, appValueStyle=0, autoType=0
    有 key、src、scId 等触发字段
"""

import requests
import json
import time

# ─── API 端点 ─────────────────────────────────────────────────

BASE = "https://iot.mi.com"
LIST_API = f"{BASE}/cgi-op/api/v1/productcenter/automation/list"
CHECK_API = f"{BASE}/cgi-op/api/v1/productcenter/automation/check/standard/automation"
SAVE_IF_API = f"{BASE}/cgi-op/api/v1/productcenter/automation/launch/save"
SAVE_THEN_API = f"{BASE}/cgi-op/api/v1/productcenter/automation/action/save"          # 普通动作
SAVE_THEN_GROUP_API = f"{BASE}/cgi-op/api/v1/productcenter/automation/group/action/save"  # 组合动作


# ─── 请求辅助 ─────────────────────────────────────────────────

def _headers():
    return {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36",
        "Origin": BASE,
        "Referer": f"{BASE}/",
    }

def _cookies(config: dict) -> dict:
    return {
        "serviceToken": config["serviceToken"],
        "userId": str(config["userId"]),
        "xiaomiiot_ph": config["xiaomiiot_ph"],
    }

def _params(config: dict) -> dict:
    return {
        "userId": str(config["userId"]),
        "xiaomiiot_ph": config["xiaomiiot_ph"],
    }


# ─── 从 specRelate 解析 specType ─────────────────────────────

def _replace_source_model(value: str, source_model: str, target_model: str) -> str:
    """将字符串中的源 model 替换为目标 model
    如 "xhuan.switch.4prz03.set_properties" + ("xhuan.switch.4prz03", "gudi.switch.swy007")
    → "gudi.switch.swy007.set_properties"
    """
    if not value or not source_model or not target_model or source_model == target_model:
        return value
    return value.replace(source_model, target_model)


def _fix_item_model(config: dict, item: dict):
    """预处理：确保 item 中所有 model 相关字段使用目标产品的 model。
    从 key/command 中提取源 model，替换 key/value/command/groupSceneDto 中的引用。
    同时修正 item.model 为空的问题。
    """
    target_model = config.get("model", "")
    if not target_model:
        return

    # 尝试从多个来源提取源 model
    source_model = ""
    command = item.get("command", "")
    if command:
        cmd_base = command.split(".set_properties")[0].split(".action")[0]
        if cmd_base and cmd_base != target_model:
            source_model = cmd_base
    if not source_model:
        key = item.get("key", "")
        if key:
            key_parts = key.split(".")
            if len(key_parts) >= 4:
                source_model_in_key = ".".join(key_parts[1:-2])
                if source_model_in_key and source_model_in_key != target_model:
                    source_model = source_model_in_key
    if not source_model:
        item_model = item.get("model", "")
        if item_model and item_model != target_model:
            source_model = item_model

    if source_model:
        # 替换 item 中所有字符串字段里的源 model
        for k in ("command", "key", "model", "specRelate", "value"):
            val = item.get(k, "")
            if isinstance(val, str) and source_model in val:
                item[k] = _replace_source_model(val, source_model, target_model)
        # 替换 groupSceneDto 中的字段
        gsd = item.get("groupSceneDto")
        if isinstance(gsd, dict):
            for k in ("command", "key", "model", "specRelate", "value"):
                val = gsd.get(k, "")
                if isinstance(val, str) and source_model in val:
                    gsd[k] = _replace_source_model(val, source_model, target_model)
        # 替换 actionList 中的 model
        al = item.get("actionList")
        if isinstance(al, list):
            for action in al:
                if isinstance(action, dict):
                    for k in ("command", "model", "specRelate", "value"):
                        val = action.get(k, "")
                        if isinstance(val, str) and source_model in val:
                            action[k] = _replace_source_model(val, source_model, target_model)
                    if "pdId" in action:
                        action["pdId"] = int(config.get("pdId", 0))

    # 确保 model 字段是目标 model（即使没有源 model 也补上）
    item_model = item.get("model", "")
    if not item_model or item_model != target_model:
        item["model"] = target_model


def _parse_spec_type(spec_relate: str) -> str:
    """从 specRelate 字段推断 specType
    property.25.1 -> prop
    event.6.2 -> event
    """
    if not spec_relate:
        return "event"
    prefix = spec_relate.split(".")[0].lower()
    if prefix == "property":
        return "prop"
    elif prefix == "event":
        return "event"
    return prefix


# ─── 从 specRelate 解析 siId 和 subIid ────────────────────────

def _parse_spec_relate(spec_relate: str) -> tuple:
    """从 specRelate 解析 (siId, subIid)
    property.25.1 -> (25, 1)
    event.6.2 -> (6, 2)
    """
    if not spec_relate:
        return ("", "")
    parts = spec_relate.split(".")
    if len(parts) >= 3:
        return (parts[1], parts[2])
    return ("", "")


# ─── 查询自动化列表 ────────────────────────────────────────────

def get_automation_list(config: dict) -> list:
    """查询产品的自定义自动化列表，返回带 _trType 标记的列表"""
    params = {**_params(config), "pdId": str(config["pdId"])}
    resp = requests.get(
        LIST_API, params=params,
        cookies=_cookies(config), headers=_headers(),
        timeout=30,
    )
    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"查询自动化列表 API 返回非 JSON (HTTP {resp.status_code}): "
                           f"{resp.text[:300] or '(空响应)'}")

    if data.get("status") != 200 and data.get("code") != 0:
        raise RuntimeError(f"查询自动化列表失败: {data}")

    result = data.get("result") or data.get("data") or {}
    if isinstance(result, list):
        return result
    if isinstance(result, dict):
        then_list = result.get("then") or []
        if_list = result.get("if") or []
        for item in then_list:
            item["_trType"] = "then"
        for item in if_list:
            item["_trType"] = "if"
        return then_list + if_list
    return []


# ─── 构建执行动作 (then) 的 groupSceneDto ─────────────────────

def _build_then_group_scene_dto(config: dict, auto_item: dict) -> dict:
    """构建 then（执行动作）类型的 groupSceneDto
    组合动作(appValueStyle=4)：含 actionList，value="null"
    普通动作(appValueStyle=0/1)：无 actionList，value 有值
    """
    pd_id = int(config.get("pdId", 0))
    model = auto_item.get("model", config.get("model", ""))
    spec_relate = auto_item.get("specRelate", "")
    si_id, sub_iid = _parse_spec_relate(spec_relate)
    if not si_id and auto_item.get("siId"):
        si_id = str(auto_item["siId"])
    if not sub_iid and auto_item.get("subIid"):
        sub_iid = str(auto_item["subIid"])

    spec_type = auto_item.get("specType") or _parse_spec_type(spec_relate)

    # actionList 处理（先解析，用于推断 appValueStyle）
    action_list = auto_item.get("actionList")
    if action_list and isinstance(action_list, str):
        try:
            action_list = json.loads(action_list)
        except Exception:
            action_list = None
    has_action_list = action_list and isinstance(action_list, list) and len(action_list) > 0

    # appValueStyle 推断逻辑：
    #   - 如果 Excel 中明确指定了 appValueStyle，直接使用
    #   - 否则：有 actionList → 4（组合动作），无 actionList → 0（普通动作）
    #   注意：不能默认 4，因为普通 then 动作 appValueStyle=0
    if auto_item.get("appValueStyle") is not None and str(auto_item.get("appValueStyle", "")).strip():
        app_value_style = int(auto_item["appValueStyle"])
    elif has_action_list:
        app_value_style = 4
    else:
        app_value_style = 0

    # command 推断：specRelate=action.* → *.action, property.* → *.set_properties
    if auto_item.get("command"):
        command = auto_item["command"]
    elif spec_relate.startswith("action."):
        command = f"{model}.action" if model else ""
    else:
        command = f"{model}.set_properties" if model else ""

    # value: 执行动作的值
    value = auto_item.get("value", "")
    if not isinstance(value, str):
        value = json.dumps(value, ensure_ascii=False) if value else ""

    group_scene_dto = {
        "pdId": pd_id,
        "intro": auto_item.get("intro", ""),
        "plugId": auto_item.get("plugId", ""),
        "fwVer": auto_item.get("fwVer", ""),
        "mcuFwVer": auto_item.get("mcuFwVer", ""),
        "platform": int(auto_item.get("platform", 0)),
        "attr": auto_item.get("attr", ""),
        "specRelate": spec_relate,
        "model": model,
        "appValueStyle": app_value_style,
        "specType": spec_type,
        "siId": int(si_id) if str(si_id).isdigit() else si_id,
        "subIid": int(sub_iid) if str(sub_iid).isdigit() else sub_iid,
        "autoType": 1,
        "command": command,
    }

    if has_action_list:
        # 组合动作：value=null, 含 actionList
        group_scene_dto["value"] = None
        group_scene_dto["actionList"] = action_list
    else:
        # 普通动作：value 有值（字符串形式）
        group_scene_dto["value"] = value

    return group_scene_dto


# ─── 构建触发条件 (if) 的 groupSceneDto ───────────────────────

def _build_if_group_scene_dto(config: dict, auto_item: dict) -> dict:
    """构建 if（触发条件）类型的 groupSceneDto"""
    pd_id = int(config.get("pdId", 0))
    model = auto_item.get("model", config.get("model", ""))
    spec_relate = auto_item.get("specRelate", "")
    si_id, sub_iid = _parse_spec_relate(spec_relate)
    if not si_id and auto_item.get("siId"):
        si_id = str(auto_item["siId"])
    if not sub_iid and auto_item.get("subIid"):
        sub_iid = str(auto_item["subIid"])

    spec_type = auto_item.get("specType") or _parse_spec_type(spec_relate)

    value = auto_item.get("value", "")
    if not isinstance(value, str):
        value = json.dumps(value, ensure_ascii=False) if value else ""

    group_scene_dto = {
        "pdId": pd_id,
        "intro": auto_item.get("intro", ""),
        "plugId": auto_item.get("plugId", ""),
        "fwVer": auto_item.get("fwVer", ""),
        "mcuFwVer": auto_item.get("mcuFwVer", ""),
        "platform": int(auto_item.get("platform", 0)),
        "attr": auto_item.get("attr", ""),
        "model": model,
        "key": auto_item.get("key", ""),
        "specRelate": spec_relate,
        "appValueStyle": int(auto_item.get("appValueStyle", 0)),
        "specType": spec_type,
        "siId": int(si_id) if str(si_id).isdigit() else si_id,
        "subIid": int(sub_iid) if str(sub_iid).isdigit() else sub_iid,
        "autoType": int(auto_item.get("autoType") or 0) if str(auto_item.get("autoType") or 0).lstrip("-").isdigit() else 0,
        "value": value,
    }
    # if 类型可能有额外字段
    # 注意：scId 是源产品的场景ID，创建新自动化时不应传入（目标产品不存在该scId）
    # 只有更新时才需要传 scId
    for extra_key in ("src",):
        if auto_item.get(extra_key) is not None:
            group_scene_dto[extra_key] = auto_item[extra_key]
    return group_scene_dto


# ─── 检查标准自动化 ────────────────────────────────────────────

def check_standard_automation(config: dict, auto_item: dict) -> dict:
    """
    检查自动化是否匹配标准场景（multipart form 上传）
    根据 _trType 自动区分 then（执行动作）和 if（触发条件）
    """
    # 预处理：确保 model 字段正确，替换 key/value 中的源 model
    _fix_item_model(config, auto_item)

    tr_type = auto_item.get("_trType", "then")
    spec_relate = auto_item.get("specRelate", "")
    si_id, sub_iid = _parse_spec_relate(spec_relate)
    if not si_id and auto_item.get("siId"):
        si_id = str(auto_item["siId"])
    if not sub_iid and auto_item.get("subIid"):
        sub_iid = str(auto_item["subIid"])

    spec_type = auto_item.get("specType") or _parse_spec_type(spec_relate)
    model = auto_item.get("model", config.get("model", ""))

    # actionList 处理（先解析，用于推断 appValueStyle）
    action_list = auto_item.get("actionList")
    if action_list and isinstance(action_list, str):
        try:
            action_list = json.loads(action_list)
        except Exception:
            action_list = None
    has_action_list = action_list and isinstance(action_list, list) and len(action_list) > 0

    if tr_type == "then":
        # appValueStyle 推断：有 actionList→4（组合），无 actionList→0（普通）
        if auto_item.get("appValueStyle") is not None and str(auto_item.get("appValueStyle", "")).strip():
            app_value_style = int(auto_item["appValueStyle"])
        elif has_action_list:
            app_value_style = 4
        else:
            app_value_style = 0
        auto_type = auto_item.get("autoType") or 1
        if isinstance(auto_type, str):
            auto_type = int(auto_type) if auto_type.strip().isdigit() else 1
        # command 推断：specRelate=action.* → *.action, property.* → *.set_properties
        if auto_item.get("command"):
            command = auto_item["command"]
        elif spec_relate.startswith("action."):
            command = f"{model}.action" if model else ""
        else:
            command = f"{model}.set_properties" if model else ""
    else:
        app_value_style = auto_item.get("appValueStyle") or 0
        if isinstance(app_value_style, str):
            app_value_style = int(app_value_style) if app_value_style.strip().isdigit() else 0
        auto_type = auto_item.get("autoType") or 0
        if isinstance(auto_type, str):
            auto_type = int(auto_type) if auto_type.strip().isdigit() else 0
        command = auto_item.get("command", "")

    value = auto_item.get("value", "")
    if not isinstance(value, str):
        value = json.dumps(value, ensure_ascii=False) if value else ""

    # 构建 multipart 字段
    fields = {
        "pdId": str(config.get("pdId", "")),
        "intro": auto_item.get("intro", ""),
        "plugId": auto_item.get("plugId", ""),
        "fwVer": auto_item.get("fwVer", ""),
        "mcuFwVer": auto_item.get("mcuFwVer", ""),
        "platform": str(auto_item.get("platform", 0)),
        "attr": auto_item.get("attr", ""),
        "model": model,
        "specRelate": spec_relate,
        "appValueStyle": str(app_value_style),
        "specType": spec_type,
        "siId": str(si_id),
        "subIid": str(sub_iid),
        "autoType": str(auto_type),
    }

    # then 类型：根据是否有 actionList 决定 value 格式
    if tr_type == "then" and command:
        fields["command"] = command

    if tr_type == "then":
        if has_action_list:
            # 组合动作（appValueStyle=4）：value=null, 有 actionList
            fields["value"] = "null"
            # actionList: multipart 中传 JSON 字符串
            fields["actionList"] = json.dumps(action_list, ensure_ascii=False)
        elif app_value_style == 4:
            # 聚合选值但没有 actionList → 自动从目标产品属性定义生成
            action_list = generate_action_list(config, auto_item)
            auto_item["actionList"] = action_list  # 缓存到 item 中供 save 使用
            action_list = _fix_action_list_for_target(config, action_list)
            fields["value"] = "null"
            fields["actionList"] = json.dumps(action_list, ensure_ascii=False)
        else:
            # 普通动作（appValueStyle=0/1）：value 有值, 无 actionList
            fields["value"] = value
    else:
        # if 类型：value 有值, 有 key 等触发字段
        fields["value"] = value
        key = auto_item.get("key", "")
        if key:
            fields["key"] = key

    # groupSceneDto
    group_scene = auto_item.get("groupSceneDto")
    if group_scene and isinstance(group_scene, dict):
        fields["groupSceneDto"] = json.dumps(group_scene, ensure_ascii=False)
    elif tr_type == "then":
        gs_dto = _build_then_group_scene_dto(config, auto_item)
        fields["groupSceneDto"] = json.dumps(gs_dto, ensure_ascii=False)
    else:
        gs_dto = _build_if_group_scene_dto(config, auto_item)
        fields["groupSceneDto"] = json.dumps(gs_dto, ensure_ascii=False)

    resp = requests.post(
        CHECK_API,
        params=_params(config),
        cookies=_cookies(config),
        headers={
            **{k: v for k, v in _headers().items() if k != "Content-Type"},
        },
        files={k: (None, v) for k, v in fields.items()},
        timeout=30,
    )
    try:
        return resp.json()
    except Exception:
        raise RuntimeError(f"检查标准自动化 API 返回非 JSON (HTTP {resp.status_code}): "
                           f"{resp.text[:300] or '(空响应)'}")


# ─── 保存自动化 ────────────────────────────────────────────────

def _fix_action_list_for_target(config: dict, action_list: list) -> list:
    """修复 actionList 中子动作的字段，使其与目标产品匹配
    - pdId → 目标 pdId
    - model / command → 目标 model
    - value 中的源 model 替换
    - siId/subIid 保留（聚合选值子动作使用相同的 siid/piid）
    """
    target_pd_id = int(config.get("pdId", 0))
    target_model = config.get("model", "")

    for act in action_list:
        if not isinstance(act, dict):
            continue
        # pdId 替换
        act["pdId"] = target_pd_id
        # model 替换
        if target_model and act.get("model"):
            source_model = act["model"]
            if source_model != target_model:
                for key in ("command", "model", "specRelate", "value"):
                    val = act.get(key, "")
                    if isinstance(val, str) and source_model in val:
                        act[key] = _replace_source_model(val, source_model, target_model)
                act["model"] = target_model
        # 清理子动作中不应传的字段（服务端会自动分配）
        for clean_key in ("saId", "tgId", "ctime", "status", "supportNative",
                          "tmpValueTypes", "labVerifyStatus", "pdIds", "type",
                          "title", "desc", "specVer"):
            if clean_key in act:
                del act[clean_key]
    return action_list


def save_automation(config: dict, auto_item: dict, is_update: bool = False) -> dict:
    """
    创建/更新自定义自动化（JSON body）
    then 普通动作 → /automation/action/save（无 groupSceneDto）
    then 组合动作 → /automation/group/action/save（扁平 payload 含 actionList）
    if（触发条件）→ /automation/launch/save（有 groupSceneDto）
    """
    # 预处理：确保 model 字段正确，替换 key/value 中的源 model
    _fix_item_model(config, auto_item)

    tr_type = auto_item.get("_trType", "then")
    spec_relate = auto_item.get("specRelate", "")
    si_id, sub_iid = _parse_spec_relate(spec_relate)
    if not si_id and auto_item.get("siId"):
        si_id = str(auto_item["siId"])
    if not sub_iid and auto_item.get("subIid"):
        sub_iid = str(auto_item["subIid"])

    spec_type = auto_item.get("specType") or _parse_spec_type(spec_relate)
    model = auto_item.get("model", config.get("model", ""))

    value = auto_item.get("value", "")
    if not isinstance(value, str):
        value = json.dumps(value, ensure_ascii=False) if value else ""

    # actionList 处理（先解析，用于推断 appValueStyle）
    action_list = auto_item.get("actionList")
    if action_list and isinstance(action_list, str):
        try:
            action_list = json.loads(action_list)
        except Exception:
            action_list = None
    has_action_list = action_list and isinstance(action_list, list) and len(action_list) > 0

    if tr_type == "then":
        # appValueStyle 推断：有 actionList→4（组合），无 actionList→0（普通）
        if auto_item.get("appValueStyle") is not None and str(auto_item.get("appValueStyle", "")).strip():
            app_value_style = int(auto_item["appValueStyle"])
        elif has_action_list:
            app_value_style = 4
        else:
            app_value_style = 0
        auto_type = auto_item.get("autoType") or 1
        if isinstance(auto_type, str):
            auto_type = int(auto_type) if auto_type.strip().isdigit() else 1

        # command 推断：
        # specRelate=action.* → command={model}.action
        # specRelate=property.* → command={model}.set_properties
        if auto_item.get("command"):
            command = auto_item["command"]
        elif spec_relate.startswith("action."):
            command = f"{model}.action" if model else ""
        else:
            command = f"{model}.set_properties" if model else ""

        # specRelate=action.* 时，value 格式为 {"siid":X,"aiid":Y,"in":[]}
        # 不需要转成 [{"siid":X,"piid":Y,"value":Z}] 格式
        if spec_relate.startswith("action.") and value:
            # 确保 value 是 JSON 字符串
            if not value.startswith("["):
                # 已经是 {"siid":...} 格式，保持原样
                pass
    else:
        app_value_style = auto_item.get("appValueStyle") or 0
        if isinstance(app_value_style, str):
            app_value_style = int(app_value_style) if app_value_style.strip().isdigit() else 0
        auto_type = auto_item.get("autoType") or 0
        if isinstance(auto_type, str):
            auto_type = int(auto_type) if auto_type.strip().isdigit() else 0
        command = auto_item.get("command", "")

    payload = {
        "pdId": int(config.get("pdId", 0)),
        "intro": auto_item.get("intro", ""),
        "plugId": auto_item.get("plugId", ""),
        "fwVer": auto_item.get("fwVer", ""),
        "mcuFwVer": auto_item.get("mcuFwVer", ""),
        "platform": int(auto_item.get("platform", 0)),
        "attr": auto_item.get("attr", ""),
        "specRelate": spec_relate,
        "model": model,
        "appValueStyle": int(app_value_style),
        "specType": spec_type,
        "siId": int(si_id) if str(si_id).isdigit() else si_id,
        "subIid": int(sub_iid) if str(sub_iid).isdigit() else sub_iid,
        "autoType": int(auto_type),
    }

    if tr_type == "then":
        # then（执行动作）分两种子类型，使用不同 API：
        # 1. 普通动作（appValueStyle=0/1，无 actionList）：
        #    → /automation/action/save，无 groupSceneDto，无 tgId
        #    子类型 a: specRelate=property.* → value=[{"siid":X,"piid":Y,"value":Z}]
        #    子类型 b: specRelate=action.* → value={"siid":X,"aiid":Y,"in":[]}, command=*.action
        # 2. 组合动作（appValueStyle=4，有 actionList）：
        #    → /automation/group/action/save，扁平 payload 含 actionList，不传 groupSceneDto/tgId
        payload["command"] = command

        if has_action_list:
            # 组合动作：value=null, 含 actionList（扁平结构，不嵌套 groupSceneDto）
            # 修复 actionList 中子动作的 pdId/model
            action_list = _fix_action_list_for_target(config, action_list)
            payload["value"] = None
            payload["actionList"] = action_list

            api = SAVE_THEN_GROUP_API
            params = _params(config)
        else:
            # 聚合选值但没有 actionList → 自动从目标产品属性定义生成
            if app_value_style == 4:
                action_list = generate_action_list(config, auto_item)
                action_list = _fix_action_list_for_target(config, action_list)
                payload["value"] = None
                payload["actionList"] = action_list

                api = SAVE_THEN_GROUP_API
                params = _params(config)
            else:
                # 普通动作：不传 groupSceneDto、tgId、actionList
                payload["value"] = value

                api = SAVE_THEN_API
                params = {**_params(config), "isUpdate": str(is_update).lower()}
    else:
        # if 类型: value 有值, 有 key 等触发字段
        payload["key"] = auto_item.get("key", "")
        payload["value"] = value
        # scId 是源产品场景ID，创建时不传（目标产品不存在），只传 src
        if auto_item.get("src") is not None:
            payload["src"] = auto_item["src"]
        # groupSceneDto
        group_scene = auto_item.get("groupSceneDto")
        if group_scene and isinstance(group_scene, dict):
            payload["groupSceneDto"] = group_scene
        else:
            payload["groupSceneDto"] = _build_if_group_scene_dto(config, auto_item)

        api = SAVE_IF_API
        params = {**_params(config), "isUpdate": str(is_update).lower()}

    # 调试：写入临时文件
    try:
        with open(f"/tmp/miot_save_debug_{tr_type}.json", "w") as f:
            json.dump({"tr_type": tr_type, "api": api, "payload": payload}, f, ensure_ascii=False, default=str, indent=2)
    except Exception:
        pass

    resp = requests.post(
        api, params=params,
        cookies=_cookies(config), headers=_headers(),
        json=payload, timeout=30,
    )
    try:
        return resp.json()
    except Exception:
        raise RuntimeError(f"保存自动化 API 返回非 JSON (HTTP {resp.status_code}): "
                           f"{resp.text[:300] or '(空响应)'}")


# ─── 查询属性定义（用于自动生成聚合选值的 actionList）──────────

PROP_DEF_API = f"{BASE}/cgi-std/api/v1/functionDefine/getInstanceProperties"
_prop_def_cache = {}  # {(userId, pdId): {(siid, piid): prop_info}}


def get_property_definitions(config: dict, use_cache: bool = True) -> dict:
    """查询产品的属性定义，返回 {(siid, piid): prop_info} 字典
    prop_info 包含 name, type, format, valueList, valueRange 等
    支持缓存，同一用户同一产品只查询一次
    """
    global _prop_def_cache
    cache_key = (str(config.get("userId", "")), str(config.get("pdId", "")))
    if use_cache and cache_key in _prop_def_cache:
        return _prop_def_cache[cache_key]

    resp = requests.get(
        PROP_DEF_API,
        params=_params(config),
        cookies=_cookies(config),
        headers=_headers(),
        timeout=30,
    )
    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"查询属性定义 API 返回非 JSON (HTTP {resp.status_code}): "
                           f"{resp.text[:300] or '(空响应)'}")

    if data.get("status") != 200:
        raise RuntimeError(f"查询属性定义失败: {data.get('message', data)}")

    result = data.get("result") or data.get("data") or []
    prop_map = {}
    for p in result:
        svcs = p.get("services", [])
        siid = svcs[0].get("iid") if svcs else p.get("siid", 0)
        piid = p.get("iid")
        prop_map[(int(siid), int(piid))] = p
    _prop_def_cache[cache_key] = prop_map
    return prop_map


def generate_action_list(config: dict, auto_item: dict) -> list:
    """为聚合选值（appValueStyle=4）自动生成 actionList
    基于 specRelate 中的 siid.piid 查询属性定义的枚举值列表
    """
    spec_relate = auto_item.get("specRelate", "")
    parts = spec_relate.split(".")
    if len(parts) < 3:
        raise RuntimeError(f"specRelate 格式错误: {spec_relate}")

    siid, piid = int(parts[1]), int(parts[2])
    model = auto_item.get("model", config.get("model", ""))
    pd_id = int(config.get("pdId", 0))

    # 查询属性定义
    prop_map = get_property_definitions(config)
    prop_info = prop_map.get((siid, piid))
    if not prop_info:
        raise RuntimeError(f"未找到 siid={siid}, piid={piid} 的属性定义，无法生成 actionList")

    value_list = prop_info.get("valueList", [])
    if not value_list:
        # 可能是数值型属性，不需要聚合选值
        raise RuntimeError(f"siid={siid}, piid={piid} 没有枚举值列表(valueList)，不是聚合选值属性")

    # 构建 actionList
    action_list = []
    for idx, v in enumerate(value_list):
        action_item = {
            "pdId": pd_id,
            "intro": str(v.get("description", v.get("value", idx))),
            "plugId": "",
            "fwVer": "",
            "mcuFwVer": "",
            "platform": 0,
            "attr": "",
            "specRelate": spec_relate,
            "model": model,
            "appValueStyle": 4,
            "specType": _parse_spec_type(spec_relate),
            "siId": siid,
            "subIid": piid,
            "autoType": 1,
            "command": f"{model}.set_properties" if model else "",
            "value": json.dumps([{"siid": siid, "piid": piid, "value": v.get("value")}]),
            "gid": 0,
            "extra": "{}",
            "rank": 10,
            "ruleId": 0,
        }
        action_list.append(action_item)

    return action_list


# ─── 批量同步自动化 ────────────────────────────────────────────

def sync_automations(config: dict, auto_items: list,
                     dry_run: bool = False,
                     delay: float = 0.5,
                     log_fn=None,
                     cancelled_fn=None) -> dict:
    """
    批量创建自定义自动化
    每条 item 的 _trType 决定走 then（执行动作）还是 if（触发条件）流程
    返回 {"success": [...], "failed": [...], "skipped": [...]}
    """
    results = {"success": [], "failed": [], "skipped": []}

    # 预处理：确保 command/key 中的 model 与目标 model 一致
    target_model = config.get("model", "")
    for item in auto_items:
        if not target_model:
            continue
        # 从 command 中提取源 model 前缀（最可靠的来源）
        command = item.get("command", "")
        source_model = ""
        if command:
            # command 格式: {source_model}.set_properties 或 {source_model}.action
            cmd_base = command.split(".set_properties")[0].split(".action")[0]
            if cmd_base and cmd_base != target_model:
                source_model = cmd_base
        # 兜底：从 key 中提取
        if not source_model:
            key = item.get("key", "")
            if key:
                key_parts = key.split(".")
                if len(key_parts) >= 4:
                    source_model_in_key = ".".join(key_parts[1:-2])
                    if source_model_in_key and source_model_in_key != target_model:
                        source_model = source_model_in_key
        # 兜底：从 item.model 中提取
        if not source_model:
            item_model = item.get("model", "")
            if item_model and item_model != target_model:
                source_model = item_model

        if source_model:
            # 替换 item 中所有字符串字段里的源 model
            for key in ("command", "key", "model", "specRelate", "value"):
                val = item.get(key, "")
                if isinstance(val, str) and source_model in val:
                    item[key] = _replace_source_model(val, source_model, target_model)
            # 替换 groupSceneDto 中的字段
            gsd = item.get("groupSceneDto")
            if isinstance(gsd, dict):
                for key in ("command", "key", "model", "specRelate", "value"):
                    val = gsd.get(key, "")
                    if isinstance(val, str) and source_model in val:
                        gsd[key] = _replace_source_model(val, source_model, target_model)
            # 替换 actionList 中的 model 和 value
            al = item.get("actionList")
            if isinstance(al, list):
                for action in al:
                    if isinstance(action, dict):
                        for key in ("command", "model", "specRelate", "value"):
                            val = action.get(key, "")
                            if isinstance(val, str) and source_model in val:
                                action[key] = _replace_source_model(val, source_model, target_model)
                        # pdId 替换为目标产品
                        if "pdId" in action:
                            action["pdId"] = int(config.get("pdId", 0))
            # 确保 model 字段是目标 model
            item["model"] = target_model
            log_fn and log_fn(f"  🔄 Model 替换: {source_model} → {target_model}")

    for i, item in enumerate(auto_items):
        if cancelled_fn and cancelled_fn():
            log_fn and log_fn("⚠️ 已取消")
            break

        intro = item.get("intro", f"自动化{i+1}")
        tr_type = item.get("_trType", "then")
        type_label = "执行动作" if tr_type == "then" else "触发条件"
        log_fn and log_fn(f"[{i+1}/{len(auto_items)}] [{type_label}] 处理: {intro}")

        if dry_run:
            log_fn and log_fn(f"  🔍 [dry-run] 将创建{type_label}: {intro}")
            results["skipped"].append({"intro": intro, "type": tr_type, "reason": "dry-run"})
            continue

        try:
            # 先检查标准自动化
            log_fn and log_fn(f"  🔍 检查标准自动化...")
            check_result = check_standard_automation(config, item)
            log_fn and log_fn(f"  📋 检查结果: {json.dumps(check_result, ensure_ascii=False)[:200]}")

            # 保存
            save_result = save_automation(config, item)
            save_status = save_result.get("status") or save_result.get("code")
            log_fn and log_fn(f"  📋 保存结果: {json.dumps(save_result, ensure_ascii=False)[:300]}")
            if save_result.get("status") == 200 or save_result.get("code") == 0:
                log_fn and log_fn(f"  ✅ 创建成功: {intro}")
                results["success"].append({"intro": intro, "type": tr_type, "result": save_result})
            else:
                msg = save_result.get("message", save_result.get("msg", json.dumps(save_result, ensure_ascii=False)))
                log_fn and log_fn(f"  ❌ 创建失败: {intro} ({msg})")
                results["failed"].append({"intro": intro, "type": tr_type, "error": msg, "result": save_result})
        except Exception as e:
            log_fn and log_fn(f"  ❌ 异常: {intro} ({e})")
            results["failed"].append({"intro": intro, "type": tr_type, "error": str(e)})

        if delay > 0:
            time.sleep(delay)

    return results


# ─── Excel 读写 ────────────────────────────────────────────────

def read_automation_excel(path: str) -> tuple:
    """
    读取自动化 Excel
    Sheet1: 产品配置 (参数名/值)
    Sheet2: 执行动作(then)
    Sheet3: 触发条件(if)  （可选，旧格式可能只有 Sheet2）
    返回 (config_dict, automation_list)
    """
    import pandas as pd

    # 读取配置
    df_config = pd.read_excel(path, sheet_name=0, dtype=str)
    config = {}
    for _, row in df_config.iterrows():
        key = str(row.iloc[0]).strip()
        val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        config[key] = val

    # 读取所有 Sheet 名称
    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names

    auto_items = []

    def _parse_sheet(sheet_name, tr_type):
        """解析单个 Sheet 的自动化数据"""
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
        items = []
        for _, row in df.iterrows():
            item = {}
            for col in df.columns:
                val = row[col]
                item[col] = str(val) if pd.notna(val) else ""
            # 数值字段转 int
            for int_key in ["siId", "subIid", "platform", "appValueStyle", "autoType",
                           "pdId", "trId", "saId", "scId", "gid", "rank", "ruleId"]:
                if int_key in item and item[int_key].isdigit():
                    item[int_key] = int(item[int_key])
            # actionList: 如果是 JSON 字符串则解析
            if "actionList" in item and item["actionList"]:
                try:
                    item["actionList"] = json.loads(item["actionList"])
                except Exception:
                    pass
            # groupSceneDto: 如果是 JSON 字符串则解析
            if "groupSceneDto" in item and item["groupSceneDto"]:
                try:
                    item["groupSceneDto"] = json.loads(item["groupSceneDto"])
                except Exception:
                    pass
            item["_trType"] = tr_type
            items.append(item)
        return items

    # 根据实际 Sheet 名称读取
    # 新格式：Sheet2=执行动作(then), Sheet3=触发条件(if)
    # 旧格式：Sheet2=自动化列表（无分类）
    then_sheet = None
    if_sheet = None
    for name in sheet_names:
        name_lower = name.lower()
        if "then" in name_lower or "执行动作" in name:
            then_sheet = name
        elif name_lower == "if" or "触发条件" in name:
            if_sheet = name

    if then_sheet or if_sheet:
        # 新格式：分 Sheet 读取
        if then_sheet:
            auto_items.extend(_parse_sheet(then_sheet, "then"))
        if if_sheet:
            auto_items.extend(_parse_sheet(if_sheet, "if"))
    else:
        # 旧格式：Sheet2 统一读取，根据 specRelate 推断类型
        if len(sheet_names) >= 2:
            df_auto = pd.read_excel(path, sheet_name=1, dtype=str)
            for _, row in df_auto.iterrows():
                item = {}
                for col in df_auto.columns:
                    val = row[col]
                    item[col] = str(val) if pd.notna(val) else ""
                # 数值字段转 int
                for int_key in ["siId", "subIid", "platform", "appValueStyle", "autoType",
                               "pdId", "trId", "saId", "scId", "gid", "rank", "ruleId"]:
                    if int_key in item and item[int_key].isdigit():
                        item[int_key] = int(item[int_key])
                # actionList: 如果是 JSON 字符串则解析
                if "actionList" in item and item["actionList"]:
                    try:
                        item["actionList"] = json.loads(item["actionList"])
                    except Exception:
                        pass
                # groupSceneDto: 如果是 JSON 字符串则解析
                if "groupSceneDto" in item and item["groupSceneDto"]:
                    try:
                        item["groupSceneDto"] = json.loads(item["groupSceneDto"])
                    except Exception:
                        pass
                # 推断 _trType
                # trId 是最可靠的判断依据：
                #   201 → then（执行动作）
                #   101 → if（事件触发）
                #   102 → if（属性变化触发）
                # 如果没有 trId，用 specRelate + key 推断
                tr_id = item.get("trId")
                spec_relate = item.get("specRelate", "")
                spec_type = item.get("specType", "")
                key = item.get("key", "")
                if tr_id:
                    if str(tr_id) == "201":
                        item["_trType"] = "then"
                    elif str(tr_id) in ("101", "102"):
                        item["_trType"] = "if"
                    else:
                        item["_trType"] = "then"
                elif key and key.startswith("prop."):
                    # key=prop.* 是属性变化触发，属于 if
                    item["_trType"] = "if"
                elif key and key.startswith("event."):
                    # key=event.* 是事件触发，属于 if
                    item["_trType"] = "if"
                elif spec_type == "event" or (not spec_type and spec_relate.startswith("event")):
                    item["_trType"] = "if"
                else:
                    item["_trType"] = "then"  # 默认
                auto_items.append(item)

    return config, auto_items


def write_automation_export_excel(path: str, config: dict, auto_list: list):
    """
    导出自动化列表到 Excel
    Sheet1: 产品配置
    Sheet2: 执行动作 (then)  — 聚合选值的 actionList 序列化为 JSON 字符串
    Sheet3: 触发条件 (if)
    """
    import pandas as pd

    # Sheet1: 配置
    config_rows = [
        {"参数名": k, "值": v}
        for k, v in [("userId", config.get("userId", "")),
                      ("xiaomiiot_ph", config.get("xiaomiiot_ph", "")),
                      ("serviceToken", config.get("serviceToken", "")),
                      ("pdId", config.get("pdId", "")),
                      ("model", config.get("model", ""))]
    ]
    df_config = pd.DataFrame(config_rows)

    # 分离 then 和 if
    then_list = [item for item in auto_list if item.get("_trType") == "then"]
    if_list = [item for item in auto_list if item.get("_trType") == "if"]

    # then 的关键字段（去掉内部标记 _trType）
    then_keys = ["intro", "specRelate", "specType", "siId", "subIid",
                 "appValueStyle", "autoType", "value", "command", "tgId", "trId", "saId",
                 "pdId", "model", "platform", "plugId", "fwVer", "mcuFwVer",
                 "attr", "gid", "extra", "ctime", "rank", "ruleId", "specVer",
                 "actionList", "groupSceneDto"]

    # if 的关键字段
    if_keys = ["intro", "specRelate", "specType", "siId", "subIid",
               "appValueStyle", "autoType", "value", "key", "trId", "scId",
               "pdId", "model", "platform", "plugId", "fwVer", "mcuFwVer",
               "attr", "gid", "extra", "ctime", "rank", "ruleId", "specVer",
               "src", "scIds", "groupSceneDto",
               "ifV2Name", "whenV2Name", "diffAutomationName", "advancedConf",
               "valueType", "valueOperation"]

    # actionList 子动作导出时需要保留的关键字段
    ACTION_ITEM_KEYS = ["pdId", "intro", "plugId", "fwVer", "mcuFwVer", "platform",
                        "attr", "specRelate", "model", "appValueStyle", "specType",
                        "siId", "subIid", "autoType", "command", "value", "tgId",
                        "trId", "saId", "gid", "extra", "rank", "ruleId", "specVer"]

    def _build_rows(items, keys):
        rows = []
        # 合并所有 item 中出现的字段，确保 actionList 等可选列不丢失
        all_keys_set = []
        for k in keys:
            if k not in all_keys_set:
                all_keys_set.append(k)
        for item in items:
            for k in item:
                if k not in all_keys_set and k != "_trType":
                    all_keys_set.append(k)

        for item in items:
            row = {}
            for k in all_keys_set:
                v = item.get(k, "")
                # actionList 特殊处理：只保留关键字段，减少 Excel 体积
                if k == "actionList" and isinstance(v, list):
                    cleaned = []
                    for act in v:
                        if isinstance(act, dict):
                            cleaned_act = {}
                            for ak in ACTION_ITEM_KEYS:
                                if ak in act:
                                    cleaned_act[ak] = act[ak]
                            # 保留不在预定义列表但有值的字段
                            for ak, av in act.items():
                                if ak not in cleaned_act and av not in (None, "", 0):
                                    cleaned_act[ak] = av
                            cleaned.append(cleaned_act)
                        else:
                            cleaned.append(act)
                    v = json.dumps(cleaned, ensure_ascii=False)
                elif isinstance(v, (dict, list)):
                    v = json.dumps(v, ensure_ascii=False)
                row[k] = str(v) if v is not None else ""
            rows.append(row)
        return rows, all_keys_set

    then_rows, then_cols = _build_rows(then_list, then_keys)
    if_rows, if_cols = _build_rows(if_list, if_keys)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_config.to_excel(writer, sheet_name="产品配置", index=False)
        if then_rows:
            pd.DataFrame(then_rows, columns=then_cols).to_excel(
                writer, sheet_name="执行动作(then)", index=False)
        if if_rows:
            pd.DataFrame(if_rows, columns=if_cols).to_excel(
                writer, sheet_name="触发条件(if)", index=False)

    return path
