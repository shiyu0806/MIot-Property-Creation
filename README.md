# MIoT Property Creation

小米 MIoT 平台设备属性批量创建工具。从 Excel 读取属性定义，自动匹配服务并批量创建属性。

## 功能

- 🔍 自动查询产品服务列表
- 🎯 智能匹配属性到服务（service_desc / service_name / siid）
- 📋 支持 bool / 数值型 / 枚举型 / string 多种属性格式
- 🔄 通用设计 — 换产品只需改 Excel 公共配置
- 🧪 支持 dry-run 预检

## 快速开始

### 1. 安装依赖

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl requests
```

### 2. 配置 Excel

打开 `MIoT_属性创建模板.xlsx`，填写两个 Sheet：

**公共配置（必填项标红）：**

| 配置项 | 说明 |
|--------|------|
| userId | 小米账号用户ID |
| pdId | 产品ID |
| model | 设备型号，如 uwize.switch.aiswi |
| serviceToken | 浏览器 Cookie 获取 |
| xiaomiiot_ph | 浏览器 Cookie 获取 |

**属性定义：**

| 列名 | 必填 | 说明 |
|------|------|------|
| name | ✅ | 属性英文名，如 on、mode |
| description | ✅ | 属性中文描述，如 开关、模式 |
| format | ✅ | 数据格式：bool/uint8/uint16/uint32/string |
| service_desc | ✅ | 服务中文名，如「开关一键」「按键1点动毫秒数」 |
| value_list | 枚举时 | 格式：`0:关闭,1:开启,2:待机` |
| value_range_min/max/step | 数值时 | 数值范围 |
| siid | 可选 | 直接指定服务ID |
| access | 可选 | 默认 read,write,notify |

### 3. 获取 Cookie

1. 登录 [iot.mi.com](https://iot.mi.com)
2. F12 → Application → Cookies
3. 复制 `serviceToken` 和 `xiaomiiot_ph` 的值

### 4. 运行

```bash
# 查看产品服务列表
python miot_create_properties.py --list-services

# 干跑检查（不实际创建）
python miot_create_properties.py --dry-run

# 正式创建（逐条确认）
python miot_create_properties.py --skip-verify

# 跳过确认直接创建
python miot_create_properties.py --skip-verify -y

# 只创建指定 siid 下的属性
python miot_create_properties.py --skip-verify --siid 13
```

## 服务匹配优先级

1. service_name + service_desc 同时精确匹配
2. service_desc 精确匹配（推荐）
3. service_name + service_desc 模糊匹配
4. service_name 精确匹配
5. service_desc 包含匹配
6. siid 直接匹配（兜底）

> 💡 填写 `service_desc`（服务中文名）是最简单的匹配方式，换产品后无需修改 siid。

## 属性格式说明

| format | value_type | valueList | valueRange |
|--------|-----------|-----------|------------|
| bool | bool_range | `[]` | `[]` |
| uint8/uint16/uint32（数值型） | number | `[]` | `[min, max, step]` |
| uint8/uint16/uint32（枚举型） | enum | `[{value, description}]` | `[]` |
| string | string | `[]` | `[]` |

## 文件说明

| 文件 | 说明 |
|------|------|
| `miot_create_properties.py` | 主脚本 — 读取 Excel 并批量创建属性 |
| `create_template.py` | 模板生成脚本 — 生成 Excel 模板 |
| `MIoT_属性创建模板.xlsx` | Excel 模板 — 属性定义 + 公共配置 + 填写说明 |

## 注意事项

- Cookie 有有效期，过期需重新获取
- 验证接口（verify）返回 405，建议使用 `--skip-verify` 跳过
- 同一服务下属性名（name）不可重复
