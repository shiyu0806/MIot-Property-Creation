# MIoT Property Creation

小米 MIoT 平台设备属性批量创建工具。从 Excel 读取属性定义，自动匹配服务并批量创建属性。支持从已有产品导出属性模板，快速复用到新产品。

## 功能

- 🔍 自动查询产品服务列表
- 🎯 智能匹配属性到服务（service_desc / service_name / siid）
- 📋 支持 bool / 数值型 / 枚举型 / string 多种属性格式
- 🔄 通用设计 — 换产品只需改 Excel 公共配置
- 🧪 支持 dry-run 预检
- 📤 **一键导出** — 从已有产品抓取全部属性，生成可复用 Excel 模板

## 快速开始

### 1. 安装依赖

**macOS / Linux：**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl requests
```

**Windows（PowerShell）：**

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install openpyxl requests
```

> 💡 如果 `python3` 命令不存在，尝试 `python`。虚拟环境激活后，后续命令统一使用 `python`（不再需要 `.venv/bin/python` 或 `.venv\Scripts\python` 前缀）。

### 2. 获取 Cookie

1. 登录 [iot.mi.com](https://iot.mi.com)
2. F12 → Application → Cookies
3. 复制 `serviceToken` 和 `xiaomiiot_ph` 的值

---

## 方式一：从已有产品导出模板（推荐）

如果你有一个已经配好属性的产品，可以直接导出它的所有属性，生成 Excel 模板后复用到新产品。

### 导出属性

**macOS / Linux：**

```bash
python miot_export_template.py \
  --pid 33257 \
  --model uwize.switch.yzw07 \
  --token '你的serviceToken' \
  --ph '你的xiaomiiot_ph' \
  --userid 1097752639
```

**Windows（PowerShell）：**

> ⚠️ PowerShell 中参数必须写在同一行，不支持 `\` 续行。含特殊字符的 token 建议用双引号包裹。

```powershell
python miot_export_template.py --pid 33257 --model uwize.switch.yzw07 --token "你的serviceToken" --ph "你的xiaomiiot_ph" --userid 1097752639
```

> 💡 以上命令已激活虚拟环境，直接用 `python` 即可。如未激活，macOS 用 `.venv/bin/python`，Windows 用 `.venv\Scripts\python`。

**参数说明：**

| 参数 | 说明 |
|------|------|
| `--pid` | 来源产品 ID（已有属性的产品） |
| `--model` | 来源产品型号 |
| `--token` | serviceToken（Cookie 获取） |
| `--ph` | xiaomiiot_ph（Cookie 获取） |
| `--userid` | 小米账号用户ID |
| `-o` | 输出文件路径（默认自动生成） |
| `--json` | 同时输出原始 JSON 数据 |
| `--delay` | 请求间隔秒数（默认 0.3） |

### 生成的 Excel

导出后自动生成包含三个 Sheet 的 Excel 文件：

| Sheet | 内容 |
|-------|------|
| **属性定义** | 已自动填好所有属性的 name/description/format/service_desc/value_list/value_range/access |
| **公共配置** | Cookie 已预填，**需修改 pdId 和 model 为目标产品** |
| **原始数据参考** | 来源产品的服务列表 + 完整属性元数据，供参考 |

### 修改并创建

1. 打开导出的 Excel，在「公共配置」Sheet 中修改 pdId 和 model
2. 可选：删掉不需要的属性行

**macOS / Linux：**

```bash
# 干跑验证
python miot_create_properties.py --excel MIoT_模板_uwize_switch_yzw07.xlsx --dry-run

# 正式创建
python miot_create_properties.py --excel MIoT_模板_uwize_switch_yzw07.xlsx --skip-verify -y
```

**Windows（PowerShell）：**

```powershell
# 干跑验证
python miot_create_properties.py --excel MIoT_模板_uwize_switch_yzw07.xlsx --dry-run

# 正式创建
python miot_create_properties.py --excel MIoT_模板_uwize_switch_yzw07.xlsx --skip-verify -y
```

---

## 方式二：手动填写模板

如果没有参照产品，可以使用空白模板手动填写。

### 配置 Excel

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

### 运行

以下命令默认已激活虚拟环境（直接用 `python`）。如未激活，macOS 用 `.venv/bin/python`，Windows 用 `.venv\Scripts\python`。

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

# 使用指定的 Excel 文件
python miot_create_properties.py --excel MIoT_模板_uwize_switch_yzw07.xlsx --dry-run
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
| `miot_export_template.py` | 📤 导出工具 — 从已有产品抓取属性，生成可复用 Excel 模板 |
| `miot_create_properties.py` | 📥 创建工具 — 读取 Excel 并批量创建属性 |
| `create_template.py` | 📄 模板生成脚本 — 生成空白 Excel 模板 |
| `MIoT_属性创建模板.xlsx` | 📋 空白 Excel 模板 — 属性定义 + 公共配置 + 填写说明 |

## 典型工作流

```
已有产品 → miot_export_template.py → Excel 模板 → 修改 pdId/model → miot_create_properties.py → 新产品
```

## 注意事项

- Cookie 有有效期，过期需重新获取
- 验证接口（verify）返回 405，建议使用 `--skip-verify` 跳过
- 同一服务下属性名（name）不可重复
- 导出模板后，务必在公共配置中修改 pdId 和 model 为目标产品
- Windows PowerShell 中参数必须写在同一行，不支持 `\` 续行
