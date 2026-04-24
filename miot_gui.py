#!/usr/bin/env python3
"""
MIoT 属性工具 - GUI 版本
功能：导出模板 / 创建属性 / 生成空白模板
"""

import sys
import os
import json
import time
import traceback

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox,
    QFileDialog, QSpinBox, QGroupBox, QMessageBox, QProgressBar,
    QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QStatusBar, QSplitter,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor, QIcon

# 复用现有脚本的核心逻辑
from miot_export_template import (
    query_services as export_query_services,
    query_properties as export_query_properties,
    parse_prop_row,
    write_prop_sheet,
    write_config_sheet,
    write_source_sheet,
)
from miot_create_properties import (
    query_services as create_query_services,
    query_properties,
    match_service,
    build_request_body,
    create_property,
    parse_access,
    parse_value_list,
    detect_value_type,
    parse_bool,
    read_config,
    read_properties,
    HEADERS,
    BASE,
    CREATE_API,
    QUERY_SERVICES_API,
)
# create_template.py 是顶层执行脚本，无法直接导入函数，GUI 中自行实现模板生成

# ─── 样式 ────────────────────────────────────────────────────

STYLESHEET = """
QMainWindow {
    background-color: #f5f6fa;
}
QTabWidget::pane {
    border: 1px solid #dcdde1;
    border-radius: 6px;
    background: white;
    margin-top: -1px;
}
QTabBar::tab {
    background: #dcdde1;
    padding: 10px 28px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-size: 14px;
    font-weight: bold;
    color: #555;
}
QTabBar::tab:selected {
    background: white;
    color: #2c3e50;
    border-bottom: 2px solid #3498db;
}
QGroupBox {
    font-weight: bold;
    font-size: 13px;
    border: 1px solid #dcdde1;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 18px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #2c3e50;
}
QPushButton {
    background-color: #3498db;
    color: white;
    border: none;
    padding: 8px 20px;
    border-radius: 4px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #2980b9;
}
QPushButton:pressed {
    background-color: #2471a3;
}
QPushButton:disabled {
    background-color: #bdc3c7;
}
QPushButton#dangerBtn {
    background-color: #e74c3c;
}
QPushButton#dangerBtn:hover {
    background-color: #c0392b;
}
QPushButton#successBtn {
    background-color: #27ae60;
}
QPushButton#successBtn:hover {
    background-color: #229954;
}
QLineEdit, QSpinBox, QComboBox {
    padding: 6px 10px;
    border: 1px solid #dcdde1;
    border-radius: 4px;
    font-size: 13px;
    background: white;
}
QLineEdit:focus, QSpinBox:focus {
    border-color: #3498db;
}
QTextEdit {
    border: 1px solid #dcdde1;
    border-radius: 4px;
    font-family: "Menlo", "Consolas", monospace;
    font-size: 12px;
    background: #2c3e50;
    color: #ecf0f1;
    padding: 8px;
}
QProgressBar {
    border: 1px solid #dcdde1;
    border-radius: 4px;
    text-align: center;
    height: 22px;
    background: white;
}
QProgressBar::chunk {
    background-color: #3498db;
    border-radius: 3px;
}
QLabel#titleLabel {
    font-size: 18px;
    font-weight: bold;
    color: #2c3e50;
}
QLabel#subtitleLabel {
    font-size: 12px;
    color: #7f8c8d;
}
"""


# ─── Worker Threads ──────────────────────────────────────────

class ExportWorker(QThread):
    """导出模板工作线程"""
    progress = pyqtSignal(str)       # 日志消息
    finished_ok = pyqtSignal(str)    # 成功，输出文件路径
    finished_err = pyqtSignal(str)   # 失败，错误消息

    def __init__(self, config, pid, model, token, ph, userid, connect_type, output_path, save_json, delay):
        super().__init__()
        self.config = config
        self.pid = pid
        self.model = model
        self.token = token
        self.ph = ph
        self.userid = userid
        self.connect_type = connect_type
        self.output_path = output_path
        self.save_json = save_json
        self.delay = delay

    def run(self):
        try:
            import requests
            cfg = {
                "userId": str(self.userid),
                "pdId": str(self.pid),
                "model": self.model,
                "serviceToken": self.token,
                "xiaomiiot_ph": self.ph,
                "connectType": str(self.connect_type),
            }
            cookies = {
                "serviceToken": self.token,
                "userId": str(self.userid),
                "xiaomiiot_ph": self.ph,
            }
            params_base = {
                "userId": str(self.userid),
                "xiaomiiot_ph": self.ph,
                "pdId": str(self.pid),
            }

            # 1. 查询服务
            self.progress.emit("📋 正在查询产品服务列表...")
            params = dict(params_base)
            params.update({
                "model": self.model,
                "connectType": str(self.connect_type),
                "language": "zh_cn",
                "version": "1",
                "status": "0",
            })
            resp = requests.get(QUERY_SERVICES_API, params=params,
                                headers=HEADERS, cookies=cookies, timeout=15)
            data = resp.json()
            if data.get("status") != 200:
                self.finished_err.emit(f"查询服务失败: {data}")
                return

            services = data.get("result", [])
            if not services:
                self.finished_err.emit("未查到服务，请检查 Cookie 和产品信息")
                return

            self.progress.emit(f"✅ 找到 {len(services)} 个服务")

            # 2. 逐服务查属性
            all_props = []
            for i, svc in enumerate(services):
                siid = svc.get("siid", "?")
                sname = svc.get("description", svc.get("name", ""))
                stype = svc.get("type", "")
                self.progress.emit(f"🔍 [{i+1}/{len(services)}] 查询服务 siid={siid} ({sname})...")

                params = dict(params_base)
                params.update({
                    "version": "1", "status": "0",
                    "siid": str(siid), "serviceType": stype,
                    "model": self.model,
                    "connectType": str(self.connect_type),
                    "language": "zh_cn",
                })
                resp = requests.get(
                    "https://iot.mi.com/cgi-std/api/v1/functionDefine/getInstanceProperties",
                    params=params, headers=HEADERS, cookies=cookies, timeout=15)
                pdata = resp.json()
                props = pdata.get("result", []) if pdata.get("status") == 200 else []

                for p in props:
                    p["_service"] = svc
                all_props.extend(props)

                if self.delay > 0:
                    time.sleep(self.delay)

            self.progress.emit(f"✅ 共获取 {len(all_props)} 条属性")

            # 3. 生成 Excel
            if not self.output_path:
                safe_model = self.model.replace(".", "_").replace("-", "_")
                self.output_path = f"MIoT_模板_{safe_model}.xlsx"

            self.progress.emit(f"📝 正在生成 Excel 模板...")

            from openpyxl import Workbook
            wb = Workbook()

            # 属性定义 Sheet
            ws1 = wb.active
            ws1.title = "属性定义"
            rows_data = []
            for p in all_props:
                row = parse_prop_row(p, p.get("_service", {}))
                rows_data.append(row)
            write_prop_sheet(ws1, rows_data)

            # 公共配置 Sheet
            ws2 = wb.create_sheet("公共配置")
            # 构造 args-like dict for write_config_sheet
            class _Args:
                pass
            args = _Args()
            args.pid = self.pid
            args.model = self.model
            args.token = self.token
            args.ph = self.ph
            args.userid = self.userid
            args.connect_type = self.connect_type
            write_config_sheet(ws2, args)

            # 原始数据参考 Sheet
            ws3 = wb.create_sheet("原始数据参考")
            write_source_sheet(ws3, services, rows_data)

            wb.save(self.output_path)

            # 4. 可选保存 JSON
            if self.save_json:
                json_path = self.output_path.replace(".xlsx", ".json")
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump({"services": services, "properties": all_props},
                              f, ensure_ascii=False, indent=2, default=str)
                self.progress.emit(f"💾 JSON 已保存: {json_path}")

            self.finished_ok.emit(self.output_path)

        except Exception as e:
            self.finished_err.emit(f"导出失败: {traceback.format_exc()}")


class CreateWorker(QThread):
    """创建属性工作线程"""
    progress = pyqtSignal(str)
    update_progress = pyqtSignal(int, int)  # current, total
    finished_ok = pyqtSignal(int, int)      # success, failed
    finished_err = pyqtSignal(str)

    def __init__(self, config, props, services, skip_verify, delay):
        super().__init__()
        self.config = config
        self.props = props
        self.services = services
        self.skip_verify = skip_verify
        self.delay = delay

    def run(self):
        try:
            success = 0
            failed = 0
            results = []

            for i, prop in enumerate(self.props):
                name = prop.get("name", f"行{i+1}")
                fmt = prop.get("format", "?")

                svc = match_service(prop, self.services)
                if svc:
                    siid = svc["siid"]
                    sname = svc.get("description", svc.get("name", ""))
                else:
                    siid = prop.get("siid", "?")
                    sname = "未匹配"

                body = build_request_body(prop, self.config, svc)
                self.update_progress.emit(i + 1, len(self.props))
                self.progress.emit(f"  [{i+1}] {name} → siid={siid} ({sname}) ... ")

                try:
                    resp = create_property(body, self.config)
                    status = resp.get("status")
                    result_val = resp.get("result")
                    if status == 200:
                        piid = result_val
                        self.progress.emit(f"  ✅ {name} 成功 (piid={piid})")
                        success += 1
                        results.append({"name": name, "status": "success", "piid": piid, "siid": siid})
                    else:
                        msg = resp.get("message", resp.get("msg", json.dumps(resp, ensure_ascii=False)))
                        self.progress.emit(f"  ❌ {name} 失败 ({msg})")
                        failed += 1
                        results.append({"name": name, "status": "failed", "error": msg, "siid": siid})
                except Exception as e:
                    self.progress.emit(f"  ❌ {name} 异常 ({e})")
                    failed += 1
                    results.append({"name": name, "status": "error", "error": str(e), "siid": siid})

                if self.delay > 0:
                    time.sleep(self.delay)

            # 保存结果
            result_file = "miot_create_result.json"
            with open(result_file, "w", encoding="utf-8") as f:
                json.dump(results, f, ensure_ascii=False, indent=2)

            self.finished_ok.emit(success, failed)

        except Exception as e:
            self.finished_err.emit(f"创建失败: {traceback.format_exc()}")


# ─── Main Window ─────────────────────────────────────────────

class MIoTMainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MIoT 属性工具")
        self.setMinimumSize(960, 800)
        self.resize(960, 850)

        # 状态
        self.export_worker = None
        self.create_worker = None
        self._services_cache = None
        self._config_cache = None
        self._props_cache = None

        self._init_ui()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 12)

        # 标题
        title = QLabel("🔧 MIoT 属性工具")
        title.setObjectName("titleLabel")
        subtitle = QLabel("小米 IoT 平台设备属性导出 & 批量创建")
        subtitle.setObjectName("subtitleLabel")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_export_tab(), "📤 导出模板")
        self.tabs.addTab(self._build_create_tab(), "📥 创建属性")
        self.tabs.addTab(self._build_template_tab(), "📄 生成模板")
        layout.addWidget(self.tabs)

        # 状态栏
        self.statusBar().showMessage("就绪")

    # ─── 导出 Tab ─────────────────────────────────────────

    def _build_export_tab(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)

        # 左侧：表单
        left = QWidget()
        left_layout = QVBoxLayout(left)

        # 产品信息
        grp_product = QGroupBox("产品信息")
        form = QFormLayout()
        self.exp_pid = QLineEdit()
        self.exp_pid.setPlaceholderText("如 33257")
        self.exp_model = QLineEdit()
        self.exp_model.setPlaceholderText("如 uwize.switch.yzw07")
        self.exp_userid = QLineEdit()
        self.exp_userid.setPlaceholderText("如 1097752639")
        self.exp_connect_type = QSpinBox()
        self.exp_connect_type.setRange(0, 99)
        self.exp_connect_type.setValue(16)
        form.addRow("产品ID (pdId):", self.exp_pid)
        form.addRow("产品型号 (model):", self.exp_model)
        form.addRow("用户ID (userId):", self.exp_userid)
        form.addRow("连接类型:", self.exp_connect_type)
        grp_product.setLayout(form)
        left_layout.addWidget(grp_product)

        # Cookie
        grp_cookie = QGroupBox("Cookie 信息")
        form2 = QFormLayout()
        self.exp_token = QLineEdit()
        self.exp_token.setPlaceholderText("浏览器 Cookie 中的 serviceToken")
        self.exp_token.setEchoMode(QLineEdit.EchoMode.Password)
        self.exp_ph = QLineEdit()
        self.exp_ph.setPlaceholderText("浏览器 Cookie 中的 xiaomiiot_ph")
        self.exp_ph.setEchoMode(QLineEdit.EchoMode.Password)
        self.exp_show_cookie = QCheckBox("显示 Cookie")
        self.exp_show_cookie.toggled.connect(self._toggle_export_cookie_visibility)
        form2.addRow("serviceToken:", self.exp_token)
        form2.addRow("xiaomiiot_ph:", self.exp_ph)
        form2.addRow("", self.exp_show_cookie)
        grp_cookie.setLayout(form2)
        left_layout.addWidget(grp_cookie)

        # 输出选项
        grp_output = QGroupBox("输出选项")
        form3 = QFormLayout()
        self.exp_output = QLineEdit()
        self.exp_output.setPlaceholderText("留空则自动生成")
        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self._browse_export_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.exp_output)
        output_row.addWidget(btn_browse)
        form3.addRow("输出路径:", output_row)
        self.exp_save_json = QCheckBox("同时保存原始 JSON")
        self.exp_delay = QSpinBox()
        self.exp_delay.setRange(0, 5000)
        self.exp_delay.setValue(300)
        self.exp_delay.setSuffix(" ms")
        form3.addRow("", self.exp_save_json)
        form3.addRow("请求间隔:", self.exp_delay)
        grp_output.setLayout(form3)
        left_layout.addWidget(grp_output)

        # 按钮
        btn_layout = QHBoxLayout()
        self.btn_export = QPushButton("🚀 开始导出")
        self.btn_export.setObjectName("successBtn")
        self.btn_export.clicked.connect(self._start_export)
        self.btn_export_cancel = QPushButton("取消")
        self.btn_export_cancel.clicked.connect(self._cancel_export)
        self.btn_export_cancel.setEnabled(False)
        btn_layout.addWidget(self.btn_export)
        btn_layout.addWidget(self.btn_export_cancel)
        left_layout.addLayout(btn_layout)

        left_layout.addStretch()
        left.setFixedWidth(460)

        # 右侧：日志
        right = QWidget()
        right_layout = QVBoxLayout(right)
        lbl = QLabel("运行日志")
        lbl.setStyleSheet("font-weight: bold; font-size: 13px;")
        self.exp_log = QTextEdit()
        self.exp_log.setReadOnly(True)
        self.exp_progress = QProgressBar()
        self.exp_progress.setVisible(False)
        right_layout.addWidget(lbl)
        right_layout.addWidget(self.exp_log)
        right_layout.addWidget(self.exp_progress)

        layout.addWidget(left)
        layout.addWidget(right, stretch=1)

        return widget

    # ─── 创建 Tab ─────────────────────────────────────────

    def _build_create_tab(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)

        # 左侧
        left = QWidget()
        left_layout = QVBoxLayout(left)

        # Excel 文件
        grp_excel = QGroupBox("Excel 文件")
        excel_layout = QHBoxLayout()
        self.crt_excel_path = QLineEdit()
        self.crt_excel_path.setPlaceholderText("选择已填好的 Excel 文件")
        btn_choose = QPushButton("选择文件")
        btn_choose.clicked.connect(self._browse_create_excel)
        excel_layout.addWidget(self.crt_excel_path)
        excel_layout.addWidget(btn_choose)
        grp_excel.setLayout(excel_layout)
        left_layout.addWidget(grp_excel)

        # 产品信息（可覆盖 Excel 中的配置）
        grp_product = QGroupBox("产品信息（可覆盖 Excel 中的配置）")
        form_product = QFormLayout()
        self.crt_pid = QLineEdit()
        self.crt_pid.setPlaceholderText("留空则使用 Excel 中的配置")
        self.crt_model = QLineEdit()
        self.crt_model.setPlaceholderText("留空则使用 Excel 中的配置")
        form_product.addRow("产品ID (pdId):", self.crt_pid)
        form_product.addRow("产品型号 (model):", self.crt_model)
        grp_product.setLayout(form_product)
        left_layout.addWidget(grp_product)

        # Cookie（可从 Excel 读取，也可手动覆盖）
        grp_cookie = QGroupBox("Cookie（可覆盖 Excel 中的配置）")
        form_cookie = QFormLayout()
        self.crt_token = QLineEdit()
        self.crt_token.setEchoMode(QLineEdit.EchoMode.Password)
        self.crt_token.setPlaceholderText("留空则使用 Excel 中的配置")
        self.crt_ph = QLineEdit()
        self.crt_ph.setEchoMode(QLineEdit.EchoMode.Password)
        self.crt_ph.setPlaceholderText("留空则使用 Excel 中的配置")
        self.crt_userid = QLineEdit()
        self.crt_userid.setPlaceholderText("留空则使用 Excel 中的配置")
        self.crt_show_cookie = QCheckBox("显示 Cookie")
        self.crt_show_cookie.toggled.connect(self._toggle_create_cookie_visibility)
        form_cookie.addRow("serviceToken:", self.crt_token)
        form_cookie.addRow("xiaomiiot_ph:", self.crt_ph)
        form_cookie.addRow("userId:", self.crt_userid)
        form_cookie.addRow("", self.crt_show_cookie)
        grp_cookie.setLayout(form_cookie)
        left_layout.addWidget(grp_cookie)

        # 选项
        grp_opts = QGroupBox("选项")
        form_opts = QFormLayout()
        self.crt_skip_verify = QCheckBox("跳过验证（推荐）")
        self.crt_skip_verify.setChecked(True)
        self.crt_delay = QSpinBox()
        self.crt_delay.setRange(0, 10000)
        self.crt_delay.setValue(500)
        self.crt_delay.setSuffix(" ms")
        self.crt_siid = QSpinBox()
        self.crt_siid.setRange(0, 999)
        self.crt_siid.setValue(0)
        self.crt_siid.setSpecialValueText("全部")
        form_opts.addRow("", self.crt_skip_verify)
        form_opts.addRow("请求间隔:", self.crt_delay)
        form_opts.addRow("指定 siid:", self.crt_siid)
        grp_opts.setLayout(form_opts)
        left_layout.addWidget(grp_opts)

        # 按钮
        btn_layout = QVBoxLayout()
        row1 = QHBoxLayout()
        self.btn_dryrun = QPushButton("🧪 干跑检查")
        self.btn_dryrun.clicked.connect(self._start_dryrun)
        self.btn_list_services = QPushButton("📋 查看服务")
        self.btn_list_services.clicked.connect(self._list_services)
        row1.addWidget(self.btn_dryrun)
        row1.addWidget(self.btn_list_services)
        btn_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.btn_create = QPushButton("🚀 开始创建")
        self.btn_create.setObjectName("dangerBtn")
        self.btn_create.clicked.connect(self._start_create)
        self.btn_create_cancel = QPushButton("取消")
        self.btn_create_cancel.clicked.connect(self._cancel_create)
        self.btn_create_cancel.setEnabled(False)
        row2.addWidget(self.btn_create)
        row2.addWidget(self.btn_create_cancel)
        btn_layout.addLayout(row2)

        left_layout.addLayout(btn_layout)
        left_layout.addStretch()
        left.setFixedWidth(460)

        # 右侧：日志
        right = QWidget()
        right_layout = QVBoxLayout(right)
        lbl = QLabel("运行日志")
        lbl.setStyleSheet("font-weight: bold; font-size: 13px;")
        self.crt_log = QTextEdit()
        self.crt_log.setReadOnly(True)
        self.crt_progress = QProgressBar()
        self.crt_progress.setVisible(False)
        right_layout.addWidget(lbl)
        right_layout.addWidget(self.crt_log)
        right_layout.addWidget(self.crt_progress)

        layout.addWidget(left)
        layout.addWidget(right, stretch=1)

        return widget

    # ─── 模板 Tab ─────────────────────────────────────────

    def _build_template_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        grp = QGroupBox("生成空白 Excel 模板")
        form = QFormLayout()

        self.tpl_output = QLineEdit("MIoT_属性创建模板.xlsx")
        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self._browse_template_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.tpl_output)
        output_row.addWidget(btn_browse)
        form.addRow("输出路径:", output_row)

        btn_gen = QPushButton("📄 生成模板")
        btn_gen.setObjectName("successBtn")
        btn_gen.clicked.connect(self._generate_template)
        form.addRow("", btn_gen)

        grp.setLayout(form)
        layout.addWidget(grp)

        layout.addStretch()
        return widget

    # ─── Cookie 可见性 ───────────────────────────────────

    def _toggle_export_cookie_visibility(self, checked):
        mode = QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        self.exp_token.setEchoMode(mode)
        self.exp_ph.setEchoMode(mode)

    def _toggle_create_cookie_visibility(self, checked):
        mode = QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        self.crt_token.setEchoMode(mode)
        self.crt_ph.setEchoMode(mode)

    # ─── 文件浏览 ─────────────────────────────────────────

    def _browse_export_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "选择输出路径", "", "Excel (*.xlsx)")
        if path:
            self.exp_output.setText(path)

    def _browse_create_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel (*.xlsx *.xls)")
        if path:
            self.crt_excel_path.setText(path)

    def _browse_template_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "选择输出路径", "MIoT_属性创建模板.xlsx", "Excel (*.xlsx)")
        if path:
            self.tpl_output.setText(path)

    # ─── 导出功能 ─────────────────────────────────────────

    def _start_export(self):
        # 校验
        pid = self.exp_pid.text().strip()
        model = self.exp_model.text().strip()
        token = self.exp_token.text().strip()
        ph = self.exp_ph.text().strip()
        userid = self.exp_userid.text().strip()

        if not all([pid, model, token, ph, userid]):
            QMessageBox.warning(self, "提示", "请填写产品信息和 Cookie")
            return

        self.exp_log.clear()
        self.btn_export.setEnabled(False)
        self.btn_export_cancel.setEnabled(True)
        self.exp_progress.setVisible(True)
        self.exp_progress.setRange(0, 0)  # 不确定进度

        self.export_worker = ExportWorker(
            config={},
            pid=pid,
            model=model,
            token=token,
            ph=ph,
            userid=userid,
            connect_type=self.exp_connect_type.value(),
            output_path=self.exp_output.text().strip() or None,
            save_json=self.exp_save_json.isChecked(),
            delay=self.exp_delay.value() / 1000.0,
        )
        self.export_worker.progress.connect(self._export_log)
        self.export_worker.finished_ok.connect(self._export_done_ok)
        self.export_worker.finished_err.connect(self._export_done_err)
        self.export_worker.start()

    def _cancel_export(self):
        if self.export_worker and self.export_worker.isRunning():
            self.export_worker.terminate()
            self.exp_log.append("⚠️ 已取消")
        self._reset_export_ui()

    def _export_log(self, msg):
        self.exp_log.append(msg)
        self.statusBar().showMessage(msg)

    def _export_done_ok(self, path):
        self._reset_export_ui()
        self.exp_log.append(f"\n🎉 导出成功！文件: {path}")
        self.statusBar().showMessage(f"导出成功: {path}")
        QMessageBox.information(self, "导出成功", f"模板已保存到:\n{path}")

    def _export_done_err(self, msg):
        self._reset_export_ui()
        self.exp_log.append(f"\n❌ {msg}")
        self.statusBar().showMessage("导出失败")
        QMessageBox.critical(self, "导出失败", msg)

    def _reset_export_ui(self):
        self.btn_export.setEnabled(True)
        self.btn_export_cancel.setEnabled(False)
        self.exp_progress.setVisible(False)

    # ─── 创建功能 ─────────────────────────────────────────

    def _load_create_config(self):
        """加载 Excel 配置，Cookie 支持手动覆盖"""
        path = self.crt_excel_path.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "提示", "请选择有效的 Excel 文件")
            return None, None

        from openpyxl import load_workbook
        wb = load_workbook(path)
        config = read_config(wb["公共配置"])
        props = read_properties(wb["属性定义"])

        # 手动 Cookie 覆盖
        if self.crt_token.text().strip():
            config["serviceToken"] = self.crt_token.text().strip()
        if self.crt_ph.text().strip():
            config["xiaomiiot_ph"] = self.crt_ph.text().strip()
        if self.crt_userid.text().strip():
            config["userId"] = self.crt_userid.text().strip()

        # 手动产品信息覆盖
        if self.crt_pid.text().strip():
            config["pdId"] = self.crt_pid.text().strip()
        if self.crt_model.text().strip():
            config["model"] = self.crt_model.text().strip()

        # 检查必填
        missing = []
        for key in ["serviceToken", "xiaomiiot_ph", "userId", "pdId", "model"]:
            if not config.get(key):
                missing.append(key)
        if missing:
            QMessageBox.warning(self, "配置缺失", f"公共配置缺少必填项:\n{', '.join(missing)}")
            return None, None

        if not props:
            QMessageBox.warning(self, "提示", "属性定义为空")
            return None, None

        return config, props

    def _list_services(self):
        config, _ = self._load_create_config()
        if not config:
            return

        self.crt_log.clear()
        self.crt_log.append("📋 正在查询服务列表...")

        try:
            services = create_query_services(config)
            if not services:
                self.crt_log.append("❌ 未查到服务，请检查 Cookie 和产品信息")
                return

            self._services_cache = services
            self._config_cache = config

            self.crt_log.append(f"\n📋 产品服务列表（共 {len(services)} 个）:")
            self.crt_log.append("-" * 60)
            self.crt_log.append(f"{'siid':>4} | {'name':<24} | {'description':<20} | type")
            self.crt_log.append("-" * 60)
            for svc in services:
                siid = svc.get("siid", "?")
                name = svc.get("name", "?")
                desc = svc.get("description", "")
                stype = svc.get("type", "")
                self.crt_log.append(f"{siid:>4} | {name:<24} | {desc:<20} | {stype}")

        except Exception as e:
            self.crt_log.append(f"❌ 查询失败: {e}")

    def _start_dryrun(self):
        config, props = self._load_create_config()
        if not config or not props:
            return

        self.crt_log.clear()
        self.crt_log.append("🧪 干跑模式 - 正在匹配服务...\n")

        try:
            services = create_query_services(config)
            if not services:
                self.crt_log.append("❌ 未查到服务列表")
                return

            self._services_cache = services
            self._config_cache = config
            self._props_cache = props

            target_siid = self.crt_siid.value()
            tasks = []
            for i, prop in enumerate(props):
                name = prop.get("name", f"行{i+1}")
                fmt = prop.get("format", "?")
                vtype = detect_value_type(fmt, prop)

                svc = match_service(prop, services)
                if svc:
                    siid = svc["siid"]
                    sname = svc.get("description", svc.get("name", ""))
                else:
                    siid = prop.get("siid", "?")
                    sname = "❌ 未匹配"

                if target_siid > 0 and int(siid) != target_siid:
                    continue

                body = build_request_body(prop, config, svc)
                tasks.append({
                    "index": i + 1,
                    "name": name,
                    "format": fmt,
                    "vtype": vtype,
                    "siid": siid,
                    "service_name": sname,
                })

            if not tasks:
                self.crt_log.append("❌ 没有匹配的属性")
                return

            self.crt_log.append(f"{'#':>3} | {'name':<20} | {'format':<8} | {'vtype':<10} | siid | 服务")
            self.crt_log.append("-" * 80)
            for t in tasks:
                self.crt_log.append(
                    f"{t['index']:>3} | {t['name']:<20} | {t['format']:<8} | "
                    f"{t['vtype']:<10} | {t['siid']:<4} | {t['service_name']}")

            self.crt_log.append(f"\n🏁 共 {len(tasks)} 条属性待创建（干跑模式，未实际执行）")

            # 显示第一条请求体
            if tasks:
                body = build_request_body(props[0], config,
                                         match_service(props[0], services))
                self.crt_log.append(f"\n📄 第1条请求体示例:")
                self.crt_log.append(json.dumps(body, ensure_ascii=False, indent=2))

        except Exception as e:
            self.crt_log.append(f"❌ 错误: {traceback.format_exc()}")

    def _start_create(self):
        config, props = self._load_create_config()
        if not config or not props:
            return

        # 确认
        reply = QMessageBox.question(
            self, "确认创建",
            f"即将创建 {len(props)} 条属性，是否继续？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # 先查服务
        try:
            services = create_query_services(config)
            if not services:
                QMessageBox.critical(self, "错误", "未查到服务列表，请检查 Cookie 和产品信息")
                return
        except Exception as e:
            QMessageBox.critical(self, "错误", f"查询服务失败: {e}")
            return

        self.crt_log.clear()
        self.crt_log.append(f"🚀 开始创建 {len(props)} 条属性...\n")
        self.btn_create.setEnabled(False)
        self.btn_dryrun.setEnabled(False)
        self.btn_list_services.setEnabled(False)
        self.btn_create_cancel.setEnabled(True)
        self.crt_progress.setVisible(True)
        self.crt_progress.setRange(0, len(props))

        # 过滤 siid
        target_siid = self.crt_siid.value()
        if target_siid > 0:
            filtered_props = []
            for prop in props:
                svc = match_service(prop, services)
                siid = svc["siid"] if svc else prop.get("siid", 0)
                if int(siid) == target_siid:
                    filtered_props.append(prop)
            props = filtered_props

        self.create_worker = CreateWorker(
            config=config,
            props=props,
            services=services,
            skip_verify=self.crt_skip_verify.isChecked(),
            delay=self.crt_delay.value() / 1000.0,
        )
        self.create_worker.progress.connect(self._create_log)
        self.create_worker.update_progress.connect(self._create_update_progress)
        self.create_worker.finished_ok.connect(self._create_done_ok)
        self.create_worker.finished_err.connect(self._create_done_err)
        self.create_worker.start()

    def _cancel_create(self):
        if self.create_worker and self.create_worker.isRunning():
            self.create_worker.terminate()
            self.crt_log.append("⚠️ 已取消")
        self._reset_create_ui()

    def _create_log(self, msg):
        self.crt_log.append(msg)
        self.statusBar().showMessage(msg)

    def _create_update_progress(self, current, total):
        self.crt_progress.setValue(current)

    def _create_done_ok(self, success, failed):
        self._reset_create_ui()
        self.crt_log.append(f"\n{'='*50}")
        self.crt_log.append(f"📊 创建完成: 成功 {success} / 失败 {failed} / 共 {success + failed}")
        self.statusBar().showMessage(f"创建完成: {success} 成功, {failed} 失败")
        if failed > 0:
            QMessageBox.warning(self, "创建完成", f"成功 {success}, 失败 {failed}\n详情见日志")
        else:
            QMessageBox.information(self, "创建完成", f"🎉 全部 {success} 条属性创建成功！")

    def _create_done_err(self, msg):
        self._reset_create_ui()
        self.crt_log.append(f"\n❌ {msg}")
        QMessageBox.critical(self, "创建失败", msg)

    def _reset_create_ui(self):
        self.btn_create.setEnabled(True)
        self.btn_dryrun.setEnabled(True)
        self.btn_list_services.setEnabled(True)
        self.btn_create_cancel.setEnabled(False)
        self.crt_progress.setVisible(False)

    # ─── 生成模板 ─────────────────────────────────────────

    def _generate_template(self):
        path = self.tpl_output.text().strip()
        if not path:
            QMessageBox.warning(self, "提示", "请填写输出路径")
            return
        try:
            self._generate_blank_template(path)
            QMessageBox.information(self, "成功", f"模板已生成:\n{path}")
            self.statusBar().showMessage(f"模板已生成: {path}")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"生成模板失败:\n{e}")

    @staticmethod
    def _generate_blank_template(output_path: str):
        """生成空白 Excel 模板"""
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.worksheet.datavalidation import DataValidation

        wb = Workbook()

        # 样式
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill("solid", fgColor="4472C4")
        opt_header_fill = PatternFill("solid", fgColor="8DB4E2")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"))
        desc_font = Font(name="Arial", size=9, color="666666")
        desc_fill = PatternFill("solid", fgColor="D9E2F3")
        opt_desc_fill = PatternFill("solid", fgColor="E8F0FE")

        # Sheet 1: 属性定义
        ws = wb.active
        ws.title = "属性定义"

        columns = [
            ("name",              20, "属性英文名\n如 on、mode",                                True),
            ("description",       20, "属性中文描述\n如 开关、模式",                            True),
            ("format",            12, "数据格式\nbool/uint8/uint16/uint32/string",             True),
            ("service_desc",      22, "服务中文名\n如「开关一键」「按键1点动毫秒数」",              True),
            ("value_list",        28, "枚举值\n格式: 0:关闭,1:开启,2:待机",                     False),
            ("value_range_min",   14, "数值范围-最小值",                                        False),
            ("value_range_max",   14, "数值范围-最大值",                                        False),
            ("value_range_step",  14, "数值范围-步长",                                          False),
            ("siid",              8,  "服务ID（备选）\n直接指定siid，填了则忽略service匹配",    False),
            ("access",            20, "访问权限\n默认: read,write,notify\n（gattAccess自动等同于access）", False),
            ("service_name",      20, "服务英文名（可选）\n配合service_desc精确区分",            False),
        ]

        for i, (col_name, width, desc, required) in enumerate(columns, 1):
            col_letter = chr(64 + i)
            ws.column_dimensions[col_letter].width = width

            fill = header_fill if required else opt_header_fill
            cell = ws.cell(row=1, column=i, value=col_name)
            cell.font = header_font
            cell.fill = fill
            cell.alignment = header_align
            cell.border = thin_border

            desc_fill_use = desc_fill if required else opt_desc_fill
            desc_cell = ws.cell(row=2, column=i, value=desc)
            desc_cell.font = desc_font
            desc_cell.fill = desc_fill_use
            desc_cell.alignment = Alignment(vertical="center", wrap_text=True)
            desc_cell.border = thin_border

        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 50

        # format 下拉
        dv = DataValidation(type="list", formula1='"bool,uint8,uint16,uint32,string"', allow_blank=True)
        dv.error = "请选择有效的数据格式"
        dv.errorTitle = "无效格式"
        ws.add_data_validation(dv)
        dv.add(f"C3:C1000")

        # Sheet 2: 公共配置
        ws2 = wb.create_sheet("公共配置")
        ws2.column_dimensions["A"].width = 22
        ws2.column_dimensions["B"].width = 65
        ws2.column_dimensions["C"].width = 40

        config_items = [
            ("userId",          "",          "小米账号用户ID（必填）",                           True),
            ("pdId",            "",          "产品ID（必填）",                                   True),
            ("model",           "",          "设备型号，如 uwize.switch.aiswi（必填）",           True),
            ("serviceToken",    "",          "浏览器 Cookie 获取（必填）",                        True),
            ("xiaomiiot_ph",    "",          "浏览器 Cookie 获取（必填）",                        True),
            ("connectType",     "16",        "连接类型（默认16）",                                False),
            ("language",        "zh_cn",     "语言（默认zh_cn）",                                 False),
            ("version",         "1",         "版本（默认1）",                                     False),
            ("status",          "0",         "状态（默认0）",                                     False),
            ("source",          "4",         "来源（默认4）",                                     False),
            ("standard",        "false",     "标准属性（默认false）",                              False),
            ("access",          "read,write,notify", "默认访问权限（gattAccess自动等同于access）",  False),
        ]

        for i, (key, val, desc, required) in enumerate(config_items, 1):
            key_cell = ws2.cell(row=i, column=1, value=key)
            val_cell = ws2.cell(row=i, column=2, value=val)
            desc_cell = ws2.cell(row=i, column=3, value=desc)

            if required:
                key_cell.font = Font(name="Arial", bold=True, color="CC0000")
            val_cell.border = thin_border
            desc_cell.font = desc_font

        # Sheet 3: 填写说明
        ws3 = wb.create_sheet("填写说明")
        ws3.column_dimensions["A"].width = 20
        ws3.column_dimensions["B"].width = 80

        instructions = [
            ("必填列", "name / description / format / service_desc"),
            ("枚举属性", "value_list 列填写格式: 0:关闭,1:开启,2:待机"),
            ("数值属性", "value_range_min / value_range_max / value_range_step 三列"),
            ("bool 属性", "format 填 bool，value_list 和 value_range 都留空"),
            ("string 属性", "format 填 string，value_list 和 value_range 都留空"),
            ("服务匹配", "优先用 service_desc（服务中文名）匹配，换产品后 siid 会变"),
            ("siid 列", "可选，填了则忽略 service 匹配直接使用（换产品后不可靠）"),
            ("access 列", "可选，默认 read,write,notify（gattAccess 自动等同于 access）"),
            ("公共配置", "标红的为必填项，其他有默认值可不改"),
        ]
        ws3.cell(row=1, column=1, value="项目").font = Font(bold=True, size=12)
        ws3.cell(row=1, column=2, value="说明").font = Font(bold=True, size=12)
        for i, (item, desc) in enumerate(instructions, 2):
            ws3.cell(row=i, column=1, value=item).font = Font(bold=True)
            ws3.cell(row=i, column=2, value=desc)

        wb.save(output_path)


# ─── Main ────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)

    window = MIoTMainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
