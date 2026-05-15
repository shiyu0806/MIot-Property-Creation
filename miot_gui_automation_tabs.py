#!/usr/bin/env python3
"""Automation tabs for the MIoT GUI."""

import os

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QPushButton, QCheckBox, QFileDialog,
    QSpinBox, QGroupBox, QMessageBox,
)

from miot_service_core import check_product_status
from miot_automation_core import read_automation_excel
from miot_gui_common import (
    _make_log_panel,
    _make_progress,
    _make_left_panel,
    _inject_group_id,
    _cookie_group,
    _polish_group,
)
from miot_gui_workers import ExportAutomationWorker, CreateAutomationWorker


class ExportAutomationTab(QWidget):
    """导出自定义自动化列表"""
    def __init__(self):
        super().__init__()
        self._worker = None
        self._build()

    def _build(self):
        layout = QHBoxLayout(self)

        left, _, lv = _make_left_panel()

        # 产品信息
        grp_prod = QGroupBox("产品信息")
        form_prod = QFormLayout()
        self.pid_edit = QLineEdit(); self.pid_edit.setPlaceholderText("留空使用 Excel 配置")
        self.model_edit = QLineEdit(); self.model_edit.setPlaceholderText("留空使用 Excel 配置")
        form_prod.addRow("产品ID (pdId):", self.pid_edit)
        form_prod.addRow("产品型号 (model):", self.model_edit)
        _polish_group(grp_prod, form_prod)
        lv.addWidget(grp_prod)

        # Cookie 信息（自动填充 + 手动覆盖）
        _, self.token_edit, self.ph_edit, self.userid_edit = _cookie_group(lv, "auto_exp")
        self.token_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")
        self.ph_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")
        self.userid_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")

        # 导出文件夹
        grp_out = QGroupBox("导出选项")
        ov = QFormLayout()
        self.out_edit = QLineEdit(); self.out_edit.setPlaceholderText("点击浏览选择导出文件夹")
        self.out_edit.setReadOnly(True)
        btn_br = QPushButton("浏览...")
        btn_br.clicked.connect(self._browse_out)
        row = QHBoxLayout(); row.addWidget(self.out_edit); row.addWidget(btn_br)
        ov.addRow("导出文件夹:", row)
        _polish_group(grp_out, ov)
        lv.addWidget(grp_out)

        # 按钮
        btn_row = QHBoxLayout()
        self.btn_export = QPushButton("📤 导出自动化")
        self.btn_export.setObjectName("successBtn")
        self.btn_export.clicked.connect(self._start)
        btn_row.addWidget(self.btn_export)
        lv.addLayout(btn_row)
        lv.addStretch()

        # 右侧日志
        right = QWidget(); rv = QVBoxLayout(right)
        self.log = _make_log_panel(rv)
        self.progress = _make_progress(rv)

        layout.addWidget(left)
        layout.addWidget(right, stretch=1)

    def _browse_out(self):
        path = QFileDialog.getExistingDirectory(self, "选择导出文件夹")
        if path:
            self.out_edit.setText(path)

    def _build_config(self):
        config = {
            "serviceToken": self.token_edit.text().strip(),
            "xiaomiiot_ph": self.ph_edit.text().strip(),
            "userId":       self.userid_edit.text().strip(),
            "pdId":         self.pid_edit.text().strip(),
            "model":        self.model_edit.text().strip(),
        }
        missing = [k for k in ("userId", "xiaomiiot_ph", "pdId") if not config.get(k)]
        if missing:
            QMessageBox.warning(self, "提示", f"缺少必填项:\n{', '.join(missing)}")
            return None

        # 自动注入 groupId
        _inject_group_id(config)

        return config

    def _start(self):
        config = self._build_config()
        if not config:
            return

        out_dir = self.out_edit.text().strip()
        if not out_dir:
            out_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        safe_model = config.get("model", "unknown").replace(".", "_").replace("-", "_")
        output_path = os.path.join(out_dir, f"{safe_model}_automation_export.xlsx")

        self.log.clear()
        self.btn_export.setEnabled(False)
        self.progress.setVisible(True); self.progress.setRange(0, 0)

        self._worker = ExportAutomationWorker(config, output_path)
        self._worker.progress.connect(self.log.append)
        self._worker.finished_ok.connect(self._done_ok)
        self._worker.finished_err.connect(self._done_err)
        self._worker.start()

    def _done_ok(self, path):
        self.btn_export.setEnabled(True)
        self.progress.setVisible(False)
        self.log.append(f"\n🎉 导出成功: {path}")
        QMessageBox.information(self, "导出成功", f"文件已保存:\n{path}")

    def _done_err(self, msg):
        self.btn_export.setEnabled(True)
        self.progress.setVisible(False)
        self.log.append(f"\n❌ {msg}")
        QMessageBox.critical(self, "导出失败", msg)


class CreateAutomationTab(QWidget):
    """批量创建自定义自动化"""
    def __init__(self):
        super().__init__()
        self._worker = None
        self._build()

    def _build(self):
        layout = QHBoxLayout(self)

        left, _, lv = _make_left_panel()

        # Excel 文件
        grp_excel = QGroupBox("自动化 Excel（包含配置 + 自动化列表）")
        ev = QHBoxLayout()
        self.excel_edit = QLineEdit(); self.excel_edit.setPlaceholderText("选择自动化 Excel 文件")
        btn_xl = QPushButton("浏览...")
        btn_xl.clicked.connect(self._browse_file)
        ev.addWidget(self.excel_edit); ev.addWidget(btn_xl)
        _polish_group(grp_excel, ev)
        lv.addWidget(grp_excel)

        # 产品信息覆盖
        grp_prod = QGroupBox("产品信息（可覆盖 Excel 配置）")
        form_prod = QFormLayout()
        self.pid_edit = QLineEdit(); self.pid_edit.setPlaceholderText("留空使用 Excel 配置")
        self.model_edit = QLineEdit(); self.model_edit.setPlaceholderText("留空使用 Excel 配置")
        form_prod.addRow("产品ID (pdId):", self.pid_edit)
        form_prod.addRow("产品型号 (model):", self.model_edit)
        _polish_group(grp_prod, form_prod)
        lv.addWidget(grp_prod)

        # Cookie 信息（自动填充 + 手动覆盖）
        _, self.token_edit, self.ph_edit, self.userid_edit = _cookie_group(lv, "auto_crt")
        self.token_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")
        self.ph_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")
        self.userid_edit.setPlaceholderText("留空使用 Excel 配置或已登录账号")

        # 选项
        grp_opt = QGroupBox("选项")
        ov = QFormLayout()
        self.chk_dryrun = QCheckBox("Dry-run（仅预检，不实际创建）")
        ov.addRow("", self.chk_dryrun)
        self.delay_spin = QSpinBox(); self.delay_spin.setRange(100, 2000)
        self.delay_spin.setValue(500); self.delay_spin.setSingleStep(100)
        self.delay_spin.setSuffix(" ms")
        ov.addRow("请求间隔:", self.delay_spin)
        _polish_group(grp_opt, ov)
        lv.addWidget(grp_opt)

        # 按钮
        btn_row = QHBoxLayout()
        self.btn_create = QPushButton("🚀 创建自动化")
        self.btn_create.setObjectName("successBtn")
        self.btn_create.clicked.connect(self._start)
        self.btn_cancel = QPushButton("⏹ 取消")
        self.btn_cancel.clicked.connect(self._cancel)
        self.btn_cancel.setEnabled(False)
        btn_row.addWidget(self.btn_create)
        btn_row.addWidget(self.btn_cancel)
        lv.addLayout(btn_row)
        lv.addStretch()

        # 右侧日志
        right = QWidget(); rv = QVBoxLayout(right)
        self.log = _make_log_panel(rv)
        self.progress = _make_progress(rv)

        layout.addWidget(left)
        layout.addWidget(right, stretch=1)

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择自动化 Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            self.excel_edit.setText(path)

    def _start(self):
        excel_path = self.excel_edit.text().strip()
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "提示", "请先选择自动化 Excel 文件")
            return

        try:
            config, auto_items = read_automation_excel(excel_path)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"Excel 读取错误:\n{e}")
            return

        # 手动输入覆盖 Excel 中的配置
        manual = {
            "serviceToken": self.token_edit.text().strip(),
            "xiaomiiot_ph": self.ph_edit.text().strip(),
            "userId":       self.userid_edit.text().strip(),
            "pdId":         self.pid_edit.text().strip(),
            "model":        self.model_edit.text().strip(),
        }
        for k, v in manual.items():
            if v:
                config[k] = v

        missing = [k for k in ("userId", "xiaomiiot_ph", "pdId") if not config.get(k)]
        if missing:
            QMessageBox.warning(self, "提示", f"缺少必填项:\n{', '.join(missing)}")
            return

        # 自动注入 groupId
        _inject_group_id(config)

        if not auto_items:
            QMessageBox.warning(self, "提示", "自动化列表为空")
            return

        # ─── 优先检查产品状态 ──────────────────────────────────
        self.log.clear()
        self.log.append("🔍 正在检查产品状态...")
        try:
            is_ok, status, status_name, msg = check_product_status(config)
            if is_ok:
                self.log.append(f"✅ {msg}")
            else:
                self.log.append(f"❌ {msg}")
                QMessageBox.critical(self, "产品状态检查失败", msg)
                return
        except Exception as e:
            self.log.append(f"⚠️ 产品状态检查异常: {e}（继续执行）")

        dry = self.chk_dryrun.isChecked()
        delay = self.delay_spin.value() / 1000.0

        self.log.append(f"📋 共 {len(auto_items)} 个自动化待创建" + (" (dry-run)" if dry else ""))
        self.btn_create.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress.setVisible(True); self.progress.setRange(0, len(auto_items))

        self._worker = CreateAutomationWorker(
            config, auto_items, dry_run=dry, delay=delay)
        self._worker.progress.connect(self.log.append)
        self._worker.update_progress.connect(lambda c, t: self.progress.setValue(c))
        self._worker.finished_ok.connect(self._done_ok)
        self._worker.finished_err.connect(self._done_err)
        self._worker.start()

    def _cancel(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
        self.log.append("⚠️ 取消请求已发送")

    def _done_ok(self, success, failed):
        self.btn_create.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress.setVisible(False)
        self.log.append(f"\n🎉 完成！成功: {success}, 失败: {failed}")
        QMessageBox.information(self, "创建完成", f"成功: {success}\n失败: {failed}")

    def _done_err(self, msg):
        self.btn_create.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress.setVisible(False)
        self.log.append(f"\n❌ {msg}")
        QMessageBox.critical(self, "创建失败", msg)
