#!/usr/bin/env python3
"""Service management tabs for the MIoT GUI."""

import os

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit, QPushButton,
    QCheckBox, QFileDialog, QSpinBox, QGroupBox, QMessageBox,
)

from miot_service_core import (
    check_product_status,
    read_service_config_excel,
    read_service_list_excel,
)
from miot_gui_common import (
    _make_log_panel,
    _make_progress,
    _make_left_panel,
    _make_cookie_fields,
    _inject_group_id,
    _polish_group,
)
from miot_gui_workers import SyncServiceWorker, ExportServiceWorker


class CreateServiceTab(QWidget):
    def __init__(self):
        super().__init__()
        self._worker = None
        self._build()

    def _build(self):
        layout = QHBoxLayout(self)

        # 左侧表单
        left, _, lv = _make_left_panel()

        # Excel 文件
        grp_file = QGroupBox("Excel 文件（服务模板）")
        fv = QHBoxLayout()
        self.file_edit = QLineEdit(); self.file_edit.setPlaceholderText("选择服务 Excel 模板")
        btn_browse = QPushButton("浏览...")
        btn_browse.clicked.connect(self._browse_file)
        fv.addWidget(self.file_edit); fv.addWidget(btn_browse)
        _polish_group(grp_file, fv)
        lv.addWidget(grp_file)

        # 产品信息覆盖
        grp_prod = QGroupBox("产品信息")
        form_prod = QFormLayout()
        self.pid_edit = QLineEdit(); self.pid_edit.setPlaceholderText("可选")
        self.model_edit = QLineEdit(); self.model_edit.setPlaceholderText("")
        form_prod.addRow("产品ID (pdId):", self.pid_edit)
        form_prod.addRow("产品型号 (model):", self.model_edit)
        _polish_group(grp_prod, form_prod)
        lv.addWidget(grp_prod)

        self.token_edit, self.ph_edit, self.userid_edit = _make_cookie_fields(self, "svc_crt")

        # 选项
        grp_opt = QGroupBox("选项")
        form_opt = QFormLayout()
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(100, 2000); self.delay_spin.setValue(500)
        self.delay_spin.setSingleStep(100); self.delay_spin.setSuffix(" ms")
        form_opt.addRow("请求间隔:", self.delay_spin)
        _polish_group(grp_opt, form_opt)
        lv.addWidget(grp_opt)

        # 按钮
        btn_row = QHBoxLayout()
        self.btn_dry = QPushButton("🧪 干跑检查")
        self.btn_dry.clicked.connect(self._start_dry)
        self.btn_run = QPushButton("🚀 开始创建")
        self.btn_run.setObjectName("successBtn")
        self.btn_run.clicked.connect(self._start_create)
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.clicked.connect(self._cancel)
        self.btn_cancel.setEnabled(False)
        btn_row.addWidget(self.btn_dry); btn_row.addWidget(self.btn_run)
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
        path, _ = QFileDialog.getOpenFileName(self, "选择服务 Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            self.file_edit.setText(path)

    def _build_config(self):
        path = self.file_edit.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "提示", "请选择有效的 Excel 文件")
            return None, None
        try:
            config = read_service_config_excel(path)
            rows   = read_service_list_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))
            return None, None

        # 覆盖
        for key, widget in [
            ("pdId", self.pid_edit), ("model", self.model_edit),
            ("serviceToken", self.token_edit),
            ("xiaomiiot_ph", self.ph_edit), ("userId", self.userid_edit),
        ]:
            val = widget.text().strip()
            if val:
                config[key] = val

        missing = [k for k in ("userId", "xiaomiiot_ph", "serviceToken", "pdId", "model")
                   if not config.get(k)]
        if missing:
            QMessageBox.warning(self, "配置缺失", f"缺少必填项:\n{', '.join(missing)}")
            return None, None
        if not rows:
            QMessageBox.warning(self, "提示", "服务列表为空")
            return None, None

        # 自动注入 groupId（从当前登录用户）
        _inject_group_id(config)

        return config, rows

    def _start_dry(self):  self._run(dry_run=True)
    def _start_create(self): self._run(dry_run=False)

    def _run(self, dry_run):
        config, rows = self._build_config()
        if not config:
            return

        # ─── 优先检查产品状态（仅正式创建时）─────────────────────
        if not dry_run:
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

            reply = QMessageBox.question(
                self, "确认创建", f"即将同步 {len(rows)} 个服务，是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return
        else:
            self.log.clear()

        self.log.append(f"{'🧪 干跑模式' if dry_run else '🚀 正式创建'} - {len(rows)} 个服务\n")
        self._set_btns(running=True)
        self.progress.setVisible(True); self.progress.setRange(0, 0)

        self._worker = SyncServiceWorker(config, rows, dry_run,
                                          self.delay_spin.value() / 1000.0)
        self._worker.progress.connect(self.log.append)
        self._worker.finished_ok.connect(self._done_ok)
        self._worker.finished_err.connect(self._done_err)
        self._worker.start()

    def _cancel(self):
        if self._worker:
            self._worker.cancel()
        self.log.append("⚠️ 取消请求已发送")

    def _done_ok(self, res):
        self._set_btns(running=False)
        summary = f"创建:{res['created']} 跳过:{res['skipped']} 修正:{res['fixed']} 错误:{res['errors']}"
        self.log.append(f"\n📊 {summary}")
        if res["errors"]:
            QMessageBox.warning(self, "完成（有错误）", summary)
        else:
            QMessageBox.information(self, "完成", f"🎉 {summary}")

    def _done_err(self, msg):
        self._set_btns(running=False)
        self.log.append(f"\n❌ {msg}")
        QMessageBox.critical(self, "失败", msg)

    def _set_btns(self, running):
        self.btn_dry.setEnabled(not running)
        self.btn_run.setEnabled(not running)
        self.btn_cancel.setEnabled(running)
        self.progress.setVisible(running)


# ─── Tab: 导出服务 ────────────────────────────────────────────

class ExportServiceTab(QWidget):
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
        self.pid_edit = QLineEdit(); self.pid_edit.setPlaceholderText("如 33257（可选）")
        self.model_edit = QLineEdit(); self.model_edit.setPlaceholderText("如 uwize.switch.aiswi")
        form_prod.addRow("产品ID (pdId):", self.pid_edit)
        form_prod.addRow("产品型号 (model):", self.model_edit)
        _polish_group(grp_prod, form_prod)
        lv.addWidget(grp_prod)

        self.token_edit, self.ph_edit, self.userid_edit = _make_cookie_fields(self, "svc_exp")

        # 选项
        grp_opt = QGroupBox("导出选项")
        ov = QFormLayout()
        self.out_edit = QLineEdit(); self.out_edit.setPlaceholderText("点击浏览选择导出文件夹")
        self.out_edit.setReadOnly(True)
        btn_br = QPushButton("浏览...")
        btn_br.clicked.connect(self._browse_out)
        row = QHBoxLayout(); row.addWidget(self.out_edit); row.addWidget(btn_br)
        ov.addRow("导出文件夹:", row)
        _polish_group(grp_opt, ov)
        lv.addWidget(grp_opt)

        # 按钮
        btn_row = QHBoxLayout()
        self.btn_export = QPushButton("📤 导出服务")
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
        missing = [k for k in ("userId", "xiaomiiot_ph", "model") if not config.get(k)]
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
        safe_model = config["model"].replace(".", "_").replace("-", "_")
        output_path = os.path.join(out_dir, f"{safe_model}_services_export.xlsx")

        self.log.clear()
        self.btn_export.setEnabled(False)
        self.progress.setVisible(True); self.progress.setRange(0, 0)

        self._worker = ExportServiceWorker(
            config, output_path)
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
