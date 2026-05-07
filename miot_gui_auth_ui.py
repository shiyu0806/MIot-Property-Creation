#!/usr/bin/env python3
"""Authentication UI helpers for the MIoT GUI."""

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import QComboBox, QDialog, QVBoxLayout

from miot_auth import MiLoginBrowser, save_user


class LoginDialog(QDialog):
    """小米账号登录对话框 - 内嵌浏览器"""
    login_success = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("登录小米账号")
        self.setMinimumSize(480, 640)
        self.resize(520, 700)
        self._browser = MiLoginBrowser(self)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 浏览器
        view = self._browser.create_view()
        layout.addWidget(view)

        # 连接信号
        self._browser.login_success.connect(self._on_login_success)

        # 启动登录
        self._browser.start_login()

    def _on_login_success(self, user_info: dict):
        """登录成功"""
        save_user(
            user_id=user_info["userId"],
            service_token=user_info["serviceToken"],
            xiaomiiot_ph=user_info["xiaomiiot_ph"],
            name=user_info.get("userId", ""),
            group_id=user_info.get("groupId", ""),
        )
        self.login_success.emit(user_info)
        self.accept()

    def closeEvent(self, event):
        self._browser.cleanup()
        super().closeEvent(event)


# ─── 自定义企业下拉框（弹出列表自动展开宽度） ────────────────

class EnterpriseComboBox(QComboBox):
    """下拉弹出列表宽度根据内容自动展开，不受控件本身固定宽度限制"""

    def showPopup(self):
        super().showPopup()
        view = self.view()
        if view is None:
            return
        fm = self.fontMetrics()
        max_text_width = 0
        for i in range(self.count()):
            text = self.itemText(i)
            max_text_width = max(max_text_width, fm.horizontalAdvance(text))
        needed_width = max_text_width + 60
        combo_width = self.width()
        popup_width = max(needed_width, combo_width)
        popup = view.parentWidget()
        if popup:
            popup.setFixedWidth(popup_width)


