#!/usr/bin/env python3
"""
MIoT 平台工具（整合版）
功能：
  服务层：创建服务 / 导出服务
  属性层：导出模板 / 创建属性 / 生成模板
"""

import sys

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox,
    QComboBox, QMenu,
)
from PyQt6.QtCore import Qt

# ── 登录模块
from miot_auth import (
    get_current_user, get_all_users, switch_user,
    remove_user, logout_current,
    update_user_group, get_curr_enterprise, get_enterprise_list,
    set_curr_enterprise,
)
# ─── 外部 GUI 模块 ─────────────────────────────────────────────

from miot_gui_styles import STYLESHEET
from miot_gui_auth_ui import LoginDialog, EnterpriseComboBox


# ─── Tab 模块 ─────────────────────────────────────────────────

from miot_gui_service_tabs import CreateServiceTab, ExportServiceTab
from miot_gui_property_tabs import ExportPropTab, CreatePropTab, TemplatePropTab
from miot_gui_automation_tabs import ExportAutomationTab, CreateAutomationTab


# ─── Main Window ──────────────────────────────────────────────

class MIoTMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MIoT 平台工具")
        self.setMinimumSize(1020, 820)
        self.resize(1060, 860)
        self._current_user = None
        self._ent_loading = False
        self._init_ui()
        self._update_user_ui()  # 初始化用户区域状态
        self._check_saved_login()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 12)

        # ── 标题行 + 用户区域
        header = QHBoxLayout()

        title = QLabel("🔧 MIoT 平台工具")
        title.setObjectName("titleLabel")
        subtitle = QLabel("小米 IoT 平台  —  服务层管理 & 属性层管理 & 自定义自动化（整合版）")
        subtitle.setObjectName("subtitleLabel")

        header_left = QVBoxLayout()
        header_left.addWidget(title)
        header_left.addWidget(subtitle)
        header_left.setSpacing(2)
        header.addLayout(header_left, 1)

        # 右上角用户区域
        self._build_user_area(header)

        layout.addLayout(header)

        # 外层 Tabs：服务层 / 属性层
        self.outer_tabs = QTabWidget()
        layout.addWidget(self.outer_tabs)

        # ── 服务层
        svc_widget = QWidget()
        svc_layout = QVBoxLayout(svc_widget)
        svc_layout.setContentsMargins(0, 0, 0, 0)
        svc_inner = QTabWidget()
        svc_inner.addTab(ExportServiceTab(), "📤 导出服务")
        svc_inner.addTab(CreateServiceTab(), "📋 创建服务")
        svc_layout.addWidget(svc_inner)
        self.outer_tabs.addTab(svc_widget, "🏗️ 服务层")

        # ── 属性层
        prop_widget = QWidget()
        prop_layout = QVBoxLayout(prop_widget)
        prop_layout.setContentsMargins(0, 0, 0, 0)
        prop_inner = QTabWidget()
        prop_inner.addTab(ExportPropTab(),   "📤 导出模板")
        prop_inner.addTab(CreatePropTab(),   "📥 创建属性")
        prop_inner.addTab(TemplatePropTab(), "📄 生成模板")
        prop_layout.addWidget(prop_inner)
        self.outer_tabs.addTab(prop_widget, "⚙️ 属性层")

        # ── 自动化
        auto_widget = QWidget()
        auto_layout = QVBoxLayout(auto_widget)
        auto_layout.setContentsMargins(0, 0, 0, 0)
        auto_inner = QTabWidget()
        auto_inner.addTab(ExportAutomationTab(), "📤 导出自动化")
        auto_inner.addTab(CreateAutomationTab(), "📥 创建自动化")
        auto_layout.addWidget(auto_inner)
        self.outer_tabs.addTab(auto_widget, "🤖 自动化")

        self.statusBar().showMessage("就绪")

    # ─── 用户区域 ─────────────────────────────────────────────

    def _build_user_area(self, parent_layout):
        """构建右上角用户区域：[企业下拉] [用户按钮]"""
        user_row = QHBoxLayout()
        user_row.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        user_row.setSpacing(8)

        # 企业刷新按钮（最左侧）
        self.ent_refresh_btn = QPushButton("🔄")
        self.ent_refresh_btn.setObjectName("entRefreshBtn")
        self.ent_refresh_btn.setFixedSize(32, 32)
        self.ent_refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.ent_refresh_btn.setToolTip("刷新企业列表")
        self.ent_refresh_btn.setVisible(False)  # 登录后才显示
        self.ent_refresh_btn.clicked.connect(self._on_ent_refresh)
        user_row.addWidget(self.ent_refresh_btn)

        # 企业下拉
        self.ent_combo = EnterpriseComboBox()
        self.ent_combo.setObjectName("entCombo")
        self.ent_combo.setCursor(Qt.CursorShape.PointingHandCursor)
        self.ent_combo.setVisible(False)  # 登录后才显示
        self.ent_combo.setToolTip("切换当前企业")
        self.ent_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.ent_combo.setMinimumContentsLength(10)
        self.ent_combo.setFixedWidth(280)
        self.ent_combo.currentIndexChanged.connect(self._on_ent_combo_changed)
        user_row.addWidget(self.ent_combo)

        # 用户按钮（右侧）
        self.user_btn = QPushButton("🔑 点击登录")
        self.user_btn.setObjectName("userBtn")
        self.user_btn.setProperty("loggedIn", "false")
        self.user_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.user_btn.clicked.connect(self._on_user_btn_clicked)
        user_row.addWidget(self.user_btn)

        parent_layout.addLayout(user_row)

    def _on_user_btn_clicked(self):
        """点击用户按钮 - 弹出菜单"""
        if not self._current_user:
            self._open_login()
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { font-size: 13px; padding: 4px; }
            QMenu::item { padding: 6px 20px; }
            QMenu::item:selected { background-color: #eaf2f8; }
        """)

        # 切换用户
        all_users = get_all_users()
        if len(all_users) > 1:
            switch_menu = menu.addMenu("🔄 切换用户")
            for u in all_users:
                uid = u["userId"]
                is_current = (uid == str(self._current_user.get("userId", "")))
                label = f"{'✅ ' if is_current else ''}{u['name']} ({uid})"
                act = switch_menu.addAction(label)
                act.setData(uid)
                if is_current:
                    act.setEnabled(False)
            switch_menu.triggered.connect(self._on_switch_user)
            menu.addSeparator()

        # 退出登录
        menu.addAction("🚪 退出登录", self._on_logout)
        # 删除用户
        menu.addAction("🗑️ 删除此账号", self._on_delete_user)

        # 在按钮下方弹出
        pos = self.user_btn.mapToGlobal(self.user_btn.rect().bottomLeft())
        menu.exec(pos)

    def _on_ent_refresh(self):
        """刷新企业列表"""
        if not self._current_user:
            return
        self.statusBar().showMessage("正在刷新企业列表...", 2000)
        self._refresh_ent_combo()
        self.statusBar().showMessage("✅ 企业列表已刷新", 3000)

    def _on_ent_combo_changed(self, index):
        """企业下拉切换"""
        if index < 0 or self._ent_loading:
            return
        ent = self.ent_combo.currentData()
        if not ent or not isinstance(ent, dict):
            return
        gid = ent.get("groupId", "")
        if not gid:
            return

        cur_gid = str(self._current_user.get("groupId", ""))
        if gid == cur_gid:
            return  # 没有变化

        # 调用 API 切换企业
        self.statusBar().showMessage(f"正在切换到 {ent.get('shortName', gid)}...", 3000)
        ok = set_curr_enterprise(
            self._current_user.get("userId", ""),
            self._current_user.get("xiaomiiot_ph", ""),
            self._current_user.get("serviceToken", ""),
            gid,
            ent.get("shortName", ""),
            ent.get("compName", ""),
        )
        if ok:
            update_user_group(self._current_user.get("userId", ""), gid)
            self._current_user["groupId"] = gid
            self._current_user["groupName"] = ent.get("compName", "")
            self.statusBar().showMessage(
                f"✅ 已切换到 {ent.get('shortName', gid)} ({ent.get('compName', '')})", 5000)
        else:
            QMessageBox.warning(self, "切换失败", "切换企业失败，请重试")
            # 回滚选择
            self._ent_loading = True
            self._select_current_enterprise()
            self._ent_loading = False

    def _on_switch_user(self, action):
        """切换用户"""
        uid = action.data()
        if uid:
            user = switch_user(uid)
            if user:
                self._current_user = user
                self._update_user_ui()
                self._fill_cookies()

    def _on_logout(self):
        """退出当前用户"""
        logout_current()
        self._current_user = None
        self._update_user_ui()
        self._clear_cookies()

    def _on_delete_user(self):
        """删除当前用户"""
        if self._current_user:
            uid = str(self._current_user.get("userId", ""))
            ret = QMessageBox.question(
                self, "确认删除",
                f"确定要删除用户 {uid} 的登录信息吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ret == QMessageBox.StandardButton.Yes:
                remove_user(uid)
                self._current_user = get_current_user()
                self._update_user_ui()
                if self._current_user:
                    self._fill_cookies()
                else:
                    self._clear_cookies()

    def _open_login(self):
        """打开登录对话框"""
        dlg = LoginDialog(self)
        dlg.login_success.connect(self._on_login_success)
        dlg.exec()

    def _on_login_success(self, user_info: dict):
        """登录成功回调"""
        self._current_user = user_info
        self._update_user_ui()
        self._fill_cookies()

    def _update_user_ui(self):
        """更新用户区域和企业下拉 UI"""
        if self._current_user:
            self.ent_combo.setVisible(True)
            self.ent_combo.setEnabled(True)
            self.ent_refresh_btn.setVisible(True)
            uid = str(self._current_user.get("userId", ""))
            name = self._current_user.get("name", uid)
            self.user_btn.setText(f"👤 {name}")
            self.user_btn.setProperty("loggedIn", "true")
            status_msg = f"已登录: {name} ({uid})"
            self._refresh_ent_combo()
            gid = self._current_user.get("groupId", "")
            if gid:
                status_msg += f" | 企业: {gid}"
            self.statusBar().showMessage(status_msg, 5000)
        else:
            self.ent_combo.setVisible(False)
            self.ent_combo.clear()
            self.ent_refresh_btn.setVisible(False)
            self.user_btn.setText("🔑 点击登录")
            self.user_btn.setProperty("loggedIn", "false")
            self.statusBar().showMessage("未登录", 3000)
        # 刷新 QSS（property 变化需要重新应用样式）
        self.user_btn.style().unpolish(self.user_btn)
        self.user_btn.style().polish(self.user_btn)

    def _refresh_ent_combo(self):
        """刷新企业下拉列表并选中当前企业"""
        if not self._current_user:
            return
        self._ent_loading = True
        self.ent_combo.blockSignals(True)
        self.ent_combo.clear()

        try:
            enterprises = get_enterprise_list(
                self._current_user.get("userId", ""),
                self._current_user.get("xiaomiiot_ph", ""),
                self._current_user.get("serviceToken", ""),
            )
            cur_gid = str(self._current_user.get("groupId", ""))
            current_index = -1
            for i, ent in enumerate(enterprises):
                gid = ent.get("groupId", "")
                # 只显示中文企业名称
                display_name = ent.get("compName", "") or ent.get("shortName", gid)
                self.ent_combo.addItem(display_name, ent)
                if gid == cur_gid:
                    current_index = i

            if current_index >= 0:
                self.ent_combo.setCurrentIndex(current_index)
            elif enterprises:
                self.ent_combo.setCurrentIndex(0)
            else:
                self.ent_combo.addItem("（无企业）")
                self.ent_combo.setEnabled(False)
        except Exception:
            self.ent_combo.addItem("（查询失败）")
            self.ent_combo.setEnabled(False)

        self.ent_combo.blockSignals(False)
        self._ent_loading = False

    def _select_current_enterprise(self):
        """根据 _current_user 中的 groupId 选中对应项（不重新请求API）"""
        cur_gid = str(self._current_user.get("groupId", ""))
        self.ent_combo.blockSignals(True)
        for i in range(self.ent_combo.count()):
            ent = self.ent_combo.itemData(i)
            if ent and isinstance(ent, dict) and str(ent.get("groupId", "")) == cur_gid:
                self.ent_combo.setCurrentIndex(i)
                break
        self.ent_combo.blockSignals(False)

    def _fill_cookies(self):
        """自动填充所有 Tab 中的 Cookie 字段"""
        if not self._current_user:
            return
        token = self._current_user.get("serviceToken", "")
        ph = self._current_user.get("xiaomiiot_ph", "")
        uid = str(self._current_user.get("userId", ""))

        # 遍历所有 Tab 中的 Cookie 字段
        for tab_widget in self._find_all_tabs():
            self._fill_tab_cookies(tab_widget, token, ph, uid)

    def _find_all_tabs(self) -> list:
        """找到所有内层 Tab 页"""
        tabs = []
        for i in range(self.outer_tabs.count()):
            outer_page = self.outer_tabs.widget(i)
            inner_tabs = outer_page.findChild(QTabWidget)
            if inner_tabs:
                for j in range(inner_tabs.count()):
                    tabs.append(inner_tabs.widget(j))
        return tabs

    def _fill_tab_cookies(self, widget, token, ph, uid):
        """填充单个 Tab 中的 Cookie 字段"""
        # 查找所有 QLineEdit，按 placeholder 或 objectName 识别
        for edit in widget.findChildren(QLineEdit):
            name = edit.placeholderText().lower()
            obj = edit.objectName().lower() if edit.objectName() else ""
            if "servicetoken" in name or "token" in obj:
                if not edit.text().strip():
                    edit.setText(token)
            elif "xiaomiiot_ph" in name or "ph" in obj:
                if not edit.text().strip():
                    edit.setText(ph)
            elif "userid" in name or "uid" in obj:
                if not edit.text().strip():
                    edit.setText(uid)
        # 也通过变量名模式匹配（更可靠）
        for attr_name in dir(widget):
            if attr_name.startswith("_"):
                continue
            attr = getattr(widget, attr_name, None)
            if not isinstance(attr, QLineEdit):
                continue
            al = attr_name.lower()
            if "token" in al and not attr.text().strip():
                attr.setText(token)
            elif "ph" in al and "edit" in al and not attr.text().strip():
                attr.setText(ph)
            elif ("userid" in al or "uid" in al) and "edit" in al and not attr.text().strip():
                attr.setText(uid)

    def _clear_cookies(self):
        """清空所有 Tab 中的 Cookie 字段"""
        for tab_widget in self._find_all_tabs():
            for edit in tab_widget.findChildren(QLineEdit):
                al = edit.objectName().lower() if edit.objectName() else ""
                attr_name = ""
                # 通过变量名模式查找
                for an in dir(tab_widget):
                    if getattr(tab_widget, an, None) is edit:
                        attr_name = an.lower()
                        break
                if any(k in attr_name for k in ("token", "ph_edit", "userid_edit", "uid_edit")):
                    edit.clear()

    def _check_saved_login(self):
        """检查是否有已保存的登录信息，并验证 token 是否有效"""
        user = get_current_user()
        if not user:
            return

        # 验证 token 是否仍然有效（调用企业列表 API）
        try:
            enterprises = get_enterprise_list(
                user.get("userId", ""),
                user.get("xiaomiiot_ph", ""),
                user.get("serviceToken", ""),
            )
        except Exception:
            enterprises = None

        if enterprises is None:
            # API 请求异常（网络问题等），不做判定，按本地状态登录
            self._current_user = user
        elif not enterprises:
            # API 返回空列表 → token 大概率已过期，清除登录状态
            logout_current()
            self._current_user = None
            self._update_user_ui()
            QMessageBox.information(
                self, "登录已过期",
                "检测到您的登录信息已过期，请重新登录。"
            )
            return
        else:
            self._current_user = user

        # 如果本地没有 groupId，尝试从 API 获取
        if not user.get("groupId"):
            try:
                ent = get_curr_enterprise(
                    user.get("userId", ""),
                    user.get("xiaomiiot_ph", ""),
                    user.get("serviceToken", ""),
                )
                if ent.get("groupId"):
                    user["groupId"] = ent["groupId"]
                    user["groupName"] = ent.get("compName", "")
                    update_user_group(user.get("userId", ""), ent["groupId"])
            except Exception:
                pass
        self._update_user_ui()
        self._fill_cookies()


# ─── Entry ────────────────────────────────────────────────────

def main():
    # ── PyInstaller frozen 环境下的 GPU/RHI 安全回退 ──
    # onefile 模式或部分 macOS 机型上，Qt 的 Metal 渲染管线在临时目录中
    # 编译 shader 会导致堆内存损坏 (free_list_checksum_botch → SIGABRT)。
    # 强制使用 Software OpenGL 或禁用 RHI shader 缓存可避免此问题。
    import os
    if getattr(sys, 'frozen', False):
        os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")
        os.environ.setdefault("QT_QUICK_BACKEND", "software")
        os.environ.setdefault("QT_OPENGL", "software")

    # WebEngine 必须在 QApplication 创建前导入
    from PyQt6.QtWebEngineWidgets import QWebEngineView  # noqa: F401

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)
    window = MIoTMainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
