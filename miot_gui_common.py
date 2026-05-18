#!/usr/bin/env python3
"""Common widgets and helpers for the MIoT GUI."""

from html import escape
from datetime import datetime
import re

from PyQt6.QtWidgets import (
    QLabel, QTextEdit, QProgressBar, QGroupBox, QFormLayout,
    QLineEdit, QWidget, QVBoxLayout, QScrollArea,
    QHBoxLayout, QPushButton,
)
from PyQt6.QtCore import Qt

from miot_auth import get_current_user


class ColorLogTextEdit(QTextEdit):
    """Dark log text area with formatted, color-coded lines."""

    LEVELS = (
        ("ERROR", ("❌", "失败", "错误", "ERROR", "Error", "Traceback"), "#ff453a"),
        ("WARN", ("⚠️", "警告", "WARN", "Warning", "取消"), "#ff9f0a"),
        ("SUCCESS", ("✅", "🎉", "成功", "完成", "SUCCESS"), "#34c759"),
        ("INFO", ("🚀", "🔍", "📋", "🧪", "INFO", "开始", "查询"), "#0a84ff"),
    )

    MAX_LINES = 1000
    TRIM_LINES = 200

    def append(self, text):  # noqa: D401 - keep QTextEdit-compatible API
        raw = "" if text is None else str(text)
        for line in raw.splitlines() or [""]:
            self._append_line(line)
        self._trim_if_needed()
        self._scroll_to_bottom()

    def _append_line(self, line: str):
        level = "LOG"
        color = "#8e8e93"
        prefix = re.match(r"^\[(INFO|SUCCESS|WARN|ERROR|LOG)\]\s*(.*)$", line)
        if prefix:
            line = prefix.group(2)
            level = prefix.group(1)
        for candidate_level, markers, candidate_color in self.LEVELS:
            if candidate_level == level or any(marker in line for marker in markers):
                level = candidate_level
                color = candidate_color
                break
        timestamp = datetime.now().strftime("%H:%M:%S")
        super().append(
            '<span style="color:#48484a;">'
            f'{timestamp}</span> '
            f'<span style="color:{color}; font-weight:600;">[{level}]</span> '
            f'<span style="color:{color};">{escape(line)}</span>'
        )

    def _trim_if_needed(self):
        if self.document().lineCount() <= self.MAX_LINES:
            return
        cursor = self.textCursor()
        cursor.movePosition(cursor.MoveOperation.Start)
        for _ in range(self.TRIM_LINES):
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor)
        cursor.removeSelectedText()
        cursor.deleteChar()

    def _scroll_to_bottom(self):
        bar = self.verticalScrollBar()
        bar.setValue(bar.maximum())


class LogPanel(QWidget):
    """Design-spec log panel with toolbar, copy/clear, and append compatibility."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("logPanel")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        toolbar = QWidget()
        toolbar.setObjectName("logToolbar")
        toolbar.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        toolbar.setFixedHeight(36)
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(12, 0, 12, 0)
        toolbar_layout.setSpacing(8)

        self._status_dot = QLabel()
        self._status_dot.setObjectName("logStatusDot")
        self._status_dot.setFixedSize(8, 8)
        self._set_status_color("#34c759")

        title = QLabel("运行日志")
        title.setObjectName("logTitle")
        toolbar_layout.addWidget(self._status_dot)
        toolbar_layout.addWidget(title)
        toolbar_layout.addStretch()

        self.copy_btn = QPushButton("复制")
        self.copy_btn.setObjectName("logToolButton")
        self.copy_btn.setFixedSize(56, 26)
        self.copy_btn.clicked.connect(self.copy_log)

        self.clear_btn = QPushButton("清空")
        self.clear_btn.setObjectName("logToolButton")
        self.clear_btn.setFixedSize(56, 26)
        self.clear_btn.clicked.connect(self.clear)

        toolbar_layout.addWidget(self.copy_btn)
        toolbar_layout.addWidget(self.clear_btn)
        layout.addWidget(toolbar)

        self._text = ColorLogTextEdit()
        self._text.setObjectName("logText")
        self._text.setReadOnly(True)
        self._text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        layout.addWidget(self._text)

    def append(self, text):
        raw = "" if text is None else str(text)
        self._text.append(raw)
        self._set_status_color(self._status_color_for(raw))

    def log(self, message: str, level: str = "INFO"):
        self.append(f"[{level}] {message}")

    def clear(self):
        self._text.clear()
        self._set_status_color("#34c759")

    def copy_log(self):
        cursor = self._text.textCursor()
        selected = cursor.hasSelection()
        if not selected:
            self._text.selectAll()
        self._text.copy()
        if not selected:
            cursor.clearSelection()
            self._text.setTextCursor(cursor)
        self.append("日志已复制到剪贴板")

    def toPlainText(self):
        return self._text.toPlainText()

    def _status_color_for(self, text: str):
        for _, markers, color in ColorLogTextEdit.LEVELS:
            if any(marker in text for marker in markers):
                return color
        return "#8e8e93"

    def _set_status_color(self, color: str):
        self._status_dot.setStyleSheet(
            f"QLabel#logStatusDot {{ background-color: {color}; border-radius: 4px; }}"
        )


def _make_log_panel(parent_layout) -> LogPanel:
    log = LogPanel()
    parent_layout.addWidget(log)
    return log

def _make_progress(parent_layout) -> QProgressBar:
    pb = QProgressBar()
    pb.setVisible(False)
    parent_layout.addWidget(pb)
    return pb

def _make_left_panel(width: int = 420):
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFixedWidth(width)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    panel = QWidget()
    layout = QVBoxLayout(panel)
    layout.setContentsMargins(0, 0, 10, 0)
    layout.setSpacing(16)
    scroll.setWidget(panel)
    return scroll, panel, layout

def _polish_group(group: QGroupBox, layout):
    """Apply roomy card spacing so modern controls do not overlap."""
    layout.setContentsMargins(20, 28, 20, 20)
    layout.setSpacing(12)
    if isinstance(layout, QFormLayout):
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(10)
        layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        layout.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        layout.setFormAlignment(Qt.AlignmentFlag.AlignTop)
    group.setLayout(layout)

def _inject_group_id(config: dict):
    """从当前登录用户自动注入 groupId 到 config（如果 config 中没有的话）"""
    if config.get("groupId"):
        return
    cur = get_current_user()
    if cur and cur.get("groupId"):
        config["groupId"] = cur["groupId"]

def _make_cookie_fields(parent: QWidget, prefix: str, show_userid=True):
    """
    Create hidden Cookie fields for code paths that still need credential values.

    The UI no longer shows a Cookie card, but login auto-fill and Excel override
    flows still read these QLineEdit instances.
    """
    token = QLineEdit(parent)
    token.setObjectName(f"{prefix}_token_edit")
    token.setEchoMode(QLineEdit.EchoMode.Password)
    token.setVisible(False)

    ph = QLineEdit(parent)
    ph.setObjectName(f"{prefix}_ph_edit")
    ph.setEchoMode(QLineEdit.EchoMode.Password)
    ph.setVisible(False)

    userid_edit = None
    if show_userid:
        userid_edit = QLineEdit(parent)
        userid_edit.setObjectName(f"{prefix}_userid_edit")
        userid_edit.setVisible(False)

    cur = get_current_user()
    if cur:
        token.setText(cur.get("serviceToken", ""))
        ph.setText(cur.get("xiaomiiot_ph", ""))
        if userid_edit:
            userid_edit.setText(str(cur.get("userId", "")))

    return token, ph, userid_edit
