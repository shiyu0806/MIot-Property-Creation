#!/usr/bin/env python3
"""Common widgets and helpers for the MIoT GUI."""

from PyQt6.QtWidgets import (
    QLabel, QTextEdit, QProgressBar, QGroupBox, QFormLayout,
    QLineEdit, QCheckBox,
)

from miot_auth import get_current_user


def _make_log_panel(parent_layout) -> QTextEdit:
    lbl = QLabel("运行日志")
    lbl.setStyleSheet("font-weight: bold; font-size: 13px;")
    log = QTextEdit()
    log.setReadOnly(True)
    parent_layout.addWidget(lbl)
    parent_layout.addWidget(log)
    return log

def _make_progress(parent_layout) -> QProgressBar:
    pb = QProgressBar()
    pb.setVisible(False)
    parent_layout.addWidget(pb)
    return pb

def _inject_group_id(config: dict):
    """从当前登录用户自动注入 groupId 到 config（如果 config 中没有的话）"""
    if config.get("groupId"):
        return
    cur = get_current_user()
    if cur and cur.get("groupId"):
        config["groupId"] = cur["groupId"]

def _cookie_group(parent_layout, prefix: str, show_userid=True):
    """
    返回 (grp, token_edit, ph_edit, userid_edit_or_None)
    prefix 用于内部区分，不展示给用户
    如果已登录，自动填充 Cookie 字段
    """
    grp = QGroupBox("Cookie 信息")
    form = QFormLayout()
    token = QLineEdit(); token.setEchoMode(QLineEdit.EchoMode.Password)
    token.setPlaceholderText("浏览器 Cookie 中的 serviceToken")
    ph = QLineEdit(); ph.setEchoMode(QLineEdit.EchoMode.Password)
    ph.setPlaceholderText("浏览器 Cookie 中的 xiaomiiot_ph")
    userid_edit = None
    form.addRow("serviceToken:", token)
    form.addRow("xiaomiiot_ph:", ph)
    if show_userid:
        userid_edit = QLineEdit()
        userid_edit.setPlaceholderText("如 1097752639")
        form.addRow("userId:", userid_edit)

    # 自动填充当前用户的 Cookie
    cur = get_current_user()
    if cur:
        token.setText(cur.get("serviceToken", ""))
        ph.setText(cur.get("xiaomiiot_ph", ""))
        if userid_edit:
            userid_edit.setText(cur.get("userId", ""))

    chk = QCheckBox("显示 Cookie")
    def toggle(checked):
        mode = QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        token.setEchoMode(mode); ph.setEchoMode(mode)
    chk.toggled.connect(toggle)
    form.addRow("", chk)
    grp.setLayout(form)
    parent_layout.addWidget(grp)
    return grp, token, ph, userid_edit


