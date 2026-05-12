#!/usr/bin/env python3
"""Stylesheet for the MIoT GUI."""

STYLESHEET = """
QMainWindow { background-color: #f5f6fa; }
QTabWidget::pane {
    border: 1px solid #dcdde1;
    border-radius: 6px;
    background: white;
    margin-top: -1px;
}
QTabBar::tab {
    background: #dcdde1;
    padding: 10px 22px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-size: 13px;
    font-weight: bold;
    color: #000;
}
QTabBar::tab:selected {
    background: white;
    color: #2c3e50;
    border-bottom: 2px solid #e67e22;
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
QPushButton:hover  { background-color: #2980b9; }
QPushButton:pressed { background-color: #2471a3; }
QPushButton:disabled { background-color: #bdc3c7; }
QPushButton#dangerBtn  { background-color: #e74c3c; }
QPushButton#dangerBtn:hover { background-color: #c0392b; }
QPushButton#successBtn { background-color: #27ae60; }
QPushButton#successBtn:hover { background-color: #229954; }
QPushButton#warnBtn  { background-color: #e67e22; }
QPushButton#warnBtn:hover { background-color: #ca6f1e; }
QLineEdit, QSpinBox, QComboBox {
    padding: 6px 10px;
    border: 1px solid #dcdde1;
    border-radius: 4px;
    font-size: 13px;
    background: white;
}
QLineEdit:focus, QSpinBox:focus { border-color: #3498db; }
QTextEdit {
    border: 1px solid #dcdde1;
    border-radius: 4px;
    font-family: "Menlo", "Monaco", monospace;
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
QProgressBar::chunk { background-color: #3498db; border-radius: 3px; }
QLabel#titleLabel   { font-size: 18px; font-weight: bold; color: #2c3e50; }
QLabel#subtitleLabel { font-size: 12px; color: #000; }
/* 企业下拉 */
QComboBox#entCombo {
    background-color: transparent;
    color: #2980b9;
    border: 2px solid #dcdde1;
    border-radius: 18px;
    padding: 4px 30px 4px 10px;
    font-size: 13px;
    font-weight: bold;
    min-height: 28px;
    max-width: 280px;
}
QComboBox#entCombo:hover { border-color: #3498db; }
QComboBox#entCombo::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: right center;
    border: none;
    width: 24px;
}
QComboBox#entCombo::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #2980b9;
    margin-right: 8px;
}
QComboBox#entCombo QAbstractItemView {
    font-size: 13px;
    border: 1px solid #dcdde1;
    border-radius: 6px;
    selection-background-color: #eaf2f8;
    selection-color: #2c3e50;
    padding: 4px;
}
/* 用户区域 */
QPushButton#userBtn {
    background-color: transparent;
    color: #2c3e50;
    border: 2px solid #dcdde1;
    border-radius: 18px;
    padding: 4px 14px 4px 10px;
    font-size: 13px;
    font-weight: bold;
    min-height: 28px;
}
QPushButton#userBtn:hover { border-color: #3498db; color: #3498db; }
QPushButton#userBtn[loggedIn="true"] {
    border-color: #27ae60; color: #27ae60;
}
QPushButton#userBtn[loggedIn="true"]:hover {
    border-color: #e74c3c; color: #e74c3c;
}
QPushButton#loginBtn {
    background-color: #ff6700;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 6px 16px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton#loginBtn:hover { background-color: #e55d00; }
QTabWidget#innerTabs::pane {
    border: 1px solid #e0e0e0;
    border-radius: 4px;
    background: #fafafa;
}
QTabBar#innerTabs::tab {
    background: #e8e8e8;
    padding: 7px 18px;
    font-size: 12px;
    font-weight: normal;
}
QTabBar#innerTabs::tab:selected {
    background: #fafafa;
    color: #e67e22;
    border-bottom: 2px solid #e67e22;
}
"""


