#!/usr/bin/env python3
"""Stylesheet for the MIoT GUI."""

STYLESHEET = """
QMainWindow, QWidget#appRoot {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #e8f5e9, stop:0.52 #f5f5f7, stop:1 #e3f2fd);
    color: #1d1d1f;
}

QWidget#appHeader {
    background-color: rgba(255, 255, 255, 0.46);
    border-bottom: 1px solid rgba(0, 0, 0, 0.08);
}

QLabel#brandIcon {
    color: white;
    font-size: 20px;
    border-radius: 10px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #34c759, stop:1 #30b350);
}

QLabel#titleLabel {
    font-size: 18px;
    font-weight: 700;
    color: #1d1d1f;
}

QLabel#subtitleLabel {
    font-size: 12px;
    color: #6e6e73;
}

QTabWidget::pane {
    border: none;
    background: transparent;
}

QTabWidget#outerTabs::pane {
    border-top: 1px solid rgba(0, 0, 0, 0.08);
}

QTabBar::tab {
    background: transparent;
    border: none;
    color: #6e6e73;
    font-size: 14px;
    font-weight: 600;
    min-height: 44px;
    padding: 0 20px;
    margin-right: 4px;
}

QTabBar::tab:hover {
    color: #1d1d1f;
}

QTabBar::tab:selected {
    color: #30b350;
    border-bottom: 2px solid #34c759;
}

QTabWidget#innerTabs::pane {
    border: none;
    background: transparent;
    margin-top: 10px;
}

QTabWidget#innerTabs QTabBar::tab {
    border: 1px solid transparent;
    border-radius: 10px;
    color: #6e6e73;
    font-size: 13px;
    min-height: 36px;
    padding: 0 16px;
    margin-right: 8px;
}

QTabWidget#innerTabs QTabBar::tab:hover {
    background: rgba(0, 0, 0, 0.04);
}

QTabWidget#innerTabs QTabBar::tab:selected {
    background: #e8f5e9;
    border: 1px solid #34c759;
    color: #30b350;
}

QGroupBox {
    background-color: rgba(255, 255, 255, 0.72);
    border: 1px solid rgba(255, 255, 255, 0.56);
    border-radius: 14px;
    font-size: 14px;
    font-weight: 700;
    margin-top: 18px;
    padding-top: 12px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 20px;
    padding: 0 10px;
    color: #1d1d1f;
    top: 5px;
}

QLabel {
    color: #1d1d1f;
    font-size: 13px;
}

QLineEdit, QSpinBox, QComboBox {
    background-color: rgba(255, 255, 255, 0.66);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 8px;
    color: #1d1d1f;
    font-size: 13px;
    min-height: 40px;
    padding: 0 14px;
}

QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
    background-color: white;
    border: 1px solid #34c759;
}

QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled {
    background-color: rgba(255, 255, 255, 0.35);
    color: #8e8e93;
}

QComboBox::drop-down {
    border: none;
    width: 24px;
}

QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #6e6e73;
    margin-right: 8px;
}

QComboBox QAbstractItemView {
    background: white;
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 10px;
    color: #1d1d1f;
    outline: none;
    padding: 4px;
    selection-background-color: #e8f5e9;
    selection-color: #1d1d1f;
}

QPushButton {
    background: rgba(0, 0, 0, 0.06);
    border: none;
    border-radius: 10px;
    color: #1d1d1f;
    font-size: 13px;
    font-weight: 600;
    min-height: 40px;
    padding: 0 20px;
}

QPushButton:hover {
    background: rgba(0, 0, 0, 0.10);
}

QPushButton:pressed {
    background: rgba(0, 0, 0, 0.16);
}

QPushButton:disabled {
    background: rgba(0, 0, 0, 0.05);
    color: #a1a1a6;
}

QPushButton#successBtn {
    color: white;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #34c759, stop:1 #30b350);
    min-height: 48px;
}

QPushButton#successBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #30d158, stop:1 #28a745);
}

QPushButton#warnBtn {
    color: white;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #ff9f0a, stop:1 #ff7a00);
    min-height: 48px;
}

QPushButton#warnBtn:hover {
    background: #ff8a00;
}

QPushButton#dangerBtn {
    color: white;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #ff453a, stop:1 #d92d20);
    min-height: 48px;
}

QPushButton#dangerBtn:hover {
    background: #d92d20;
}

QPushButton#entRefreshBtn {
    color: white;
    border-radius: 16px;
    font-size: 15px;
    min-width: 32px;
    max-width: 32px;
    min-height: 32px;
    max-height: 32px;
    padding: 0;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #34c759, stop:1 #30b350);
}

QComboBox#entCombo {
    background-color: rgba(255, 255, 255, 0.66);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 18px;
    color: #1d1d1f;
    font-size: 13px;
    font-weight: 600;
    max-width: 280px;
    min-height: 36px;
    max-height: 36px;
    padding: 0 30px 0 16px;
}

QComboBox#entCombo:hover {
    background-color: rgba(255, 255, 255, 0.92);
}

QPushButton#userBtn {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #34c759, stop:1 #30b350);
    border: none;
    border-radius: 18px;
    color: white;
    font-size: 13px;
    font-weight: 700;
    min-height: 36px;
    max-height: 36px;
    padding: 0 16px;
}

QPushButton#userBtn:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #30d158, stop:1 #28a745);
}

QPushButton#userBtn[loggedIn="false"] {
    background: rgba(255, 255, 255, 0.66);
    border: 1px solid rgba(0, 0, 0, 0.08);
    color: #1d1d1f;
}

QPushButton#loginBtn {
    color: white;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #34c759, stop:1 #30b350);
}

QCheckBox {
    color: #6e6e73;
    font-size: 13px;
    spacing: 8px;
}

QCheckBox::indicator {
    height: 18px;
    width: 18px;
}

QTextEdit {
    background-color: #1c1c1e;
    border: none;
    border-radius: 14px;
    color: #8e8e93;
    font-family: "SF Mono", "Menlo", "Monaco", monospace;
    font-size: 12px;
    padding: 16px;
}

QWidget#logPanel {
    background-color: #1c1c1e;
    border: none;
    border-radius: 14px;
}

QWidget#logToolbar {
    background-color: rgba(255, 255, 255, 0.06);
    border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    border-top-left-radius: 14px;
    border-top-right-radius: 14px;
}

QLabel#logTitle {
    color: #8e8e93;
    font-size: 12px;
    font-weight: 600;
}

QPushButton#logToolButton {
    background: rgba(255, 255, 255, 0.08);
    border: none;
    border-radius: 6px;
    color: #8e8e93;
    font-size: 11px;
    font-weight: 600;
    min-height: 26px;
    padding: 0;
}

QPushButton#logToolButton:hover {
    background: rgba(255, 255, 255, 0.14);
}

QTextEdit#logText {
    border-radius: 0 0 14px 14px;
}

QProgressBar {
    background: rgba(255, 255, 255, 0.66);
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 10px;
    color: #1d1d1f;
    height: 22px;
    text-align: center;
}

QProgressBar::chunk {
    border-radius: 9px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #34c759, stop:1 #0a84ff);
}

QScrollBar:vertical {
    background: transparent;
    border: none;
    margin: 2px;
    width: 8px;
}

QScrollBar::handle:vertical {
    background: rgba(0, 0, 0, 0.16);
    border-radius: 4px;
    min-height: 24px;
}

QScrollBar::handle:vertical:hover {
    background: rgba(0, 0, 0, 0.26);
}

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical {
    background: transparent;
    border: none;
    height: 0;
}

QStatusBar {
    background-color: rgba(255, 255, 255, 0.46);
    border-top: 1px solid rgba(0, 0, 0, 0.08);
    color: #6e6e73;
    font-size: 12px;
}

QMenu {
    background: white;
    border: 1px solid rgba(0, 0, 0, 0.08);
    border-radius: 10px;
    color: #1d1d1f;
    font-size: 13px;
    padding: 6px;
}

QMenu::item {
    border-radius: 6px;
    padding: 7px 22px;
}

QMenu::item:selected {
    background: #e8f5e9;
}
"""
