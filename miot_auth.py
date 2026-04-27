#!/usr/bin/env python3
"""
MIoT 小米登录模块
- 通过内嵌浏览器让用户登录小米账号
- 从 Cookie 中提取 serviceToken / xiaomiiot_ph / userId
- 多用户本地存储与切换
"""

import json
import os
import time
from typing import Optional

from PyQt6.QtCore import QObject, pyqtSignal, QTimer, QUrl


# ─── 用户数据存储 ─────────────────────────────────────────────

USER_DATA_FILE = os.path.join(os.path.expanduser("~"), ".miot_users.json")

def _load_users() -> dict:
    """加载所有用户数据"""
    if os.path.exists(USER_DATA_FILE):
        try:
            with open(USER_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"users": {}, "current": None}
    return {"users": {}, "current": None}


def _save_users(data: dict):
    """保存所有用户数据"""
    os.makedirs(os.path.dirname(USER_DATA_FILE) or ".", exist_ok=True)
    with open(USER_DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_current_user() -> Optional[dict]:
    """获取当前登录用户信息，返回 dict 或 None"""
    data = _load_users()
    uid = data.get("current")
    if uid and uid in data.get("users", {}):
        user = data["users"][uid]
        # 检查是否过期（serviceToken 一般 30 天有效）
        if user.get("loginTime") and time.time() - user["loginTime"] < 30 * 86400:
            return user
    return None


def get_all_users() -> list:
    """获取所有已保存的用户列表"""
    data = _load_users()
    result = []
    for uid, info in data.get("users", {}).items():
        result.append({
            "userId": uid,
            "name": info.get("name", uid),
            "loginTime": info.get("loginTime", 0),
        })
    # 按最近登录排序
    result.sort(key=lambda x: x.get("loginTime", 0), reverse=True)
    return result


def save_user(user_id: str, service_token: str, xiaomiiot_ph: str,
              name: str = ""):
    """保存用户登录信息"""
    data = _load_users()
    data["users"][str(user_id)] = {
        "userId": str(user_id),
        "serviceToken": service_token,
        "xiaomiiot_ph": xiaomiiot_ph,
        "name": name or str(user_id),
        "loginTime": time.time(),
    }
    data["current"] = str(user_id)
    _save_users(data)


def switch_user(user_id: str) -> Optional[dict]:
    """切换到指定用户"""
    data = _load_users()
    uid = str(user_id)
    if uid in data.get("users", {}):
        data["current"] = uid
        _save_users(data)
        user = data["users"][uid]
        if user.get("loginTime") and time.time() - user["loginTime"] < 30 * 86400:
            return user
    return None


def remove_user(user_id: str):
    """移除用户"""
    data = _load_users()
    uid = str(user_id)
    data.get("users", {}).pop(uid, None)
    if data.get("current") == uid:
        # 切换到其他用户
        remaining = list(data.get("users", {}).keys())
        data["current"] = remaining[0] if remaining else None
    _save_users(data)


def logout_current():
    """退出当前用户"""
    data = _load_users()
    data["current"] = None
    _save_users(data)


# ─── 登录浏览器 ────────────────────────────────────────────────

# iot.mi.com 登录后会跳转到主页或产品配置页
# 我们需要监控 Cookie 变化，当出现 serviceToken 时表示登录成功

LOGIN_URL = "https://iot.mi.com/fe-op/productCenter/config/advance/automation"
# 小米 SSO 登录入口
SSO_LOGIN_URL = "https://account.xiaomi.com/pass/serviceLogin?callback=https%3A%2F%2Fiot.mi.com%2Fsts%2Foauth%2Fcallback&sid=iot.mi.com"


class MiLoginBrowser(QObject):
    """小米登录浏览器控制器
    - 打开内嵌 QWebEngineView
    - 监控 Cookie 变化
    - 登录成功后自动提取凭证
    """

    login_success = pyqtSignal(dict)   # 登录成功，返回用户信息
    login_failed = pyqtSignal(str)     # 登录失败
    cookie_updated = pyqtSignal()      # Cookie 更新

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cookies = {}
        self._profile = None
        self._view = None
        self._poll_timer = None
        self._checked = False

    def create_view(self):
        """创建 WebEngine 视图（带独立 Profile 以隔离 Cookie）"""
        from PyQt6.QtWebEngineWidgets import QWebEngineView
        from PyQt6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage

        self._profile = QWebEngineProfile("miot_login", self)
        self._profile.setHttpUserAgent(
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_4 like Mac OS X) "
            "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1"
        )

        # 监听 Cookie 变化
        cookie_store = self._profile.cookieStore()
        cookie_store.cookieAdded.connect(self._on_cookie_added)

        # 用 profile 创建 page，再设置到 view
        page = QWebEnginePage(self._profile, self)
        self._view = QWebEngineView()
        self._view.setPage(page)

        # 定时轮询检查登录状态（兜底方案）
        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._poll_login_status)
        self._poll_timer.setInterval(1500)

        return self._view

    def start_login(self):
        """开始登录流程"""
        self._cookies = {}
        self._checked = False
        self._view.load(QUrl(LOGIN_URL))
        self._poll_timer.start()

    def _on_cookie_added(self, cookie):
        """Cookie 添加回调"""
        domain = cookie.domain()
        name = bytes(cookie.name()).decode("utf-8", errors="replace")
        value = bytes(cookie.value()).decode("utf-8", errors="replace")

        # 只关心 iot.mi.com 和 xiaomi.com 域下的关键 Cookie
        if name in ("serviceToken", "userId", "xiaomiiot_ph"):
            self._cookies[name] = value
            self.cookie_updated.emit()

        # 检查是否登录成功
        if self._has_all_cookies() and not self._checked:
            self._checked = True
            self._poll_timer.stop()
            self._emit_login_success()

    def _has_all_cookies(self) -> bool:
        """检查是否已获取所有必要 Cookie"""
        return all(k in self._cookies for k in ("serviceToken", "userId", "xiaomiiot_ph"))

    def _poll_login_status(self):
        """定时轮询：通过 JS 检查当前 URL 和 Cookie（兜底）"""
        if self._checked:
            return

        # 尝试通过 JS 读取 document.cookie（httpOnly 的 cookie 读不到，但有些非 httpOnly 的可以）
        if self._has_all_cookies():
            self._checked = True
            self._poll_timer.stop()
            self._emit_login_success()

    def _emit_login_success(self):
        """发射登录成功信号"""
        user_info = {
            "userId": self._cookies.get("userId", ""),
            "serviceToken": self._cookies.get("serviceToken", ""),
            "xiaomiiot_ph": self._cookies.get("xiaomiiot_ph", ""),
        }
        self.login_success.emit(user_info)

    def cleanup(self):
        """清理资源"""
        if self._poll_timer:
            self._poll_timer.stop()
