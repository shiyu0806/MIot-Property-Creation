#!/usr/bin/env python3
"""
MIoT 平台抓包工具
启动有头浏览器，打开 MIoT 平台，用户操作后抓取相关 API 请求
"""
import json
import os
import sys
import time
from playwright.sync_api import sync_playwright

TARGET_URL = "https://iot.mi.com/fe-op/productCenter"
CAPTURED = []


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/146.0.0.0 Safari/537.36",
        )
        page = context.new_page()

        # 抓取所有 API 请求
        def on_request(request):
            url = request.url
            # 只抓 MIoT 相关 API
            if "iot.mi.com" in url and ("/api/" in url or "/cgi-" in url):
                entry = {
                    "method": request.method,
                    "url": url,
                    "headers": dict(request.headers),
                }
                if request.post_data:
                    try:
                        entry["body"] = json.loads(request.post_data)
                    except (json.JSONDecodeError, TypeError):
                        entry["body"] = request.post_data
                CAPTURED.append(entry)
                print(f"📡 [{request.method}] {url[:100]}")

        def on_response(response):
            url = response.url
            if "iot.mi.com" in url and ("/api/" in url or "/cgi-" in url):
                for entry in CAPTURED:
                    if entry["url"] == url:
                        entry["status"] = response.status
                        try:
                            entry["response"] = response.json()
                        except Exception:
                            entry["response_text"] = response.text()[:500]
                        break

        page.on("request", on_request)
        page.on("response", on_response)

        print(f"🌐 正在打开 {TARGET_URL} ...")
        page.goto(TARGET_URL, wait_until="networkidle", timeout=30000)
        print(f"✅ 页面已加载，请在浏览器中操作")
        print(f"   操作完成后，按 Enter 键保存抓包结果并退出")

        try:
            input()
        except (EOFError, KeyboardInterrupt):
            pass

        # 保存结果
        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "captured_api.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(CAPTURED, f, ensure_ascii=False, indent=2)
        print(f"\n✅ 共抓取 {len(CAPTURED)} 条 API 请求")
        print(f"💾 结果已保存: {output_path}")

        browser.close()


if __name__ == "__main__":
    main()
