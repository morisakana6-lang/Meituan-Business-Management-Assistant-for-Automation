"""
美团经营宝凭证获取模块

=== 功能说明 ===
通过 Playwright 启动浏览器，拦截美团经营宝页面的网络请求，
从请求 URL 中正则提取 mtgsig 签名参数，并收集登录 Cookie。

=== 凭证文件 ===
| 文件 | 路径 | 作用 | 有效期 |
|------|------|------|--------|
| auth.json | auth/ | Playwright 登录态，可复用 | 数天~数周 |
| credentials.json | auth/ | Cookie + mtgsig | 会过期，需定期更新 |

=== 凭证说明 ===
- Cookie: 登录态验证，包含用户身份信息
- mtgsig: 签名参数，内含时间戳，每次请求都会变化

=== 命令行用法 ===
    python common/credentials.py              # 首次运行：扫码登录 → 获取凭证
    python common/credentials.py --force     # 强制重新登录（清除已有登录态）
    python common/credentials.py -f           # 同上，简写

环境变量：
    HEADLESS=true python common/credentials.py    # 无头模式（服务器部署）

=== 核心流程 ===
1. 检查 auth.json 是否存在
   - 存在 → 加载登录态，访问报表页面
   - 不存在 → 打开浏览器登录页面，等待用户扫码
2. 设置网络响应拦截器，监听 getBoardReport/getBigBoardData 请求
3. 从捕获的请求 URL 中正则提取 mtgsig
4. 收集 dianping.com/meituan.com 域名的所有 Cookie
5. 保存登录态到 auth.json，凭证到 credentials.json

=== 被其他模块调用方式 ===
from common.credentials import load_credentials, save_credentials

cookie, mtgsig = load_credentials()     # 加载凭证
save_credentials(cookie, mtgsig)        # 保存凭证
"""

import json
import time
import re
import os
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth


# 获取项目根目录（meituan/）
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDENTIALS_FILE = os.path.join(_PROJECT_ROOT, "auth", "credentials.json")
AUTH_STATE_FILE = os.path.join(_PROJECT_ROOT, "auth", "auth.json")


def stealth_config(page):
    """
    配置 stealth 模式

    使用 playwright-stealth 隐藏自动化特征：
    - 伪造 webdriver 属性
    - 伪造 navigator 属性
    - 伪造浏览器指纹
    - 隐藏自动化控制条

    目的：防止美团风控检测到自动化访问
    """
    stealth = Stealth()
    stealth.apply_stealth_sync(page)


def check_login_status(context) -> bool:
    """
    检查是否已登录

    通过检查 Cookie 中是否包含登录凭证：
    - edper: 大众点评登录凭证
    - JSESSIONID: 通用会话标识

    Returns:
        True: 已登录
        False: 未登录
    """
    cookies = context.cookies()
    cookie_dict = {c['name']: c['value'] for c in cookies}
    return 'edper' in cookie_dict or 'JSESSIONID' in cookie_dict


def fetch_credentials():
    """
    获取 Cookie 和 mtgsig 的核心逻辑

    Returns:
        tuple: (cookie_str, mtgsig, need_relogin)
            - cookie_str: Cookie 字符串
            - mtgsig: mtgsig 签名参数
            - need_relogin: 是否需要重新登录（登录态过期时）
    """
    use_existing_auth = os.path.exists(AUTH_STATE_FILE)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        if use_existing_auth:
            print("检测到已保存的登录态，正在加载...")
            context = browser.new_context(storage_state=AUTH_STATE_FILE)
        else:
            context = browser.new_context(
                viewport={"width": 1280, "height": 720}
            )

        page = context.new_page()
        stealth_config(page)

        mtgsig_captured = []

        def handle_response(response):
            url = response.url
            if "getBoardReport" in url or "getBigBoardData" in url:
                print(f"捕获到目标请求: {url[:80]}...")
                mtgsig_match = re.search(r'mtgsig=([^&]+)', url)
                if mtgsig_match:
                    mtgsig_captured.append(mtgsig_match.group(1))
                    print(f"[OK] mtgsig 捕获成功 (长度: {len(mtgsig_captured[-1])})")

        page.on("response", handle_response)
        print("拦截器已设置，等待捕获 mtgsig...")

        if use_existing_auth:
            print("正在访问页面...")
            page.goto("https://e.dianping.com/shopdiy-node/report", wait_until="domcontentloaded", timeout=30000)

            if not check_login_status(context):
                print("登录态已过期，需要重新登录")
                browser.close()
                return None, None, True

            print("登录态有效，等待页面加载...")
            time.sleep(5)
        else:
            print("正在访问登录页面...")
            page.goto("https://e.dianping.com/shopdiy-node/report", wait_until="domcontentloaded", timeout=60000)

            print("\n" + "="*60)
            print("请在浏览器中完成登录（扫码或账号密码）")
            print("="*60)

            max_wait = 300
            start_time = time.time()

            while time.time() - start_time < max_wait:
                if check_login_status(context):
                    print("\n登录成功！正在跳转报表页面...")
                    page.goto("https://e.dianping.com/shopdiy-node/report", wait_until="networkidle", timeout=60000)
                    time.sleep(5)
                    break
                time.sleep(1)
            else:
                print("登录超时！")
                browser.close()
                return None, None, False

        cookies = context.cookies()
        cookie_list = []
        for c in cookies:
            domain = c.get('domain', '')
            if 'dianping.com' in domain or 'meituan.com' in domain:
                cookie_list.append(f"{c['name']}={c['value']}")
        cookie_str = "; ".join(cookie_list)

        mtgsig = mtgsig_captured[0] if mtgsig_captured else None

        print("正在保存登录态...")
        context.storage_state(path=AUTH_STATE_FILE)

        browser.close()

        return cookie_str, mtgsig, False


def save_credentials(cookie: str, mtgsig: str = ""):
    """
    保存凭证到 auth/credentials.json

    Args:
        cookie: Cookie 字符串
        mtgsig: mtgsig 签名参数（可选）

    注意：如果 mtgsig 为空，会保留文件中已有的 mtgsig
    """
    existing = {}
    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except:
            pass

    data = {
        "cookie": cookie,
        "mtgsig": mtgsig or existing.get("mtgsig", ""),
        "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_credentials():
    """
    从 auth/credentials.json 加载凭证

    Returns:
        tuple: (cookie, mtgsig)
            - 如果文件不存在或读取失败，返回 (None, None)

    被 tuiguangtong/client.py 等模块调用：
        from common import load_credentials
        cookie, mtgsig = load_credentials()
    """
    if not os.path.exists(CREDENTIALS_FILE):
        return None, None

    try:
        with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
            creds = json.load(f)
        return creds.get("cookie", ""), creds.get("mtgsig", "")
    except:
        return None, None


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="美团经营宝凭证获取工具")
    parser.add_argument("--force", "-f", action="store_true", help="强制重新登录")
    args = parser.parse_args()

    print("="*60)
    print("美团经营宝凭证获取工具 (网络拦截方案)")
    print("="*60)

    if args.force and os.path.exists(AUTH_STATE_FILE):
        print("强制重新登录，删除旧登录态...")
        os.remove(AUTH_STATE_FILE)

    cookie, mtgsig, need_relogin = fetch_credentials()

    if need_relogin:
        print("\n登录态过期，正在重新登录...")
        cookie, mtgsig, _ = fetch_credentials()

    if cookie:
        print(f"\n[OK] Cookie 获取成功 (长度: {len(cookie)})")
        save_credentials(cookie, mtgsig)

        print("\n" + "="*60)
        print("凭证获取完成！")
        print("="*60)

        if mtgsig:
            print(f"[OK] mtgsig 获取成功 (长度: {len(mtgsig)})")
            print(f"mtgsig: {mtgsig[:80]}...")
        else:
            print("[WARNING] mtgsig 未捕获到，请手动刷新报表页面后重试")
    else:
        print("\n[ERROR] Cookie 获取失败，请重试！")
