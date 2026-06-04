"""bit_api 模块 - BitBrowser API 调用"""
import requests
import json

API_URL = "http://127.0.0.1:54345"

# 默认超时时间（秒）
DEFAULT_TIMEOUT = 120


def openBrowser(window_id, timeout=None):
    """打开 BitBrowser 浏览器窗口

    参数:
        window_id: BitBrowser窗口ID
        timeout: 超时时间（秒），默认120秒
    """
    payload = {"id": window_id}
    timeout = timeout or DEFAULT_TIMEOUT

    try:
        resp = requests.post(f"{API_URL}/browser/open", json=payload, timeout=timeout)
        data = resp.json()

        if data.get("success"):
            result_data = data.get("data", {})

            # 直接使用 API 返回的 http 和 driver
            http_port = result_data.get("http", "")
            driver_path = result_data.get("driver", "")

            return {
                "success": True,
                "data": {
                    "driver": driver_path,
                    "http": http_port  # 已经是 127.0.0.1:端口 格式
                }
            }
        else:
            print(f"BitBrowser 打开失败: {data}")
            return {"success": False, "msg": str(data)}

    except requests.exceptions.Timeout:
        print(f"BitBrowser API 请求超时（{timeout}秒），请检查BitBrowser响应速度")
        return {"success": False, "msg": f"请求超时({timeout}s)"}
    except requests.exceptions.ConnectionError:
        print("BitBrowser API 连接失败，请确认 BitBrowser 是否已启动")
        return {"success": False, "msg": "API连接失败"}
    except Exception as e:
        print(f"BitBrowser 调用异常: {e}")
        return {"success": False, "msg": str(e)}


def closeBrowser(window_id):
    """关闭 BitBrowser 浏览器窗口"""
    payload = {"id": window_id}

    try:
        resp = requests.post(f"{API_URL}/browser/close", json=payload, timeout=30)
        data = resp.json()

        if data.get("success"):
            return {"success": True}
        else:
            return {"success": False, "msg": str(data)}

    except requests.exceptions.ConnectionError:
        return {"success": False, "msg": "API连接失败"}
    except Exception as e:
        return {"success": False, "msg": str(e)}


def updateBrowserCookies(window_id, cookies):
    """更新浏览器 Cookies 到 BitBrowser"""
    if not cookies:
        return None

    try:
        cookie_list = []

        if isinstance(cookies, list):
            # Selenium get_cookies() 返回的列表
            for ck in cookies:
                if isinstance(ck, dict) and ck.get('name') and ck.get('value'):
                    clean_cookie = {
                        "name": ck['name'],
                        "value": ck['value'],
                        "domain": ck.get('domain', ''),
                        "path": ck.get('path', '/'),
                    }
                    if ck.get('expires') and isinstance(ck.get('expires'), (int, float)) and ck['expires'] > 0:
                        clean_cookie['expires'] = int(ck['expires'])
                    if ck.get('secure'):
                        clean_cookie['secure'] = ck['secure']
                    if ck.get('httpOnly'):
                        clean_cookie['httpOnly'] = ck['httpOnly']
                    cookie_list.append(clean_cookie)

        payload = {
            "id": window_id,
            "cookies": cookie_list
        }

        resp = requests.post(f"{API_URL}/browser/cookies/update", json=payload, timeout=30)
        result = resp.json()
        return result

    except Exception as e:
        print(f"Cookies 更新异常: {e}")
        return None