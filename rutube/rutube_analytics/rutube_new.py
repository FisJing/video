# -*- coding: utf-8 -*-
"""
Rutube视频数据抓取 - 改进版
抓取频道下所有视频的详细数据（观看量、点赞数、评论数等）

改进点:
  - 移除硬编码Cookie，使用浏览器登录状态
  - 添加环境检测和登录状态检测
  - 统一输出目录

使用方式:
  python rutube_new.py <window_id>          # 抓取所有视频
  python rutube_new.py <window_id> --shorts # 只抓Shorts
  python rutube_new.py <window_id> --videos # 只抓普通视频

输出:
  分析结果/rutube_videos.xlsx
"""

import sys
import os
import time
import json
import openpyxl
import requests
from datetime import datetime

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "分析结果")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bit_api import openBrowser


def parse_number(num_str):
    """解析数字（支持K、M、тыс、млн等单位）"""
    if not num_str:
        return 0

    num_str = str(num_str).upper().strip().replace(',', '.').replace(' ', '')

    try:
        if 'ТЫС' in num_str:
            return int(float(num_str.replace('ТЫС', '').strip()) * 1000)
        if 'МЛН' in num_str:
            return int(float(num_str.replace('МЛН', '').strip()) * 1000000)
        if 'K' in num_str:
            return int(float(num_str.replace('K', '').strip()) * 1000)
        if 'M' in num_str:
            return int(float(num_str.replace('M', '').strip()) * 1000000)
        if '万' in num_str:
            return int(float(num_str.replace('万', '').strip()) * 10000)

        return int(float(num_str.replace('.', '').replace(',', '')))
    except:
        return 0


def check_bitbrowser():
    """检测BitBrowser是否运行"""
    api_url = "http://127.0.0.1:54345"

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                api_url = config.get('bitbrowser_api', api_url)
        except:
            pass

    try:
        resp = requests.get(api_url, timeout=5)
        return True
    except:
        return False


def check_login(driver):
    """检测是否已登录Rutube"""
    try:
        result = driver.execute_script("""
            var avatar = document.querySelector('[class*="avatar"], [class*="user-avatar"], img[class*="avatar"]');
            var loginBtn = document.querySelector('[class*="login"], [href*="login"], a[href*="signin"]');
            return {
                hasAvatar: avatar !== null,
                hasLoginBtn: loginBtn !== null
            };
        """)
        return result.get('hasAvatar', False)
    except:
        return False


def connect_browser(window_id):
    """连接浏览器"""
    res = openBrowser(window_id)
    if not res.get("success"):
        print(f"❌ 浏览器启动失败: {res.get('msg', '')}")
        return None

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", res["data"]["http"])
    options.add_argument("--disable-backgrounding-occluded-windows")

    driver = webdriver.Chrome(service=Service(res["data"]["driver"]), options=options)
    return driver


def scroll_and_get_shorts(driver, max_scrolls=150):
    """滚动获取Shorts列表"""
    all_shorts = {}
    no_new_count = 0

    print("滚动获取Shorts列表...")

    for i in range(max_scrolls):
        time.sleep(4)

        shorts = driver.execute_script("""
            var shorts = [];
            var links = document.querySelectorAll('a[href*="/shorts/"]');
            links = Array.from(links).filter(function(l) {
                return l.href.indexOf('entry_point') === -1 &&
                       l.href.indexOf('forms') === -1 &&
                       l.href.indexOf('?') === -1 &&
                       l.href.match(/shorts\\/([a-zA-Z0-9]+)\\/?$/) !== null;
            });

            for (var link of links) {
                var idMatch = link.href.match(/shorts\\/([a-zA-Z0-9]+)\\/?$/);
                if (idMatch) {
                    var s = { id: idMatch[1], url: link.href, title: '', views: 0, type: 'shorts' };
                    var linkText = (link.innerText || '').trim();
                    if (!linkText.match(/^[0-9]+\+$/)) s.title = linkText.substring(0, 100);
                    var parent = link.parentElement;
                    if (parent) {
                        var parentText = parent.innerText || '';
                        var viewsMatch = parentText.match(/([0-9,.]+[KkMmтысмлн]*)\\s*(просмотров|просмотра)/i);
                        if (viewsMatch) s.viewsStr = viewsMatch[1];
                    }
                    if (!shorts.some(v => v.id === s.id)) shorts.push(s);
                }
            }
            return shorts;
        """)

        new_count = 0
        for s in shorts:
            sid = s.get("id")
            if sid and sid not in all_shorts:
                s["views"] = parse_number(s.get("viewsStr", "0"))
                all_shorts[sid] = s
                new_count += 1

        print(f"  滚动{i+1}: 累计{len(all_shorts)}个 (新增{new_count})")

        if new_count == 0:
            no_new_count += 1
            if no_new_count >= 10:
                print("连续10次无新内容，完成")
                break
        else:
            no_new_count = 0

        driver.execute_script("window.scrollBy(0, 800);")
        time.sleep(1)

    return list(all_shorts.values())


def scroll_and_get_videos(driver, max_scrolls=50):
    """滚动获取普通视频列表"""
    all_videos = {}
    no_new_count = 0

    print("滚动获取视频列表...")

    for i in range(max_scrolls):
        time.sleep(3)

        videos = driver.execute_script("""
            var videos = [];
            var links = document.querySelectorAll('a[href*="/video/"]');
            links = Array.from(links).filter(function(l) {
                return l.href.match(/video\\/([a-zA-Z0-9]+)/) !== null;
            });

            for (var link of links) {
                var idMatch = link.href.match(/video\\/([a-zA-Z0-9]+)/);
                if (idMatch) {
                    var v = { id: idMatch[1], url: link.href, title: '', views: 0, type: 'video' };
                    v.title = (link.innerText || '').trim().substring(0, 100);
                    var container = link.closest('[class*="video-item"], [class*="card"]') || link.parentElement;
                    if (container) {
                        var containerText = container.innerText || '';
                        var viewsMatch = containerText.match(/([0-9,.]+[KkMmтысмлн]*)\\s*(просмотров|просмотра)/i);
                        if (viewsMatch) v.viewsStr = viewsMatch[1];
                    }
                    if (!videos.some(v => v.id === v.id)) videos.push(v);
                }
            }
            return videos;
        """)

        new_count = 0
        for v in videos:
            vid = v.get("id")
            if vid and vid not in all_videos:
                v["views"] = parse_number(v.get("viewsStr", "0"))
                all_videos[vid] = v
                new_count += 1

        print(f"  滚动{i+1}: 累计{len(all_videos)}个 (新增{new_count})")

        if new_count == 0:
            no_new_count += 1
            if no_new_count >= 10:
                print("连续10次无新内容，完成")
                break
        else:
            no_new_count = 0

        driver.execute_script("window.scrollBy(0, 800);")
        time.sleep(1)

    return list(all_videos.values())


def get_detail(driver, video_url):
    """进入详情页获取完整数据"""
    try:
        driver.get(video_url)
        time.sleep(5)

        data = driver.execute_script("""
            var result = { title: '', views: 0, likes: 0, comments: 0, favorites: 0, shares: 0, duration: '', publishDate: '' };

            function parseNum(str) {
                str = str.toString().toUpperCase().replace(',', '.');
                if (str.indexOf('ТЫС') !== -1) return parseFloat(str.replace('ТЫС', '').trim()) * 1000;
                if (str.indexOf('МЛН') !== -1) return parseFloat(str.replace('МЛН', '').trim()) * 1000000;
                if (str.indexOf('K') !== -1) return parseFloat(str.replace('K', '').trim()) * 1000;
                if (str.indexOf('M') !== -1) return parseFloat(str.replace('M', '').trim()) * 1000000;
                return parseInt(str.replace(/[^0-9]/g, '')) || 0;
            }

            try {
                // 获取标题 - 多种方式尝试
                var metaTitle = document.querySelector('meta[property="og:title"]');
                if (metaTitle && metaTitle.getAttribute('content')) {
                    result.title = metaTitle.getAttribute('content').substring(0, 100);
                }

                // 如果og:title无效，尝试其他方式
                if (!result.title || result.title === 'Главная' || result.title === 'Rutube') {
                    // 尝试document.title
                    var docTitle = document.title || '';
                    if (docTitle && docTitle.indexOf('Rutube') === -1 && docTitle.indexOf('Главная') === -1) {
                        result.title = docTitle.split(' — ')[0].split(' - ')[0].substring(0, 100);
                    }
                }

                if (!result.title || result.title === 'Главная' || result.title === 'Rutube') {
                    // 尝试页面内的标题元素
                    var titleEl = document.querySelector('[class*="video-title"], [class*="video-name"], .video-info-title, h1:not(.logo)');
                    if (titleEl) {
                        var t = (titleEl.innerText || titleEl.textContent || '').trim();
                        if (t && t !== 'Главная' && t !== 'Rutube') result.title = t.substring(0, 100);
                    }
                }

                // 最后尝试从URL路径提取
                if (!result.title || result.title === 'Главная' || result.title === 'Rutube') {
                    var urlPath = window.location.pathname;
                    var parts = urlPath.split('/');
                    if (parts.length > 2) {
                        result.title = 'Video ' + parts[parts.length - 1].substring(0, 20);
                    }
                }

                var pageText = document.body.innerText || '';
                var viewsMatch = pageText.match(/([0-9,.]+[KkMmтысмлн]*)\\s*(просмотров|просмотра|views)/i);
                if (viewsMatch) result.views = parseNum(viewsMatch[1]);

                var likeButtons = document.querySelectorAll('[class*="like"], [aria-label*="лайк"], [aria-label*="like"]');
                for (var btn of likeButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) { result.likes = parseNum(txt.trim()); break; }
                }

                var commentSection = document.querySelector('[class*="comment"]');
                if (commentSection) {
                    var txt = commentSection.innerText || '';
                    var numMatch = txt.match(/[0-9]+/);
                    if (numMatch) result.comments = parseInt(numMatch[0]) || 0;
                }

                var favButtons = document.querySelectorAll('[class*="favorite"], [class*="bookmark"], [aria-label*="Избранное"]');
                for (var btn of favButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) { result.favorites = parseNum(txt.trim()); break; }
                }

                var shareButtons = document.querySelectorAll('[class*="share"], [aria-label*="share"], [aria-label*="Поделиться"]');
                for (var btn of shareButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) { result.shares = parseNum(txt.trim()); break; }
                }

                var timeEl = document.querySelector('time');
                if (timeEl) {
                    var dt = timeEl.getAttribute('datetime');
                    if (dt) result.publishDate = dt.substring(0, 10);
                    else { var txt = (timeEl.innerText || '').trim(); if (txt.length < 50) result.publishDate = txt; }
                }

                var durationEl = document.querySelector('[class*="duration"]');
                if (durationEl) result.duration = durationEl.innerText.trim();
            } catch(e) {}
            return result;
        """)

        return data

    except Exception as e:
        print(f"    获取详情失败: {e}")
        return {"title": "", "views": 0, "likes": 0, "comments": 0, "favorites": 0, "shares": 0, "duration": "", "publishDate": ""}


def save_excel(videos, output_file):
    """保存到Excel"""
    if not videos:
        print("无数据")
        return

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rutube视频"

    headers = ["序号", "视频ID", "URL", "标题", "类型", "观看量", "点赞数", "评论数", "收藏数", "分享数", "发布时间", "平台", "抓取时间"]
    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)

    for i, v in enumerate(videos, 2):
        ws.cell(i, 1, i-1)
        ws.cell(i, 2, v.get("id", ""))
        ws.cell(i, 3, v.get("url", ""))
        ws.cell(i, 4, v.get("title", ""))
        ws.cell(i, 5, v.get("type", ""))
        ws.cell(i, 6, v.get("views", 0))
        ws.cell(i, 7, v.get("likes", 0))
        ws.cell(i, 8, v.get("comments", 0))
        ws.cell(i, 9, v.get("favorites", 0))
        ws.cell(i, 10, v.get("shares", 0))
        ws.cell(i, 11, v.get("publishDate", ""))
        ws.cell(i, 12, "Rutube")
        ws.cell(i, 13, datetime.now().strftime("%Y-%m-%d %H:%M"))

    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 40

    wb.save(output_file)
    print(f"✅ 已保存: {output_file} ({len(videos)} 条)")


def main():
    print("=" * 60)
    print("Rutube视频数据抓取 - 改进版")
    print("=" * 60)

    if not check_bitbrowser():
        print("❌ BitBrowser 未运行，请先启动")
        print("   提示: 运行 python check_environment.py 检测环境")
        return

    if len(sys.argv) < 2:
        print("\n用法:")
        print("  python rutube_new.py <window_id>          # 抓取所有")
        print("  python rutube_new.py <window_id> --shorts # 只抓Shorts")
        print("  python rutube_new.py <window_id> --videos # 只抓普通视频")
        print("\n也可以使用config.json中的配置:")
        print("  python run_russian.py rutube")
        return

    window_id = sys.argv[1]
    mode = sys.argv[2].lstrip('--') if len(sys.argv) > 2 else "shorts"

    user_channel_id = "67859358"
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                user_channel_id = config.get('platforms', {}).get('rutube', {}).get('user_channel_id', user_channel_id)
        except:
            pass

    print(f"\n连接浏览器 window_id: {window_id}")
    driver = connect_browser(window_id)
    if not driver:
        return

    try:
        handles = driver.window_handles
        for i in range(len(handles)-1, 0, -1):
            try:
                driver.switch_to.window(handles[i])
                driver.close()
            except:
                pass
        driver.switch_to.window(driver.window_handles[0])

        print("\n访问 Rutube 主页...")
        driver.get("https://rutube.ru/")
        time.sleep(3)

        if not check_login(driver):
            print("❌ 未检测到登录状态，请先在浏览器中登录 Rutube 账号")
            print("   提示: 登录后运行 python get_cookies.py 保存Cookie")
            return

        print("✅ 已检测到登录状态")

        all_videos = []

        if mode in ["all", "shorts"]:
            shorts_url = f"https://rutube.ru/channel/{user_channel_id}/shorts/"
            print(f"\n访问Shorts页面: {shorts_url}")
            driver.get(shorts_url)
            time.sleep(5)

            shorts = scroll_and_get_shorts(driver)
            print(f"\n获取到 {len(shorts)} 个Shorts")

            if shorts:
                print("\n获取详细数据...")
                for i, s in enumerate(shorts):
                    url = s.get("url", "")
                    if not url:
                        continue

                    try:
                        print(f"  {i+1}/{len(shorts)}: {s.get('title', '')[:30]}")
                        detail = get_detail(driver, url)
                        # 如果详情页获取到标题，则更新
                        if detail.get("title"):
                            s["title"] = detail.get("title")
                        s["views"] = detail.get("views", 0) or s.get("views", 0)
                        s["likes"] = detail.get("likes", 0)
                        s["comments"] = detail.get("comments", 0)
                        s["favorites"] = detail.get("favorites", 0)
                        s["shares"] = detail.get("shares", 0)
                        s["publishDate"] = detail.get("publishDate", "")

                        print(f"    标题:{s['title'][:25] if s['title'] else 'N/A'} 观看:{s['views']} 点赞:{s['likes']}")

                        if (i + 1) % 20 == 0:
                            time.sleep(1)

                        # 返回列表页，超时时重新导航
                        try:
                            driver.back()
                            time.sleep(1)
                        except Exception as e:
                            print(f"    返回超时，重新导航...")
                            driver.get(shorts_url)
                            time.sleep(3)

                    except Exception as e:
                        print(f"    处理失败: {e}")
                        # 尝试恢复到列表页
                        try:
                            driver.get(shorts_url)
                            time.sleep(3)
                        except:
                            pass
                        continue

                all_videos.extend(shorts)

        if mode in ["all", "videos"]:
            videos_url = f"https://rutube.ru/channel/{user_channel_id}/videos/"
            print(f"\n访问视频页面: {videos_url}")
            driver.get(videos_url)
            time.sleep(5)

            videos = scroll_and_get_videos(driver)
            print(f"\n获取到 {len(videos)} 个视频")

            if videos:
                print("\n获取详细数据...")
                for i, v in enumerate(videos):
                    url = v.get("url", "")
                    if not url:
                        continue

                    print(f"  {i+1}/{len(videos)}: {v.get('title', '')[:30]}")
                    detail = get_detail(driver, url)
                    v["views"] = detail.get("views", 0) or v.get("views", 0)
                    v["likes"] = detail.get("likes", 0)
                    v["comments"] = detail.get("comments", 0)
                    v["favorites"] = detail.get("favorites", 0)
                    v["shares"] = detail.get("shares", 0)
                    v["publishDate"] = detail.get("publishDate", "")

                    print(f"    观看:{v['views']} 点赞:{v['likes']} 评论:{v['comments']}")

                    if (i + 1) % 20 == 0:
                        time.sleep(1)
                    try:
                        driver.back()
                        time.sleep(1)
                    except Exception as e:
                        print(f"    返回失败，重新导航...")
                        driver.get(shorts_url)
                        time.sleep(3)

                all_videos.extend(videos)

        if all_videos:
            output_file = os.path.join(OUTPUT_DIR, "rutube_videos.xlsx")
            save_excel(all_videos, output_file)

            print("\n数据预览:")
            for v in all_videos[:10]:
                print(f"  {v.get('title', '')[:25]} | 观看:{v.get('views', 0)} 点赞:{v.get('likes', 0)}")
        else:
            print("\n❌ 未获取到任何视频数据")

    except Exception as e:
        print(f"❌ 异常: {e}")
        import traceback
        traceback.print_exc()

        # 异常时保存已获取的数据
        if all_videos:
            print("\n保存已获取的数据...")
            output_file = os.path.join(OUTPUT_DIR, "rutube_videos_partial.xlsx")
            save_excel(all_videos, output_file)
            print(f"✅ 已保存 {len(all_videos)} 条数据到: {output_file}")

    finally:
        try:
            driver.quit()
        except:
            pass

    print("=" * 60)


if __name__ == "__main__":
    main()