# -*- coding: utf-8 -*-
"""
Rutube视频数据抓取 - 普通视频和Shorts
抓取频道下所有视频的详细数据（观看量、点赞数、评论数等）

使用方式：
  python rutube.py <window_id>          # 抓取所有视频
  python rutube.py <window_id> --shorts # 只抓Shorts
  python rutube.py <window_id> --videos # 只抓普通视频

输出：
  分析结果/rutube_videos.xlsx 或 rutube_shorts.xlsx
"""
import sys
import os
import time
import openpyxl
from datetime import datetime

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "分析结果")

sys.path.insert(0, os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, os.path.join(os.path.dirname(SCRIPT_DIR), "VideoAutoPost"))

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bit_api import openBrowser

# Cookie数据
COOKIES_DATA = [
    {"name":"jwt","value":"eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJydXBhc3NfaWQiOjYzMjEyMjczLCJlbWFpbF9jb25maXJtZWQiOnsidHlwZW9mIjoiYm9vbCIsInZhbHVlIjp0cnVlfSwiaGFzX3Bob25lIjp7InR5cGVvZiI6ImJvb2wiLCJ2YWx1ZSI6ZmFsc2V9LCJleHAiOjE3NzY5OTUzMjIsIm9yaWdfaWF0IjoxNzc0NDAzMzIyLCJ1c2VyX2lkIjo2Nzg1OTM1OCwiZGF0YSI6eyJ1c2VyX2luZm8iOnsiaWQiOjY3ODU5MzU4fX0sInVzZXJuYW1lIjoicHJjdmFsdmUyMUBnbWFpbC5jb20iLCJlbWFpbCI6InByY3ZhbHZlMjFAZ21haWwuY29tIiwibmFtZSI6IlByYyBWYWx2ZSBNZWRpYS15aSIsInBhcmVudCI6bnVsbCwidHlwZSI6ImFjY291bnQiLCJjb250ZXh0IjpudWxsLCJ0cnVzdCI6W10sInN1YnMiOltdfQ.bBXrC_TNo9ICpol--9SdNNPoADsoJuYqzjATYkXTb5k","domain":".rutube.ru","path":"/","httpOnly":True,"secure":True},
    {"name":"psid2","value":"q8lrneb113105lk4deh6e7gj8jt0978y","domain":".rutube.ru","path":"/","httpOnly":True,"secure":True},
    {"name":"refreshToken","value":"eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJydXBhc3NfaWQiOjYzMjEyMjczLCJleHAiOjE4MDU5MzkzMjIsInJlcXVlc3RfaWQiOiI2YjM5ZmI2Ny0zNDE5LTRhYTEtYWVjZi1jN2UyZjM3NzA5MjQifQ.OHWh-n9IkdAv3iwjQvhlwz65eKLyOczRakSYSN6mnow","domain":".rutube.ru","path":"/","httpOnly":True,"secure":True},
    {"name":"session_id","value":"15326383151774402983_1774402983904","domain":".rutube.ru","path":"/","secure":True},
    {"name":"visitorID","value":"67859358","domain":".rutube.ru","path":"/","secure":True},
    {"name":"cid","value":"15326383151774402983","domain":".rutube.ru","path":"/","secure":True},
    {"name":"csrftoken","value":"ZZ6FvUhb2UZcsObKemvioyL14XvFTyxN","domain":"rutube.ru","path":"/","secure":True}
]


def connect_browser(window_id):
    res = openBrowser(window_id)
    if not res.get("success"):
        print(f"浏览器启动失败: {res.get('msg', '')}")
        return None
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", res["data"]["http"])
    options.add_argument("--disable-backgrounding-occluded-windows")
    driver = webdriver.Chrome(service=Service(res["data"]["driver"]), options=options)
    return driver


def inject_cookies(driver):
    print("注入Cookie...")
    for c in COOKIES_DATA:
        try:
            driver.add_cookie({'name': c['name'], 'value': c['value'], 'domain': c.get('domain', '.rutube.ru'), 'path': c.get('path', '/')})
        except: pass
    print("Cookie注入完成")


def parse_num(str_val):
    """解析数字（支持K、M等）"""
    try:
        s = str(str_val).upper().replace(',', '.')
        if 'K' in s:
            return float(s.replace('K', '')) * 1000
        if 'M' in s:
            return float(s.replace('M', '')) * 1000000
        return int(s.replace('.', '').replace(' ', '')) or 0
    except:
        return 0


def scroll_and_get_shorts(driver, max_scrolls=100):
    """滚动获取Shorts列表"""
    all_shorts = {}
    no_new_count = 0

    print("滚动获取Shorts列表...")

    for i in range(max_scrolls):
        time.sleep(3)

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
                    var s = {
                        id: idMatch[1],
                        url: link.href,
                        title: '',
                        views: 0,
                        type: 'shorts'
                    };

                    var linkText = (link.innerText || '').trim();
                    if (!linkText.match(/^[0-9]+\+$/)) {
                        s.title = linkText.substring(0, 100);
                    }

                    var parent = link.parentElement;
                    if (parent) {
                        var parentText = parent.innerText || '';
                        var viewsMatch = parentText.match(/([0-9,.]+[KkMm]?)\\s*(просмотров|просмотра)/i);
                        if (viewsMatch) s.viewsStr = viewsMatch[1];
                    }

                    if (!shorts.some(v => v.id === s.id)) {
                        shorts.push(s);
                    }
                }
            }
            return shorts;
        """)

        new_count = 0
        for s in shorts:
            sid = s.get("id")
            if sid and sid not in all_shorts:
                s["views"] = parse_num(s.get("viewsStr", "0"))
                all_shorts[sid] = s
                new_count += 1

        print(f"  滚动{i+1}: 累计{len(all_shorts)}个 (新增{new_count})")

        if new_count == 0:
            no_new_count += 1
            if no_new_count >= 5:
                print("连续5次无新内容")
                break
        else:
            no_new_count = 0

        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(2)

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
                    var v = {
                        id: idMatch[1],
                        url: link.href,
                        title: '',
                        views: 0,
                        type: 'video'
                    };

                    v.title = (link.innerText || '').trim().substring(0, 100);

                    var container = link.closest('[class*="video-item"], [class*="card"]') || link.parentElement;
                    if (container) {
                        var containerText = container.innerText || '';
                        var viewsMatch = containerText.match(/([0-9,.]+[KkMm]?)\\s*(просмотров|просмотра)/i);
                        if (viewsMatch) v.viewsStr = viewsMatch[1];
                    }

                    if (!videos.some(v => v.id === v.id)) {
                        videos.push(v);
                    }
                }
            }
            return videos;
        """)

        new_count = 0
        for v in videos:
            vid = v.get("id")
            if vid and vid not in all_videos:
                v["views"] = parse_num(v.get("viewsStr", "0"))
                all_videos[vid] = v
                new_count += 1

        print(f"  滚动{i+1}: 累计{len(all_videos)}个 (新增{new_count})")

        if new_count == 0:
            no_new_count += 1
            if no_new_count >= 5:
                print("连续5次无新内容")
                break
        else:
            no_new_count = 0

        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(2)

    return list(all_videos.values())


def get_detail(driver, video_url):
    """进入详情页获取完整数据"""
    try:
        driver.get(video_url)
        time.sleep(6)

        data = driver.execute_script("""
            var result = {
                views: 0,
                likes: 0,
                comments: 0,
                favorites: 0,
                shares: 0,
                duration: '',
                publishDate: ''
            };

            function parseNum(str) {
                str = str.toString().toUpperCase().replace(',', '.');
                if (str.indexOf('K') !== -1) return parseFloat(str.replace('K', '').replace('K', '')) * 1000;
                if (str.indexOf('M') !== -1) return parseFloat(str.replace('M', '').replace('M', '')) * 1000000;
                if (str.indexOf('ТЫС') !== -1) return parseFloat(str.replace('ТЫС', '')) * 1000;
                return parseInt(str.replace(/[^0-9]/g, '')) || 0;
            }

            try {
                // 观看量
                var viewsPatterns = [
                    /([0-9,.]+[KkMmтыс]*)\\s*(просмотров|просмотра|views)/i
                ];
                var pageText = document.body.innerText || '';
                for (var pattern of viewsPatterns) {
                    var match = pageText.match(pattern);
                    if (match) {
                        result.views = parseNum(match[1]);
                        break;
                    }
                }

                // 点赞数
                var likeButtons = document.querySelectorAll('[class*="like"], [aria-label*="лайк"], [aria-label*="like"]');
                for (var btn of likeButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) {
                        result.likes = parseNum(txt.trim());
                        break;
                    }
                }

                // 评论数
                var commentSection = document.querySelector('[class*="comment"]');
                if (commentSection) {
                    var commentCount = commentSection.querySelector('[class*="count"]') || commentSection;
                    var txt = commentCount.innerText || '';
                    var numMatch = txt.match(/[0-9]+/);
                    if (numMatch) result.comments = parseInt(numMatch[0]) || 0;
                }

                // 收藏数
                var favButtons = document.querySelectorAll('[class*="favorite"], [class*="bookmark"], [aria-label*="Избранное"]');
                for (var btn of favButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) {
                        result.favorites = parseNum(txt.trim());
                        break;
                    }
                }

                // 分享数
                var shareButtons = document.querySelectorAll('[class*="share"], [aria-label*="share"], [aria-label*="Поделиться"]');
                for (var btn of shareButtons) {
                    var counter = btn.querySelector('[class*="count"]') || btn.nextElementSibling || btn;
                    var txt = counter.innerText || counter.textContent || '';
                    if (txt.match(/[0-9]/)) {
                        result.shares = parseNum(txt.trim());
                        break;
                    }
                }

                // 发布时间 - 多种方式获取
                var timeSelectors = [
                    'time',
                    '[class*="date"]',
                    '[class*="time"]',
                    '[class*="publish"]'
                ];
                for (var sel of timeSelectors) {
                    var el = document.querySelector(sel);
                    if (el) {
                        var dt = el.getAttribute('datetime');
                        if (dt) {
                            result.publishDate = dt.substring(0, 10);
                            break;
                        }
                        var txt = (el.innerText || el.textContent || '').trim();
                        if (txt && txt.length < 50) {
                            // 相对时间
                            var relMatch = txt.match(/(\d+[днейденьчасовчасминутминут]+ назад|\d+[dhd]+ ago)/i);
                            if (relMatch) {
                                result.publishDate = relMatch[0];
                                break;
                            }
                            // 日期格式
                            var dateMatch = txt.match(/(\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{2,4})/);
                            if (dateMatch) {
                                result.publishDate = dateMatch[0];
                                break;
                            }
                            result.publishDate = txt.substring(0, 30);
                            break;
                        }
                    }
                }

                // 时长
                var durationEl = document.querySelector('[class*="duration"]');
                if (durationEl) result.duration = durationEl.innerText.trim();

            } catch(e) {}
            return result;
        """)

        data["views"] = data.get("views", 0)
        data["likes"] = data.get("likes", 0)
        data["comments"] = data.get("comments", 0)
        data["favorites"] = data.get("favorites", 0)
        data["shares"] = data.get("shares", 0)
        return data

    except Exception as e:
        print(f"    获取详情失败: {e}")
        return {"views": 0, "likes": 0, "comments": 0, "favorites": 0, "shares": 0, "duration": "", "publishDate": ""}


def save_excel(videos, output_file, video_type):
    if not videos:
        print("无数据")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Rutube {video_type}"

    headers = ["序号", "视频ID", "URL", "标题", "观看量", "点赞数", "评论数", "收藏数", "分享数", "发布时间", "平台", "抓取时间"]
    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)

    for i, v in enumerate(videos, 2):
        ws.cell(i, 1, i-1)
        ws.cell(i, 2, v.get("id", ""))
        ws.cell(i, 3, v.get("url", ""))
        ws.cell(i, 4, v.get("title", ""))
        ws.cell(i, 5, v.get("views", 0))
        ws.cell(i, 6, v.get("likes", 0))
        ws.cell(i, 7, v.get("comments", 0))
        ws.cell(i, 8, v.get("favorites", 0))
        ws.cell(i, 9, v.get("shares", 0))
        ws.cell(i, 10, v.get("publishDate", ""))
        ws.cell(i, 11, "Rutube")
        ws.cell(i, 12, datetime.now().strftime("%Y-%m-%d %H:%M"))

    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 40
    wb.save(output_file)
    print(f"已保存: {output_file}")


def main():
    print("=" * 50)
    print("Rutube视频数据抓取")
    print("=" * 50)

    if len(sys.argv) < 2:
        print("用法:")
        print("  python rutube.py <window_id>          # 抓取所有")
        print("  python rutube.py <window_id> --shorts # 只抓Shorts")
        print("  python rutube.py <window_id> --videos # 只抓普通视频")
        return

    window_id = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "all"

    driver = connect_browser(window_id)
    if not driver:
        return

    try:
        # 清理标签页
        handles = driver.window_handles
        for i in range(len(handles)-1, 0, -1):
            try:
                driver.switch_to.window(handles[i])
                driver.close()
            except: pass
        driver.switch_to.window(driver.window_handles[0])

        # 访问主页注入Cookie
        driver.get("https://rutube.ru/")
        time.sleep(5)
        inject_cookies(driver)
        driver.refresh()
        time.sleep(5)

        user_id = "67859358"
        all_videos = []

        # 根据模式抓取
        if mode in ["all", "shorts"]:
            shorts_url = f"https://rutube.ru/channel/{user_id}/shorts/"
            print(f"\n访问Shorts页面: {shorts_url}")
            driver.get(shorts_url)
            time.sleep(8)

            shorts = scroll_and_get_shorts(driver)
            print(f"\n获取到 {len(shorts)} 个Shorts")

            if shorts:
                print("\n获取详细数据...")
                for i, s in enumerate(shorts):
                    url = s.get("url", "")
                    if not url:
                        continue

                    print(f"  {i+1}/{len(shorts)}: {s.get('title', '')[:30]}")
                    detail = get_detail(driver, url)
                    s["views"] = detail.get("views", 0) or s.get("views", 0)
                    s["likes"] = detail.get("likes", 0)
                    s["comments"] = detail.get("comments", 0)
                    s["favorites"] = detail.get("favorites", 0)
                    s["shares"] = detail.get("shares", 0)
                    s["publishDate"] = detail.get("publishDate", "")
                    s["duration"] = detail.get("duration", "")

                    print(f"    观看:{s.get('views', 0)} 点赞:{s['likes']} 评论:{s['comments']} 发布:{s['publishDate']}")

                    if (i + 1) % 15 == 0:
                        time.sleep(2)
                    driver.back()
                    time.sleep(2)

                all_videos.extend(shorts)

        if mode in ["all", "videos"]:
            videos_url = f"https://rutube.ru/channel/{user_id}/videos/"
            print(f"\n访问视频页面: {videos_url}")
            driver.get(videos_url)
            time.sleep(8)

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
                    v["duration"] = detail.get("duration", "")

                    print(f"    观看:{v.get('views', 0)} 点赞:{v['likes']} 评论:{v['comments']} 发布:{v['publishDate']}")

                    if (i + 1) % 15 == 0:
                        time.sleep(2)
                    driver.back()
                    time.sleep(2)

                all_videos.extend(videos)

        # 保存数据
        if all_videos:
            output_file = os.path.join(OUTPUT_DIR, "rutube_videos.xlsx")
            save_excel(all_videos, output_file, mode)

            print("\n数据预览:")
            for v in all_videos[:10]:
                print(f"  {v.get('title', '')[:25]} | 观看:{v.get('views', 0)} 点赞:{v.get('likes', 0)}")

    except Exception as e:
        print(f"异常: {e}")
        import traceback
        traceback.print_exc()

    finally:
        try: driver.quit()
        except: pass

    print("=" * 50)


if __name__ == "__main__":
    main()