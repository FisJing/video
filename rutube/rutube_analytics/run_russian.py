# -*- coding: utf-8 -*-
"""
Rutube视频抓取入口脚本

使用方式:
  python run_russian.py rutube  # 抓取Rutube

配置文件: config.json
"""

import sys
import os
import json
import subprocess

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")


def load_config():
    """加载配置文件"""
    if not os.path.exists(CONFIG_FILE):
        print("❌ config.json 不存在")
        return None

    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError:
        print("❌ config.json 格式错误")
        return None


def get_window_id(config):
    """获取rutube的window_id"""
    if not config:
        return None
    platforms = config.get('platforms', {})
    return platforms.get('rutube', {}).get('window_id', '')


def run_rutube(window_id=None):
    """运行rutube抓取脚本"""
    script_path = os.path.join(SCRIPT_DIR, "rutube_new.py")
    if not os.path.exists(script_path):
        print(f"❌ 脚本不存在: {script_path}")
        return False

    if not window_id:
        config = load_config()
        window_id = get_window_id(config)

    if not window_id:
        print("❌ rutube 未配置 window_id")
        return False

    print(f"\n{'='*60}")
    print(f"开始抓取 RUTUBE 视频数据")
    print(f"window_id: {window_id}")
    print(f"{'='*60}\n")

    try:
        result = subprocess.run(
            [sys.executable, script_path, window_id],
            cwd=SCRIPT_DIR,
            capture_output=False,
            text=True
        )
        return result.returncode == 0
    except Exception as e:
        print(f"❌ 运行失败: {e}")
        return False


def show_help():
    """显示帮助"""
    print("=" * 60)
    print("Rutube视频抓取工具")
    print("=" * 60)
    print("\n使用方式:")
    print("  python run_russian.py rutube")
    print("\n准备步骤:")
    print("  1. 启动 BitBrowser")
    print("  2. 登录Rutube账号")
    print("  3. 配置 config.json 中的 window_id")
    print("=" * 60)


def main():
    if len(sys.argv) < 2:
        show_help()
        return

    platform = sys.argv[1].lower()

    if platform in ["help", "--help", "-h"]:
        show_help()
        return

    if platform != "rutube":
        print(f"❌ 本工具仅支持 rutube 平台")
        return

    config = load_config()
    if not config:
        return

    window_id = get_window_id(config)
    if not window_id:
        print("❌ rutube 未配置 window_id")
        return

    success = run_rutube(window_id)
    print(f"\n{'✅' if success else '❌'} rutube 抓取{'完成' if success else '失败'}")


if __name__ == "__main__":
    main()