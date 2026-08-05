# Overseas Video Automation

VK / Rutube / Dzen 海外视频平台数据分析与自动化系统

基于 BitBrowser 指纹浏览器 + Python + Selenium 的多平台视频数据抓取系统，覆盖俄罗斯主流视频/内容平台，自动获取已发布视频的播放量、点赞、评论等统计数据，并导出 Excel 报表。

## ✨ 功能特性

- 🌍 **三平台覆盖**：VK（社交）、Rutube（视频）、Dzen（内容）
- 🤖 **指纹浏览器多开**：基于 BitBrowser 实现多账号隔离，每个平台独立浏览器环境
- 📊 **数据抓取**：自动抓取频道下所有视频的播放量、点赞数、评论数等
- 📈 **Excel 导出**：结果统一导出为 .xlsx，便于二次分析
- 🔁 **登录态复用**：使用浏览器已登录状态，无需硬编码 Cookie

## 🏗️ 架构

```
BitBrowser 指纹浏览器（多窗口隔离）
       |
       +-- dzen/         Dzen 视频数据抓取
       +-- rutube/       Rutube 视频数据抓取
       +-- vk/           VK 视频数据抓取
              |
              v
       Selenium（连接 BitBrowser 调试端口）
              |
              v
       分析结果/*.xlsx（数据导出）
```

## 📁 目录结构

```
overseas-video-automation/
|-- dzen/
|   |-- dzen_new.py        # Dzen 抓取主脚本
|   |-- bit_api.py         # BitBrowser API 封装
|   |-- config.json        # 平台配置
|   `-- 分析结果/           # 输出目录
|-- rutube/
|   `-- rutube_analytics/
|       |-- rutube_new.py  # Rutube 抓取主脚本
|       |-- run_russian.py # 入口脚本
|       |-- bit_api.py
|       |-- config.json
|       `-- 分析结果/
`-- vk/
    |-- vk_final.py        # VK 抓取主脚本
    |-- bit_api.py
    `-- 分析结果/
```

## 🚀 快速开始

### 前置要求

- Python 3.8+
- [BitBrowser 指纹浏览器](https://www.bitbrowser.cn/)（已启动，并为各平台创建浏览器窗口）
- 依赖：selenium、openpyxl、requests

### 使用方式

每个平台独立运行，传入 BitBrowser 窗口 ID：

```bash
# Dzen 视频数据抓取
python dzen/dzen_new.py <window_id>

# Rutube 抓取（支持筛选）
python rutube/rutube_analytics/rutube_new.py <window_id>           # 全部视频
python rutube/rutube_analytics/rutube_new.py <window_id> --shorts  # 仅 Shorts
python rutube/rutube_analytics/rutube_new.py <window_id> --videos  # 仅普通视频

# VK 视频数据抓取
python vk/vk_final.py <window_id>
```

结果输出到各平台 `分析结果/` 目录下的 `.xlsx` 文件。

## 🔧 配置

各平台 `config.json` 配置 BitBrowser API 地址与窗口 ID（请替换为你的实际值）：

```json
{
  "bitbrowser_api": "http://127.0.0.1:54345",
  "platforms": {
    "dzen": {
      "window_id": "<your_bitbrowser_window_id>",
      "channel_url": "<your_channel_url>"
    }
  }
}
```

## 🛠️ 技术栈

Python · Selenium · openpyxl · requests · BitBrowser 指纹浏览器

## 📄 License

MIT
