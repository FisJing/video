# VideoAnalytics 视频数据分析工具

多平台视频数据抓取分析工具，获取已发布视频的播放量、点赞、评论等统计数据。

## 支持平台

| 平台 | 方式 | 需要登录 | 依赖库 |
|------|------|----------|--------|
| **YouTube** | API | ❌ | yt-dlp |
| **TikTok** | API | ❌ | TikTok-Api, playwright |
| **Instagram** | API | ❌ | instaloader |
| **B站** | API | ❌ | bilibili-api-python |
| **抖音** | Selenium | ✅ | BitBrowser |
| **微信视频号** | Selenium | ✅ | BitBrowser |
| **Facebook** | Selenium | ✅ | BitBrowser |
| **X/Twitter** | Selenium | ✅ | BitBrowser |
| **VK** | Selenium | ✅ | BitBrowser |
| **Rutube** | Selenium | ✅ | BitBrowser |
| **Dzen** | Selenium | ✅ | BitBrowser |

## 文件结构

```
VideoAnalytics/
├── analyze_base.py              # 基础框架（数据结构、抽象类）
├── analyze_all_new.py           # 统一入口脚本
│
├── analyze_youtube_new.py       # YouTube分析（yt-dlp）
├── analyze_tiktok_new.py        # TikTok分析（TikTok-Api）
├── analyze_instagram_new.py     # Instagram分析（Instaloader）
├── analyze_bilibili_new.py      # B站分析（bilibili-api）
├── analyze_douyin_new.py        # 抖音分析（Selenium）
│
├── analyze_weixin_new.py        # 微信视频号分析（Selenium）
├── analyze_facebook_new.py      # Facebook分析（Selenium）
├── analyze_x_new.py             # X/Twitter分析（Selenium）
├── analyze_vk_new.py            # VK分析（Selenium）
├── analyze_rutube_new.py        # Rutube分析（Selenium）
├── analyze_dzen_new.py          # Dzen分析（Selenium）
│
├── *_videos.xlsx                # 输入文件（视频URL列表）
├── *_stats.xlsx                 # 输出文件（统计数据）
├── window_ids.json              # BitBrowser窗口ID配置
│
└── README.md                    # 本文档
```

## 安装依赖

### API方式（无需登录）

```bash
# YouTube
pip install yt-dlp

# TikTok
pip install TikTok-Api playwright
playwright install

# Instagram
pip install instaloader

# B站
pip install bilibili-api-python
```

### Selenium方式（需要BitBrowser）

```bash
pip install selenium openpyxl requests

# 需要安装 BitBrowser 指纹浏览器
# 需要在 BitBrowser 中登录各平台账号
```

## 使用方法

### 1. 准备输入文件

为每个平台创建 Excel 文件，第一列放视频URL：

```
youtube_videos.xlsx:
| URL |
|-----|
| https://www.youtube.com/watch?v=xxx |
| https://www.youtube.com/watch?v=yyy |

weixin_videos.xlsx:
| URL |
|-----|
| https://channels.weixin.qq.com/xxx |
```

### 2. 运行分析

#### API方式（无需登录）

```bash
# 单平台分析
python analyze_youtube_new.py
python analyze_tiktok_new.py
python analyze_instagram_new.py
python analyze_bilibili_new.py

# 使用统一入口
python analyze_all_new.py youtube
python analyze_all_new.py tiktok

# 分析所有API平台
python analyze_all_new.py all
```

#### Selenium方式（需要登录）

```bash
# 单平台分析（需要window_id）
python analyze_weixin_new.py <window_id>
python analyze_facebook_new.py <window_id>
python analyze_x_new.py <window_id>

# 使用统一入口
python analyze_all_new.py weixin <window_id>
python analyze_all_new.py facebook <window_id>

# 配置window_ids.json后批量运行
python analyze_all_new.py all_selenium
```

### 3. 配置 window_ids.json

```json
{
    "weixin": "3f62ade3059f4e4bb796f9aa77d8e9b0",
    "facebook": "xxx...",
    "x": "xxx...",
    "vk": "xxx...",
    "rutube": "xxx...",
    "dzen": "xxx...",
    "douyin": "b6a1052e23604b068ecda52858edbbd0"
}
```

## 输出数据字段

| 字段 | 说明 |
|------|------|
| url | 视频URL |
| platform | 平台名称 |
| title | 视频标题 |
| views | 播放量 |
| likes | 点赞数 |
| comments | 评论数 |
| shares | 分享数 |
| favorites | 收藏数 |
| duration | 视频时长（秒） |
| author | 作者名称 |
| author_fans | 作者粉丝数 |
| publish_time | 发布时间 |
| fetch_time | 数据抓取时间 |

## 各平台特点

### YouTube (yt-dlp)
- 最成熟的方案，支持1000+网站
- 获取完整元数据：播放量、点赞、评论、时长、作者、粉丝数
- 无需登录，速度快

### TikTok (TikTok-Api)
- 需要异步运行
- 首次使用需安装 playwright
- 获取播放量、点赞、分享、评论、收藏

### Instagram (Instaloader)
- 支持帖子、Reels、IGTV
- 登录后可获取更多数据
- 注意：播放量通常不公开显示

### B站 (bilibili-api)
- 完整的B站API封装
- 获取播放量、点赞、评论、分享、收藏、弹幕数
- 可查询作者粉丝数

### 抖音
- 反爬严格，需要Selenium+BitBrowser
- 从创作者中心获取数据
- 支持Shadow DOM页面结构

### 微信视频号
- 使用 wujie-app Shadow DOM
- 需从创作者中心获取数据
- 登录态有效期较短（12-72小时）

### Facebook
- 反爬严格，推荐指纹浏览器
- 数据解析需要处理特殊格式

### X/Twitter
- 播放量显示为 views/impressions
- 需要处理动态DOM结构

### VK / Rutube / Dzen
- 俄罗斯平台
- 需处理俄语单位（тыс、млн）
- VK类似Facebook结构

## 成熟开源方案参考

| 项目 | Stars | 平台 |
|------|-------|------|
| [yt-dlp](https://github.com/yt-dlp/yt-dlp) | 156,950 | YouTube + 1000+网站 |
| [Douyin_TikTok_Download_API](https://github.com/Evil0ctal/Douyin_TikTok_Download_API) | 17,175 | 抖音、TikTok、快手、B站 |
| [gallery-dl](https://github.com/mikf/gallery-dl) | 17,775 | 多平台图片/视频 |
| [Instaloader](https://github.com/instaloader/instaloader) | 12,132 | Instagram |
| [TikTok-Api](https://github.com/davidteather/TikTok-Api) | 6,288 | TikTok |
| [bilibili-api](https://github.com/Nemo2011/bilibili-api) | 3,813 | B站 |

## 注意事项

1. **API平台**：无登录限制，可直接运行，注意请求频率
2. **Selenium平台**：需要先在 BitBrowser 中登录账号
3. **数据准确性**：部分平台数据可能因页面更新而变化
4. **反爬处理**：建议使用 BitBrowser 指纹浏览器绕过检测
5. **请求频率**：批量分析时注意间隔，避免触发限制

## 作者

Lremi