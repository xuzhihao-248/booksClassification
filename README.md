# 中文图书数据采集系统

从豆瓣读书采集中文图书数据（书名、内容简介、中图法分类号），用于训练中文文本分类器。目标 10 万条，覆盖中图法全部 22 个一级大类。

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 测试运行（20 条，验证全流程）
python main.py --max 20

# 中等运行（500 条）
python main.py --max 500

# 正式运行（10 万条）
python main.py

# 从断点恢复
python main.py --resume
```

## 输出

数据写入 `data/output/books_100k.xlsx`，包含以下列：

| 列名 | 说明 |
|------|------|
| 书名 | 图书名称 |
| 内容简介 | 图书内容简介 |
| 中图法分类号 | CLC 分类号（如 I247.5、TP311.13） |
| 作者 | 作者 |
| 出版社 | 出版社名称 |
| 出版年 | 出版年份 |
| ISBN | 国际标准书号 |
| 豆瓣标签 | 用户标签（分号分隔） |
| 次要分类号 | 备选分类号 |
| 分类置信度 | 0.0 - 1.0 |
| 豆瓣评分 | 用户评分 |
| 来源URL | 豆瓣页面链接 |
| 抓取时间 | ISO 8601 时间戳 |

## 架构

```
main.py                    # 单文件，包含全部模块
├── AntiScrapeSession      # HTTP 会话管理 + 反爬（UA 轮换、Cookie、冷却）
├── BookParser             # 豆瓣页面 HTML 解析
├── CLCMapper              # 豆瓣标签 → 中图法分类号映射
├── CheckpointManager      # 断点续传状态管理
├── DedupFilter            # SQLite 持久化去重
├── ExcelWriter            # 缓冲批量写入 Excel
├── AdaptiveBalancer       # 自适应分类平衡调度
└── BookScraper            # 主调度器
tag_to_clc.json            # 标签 → CLC 映射表（5000+ 条）
```

## 分类策略

豆瓣不直接提供中图法分类号，通过**标签映射法**推断：

```
豆瓣用户标签 → tag_to_clc.json → 中图法分类号
```

1. **精确匹配** — 标签直接命中映射表
2. **模糊匹配** — 字符串相似度 >= 0.80
3. **多标签投票** — 每本书 5-10 个标签加权计算
4. **兜底分类** — 无法匹配时归入 Z 类

## 依赖

- Python >= 3.12
- httpx — HTTP 客户端
- beautifulsoup4 + lxml — HTML 解析
- openpyxl — Excel 读写
- fake-useragent — UA 生成

## 文档

- [需求文档](doc/proposal.md)
- [测试文档](doc/testing.md)

## 注意事项

- 爬虫按 2-5 秒间隔请求，避免对豆瓣造成压力
- 如遇封禁（HTTP 418），程序会自动冷却并更换身份
- 支持 Ctrl+C 优雅中断，进度自动保存
- 重新运行前建议清除 `data/session_cookies.json` 和 `data/seen_books.db`
