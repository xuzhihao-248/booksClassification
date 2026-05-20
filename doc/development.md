# 开发流程文档

## 1. 环境准备

### 1.1 系统要求

- Windows 10/11 或 Linux
- Python 3.10+
- 稳定的网络连接（需访问 douban.com）

### 1.2 项目初始化

```bash
# 创建项目目录
mkdir D:\code\booksClassification
cd D:\code\booksClassification

# 创建子目录
mkdir doc data\output logs

# 创建虚拟环境（推荐）
python -m venv venv
venv\Scripts\activate   # Windows
# source venv/bin/activate  # Linux

# 安装依赖
pip install httpx beautifulsoup4 lxml openpyxl
```

### 1.3 依赖说明

| 包名 | 版本 | 用途 |
|------|------|------|
| httpx | >=0.27.0 | HTTP客户端，支持连接池、超时、重定向 |
| beautifulsoup4 | >=4.12.0 | HTML解析 |
| lxml | >=5.1.0 | BS4后端解析器（C扩展，速度快） |
| openpyxl | >=3.1.0 | Excel .xlsx 文件读写 |

## 2. 架构设计

### 2.1 整体架构 (v2.0 多源版)

```
main.py  (单文件)
├── SourceSession       — 通用HTTP会话（各源独立配置）
├── DoubanParser        — 豆瓣页面HTML解析
├── DangdangParser      — 当当页面HTML解析 (GBK编码)
├── CLCMapper           — 标签→中图法映射（共用）
├── DedupFilter         — SQLite去重（共用）
├── CheckpointManager   — 断点续传（多源状态）
├── ExcelWriter         — 缓冲Excel输出（共用）
├── AdaptiveBalancer    — 分类平衡调度（多源URL发现）
├── SourceManager       — 源轮换管理器（封禁检测+切换）
└── BookScraper         — 主编排器（多源调度）
```

### 2.2 数据流

```
SourceManager.active_source() → 选择可用源 (douban / dangdang)
    ↓
Source.yield_books() → URL发现 → 列表页爬取 → 详情页爬取 → 解析 → CLC映射
    ↓ 被封
SourceManager.mark_blocked() → 标记冷却 → 切换到下一个源
    ↓ 成功
去重校验(SQLite) → 写入Excel(缓冲500条/批次) → 断点保存(每1000条)
```

### 2.3 源轮换策略

```
douban (优先) → 被封(418/超时) → dangdang → 被封 → 等待douban冷却 → douban
                    ↓ 冷却 600-1800s                ↓ 冷却 300-900s
```

### 2.4 核心类设计

#### SourceSession (通用HTTP会话)

```
每个源独立实例，有自己的cookie文件和封禁状态。

属性:
  name: str                  — 源标识 (douban/dangdang)
  client: httpx.Client       — HTTP客户端
  cookie_file: Path          — 独立Cookie文件 (cookies_{name}.json)
  last_request_time: float   — 上次请求时间戳
  consecutive_failures: int  — 连续失败计数
  blocked_until: float       — 封禁冷却截止时间

方法:
  _build_client()            — 构建客户端（随机UA+Cookie加载）
  _load_cookies()            — 加载已保存Cookie（过期>1h丢弃）
  _save_cookies()            — 保存Cookie到文件
  _enforce_delay()           — 强制延迟（2-5s随机 + 失败加权）
  _rotate_if_needed()        — 连续失败≥3次时换UA+冷却
  is_blocked() → bool        — 查询是否在封禁冷却期
  get(url, referer, encoding)— 发送GET请求（含重试+反爬+自动编码检测）
```

**反爬策略实现：**

| 层级 | 措施 | 代码实现 |
|------|------|----------|
| L1 | 随机延迟 | `_enforce_delay()`: `random(2.0, 5.0)` + 失败加权 |
| L2 | UA轮换 | `_build_client()`: `random.choice(USER_AGENTS)` |
| L3 | Cookie管理 | bid随机生成 + `_save_cookies`/`_load_cookies` + TTL过期 |
| L4 | 请求头伪装 | Accept/Accept-Language/Connection等 |
| L5 | 退避重试 | `get()`: 418/429→指数退避, 403→换身份, 最大5次 |

#### BookParser

```
静态方法:
  extract_douban_id(url) → str       — 从URL提取豆瓣ID
  parse_listing_page(html) → list    — 解析标签列表页
  parse_book_page(html, url) → dict  — 解析书籍详情页
```

**书籍详情页CSS选择器（多级回退）：**

| 字段 | 主选择器 | 回退选择器 |
|------|----------|------------|
| 书名 | `#wrapper h1 span[property="v:itemreviewed"]` | `#wrapper h1`, `[property="v:itemreviewed"]` |
| 简介 | `#link-report .intro` | `.related_info .indent .intro`, `#link-report` |
| 标签 | `#db-tags-section a.tag` | `.tags a` |
| 评分 | `strong.ll.rating_num` | `[property="v:average"]` |

**缺字段处理：**
- 书名缺失 → 整条丢弃（必填字段）
- 简介缺失 → 填充空字符串（可选字段）
- 标签缺失 → 使用来源标签页名称回退分类
- 作者/出版社缺失 → 填充空字符串

#### CLCMapper

```
方法:
  map_tags_to_clc(tags) → list[dict]  — 多标签投票→排序结果
  _lookup(tag) → list[dict]           — 单标签查找（精确→模糊）

查找策略:
  1. 精确匹配：查 tag_to_clc.json（593条映射）
  2. 模糊匹配：SequenceMatcher 相似度 >= 0.80，降权采纳
  3. 兜底：返回 Z（综合性图书），置信度0.0
```

#### ExcelWriter

```
方法:
  write(item)    — 写入缓冲（500条/批次自动刷新）
  _flush()       — 刷新缓冲到Excel
  close()        — 最终写入+自动列宽+筛选器

输出列:
  douban_id, 书名, 作者, 出版社, 出版年, ISBN,
  内容简介, 豆瓣标签, 中图法分类号, 次要分类号,
  分类置信度, 豆瓣评分, 来源URL, 抓取时间
```

#### SourceManager (源轮换管理器)

```
管理多个数据源的可用状态，被封时自动切换。

属性:
  blocked_until: dict[str, float]  — 源名 → 解封时间戳
  SOURCE_ORDER: ["douban", "dangdang"]

方法:
  active_source() → str     — 返回第一个未封源，全被封则等待
  mark_blocked(name, sec)   — 标记源被封，记录冷却时间
  get_state() → dict        — 导出状态用于断点保存
```

## 3. URL发现策略

### 3.1 豆瓣: 标签页浏览（~60%）

```
URL格式: https://book.douban.com/tag/{标签名}?start={offset}&type=T

分页: start=0, 20, 40, ... (最多到980)
每页: 20本书
```

### 3.2 豆瓣: 随机ID扫描（~40%）

```
URL格式: https://book.douban.com/subject/{douban_id}/

ID范围: 1,000,000 ~ 37,000,000
策略: 随机采样，去重过滤
```

### 3.3 当当: 分类页浏览

```
URL格式: http://category.dangdang.com/cp{cat_id}.html?page_index={N}

分页: page_index=1, 2, 3, ... (最大约100)
每页: 60本书
32个分类 → CLC映射，覆盖22大类
```

### 3.4 自适应调度算法

```python
def next_tags(n=15):
    1. 计算各分类完成比例 (current / target)
    2. 按缺口从大到小排列
    3. 每个缺口类别选取1-3个对应标签
    4. 过滤掉已知无效标签 (bad_tags)
    5. 不足n个时随机补充
```

## 4. 数据质量控制

### 4.1 必填校验

- douban_id: 非空，唯一
- title: 非空，否则整条丢弃
- clc_code: 非空，格式匹配 `^[A-Z][A-Za-z0-9.\-]*$`

### 4.2 去重策略

- 主键: douban_id（SQLite持久化）
- 内存set热缓存，SQLite为source of truth
- 每100条commit一次

### 4.3 容错机制

| 异常类型 | 处理方式 |
|----------|----------|
| HTTP 404 | 标签→标记bad_tags; 书籍→跳过 |
| HTTP 429/418 | 指数退避重试，换UA+Cookie |
| HTTP 403 | 换UA+Cookie后重试 |
| HTTP 5xx | 指数退避重试 |
| 网络超时 | 重试3次 |
| HTML解析异常 | 记录日志，跳过该书 |
| 书名缺失 | 跳过该书（必填字段） |
| 标签缺失 | 使用来源标签回退分类 |
| 磁盘满载 | 暂停爬取，flush缓冲 |

## 5. 性能指标

### 5.1 实测数据（20条测试）

| 指标 | 数值 |
|------|------|
| 请求延迟 | 2-5秒（可配置） |
| 单条耗时 | ~4.8秒 |
| 成功率 | 100%（HTTP 200） |
| 10万条预估 | 5-6天 |

### 5.2 优化空间

- 降低 REQUEST_DELAY_MIN 到 1.5秒 → 3.5天
- 使用代理池分散请求 → 2天
- 多线程/协程（谨慎，增加被封风险）

## 6. 迭代记录

### v2.0.1 (2026-05-19) — 紧急修复

**Bug: 断点恢复时Excel数据被覆盖清空**

- **现象**: 用户爬取440条后重启，Excel文件被空白工作簿覆盖，已有数据全部丢失
- **根因**: `ExcelWriter.__init__()` 总是创建新 `Workbook()`，不检查文件是否已存在
- **修复**: 
  - `ExcelWriter` 新增 `resume` 参数，为 `True` 时用 `load_workbook()` 加载已有文件并续写
  - `BookScraper` 将 `resume` 标志传递给 `ExcelWriter`
  - 兼容旧版列数不同的文件（自动补齐缺失列头）
- **恢复**: 删除 `seen_books.db` 去重库，让已爬ID可被重新采集

### v2.0 (2026-05-19) — 多源版
- 新增当当网数据源，支持三源轮换
- 重构为 SourceSession（通用HTTP）+ SourceManager（源调度）
- 新增 DangdangParser：GBK编码处理、面包屑分类→CLC映射
- 32个当当分类映射到22个中图法大类
- 封禁自动切换：豆瓣被封→当当，当当被封→等豆瓣冷却
- 断点增加 dd_cursor 和 source_states 字段
- Excel增加"数据来源"列标识源

### v1.0 (2026-05-19)

- 初始版本，支持豆瓣读书数据采集
- 实现5层反爬措施
- 实现标签→中图法映射（593条）
- 实现断点续传
- 实现自适应分类平衡调度

### 已知问题

1. 豆瓣部分标签页不存在（如"毛泽东思想"、"邓小平理论"），自动标记为bad_tags跳过
2. 豆瓣改版后部分书籍详情页不再展示用户标签，使用来源标签回退
3. Windows终端GBK编码不支持Unicode进度条，使用ASCII字符替代
4. 长时间运行可能触发豆瓣IP限流，需要冷却后自动恢复
5. 当当部分图书详情页无简介（"本商品暂无详情"），使用列表页摘要回退
6. 京东图书反爬级别过高（需Playwright+代理），暂不集成
