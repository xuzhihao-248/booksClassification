# 测试流程文档

## 1. 测试策略

分四个阶段，逐级验证系统可靠性：

```
阶段1: 连通性测试（单体，~2分钟）    → 验证网络+解析
阶段2: 小批量测试（20条，~2分钟）     → 验证全流程
阶段3: 中批量测试（200条，~16分钟）   → 验证稳定性+分类分布
阶段4: 正式运行（10万条，5-6天）      → 完整数据采集
```

## 2. 阶段1：连通性测试

### 2.1 目的

验证网络连接、HTML解析、CLC映射模块单独可用。

### 2.2 测试脚本

```bash
# 清除旧状态，避免cookie污染
Remove-Item -Force data/session_cookies.json -ErrorAction SilentlyContinue
Remove-Item -Force data/seen_books.db -ErrorAction SilentlyContinue

# 运行连通性测试
python -c "
import sys
sys.path.insert(0, '.')
from main import AntiScrapeSession, BookParser, CLCMapper

# 1. 测试标签页
s = AntiScrapeSession()
r = s.get('https://book.douban.com/tag/%E5%B0%8F%E8%AF%B4?start=0&type=T')
assert r is not None, 'FAIL: 标签页无响应'
assert r.status_code == 200, f'FAIL: 标签页HTTP {r.status_code}'
assert len(r.text) > 10000, f'FAIL: 标签页HTML过短({len(r.text)}字节)'

# 2. 测试列表解析
books = BookParser.parse_listing_page(r.text)
assert len(books) > 0, 'FAIL: 列表页解析为0条'

# 3. 测试书籍详情
b = books[0]
r2 = s.get(b['url'])
assert r2 is not None, 'FAIL: 详情页无响应'
assert r2.status_code == 200, f'FAIL: 详情页HTTP {r2.status_code}'

data = BookParser.parse_book_page(r2.text, b['url'])
assert data.get('title'), 'FAIL: 书名解析为空'

# 4. 测试CLC映射
mapper = CLCMapper()
tags = data.get('douban_tags', [])
if not tags:
    tags = ['小说']  # 回退标签
mappings = mapper.map_tags_to_clc(tags)
assert len(mappings) > 0, 'FAIL: CLC映射为空'
assert mappings[0]['clc_code'], 'FAIL: CLC编码为空'

print('阶段1 全部通过!')
print(f'  标签页: {len(books)}本书')
print(f'  书名: {data[\"title\"][:30]}')
print(f'  简介长度: {len(data.get(\"description\", \"\"))}')
print(f'  CLC: {mappings[0][\"clc_code\"]} (置信度: {mappings[0][\"confidence\"]:.2f})')
"
```

### 2.5 当当网连通性测试 (v2.0新增)

```bash
python -c "
import sys
sys.path.insert(0, '.')
from main import SourceSession, DangdangParser

# 1. 测试分类列表页
s = SourceSession('dangdang_test')
r = s.get('http://category.dangdang.com/cp01.03.00.00.00.00.html?page_index=1',
         referer='http://book.dangdang.com/')
assert r is not None, 'FAIL: 列表页无响应'
assert r.status_code == 200, f'FAIL: 列表页HTTP {r.status_code}'
assert len(r.content) > 10000, f'FAIL: 列表页过短({len(r.content)}字节)'

# 2. 测试列表页解析
books = DangdangParser.parse_listing_page(r.content)
assert len(books) > 0, 'FAIL: 列表页解析为0条'

# 3. 测试书籍详情
b = books[0]
r2 = s.get(b['url'], referer='http://category.dangdang.com/')
assert r2 is not None, 'FAIL: 详情页无响应'
assert r2.status_code == 200, f'FAIL: 详情页HTTP {r2.status_code}'

data = DangdangParser.parse_book_page(r2.content, b['url'])
assert data.get('title'), 'FAIL: 书名解析为空'
assert data.get('product_id'), 'FAIL: 产品ID提取失败'

# 4. 测试CLC映射
categories = data.get('douban_tags', [])
if not categories:
    categories = ['小说']  # 回退
mappings = DangdangParser.map_categories_to_clc(categories)
assert len(mappings) > 0, 'FAIL: CLC映射为空'
assert mappings[0]['clc_code'], 'FAIL: CLC编码为空'

print('当当网阶段1 全部通过!')
print(f'  列表页: {len(books)}本书')
print(f'  书名: {data[\"title\"][:30]}')
print(f'  分类: {categories}')
print(f'  简介长度: {len(data.get(\"description\", \"\"))}')
print(f'  CLC: {mappings[0][\"clc_code\"]} ({mappings[0][\"sub\"]})')
"
```

### 2.6 源切换测试 (v2.0新增)

```bash
# 测试源轮换逻辑
python -c "
import sys
sys.path.insert(0, '.')
from main import SourceManager, SourceSession

mgr = SourceManager()
assert mgr.active_source() == 'douban', 'FAIL: 默认源应为douban'

# 模拟豆瓣被封
mgr.mark_blocked('douban', 60)
assert mgr.active_source() == 'dangdang', 'FAIL: 应切换到dangdang'

# 模拟两个都被封
mgr.mark_blocked('dangdang', 60)
print('两个源均被封时的行为:')
# 预期: 等待冷却最快的源恢复
print('  douban 冷却:', mgr.blocked_until.get('douban', 0) - __import__('time').time(), 's')
print('  dangdang 冷却:', mgr.blocked_until.get('dangdang', 0) - __import__('time').time(), 's')
print('源切换逻辑测试通过!')
"
```

### 2.7 通过标准

- [ ] 标签页HTTP 200，HTML > 10KB
- [ ] 列表页解析出 ≥ 1 本书
- [ ] 详情页HTTP 200，书名不为空
- [ ] CLC映射返回有效分类号

### 2.4 实测记录

| 日期 | 标签页字节 | 列表数 | 书名 | CLC | 状态 |
|------|-----------|--------|------|-----|------|
| 2026-05-19 | 51,604 | 20 | 某某计划 | I24 | 通过 |

---

## 3. 阶段2：小批量测试

### 3.1 目的

验证完整爬取管道：URL发现 → 抓取 → 解析 → 分类 → 写入Excel。

### 3.2 测试命令

```bash
# 清除所有旧状态
Remove-Item -Force data/session_cookies.json -ErrorAction SilentlyContinue
Remove-Item -Force data/seen_books.db -ErrorAction SilentlyContinue
Remove-Item -Force data/checkpoint.json -ErrorAction SilentlyContinue
Remove-Item -Force data/output/books_100k.xlsx -ErrorAction SilentlyContinue

# 运行20条测试
python main.py --max 20
```

### 3.3 验证项目

#### 3.3.1 运行日志检查

观察日志输出，逐项确认：

- [ ] 无 HTTP 418/403/429 错误
- [ ] 无 "会话被限流" 警告
- [ ] 无连续3次以上请求失败
- [ ] 标签页有数据返回（非"无更多数据"）
- [ ] 每个标签页能找到书籍并爬取详情

#### 3.3.2 Excel输出检查

```bash
python -c "
import openpyxl
wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
ws = wb.active
print(f'总行数: {ws.max_row - 1}')  # 减去表头
print()

# 逐行检查
empty_title = 0
empty_desc = 0
empty_clc = 0
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1]: empty_title += 1       # 书名
    if not row[6]: empty_desc += 1         # 简介
    if not row[8]: empty_clc += 1          # CLC

print(f'书名缺失: {empty_title}')
print(f'简介缺失: {empty_desc}')
print(f'CLC缺失: {empty_clc}')
print()

# 显示前3条
for row in ws.iter_rows(min_row=2, max_row=4, values_only=True):
    print(f'[{row[8]}] {row[1][:30]}  |  {str(row[6])[:50]}...')
"
```

- [ ] 总行数 = 目标数（20）
- [ ] 书名缺失 = 0
- [ ] CLC分类号缺失 = 0
- [ ] 简介缺失 < 20%（部分书籍确实无简介）

#### 3.3.3 分类覆盖检查

- [ ] 所有CLC编码首字母在中图法22大类中
- [ ] 置信度范围在 0.0 ~ 1.0
- [ ] 无异常的编码格式（如纯数字）

### 3.4 通过标准

- [ ] 20条全部成功写入
- [ ] 0条书名缺失
- [ ] 0条CLC缺失
- [ ] 无HTTP错误（429/403/418）
- [ ] Excel文件可正常打开

### 3.5 实测记录

| 日期 | 目标数 | 实际数 | 成功率 | 耗时 | CLC覆盖 | 状态 |
|------|--------|--------|--------|------|---------|------|
| 2026-05-19 | 20 | 20 | 100% | 97s | A(18), D(1), G(1) | 通过 |
| 2026-05-19 | 5 | 5 | 100% | 25s | A(5) | 通过 |

---

## 4. 阶段3：中批量测试

### 4.1 目的

验证系统在较长时间运行下的稳定性，以及多类别覆盖能力。

### 4.2 测试命令

```bash
# 清除旧状态
Remove-Item -Force data/session_cookies.json -ErrorAction SilentlyContinue
Remove-Item -Force data/seen_books.db -ErrorAction SilentlyContinue
Remove-Item -Force data/checkpoint.json -ErrorAction SilentlyContinue
Remove-Item -Force data/output/books_100k.xlsx -ErrorAction SilentlyContinue

# 运行500条测试
python main.py --max 500
```

### 4.3 验证项目

#### 4.3.1 稳定性指标

- [ ] 成功率 >= 95%
- [ ] 无连续10次以上相同错误
- [ ] 无触发豆瓣封禁（302到/misc/sorry）
- [ ] 内存占用稳定（无持续增长）

#### 4.3.2 分类分布检查

```bash
python -c "
import openpyxl
from collections import Counter
wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
ws = wb.active
clc_counter = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[8]:
        clc_counter[row[8][0]] += 1

print('分类分布:')
for cat in sorted(clc_counter.keys()):
    count = clc_counter[cat]
    bar = '#' * (count // 5) if count >= 5 else '.'
    print(f'  {cat}: {bar} {count}')
print(f'总类别数: {len(clc_counter)}')
"
```

- [ ] 覆盖 >= 3 个不同类别
- [ ] Z类（无法分类）占比 < 30%
- [ ] 无单一类别占比 > 80%

#### 4.3.3 断点续传验证

```bash
# 1. 运行中按 Ctrl+C 中断
python main.py --max 500
# (等待爬取约100条后按 Ctrl+C)

# 2. 检查断点文件存在
Test-Path data/checkpoint.json  # Windows PowerShell
# ls data/checkpoint.json       # Linux

# 3. 从断点恢复运行
python main.py --max 500 --resume

# 4. 验证总数是否正确
```

- [ ] checkpoint.json 包含 total_scraped, category_counts, tag_cursor
- [ ] 恢复后从正确位置继续
- [ ] 无重复数据
- [ ] 总数正确
- [ ] **Excel数据不丢失** — 恢复前后行数一致（v2.0.1回归测试）

#### 4.3.4 Excel续写回归测试 (v2.0.1新增)

验证断点恢复时Excel数据不被覆盖：

```bash
# 1. 小批量爬取并中断
Remove-Item -Force data/session_cookies.json -ErrorAction SilentlyContinue
Remove-Item -Force data/seen_books.db -ErrorAction SilentlyContinue
Remove-Item -Force data/checkpoint.json -ErrorAction SilentlyContinue
Remove-Item -Force data/output/books_100k.xlsx -ErrorAction SilentlyContinue

python main.py --max 10
# (Ctrl+C 中断在第5条左右)

# 2. 记录当前Excel行数
python -c "
import openpyxl
wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
before = wb.active.max_row - 1
print(f'中断前行数: {before}')
with open('data/before_count.txt', 'w') as f:
    f.write(str(before))
"

# 3. 断点恢复
python main.py --max 10 --resume

# 4. 验证数据不丢失
python -c "
import openpyxl
wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
after = wb.active.max_row - 1
with open('data/before_count.txt') as f:
    before = int(f.read().strip())
assert after >= before, f'FAIL: 数据丢失! 中断前{before}行, 恢复后{after}行'
print(f'PASS: 中断前{before}行, 恢复后{after}行, 无数据丢失')
"
```

### 4.4 通过标准

- [ ] 成功率 >= 90%
- [ ] 分类覆盖 >= 3个类别
- [ ] 断点续传功能正常
- [ ] Z类占比 < 30%
- [ ] 无豆瓣封禁事件

### 4.5 实测记录

| 日期 | 目标 | 实际 | 成功率 | 耗时 | 类别数 | 断点恢复 | 状态 |
|------|------|------|--------|------|--------|----------|------|
| - | 500 | - | - | - | - | - | 待测试 |

---

## 5. 阶段4：正式运行

### 5.1 运行命令

```bash
# 正式运行（全量10万条）
python main.py

# 或指定其他目标数量
python main.py --max 100000

# 后台运行（Windows PowerShell）
Start-Process -NoNewWindow python -ArgumentList "main.py" -RedirectStandardOutput "logs/stdout.log" -RedirectStandardError "logs/stderr.log"

# 后台运行（Linux/Mac）
nohup python main.py > logs/stdout.log 2> logs/stderr.log &
```

### 5.2 运行监控

```bash
# 实时查看日志
Get-Content logs/scraper.log -Tail 20 -Wait   # Windows
# tail -f logs/scraper.log                     # Linux

# 查看当前进度
python -c "
import json
with open('data/checkpoint.json', 'r') as f:
    state = json.load(f)
print(f'已采集: {state[\"total_scraped\"]}')
print(f'更新时间: {state[\"last_updated\"]}')
for cat, count in sorted(state['category_counts'].items()):
    print(f'  {cat}: {count}')
"

# 查看Excel行数
python -c "
import openpyxl
wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
print(f'当前行数: {wb.active.max_row - 1}')
"
```

### 5.3 中间检查点

| 进度 | 检查项 |
|------|--------|
| 1,000条 | 验证分类分布，调整 tag_to_clc.json |
| 5,000条 | 检查封禁频率，必要时调大 REQUEST_DELAY_MIN |
| 25,000条 | 评估完成时间，决定是否需要优化 |
| 50,000条 | 全面数据质量审查 |
| 75,000条 | 检查冷门类别覆盖情况 |
| 100,000条 | 最终汇总统计 |

### 5.4 突发事件处理

#### 被封禁（连续418/302到sorry页）

```bash
# 1. Ctrl+C 停止
# 2. 等待 10-30 分钟冷却
# 3. 清除cookie重新开始
Remove-Item -Force data/session_cookies.json
python main.py --resume
```

#### 网络中断

```bash
# 直接重新运行（断点续传自动生效）
python main.py --resume
```

#### Excel文件过大

```bash
# 修改 main.py 中的文件名，分批输出
# OUTPUT_FILE = OUTPUT_DIR / "books_part1.xlsx"
```

### 5.5 最终验证

采集完成后，运行完整的数据质量报告：

```bash
python -c "
import openpyxl
from collections import Counter
import re

wb = openpyxl.load_workbook('data/output/books_100k.xlsx')
ws = wb.active

total = ws.max_row - 1
empty_title = 0
empty_desc = 0
empty_clc = 0
clc_counter = Counter()
conf_sum = 0.0
desc_lens = []

for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[1]: empty_title += 1
    if not row[6]: empty_desc += 1
    if not row[8]: empty_clc += 1
    if row[8]: clc_counter[row[8][0]] += 1
    if row[10]: conf_sum += float(row[10])
    if row[6]: desc_lens.append(len(str(row[6])))

print(f'=== 数据质量报告 ===')
print(f'总记录数: {total}')
print(f'书名完整率: {(total - empty_title) / total * 100:.1f}%')
print(f'简介完整率: {(total - empty_desc) / total * 100:.1f}%')
print(f'CLC完整率: {(total - empty_clc) / total * 100:.1f}%')
print(f'平均简介长度: {sum(desc_lens) / len(desc_lens):.0f} 字' if desc_lens else 'N/A')
print(f'平均置信度: {conf_sum / total:.2f}')
print()
print('=== 分类分布 ===')
for cat in sorted(clc_counter.keys()):
    count = clc_counter[cat]
    pct = count / total * 100
    bar = '#' * int(pct) + '.' * (20 - int(pct))
    print(f'  {cat}: {bar} {count} ({pct:.1f}%)')
print(f'类别数: {len(clc_counter)}')
"
```

### 5.6 通过标准

- [ ] 总记录数 >= 100,000
- [ ] 书名完整率 >= 99%
- [ ] CLC完整率 >= 99%
- [ ] 覆盖 >= 15 个中图法大类
- [ ] 平均置信度 >= 0.3
- [ ] Z类占比 < 20%

---

## 6. 常见问题排查

### Q1: 启动后所有请求都返回 418 或 302

**原因**: IP已被豆瓣临时封禁，或使用了被标记的Cookie。

**解决**:
```bash
Remove-Item -Force data/session_cookies.json
# 等待10-30分钟后重试
python main.py --max 20
```

### Q2: 解析到的书名/简介为空

**原因**: 豆瓣页面结构可能已变更。

**解决**:
```bash
# 调试单页
python -c "
from main import AntiScrapeSession, BookParser
s = AntiScrapeSession()
r = s.get('https://book.douban.com/subject/1007305/')  # 替换为问题ID
with open('debug.html', 'w', encoding='utf-8') as f:
    f.write(r.text)
data = BookParser.parse_book_page(r.text, '')
print('Title:', repr(data.get('title')))
print('Desc:', repr(data.get('description', '')[:100]))
"
# 检查 debug.html，更新 main.py 中的CSS选择器
```

### Q3: 标签页返回200但解析出0本书

**原因**: 该标签在豆瓣上确实没有书籍，或页面布局特殊。

**解决**: 不做处理（标签自动标记为bad_tags，后续跳过）。

### Q4: 内存占用持续增长

**原因**: ExcelWriter缓冲区可能在大量数据时增长。

**解决**: 减小 BATCH_SIZE（main.py 第68行），如改为200。

### Q5: 进度条显示乱码

**原因**: Windows终端使用GBK编码。

**解决**: 已使用ASCII字符（#和.）替代Unicode字符（█和░），不影响功能。

---

## 7. 测试记录汇总

| 阶段 | 日期 | 目标 | 结果 | 成功率 | 耗时 | 备注 |
|------|------|------|------|--------|------|------|
| 连通性 | 2026-05-19 | 1页 | 通过 | - | <10s | 标签页51604字节，20本书 |
| 小批量 | 2026-05-19 | 5条 | 通过 | 100% | 25s | 全A类 |
| 小批量 | 2026-05-19 | 20条 | 通过 | 100% | 97s | A(18), D(1), G(1) |
| 中批量 | - | 500条 | 待测 | - | - | - |
| 正式 | - | 10万条 | 待测 | - | - | - |
