---
name: 高迪行书竖排
description: 古籍书法类 Word 文档完整排版工作流。涵盖从文本提取（PDF/EPUB）到最终成品的全部步骤：文本清洗、格式设置（繁体一号42.6pt行距）、段落优化、字库字符检查、竖排标点、繁简转换、智能目录生成。当用户提到"高迪行书""书法排版""竖排排版""古籍排版"时触发此 skill。
---

# 高迪行书竖排

古籍/书法类 Word 文档的完整排版工作流，从原始文本到印刷级成品。

## 完整工作流总览

```
文本提取 → PUA检测 → 清洗切分 → 模板排版 → 段落优化 → 字库检查 → 繁简转换 → 竖排标点 → 智能目录
   Step 1     Step 2      Step 3      Step 4      Step 5      Step 6      Step 7       Step 8       Step 9
```

---

## 成品规格

基于 `黄帝内经_素问_高迪行书横版_繁体.docx` 的基础样本：

### 页面设置

| 属性 | 值 | XML |
|------|------|-----|
| 方向 | 横向（LANDSCAPE） | `w:orient="landscape"` |
| 纸张 | A4 横向 29.7 × 21.0 cm | `w:pgSz w:w="16838" w:h="11906"` |
| 边距 | 四边均 1.3 cm | `w:top/bottom/left/right="737"` |
| 页面边框 | 四边单线、深红色 `#953735`、sz=12、space=24 | `w:pgBorders` |
| 文字网格 | lines, linePitch=312, charSpace=0 | `w:docGrid` |

### 文首

| 元素 | 格式 | 对齐 |
|------|------|------|
| 段0：印章图片 | 含 DrawingML 行内图片（从模板提取或 COM 插入） | **左对齐** |
| 段1：大标题 | 书名 + 副标题，48pt，行距 240/auto | **左对齐** |
| 段2：空行 | | |
| 段3：目录标题 | "目　錄"，26pt | **左对齐** |
| 目录条目 | 智能目录 TOC 域（F9 更新） | |

### 正文格式

| 元素 | 字体 | 字号 | 行距 | 对齐 | XML 要点 |
|------|------|------|------|------|----------|
| 正文 | 高迪书法AI_行书V1（ascii/hAnsi/eastAsia 全设） | 26pt（w:sz=52） | 42.6pt 固定（852/exact） | 两端对齐（both） | 三种字体属性必须都设 |
| 章节标题 | 高迪书法AI_行书V1 | 26pt（w:sz=52） | 42.6pt 固定（852/exact） | 左对齐（left） | 颜色深红 `#953735`；`keepNext` 防孤行；篇中前空两行 |
| 译文标题 | 高迪书法AI_行书V1 | 16pt（w:sz=32） | 单倍（240/auto） | 左对齐（left） | "XX参考译文" |
| 译文正文 | 高迪书法AI_行书V1 | 16pt（w:sz=32） | 单倍（240/auto） | 两端对齐（both） | |

### 竖排标点格式

| 属性 | 值 |
|------|------|
| 字体 | 楷体（ascii/hAnsi/eastAsia 全设） |
| 字号 | 小四 12pt（w:sz=24） |
| 颜色 | 深红色 #8B0000 |
| 垂直定位 | **上标（w:vertAlign="superscript"）** — 关键！ |
| 替换规则 | 引号→『』、书名号→波浪下划线、括号→︵︶ |

### 文末

| 元素 | 格式 | 对齐 |
|------|------|------|
| 标记文字+印章 | "本文由高迪書法行書排版而成 2026" + 行内印章图片，14pt | **左对齐**，**文档最末行** |
| sectPr | 紧跟标记行之后，不能有空段 | |

---

## 段落间距与页面规则

### 段落类型识别

| 类型 | 识别条件 |
|------|----------|
| **章节标题** | 正文格式（26pt/852exact）+ 匹配 `.{1,25}第[一二三四五六七八九十百千]+$` |
| **正文段落** | 正文格式（26pt/852exact），非标题、非文末标记 |
| **译文标题** | 含"译文"或"譯文" |
| **译文段落** | 译文格式（16pt/240auto），非译文标题 |
| **文末标记** | 含"高迪書法"或"高迪书法" |

### 第一层：段落内部 — 同类段落间 0 空段

| 位置 | 空段数 |
|------|--------|
| 章节标题 → 正文段落 | 0 |
| 正文段落 ↔ 正文段落 | 0 |
| 译文标题 → 译文段落 | 0 |
| 译文段落 ↔ 译文段落 | 0 |

### 第二层：段落块边界 — 块前后 1 空段

| 位置 | 空段数 | 说明 |
|------|--------|------|
| 正文末段 → 译文标题 | 1 | 正文块与译文块之间 |
| 译文末段 → 下一内容 | 1 | 译文块结束后的间隔 |

### 第三层：章节标题间距 + 防孤行

| 位置 | 空段数 | 说明 |
|------|--------|------|
| 篇首标题前 | 0 | TOC 后第一个章节 |
| 篇中标题前 | 2 | 前补 2 空段与前文隔开 |
| 文末标记前 | 1 | 标记行前 1 空段 |

**章节标题必须设 `w:keepNext`**：标题不放在页面最后一行，必须与下一段正文同页。

**章节标题属性：** 颜色深红 `#953735`（与边框一致）；`keepNext` + `keepLines`；标题长度不超过一行。

### 第四层：页面首尾

| 位置 | 规则 |
|------|------|
| **首印章** | 文档第一段，**左对齐**，含 DrawingML 行内图片 |
| **书名标题** | 首印章后，**左对齐**，48pt |
| **目录标题** | "目　錄"，**左对齐**，26pt |
| **TOC 域** | 目录标题后，含 `TOC \f \h \z \* MERGEFORMAT`（Ctrl+A → F9 更新） |
| **篇首标题** | 必须新起一页，不能与目录同页（`w:pageBreakBefore`） |
| **文末标记+印章** | 文档最后一段，**左对齐**，14pt + 行内印章图片，紧跟 sectPr，后面不能有空段 |

### 排版示意

```
[首印章]                  ← 左对齐，DrawingML 行内图片
黃帝內經素問 原文及譯文     ← 左对齐，48pt
（空行）
目　錄                     ← 左对齐，26pt
[TOC 域占位]               ← Ctrl+A → F9 更新
┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄  ← 分页
[章节标题]                ← 篇首，前 0 空段
正文段落
正文段落
…
                          ← 1 空段（第二层）
[译文标题]
译文段落
译文段落
…
                          ← 1 空段（第二层）
                          ← 2 空段（第三层：篇中前置）
[章节标题]                ← 篇中
正文段落
…
                          ← 1 空段（第二层）
[译文标题]
译文段落
…
[文末标记+印章]           ← 最后一段，左对齐，含行内印章图片，紧跟 sectPr
```

### keepNext 副作用与近空白页

keepNext 防止标题出现在页面最后一行，但会导致副作用：当标题后紧接的正文段落很长（超过一页），keepNext 会把标题单独推到新页，造成近空白页。

**检测与修复（需 WPS COM）：**
1. 用 COM 扫描所有页面，找出字数 < 50 且仅含章节标题的页面
2. 对这些标题移除 keepNext
3. 再次验证无近空白页

**经验：** 素问 81 章中有 2 个标题（"五臟生成篇第十"、"刺禁論篇第五十二"）需要移除 keepNext。这是 python-docx 无法预测的，必须依赖渲染引擎分页结果。

### 从 txt 全流程重建（推荐）

增量修改 docx 容易引入累积错误（间距、域、印章等问题交叉影响）。推荐从纯文本重建：

1. **提取文本**：从现有 docx 提取正文为 txt，标记章节 `[CHAPTER]`、译文标题 `[TRANS_TITLE]`、译文 `[TRANS]`、文末 `[COLOPHON]`
2. **模板法构建**：复制成品 docx 为模板 → 清除正文段落（保留 sectPr + colophon+seal 段落）→ 按 txt 内容逐段插入到 colophon+seal 段落之前
3. **格式设置**：正文 26pt/852exact、标题加颜色/keepNext/TC 域、译文 16pt/240auto
4. **插入页面首尾**：书名（左对齐 48pt）+ 空段 + 目录标题 + TOC 域 → colophon 段落之前
5. **处理首印章**：模板中首印章可能是孤立图片（media 中存在但 body 未引用），用 COM `InlineShapes.AddPicture()` 插入到第一段
6. **运行竖排标点脚本**：如果源文本含 ASCII 直引号 `"`（非弯引号），需先替换为 `『』`，再运行标点脚本
7. **COM 后处理**：设置对齐（首印章/书名/目录/文末标记全部左对齐）、扫描近空白页、移除问题标题的 keepNext

### 修复工作流（按层递进，每层验证通过再处理下一层）

1. **第一层**：清除正文/译文内部、标题后的多余空段 → 验证 0 违规
2. **第二层**：确保正文→译文标题前 1 空段，译文末段后 1 空段 → 验证
3. **第三层**：篇首 0 空段、篇中 2 空段、文末标记前 1 空段、所有标题加 keepNext → 验证
4. **第四层**：篇首标题加 pageBreakBefore，文末标记压到最末尾 → 验证
5. **COM 后处理**：设置对齐（左对齐）、扫描近空白页、移除问题标题的 keepNext → 验证

---

## COM 分页检测

python-docx 无法判断分页（页码由 Word/WPS 渲染引擎计算，不在 XML 中）。必须用 COM 获取可靠结果。

### 核心方法

```python
import win32com.client
import time

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
doc = word.Documents.Open(fname)
time.sleep(5)  # 等待分页完成

# Word 常量
wdActiveEndPageNumber = 3          # 段落所在页号
wdVerticalPositionRelativeToPage = 6  # 段落在页面中的 Y 坐标（pt）

# 总页数
total_pages = doc.ComputeStatistics(2)  # wdStatisticPages

# 页面可用高度
ps = doc.Sections(1).PageSetup
usable_h = ps.PageHeight - ps.TopMargin - ps.BottomMargin

# 段落页号和 Y 坐标
para = doc.Paragraphs(i)
page = para.Range.Information(3)   # 页号
y_pos = para.Range.Information(6)  # Y 坐标（pt）
gap_to_bottom = usable_h - y_pos   # 距页面底部距离

# 设置段前间距（将段落推到页面底部）
para.Format.SpaceBefore = gap_value  # 单位：pt
```

### 典型应用

| 场景 | 方法 |
|------|------|
| 扫描空白页 | 逐段检查页号，找出某页仅含标题的近空白页 |
| 定位文末标记 | 检查最后一段的页号和 Y 坐标，用 SpaceBefore 推到底部 |
| 验证 keepNext 副作用 | 检查每个标题的下一页是否有被推走的正文 |
| 检查分页效果 | 修改后重新检查页号和 Y 坐标 |

### 注意事项

- **打开文件后必须等待**：`time.sleep(5)` 让 WPS 完成分页
- **WPS COM 的 `ParagraphFormat` 属性名不同**：用 `para.Format.SpaceBefore` 代替 `para.ParagraphFormat.SpaceBefore`
- **SpaceBefore 二分法**：设过大可能推到下一页，需二分查找最大不翻页值
- **COM 对象可能返回错误类型**：如果 `Documents.Open` 返回异常，先 `taskkill /F /IM wps.exe` 清理残留进程
- **`Information(3)` 不是 `Information(1)`**：常量 1 是 `wdActiveEndAdjustedPageNumber`（受页码起始值影响），3 是 `wdActiveEndPageNumber`（物理页号）

### 印章处理

模板中的印章有三种可能情况：

| 情况 | 识别方法 | 处理 |
|------|----------|------|
| 独立 DrawingML 段落 | body 中含 `w:drawing` 的段落 | python-docx 清理时保留该段落 |
| 行内图片（colophon+seal） | 文末标记文字+图片同一段落 | 保留该段落，不在 txt 重建时重复创建 |
| 孤立图片（media 中存在但未引用） | rels 中有 image 关系但 body 中无 `r:embed` | 从模板 media 提取图片文件，用 COM `InlineShapes.AddPicture()` 插入 |

**重要：python-docx 的 `doc.save()` 只保留被引用的图片。孤立图片会丢失。** 必须在 python-docx 操作之前从原始模板提取图片文件。

### ASCII 引号处理

竖排标点脚本的 QUOTE_MAP 只处理弯引号（`""` U+201C/201D）。如果源文本含 ASCII 直引号 `"`（U+0022），脚本不会替换。

**处理方法：**

```python
# 在运行竖排标点脚本之前，替换 ASCII 引号为角括号
for p in doc.paragraphs:
    for run in p.runs:
        if '"' in run.text:
            new_text = []
            is_open = True
            for ch in run.text:
                if ch == '"':
                    new_text.append('『' if is_open else '』')
                    is_open = not is_open
                else:
                    new_text.append(ch)
            run.text = ''.join(new_text)
```

然后再运行竖排标点脚本，角括号 `『』` 会被正确识别并设置格式。

---

## Step 1 — 文本提取

从 PDF 或 EPUB 获取原始文本。

### EPUB 提取

```bash
# 方法一：calibre 转换
ebook-convert input.epub output.txt

# 方法二：解压提取 HTML
unzip input.epub -d epub_extracted
# 用 BeautifulSoup 清理 HTML，提取纯文本
```

EPUB 提取的文本通常质量较好，主要问题：
- 注释/脚注混入正文
- 章节标记不统一（需人工确认）
- **多章合并**：EPUB 可能把两篇合并为一个 HTML 文件（如史记"季布、栾布列传—袁盎晁错列传"），需检查内部分标题

**EPUB 提取脚本结构：**

```python
import zipfile
from bs4 import BeautifulSoup

with zipfile.ZipFile(epub_path) as z:
    # 1. 读 OPF 获取 spine 顺序
    opf = BeautifulSoup(z.open('content.opf').read(), 'xml')
    manifest = {item['id']: item['href'] for item in opf.find_all('item')}
    spine = [ref['idref'] for ref in opf.find('spine').find_all('itemref')]

    # 2. 按 spine 顺序提取各章
    for iid in spine:
        html = BeautifulSoup(z.open(manifest[iid]).read(), 'html.parser')
        title = html.find('h3').get_text(strip=True)
        paragraphs = [p.get_text(strip=True) for p in html.find_all('p')]
```

**EPUB 注意事项：**
- 标题常带"正文 "前缀，需用 `re.sub(r'^正文\s*', '', title)` 清理
- 分隔线 `"---"` 会被提取为段落，需过滤
- spine 第一个通常是封面（titlepage），跳过
- 史记等大型古籍通常 100+ 章节，每章一个 HTML 文件

### PDF 提取

**脚本工具：**

```bash
python <skill-dir>/scripts/extract_pdf.py <PDF路径> [输出JSON] [--start PAGE] [--end PAGE]
```

提取逻辑：
- 识别章节标题（含"篇第X"或"XX第X"的短行）
- **两种 PDF 结构**：
  - **素问型**：原文与译文用空行分隔，需用现代口语特征检测区分
  - **灵枢型**：用【原文】/【白话解】标记明确区分，状态机解析
- 按对话标记分段（帝曰/岐伯曰等切换处分段）
- 输出 JSON：`[{title, original: [...], translation: [...]}]`

**关键 — 必须用状态机提取，不能用嵌套循环：**

```python
state = 'skip'  # skip | original | translation
for line in all_lines:
    if is_title_line(line):    # 优先级最高
        save_current_chapter()
        start_new_chapter()
        state = 'skip'
    elif line.startswith('【原文】'):
        state = 'original'
    elif line.startswith('【白话解】'):
        state = 'translation'
    elif state == 'original':
        current_ch['original'].append(line)
    elif state == 'translation':
        current_ch['translation'].append(line)
```

**章节标题变体：**

| 格式 | 示例 | 正则 |
|------|------|------|
| XX篇第X | 上古天真論篇第一 | `(.{2,10}篇第[一二三...]+)` |
| XX第X（无篇字） | 九针十二原第一 | `(.{2,15}第[一二三...]+)$)` |
| XX体裁名 | 五帝本紀、刺客列傳 | `(.{1,20}(本紀|世家|列傳|書|表|自序|年表|月表)$)` |

**PDF 字号过滤脚注（关键技巧）：**

古籍 PDF 中正文、脚注标记、脚注内容使用不同字号和字体。PyMuPDF 提取时可用 `span['size']` 精确区分：

| 内容类型 | 字号 | 字体 | 处理 |
|----------|------|------|------|
| 正文 | 18pt | 宋体 | 保留 |
| 卷/章标题 | 28-30pt | 宋体（居中） | 提取为标题 |
| 脚注标记（上标数字） | 10.5pt | 宋体 | **丢弃** |
| 脚注内容 | 12pt | 楷体 | **丢弃** |

```python
MIN_SIZE = 15.0  # 只保留 >= 15pt 的内容
for span in line['spans']:
    if span['size'] < MIN_SIZE:
        continue  # 跳过脚注标记和脚注内容
    text = span['text'].strip()
    if text:
        main_text += text
```

**为什么用字号过滤而不是位置过滤：**
- 位置过滤（y0 > page_height * 0.75）在某些页面会误判（长脚注占据大面积）
- 段落检测（`\d+[\s　]` 模式）无法区分正文中的合法数字
- **字号是最可靠的特征**：脚注标记（10.5pt）和脚注内容（12pt，楷体）与正文（18pt，宋体）字号差距显著
- 过滤后用 `re.search(r'[0-9a-zA-Z]', text)` 验证：古典文本不应含阿拉伯数字或英文字母

**验证规则：** 古典文本不应含阿拉伯数字或英文字母。提取后扫描所有段落，如有残留说明过滤不彻底。

**偏旁部首合并（生僻字处理）：**

古籍 PDF 中生僻字常被拆分为部首+部件形式，如 `(月真)→䐜`：

```python
DECOMPOSED_MAP = {
    '月真': '䐜', '雩重': '霳', '疒肙': '痋', '疒颓': '㿗',
    '月刍': '皱', '火台': '炱', '骨行': '骭', '亻亦': '俴',
    '火矣': '焠', '月囷': '腃', '骨盾': '骹', '月吕': '脔',
    '忄农': '憹', '出页': '颛', '石历': '砬',
}
```

**OCR 常见错误（ABBYY/Tesseract）：**

| 错误类型 | 示例 | 修正 |
|----------|------|------|
| 形近字混淆 | 己/已/巳、戊/戍/戎、曰/日、王/玉、未/末 | 逐字校对 |
| 标点丢失/错位 | 句末无句号、逗号变为空格 | 对照原文 |
| 繁简混杂 | 同一段落内繁简体交替 | 统一转换 |
| 段落断裂 | 一句话被拆成多段 | 后续合并处理 |
| 生僻字拆分 | (月真)→䐜 | 部首映射替换 |

**PDF vs DOCX 校对：**

```bash
python <skill-dir>/scripts/pdf_proofread.py <PDF路径> <DOCX路径> [--skip N] [--no-punct]
```

---

## Step 2 — PUA 字符检测与替换

**PUA（Private Use Area, U+E000-U+F8FF）** 字符是某些 EPUB 制作工具不规范编码产生的"私用区"字符，在标准字体中无法显示（显示为豆腐块/方框）。**所有 PUA 字符均来源于 EPUB 源文件，PDF 不含 PUA。**

### 检测方法

对 docx 文件执行 PUA 扫描（必须用 lxml 直接解析 XML，python-docx 的 `paragraph.text` 不含页眉页脚文本框等内容）：

```python
from lxml import etree
import zipfile

def detect_pua_chars(docx_path):
    """扫描 docx 中所有 XML 文件，检测 PUA 字符 (U+E000-U+F8FF)"""
    pua_chars = {}  # {pua_codepoint: count}
    pua_contexts = []  # [(codepoint, context_text)]

    with zipfile.ZipFile(docx_path) as z:
        for xml_name in z.namelist():
            if not xml_name.endswith('.xml'):
                continue
            xml_bytes = z.read(xml_name)
            try:
                tree = etree.fromstring(xml_bytes)
            except:
                continue
            # 提取所有文本节点
            for elem in tree.iter():
                if elem.text:
                    for ch in elem.text:
                        cp = ord(ch)
                        if 0xE000 <= cp <= 0xF8FF:
                            pua_chars[cp] = pua_chars.get(cp, 0) + 1
                if elem.tail:
                    for ch in elem.tail:
                        cp = ord(ch)
                        if 0xE000 <= cp <= 0xF8FF:
                            pua_chars[cp] = pua_chars.get(cp, 0) + 1

    return pua_chars
```

### 映射方法：语句匹配

PUA 字符的上下文通常是完整的古文语句，可以用来匹配 PDF 繁体原文确认对应的标准 Unicode 字符。

**映射流程：**

1. 从 docx XML 中提取每个 PUA 字符的上下文语句（前后各 10-20 字）
2. 将上下文语句转为繁体（opencc s2t），在 PDF 原文中搜索匹配
3. PDF 对应位置的字符即为 PUA 字符应映射的标准 Unicode
4. 高可信映射（PDF 直接确认）：直接替换
5. 低可信映射（推断）：列出让用户确认

```python
import opencc

def map_pua_via_context(docx_path, pdf_path):
    """通过语句上下文匹配 PDF 确认 PUA 映射"""
    converter = opencc.OpenCC('t2s')  # 繁→简，用于匹配

    # 1. 从 docx 提取 PUA 字符及上下文
    pua_contexts = extract_pua_contexts(docx_path)  # [(pua_cp, before_text, after_text)]

    # 2. 从 PDF 提取文本（需处理 CJK 兼容偏旁问题）
    pdf_text = extract_pdf_text(pdf_path)

    # 3. 逐个匹配
    mappings = {}
    for pua_cp, before, after in pua_contexts:
        # 构建搜索模式：前文 + ? + 后文（至少各 5 字）
        search_before = before[-10:] if len(before) >= 10 else before
        search_after = after[:10] if len(after) >= 10 else after
        # 在 PDF 文本中搜索，? 处的字符即为映射目标
        match = find_in_pdf(pdf_text, search_before, search_after)
        if match:
            mappings[pua_cp] = ord(match)
    return mappings
```

### 替换方法

确认映射后，直接在 docx 的 XML 中替换 PUA 字符：

```python
def replace_pua_in_docx(docx_path, pua_to_unicode, output_path):
    """在 docx XML 中替换 PUA 字符为标准 Unicode"""
    import shutil, zipfile

    pua_map = {chr(k): chr(v) for k, v in pua_to_unicode.items()}
    shutil.copy2(docx_path, output_path)

    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml'):
                    text = data.decode('utf-8')
                    for pua_char, std_char in pua_map.items():
                        text = text.replace(pua_char, std_char)
                    data = text.encode('utf-8')
                zout.writestr(item, data)
```

### 实际案例（史记 EPUB）

史记 EPUB 来源的 docx 中发现 39 个 PUA 字符，共 131 处使用。映射表示例：

| PUA | 标准 Unicode | 字 | 上下文 |
|-----|-------------|---|--------|
| U+E837 | U+7D2C | 紬 | 朕唯未能循明也，紬績日分 |
| U+E844 | U+8A22 | 訢 | 訢載武王 |
| U+E0F4 | U+84FA | 蓺 | 蓺五種 |
| U+E41D | U+8CB2 | 貲 | 攻朝貲塞 |
| U+E75C | U+882D | 蠭 | 豪傑蠭起 |

**关键注意事项：**
- PUA 扫描必须覆盖所有 XML 文件（document.xml、header/footer、textboxes 等）
- python-docx 的 `paragraph.text` 只读取正文段落，会遗漏页眉页脚中的 PUA
- 映射时优先用繁体 PDF 确认，繁体版是权威标准
- EPUB 源文件本身含有 PUA，不能作为映射参考
- PDF 文本提取会使用 CJK 兼容偏旁（如 ⽇ U+2F47 代替 日 U+65E5），匹配时需注意

### 工作流位置

**PUA 检测必须在文本提取（Step 1）之后、清洗切分（Step 3）之前执行。** 原因：
- 文本提取后才能扫描完整内容
- 清洗切分前替换 PUA，避免 PUA 字符干扰后续的段落切分和字库检查
- 如果用 PDF 作为文本源则不需要此步骤（PDF 不含 PUA）

---

## Step 3 — 文本清洗与切分

```python
CHAPTER = re.compile(r'(?=.*篇)(?=.*第).{4,29}$')
DIALOG = re.compile(r'^(帝曰|黃帝曰|岐伯曰|岐伯對曰|雷公曰|黃帝問曰|黃帝問)')
TRANSLATION = re.compile(r'參考譯文|参考译文')
```

清洗步骤：
1. 去除多余空白（连续空格→单空格，行首行尾空白）
2. 统一标点风格（全角/半角）
3. 按篇章切分：识别章节标题行作为分隔点
4. 区分正文与译文区域
5. 用 python-docx 生成结构化 .docx 文件
6. **验证：扫描所有段落，确认无阿拉伯数字和英文字母**（古典文本不应含 `[0-9a-zA-Z]`）

**古典文本验证规则：**
```python
import re
contaminated = [p for p in paragraphs if re.search(r'[0-9a-zA-Z]', p)]
if contaminated:
    print(f'警告: {len(contaminated)} 段含数字/英文，可能有脚注残留')
```

---

## Step 4 — 模板法排版

**核心原则：复制已有成品 docx 作为模板，只替换正文内容。**

手动设置页面方向、竖排、边框、印章、页脚等格式极易遗漏，模板法保留所有格式设置。

### 操作步骤

```python
import shutil
from docx import Document
from docx.oxml.ns import qn

# 1. 复制模板
shutil.copy2(TEMPLATE_DOCX, OUTPUT_DOCX)
doc = Document(OUTPUT_DOCX)

# 2. 删除模板正文段落（保留 sectPr）
body = doc.element.body
sectPr = body.find(qn('w:sectPr'))
to_remove = [el for el in body if el is not sectPr]
for el in to_remove:
    body.remove(el)

# 3. 插入新内容到 sectPr 之前
body.insert(list(body).index(sectPr), new_para)

# 4. 保存
doc.save(OUTPUT_DOCX)
```

### 印章图片恢复

python-docx 删除段落后需用 zipfile 恢复含 DrawingML 的首尾印章段落：

```python
import zipfile, re

with zipfile.ZipFile(TEMPLATE, 'r') as z:
    template_xml = z.read('word/document.xml').decode('utf-8')

# 提取首尾印章段落
paras = re.findall(r'<w:p[^>]*>.*?</w:p>', template_xml, re.DOTALL)
seal_first, seal_last = paras[0], paras[-1]

# 替换 rId 引用，插入到目标 XML
```

**关键：印章段落的 rId 必须与 rels 中的图片关系一致。**

### 关键经验 — lineRule 必须为 exact

Word 的 `w:spacing` 元素：
- `w:line="852"` `w:lineRule="auto"` → 852/240 = 3.55 倍行距（**页数暴增！**）
- `w:line="852"` `w:lineRule="exact"` → 固定 42.6pt（正确）

---

## Step 5 — 段落优化

PDF 提取的文本通常段落过于碎片化。需要两轮合并。

### 第一轮：合并非对话起首的连续段落

```python
DIALOG = re.compile(r'^(帝曰|黃帝曰|岐伯曰|岐伯對曰|雷公曰|黃帝問曰|黃帝問)')
```

规则：
- 连续的非空段落，如果不是以对话标记开头，合并到前一段
- 空段落作为分隔符，结束当前合并组
- **译文部分不参与合并**

### 第二轮：合并问答对

```python
QUESTION = re.compile(r'^(帝曰|黃帝曰|黃帝問曰|黃帝問)')
ANSWER = re.compile(r'^(岐伯曰|岐伯對曰|雷公曰)')
```

- "帝曰..." + 紧接的"岐伯曰..." → 合并为一段
- 从后往前合并，避免索引偏移

### 注意事项

- 合并后检查译文前是否有空行，如缺失需手动插入空段落
- 插入空段落：`empty_p = OxmlElement('w:p')` → `elem.addprevious(empty_p)`

---

## Step 6 — 字库字符检查

检查文档用字是否在目标字体（.ttf/.otf）中存在。

```bash
python <skill-dir>/scripts/check_missing_chars.py <DOCX路径> <字体文件路径> [输出目录] [--chunk 200]
```

功能：
- 从 docx 提取所有汉字（支持 CJK 全区段：基本区、扩展A-F、兼容区）
- 与字库 cmap 比对，找出缺失字符
- 按出现频率排序输出
- 自动分组写入文件 + Unicode 编码列表

处理缺失字符的方式：
1. **补充字库**：在 FontLab 中将缺失字符添加到字体文件
2. **替代字**：用字形相近且字库已有的字符替换
3. **标注跳过**：记录缺失字符，排版时手动处理

---

## Step 7 — 繁简转换

```bash
python <skill-dir>/scripts/s2twp.py <输入文件> [--mode auto]
```

- 将简体中文转为繁体（使用 opencc s2t 模式）
- 正确处理一对多映射（发→發/髮、干→幹/乾、面→面/麵 等）
- 古典字形保留（裏、喫、穀）
- 当原文为繁体时，只改一对多部分，智能判断语义

---

## Step 8 — 竖排标点

```bash
python <skill-dir>/scripts/fix_punctuation.py <输入文件>
```

### 处理规则

| 原字符 | 替换为 | 位置 | 说明 |
|--------|--------|------|------|
| " " (双引号) | 『 』 | 上标/下标 | 角括号 |
| ' ' (单引号) | 「 」 | 上标/下标 | 角括号 |
| 《》 (书名号) | 删除《》，对书名加波浪下划线 | `w:u w:val="wave"` | 竖排时波浪线自动显示在文字左侧（GB/T 15834-2011） |
| 〈〉 (单书名号) | 删除〈〉，对书名加波浪下划线 | `w:u w:val="wave"` | 同上 |
| （） (括号) | ︵ ︶ | 不调整 | 直式括号 |
| ，。！？；：、…·～ | 不变 | 上标 | 缩小淡化 |
| - – — (连字符) | 不变 | 不调整 | |

**竖排标点格式：** 楷体 12pt、深红 #8B0000、上标定位。圈号数字 ①②③ 同样设为楷体深红但不调整上下标。

---

## Step 9 — 智能目录

### 前提条件

- 正文标题必须是独立段落
- 标题需含可识别的章节模式（"篇"+"第"、体裁名等）
- 每个章节标题唯一（不重复）

### 技术方案：纯 XML 插入（不用 COM）

**重大教训：WPS 的 COM `Fields.Add()` 会把 TC/TOC 域变成 `DOCVARIABLE`，完全失效。必须用 OxmlElement 构造域 XML。**

```python
def insert_tc_field(para_element, title_text):
    """在段落末尾插入 TC 域"""
    # begin → instrText → end（三段式）
    instr.text = r' TC "' + title_text + r'" \f \l 1 \* MERGEFORMAT '

def insert_toc_field(after_para_element):
    """在"目　錄"段落之后插入 TOC 域"""
    # begin → instrText → separate → 占位文字 → end（五段式）
    instr.text = r' TOC \f \h \z \* MERGEFORMAT '
```

**instrText 中的反斜杠必须用 raw string**：`\f` `\l` 在 Python 中是转义序列，会报 `ValueError: All strings must be XML compatible`。必须用 `r'...'`。

### 域格式对照

| 元素 | 格式 |
|------|------|
| TC instrText | ` TC "上古天真論篇第一" \f \l 1 \* MERGEFORMAT ` |
| TOC instrText | ` TOC \f \h \z \* MERGEFORMAT ` |
| TC 域结构 | begin → instrText → end（三段式） |
| TOC 域结构 | begin → instrText → separate → 内容 → end（五段式） |

### 安全规则（必读）

1. **用 `\t\d+` 区分 TOC 条目和正文标题** — TOC 条目含 tab+页码，正文标题不含
2. **用元素引用代替索引** — `p._element` 而非 `doc.paragraphs[idx]`，避免插入/删除导致索引偏移
3. **清除域操作必须限定正文区域** — `body.findall('.//w:fldChar')` 会误删 TOC 域的 end 标记，导致目录永久失效
4. **修改颜色前先清除旧色** — 防止残留
5. **先打印诊断再执行** — 显示目标数量、首条/末条，确认后再修改

---

## 文本审核（排版后检查）

### 删除脚注标记 `[N]`

PDF 原文含脚注 `[1]` `[2]`，提取后注释内容已丢失但标记残留。

**难点：`[N]` 在 Word XML 中被拆成三个 run**（`[` + 数字 + `]`），有时还与相邻标点合并（`"。["`）。

**策略 — 两轮法：**
1. 第一轮：匹配三个连续 run，直接删除（处理 ~95%）
2. 第二轮：段落级文本定位法，映射回 run 区间裁剪（处理残留）

### 清除混入的【提要】【题解】【注释】

```python
CONTAMINATION = re.compile(r'【(提要|题解|注释)】')
```

检测后保留标记之前的正文，删除标记及之后内容。

### 审核清单

| 检查项 | 方法 |
|--------|------|
| `[N]` 脚注残留 | `re.search(r'\[\d+\]', p.text)` |
| 【提要/题解/注释】混入 | `re.search(r'【(提要|题解|注释)】', p.text)` |
| 译文格式错误 | 检查行距 852/exact 段落的现代语特征 |
| 章节数量 | 统计标题正则匹配数 |

---

## 常见问题

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| 页数暴增（700+页） | lineRule=auto 使行距变为 3.55 倍 | 改为 exact |
| TOC 条目重复/三倍 | 旧 TC 域未清除就插入新域 | 先删所有域再重建 |
| WPS COM 域变 DOCVARIABLE | WPS `Fields.Add()` 行为与 Word 不同 | 用纯 XML（OxmlElement）插入域 |
| instrText 转义报错 | `\f` `\l` 被当 Python 转义字符 | 用 `r'...'` raw string |
| 格式丢失（边框/印章/页码） | 从零创建格式时遗漏 | 用模板法 |
| 印章图片丢失 | 删除段落时把含 DrawingML 的段也删了 | zipfile 恢复印章段落 |
| 正文内容被误删 | 删段落后用旧索引操作 | 重新收集索引 |
| **TOC 域 end 标记被清除** | `findall('.//w:fldChar')` 全文档扫描误删 TOC end | 清除域必须限定正文区域 |
| TOC 条目被误当标题 | 标题正则匹配了含章节名的 TOC 条目 | 用 `\t\d+` 排除 TOC 条目 |
| 索引偏移导致正文被改红 | 插入空段后 `doc.paragraphs[idx]` 指向错误 | 用 `p._element` 元素引用 |
| 篇首出现多余空行 | TOC 更新后首篇前残留空段 | 篇首禁止空行 |
| 标题长度判断过宽 | `len > 25` 无法排除假标题 | 用 `页面宽度 ÷ 字号` 计算最大字数 |
| 译文前空行消失 | 段落合并吞掉空段 | 合并后手动插入空段落 |
| 译文以正文格式排版 | is_likely_modern() 漏判 | 检查 852/exact 段落内容 |
| EPUB 多章合并 | 两篇合一个 HTML | 检查内部分标题 |
| **PUA 字符（豆腐块）** | EPUB 源文件使用私用区编码 U+E000-U+F8FF | Step 2 检测并通过语句匹配 PDF 映射为标准 Unicode |
| PUA 映射无法确认 | PDF 文本提取含兼容偏旁（如 ⽇≠日） | 放宽匹配条件，或人工确认 |
| 竖排标点误转目录占位 | 括号被转为︵︶ | 标点处理前标记目录段落 |
| 弯引号未被替换 | QUOTE_MAP 写了 ASCII 直引号 | 用 `\u201c`/`\u201d` 转义 |
| 首印章丢失 | 模板中孤立图片（media 存在但未引用） | python-docx 操作前提取图片，用 COM 插入 |
| 对齐全变居中 | 误用 `jc=center` | 首印章/书名/目录/文末标记全部 **左对齐** |
| COM 返回错误对象类型 | 残留 WPS 进程 | `taskkill /F /IM wps.exe` 后重试 |
| python-docx 丢失图片 | 未引用的图片在 save() 时被清除 | 操作前从模板 zipfile 提取图片文件 |
| ASCII 引号未被竖排标点脚本处理 | 脚本只映射弯引号 U+201C/201D | 先用脚本将 ASCII `"` 替换为 `『』`，再运行标点脚本 |
| WPS COM 无 ParagraphFormat | WPS 接口与 Word 不同 | 用 `para.Format.SpaceBefore` 代替 |
| 无法判断分页 | python-docx 无布局引擎 | 用 COM: `Information(3)` 页号, `Information(6)` Y坐标 |

---

## 依赖

- Python 3 + python-docx + lxml + opencc + fontTools
- pywin32（仅 COM 方案需要，纯 XML 方案不需要）
- WPS 或 Microsoft Word（最终打开验证用）
- calibre（EPUB 转换，可选）

## 子 skill 索引

| Skill | 用途 | 命令 |
|-------|------|------|
| [高迪繁简转换](../高迪繁简转换/SKILL.md) | 简→繁转换 | `/高迪繁简转换` |
| [高迪书法竖排标点](../高迪书法竖排标点/SKILL.md) | 竖排标点替换 | `/高迪书法竖排标点` |
| [高迪智能目录](../高迪智能目录/SKILL.md) | 智能目录生成 | `/高迪智能目录` |

## 本 skill 脚本

| 脚本 | 用途 |
|------|------|
| `scripts/extract_pdf.py` | 从"原文+白话文"PDF 提取章节、原文、译文 |
| `scripts/pdf_proofread.py` | PDF vs DOCX 逐章校对 |
| `scripts/check_missing_chars.py` | 比较文档字符与字库覆盖范围 |
