---
name: 高迪行书竖排
description: 古籍书法类 Word 文档完整排版工作流。涵盖从文本提取（PDF/EPUB）到最终成品的全部步骤：文本清洗、格式设置（繁体一号42.6pt行距）、段落优化、字库字符检查、竖排标点、繁简转换、智能目录生成。当用户提到"高迪行书""书法排版""竖排排版""古籍排版"时触发此 skill。
---

# 高迪行书竖排

古籍/书法类 Word 文档的完整排版工作流，从原始文本到印刷级成品。

## 完整工作流总览

```
文本提取 → 清洗切分 → 模板排版 → 段落优化 → 字库检查 → 繁简转换 → 竖排标点 → 智能目录
   Step 1      Step 2      Step 3      Step 4      Step 5      Step 6       Step 7       Step 8
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

| 元素 | 格式 |
|------|------|
| 段0：印章图片 | 居中（含 DrawingML 图片） |
| 段1：大标题 | 书名 + 副标题，48pt，行距 240/auto，默认对齐 |
| 段2：空行 | 行距 852/exact |
| 段3：目录标题 | "目　錄"，26pt，居中 |
| 目录条目 | 智能目录 TOC 域（F9 更新） |

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

| 元素 | 格式 |
|------|------|
| 标记文字 | "本文由高迪書法行書排版而成 2026"，14pt，居中 |
| 印章图片 | 居中（含 DrawingML 图片） |

---

## 章节标题规则

| 属性 | 值 |
|------|-----|
| 颜色 | 深红 `#953735`（与页面边框一致） |
| 防孤行 | `w:keepNext` + `w:keepLines`（标题不单独留在页底） |
| 篇首标题 | **禁止空行** — 首篇（TOC 后第一个章节）前不得有空段 |
| 篇中标题 | 前面补空段至 2 个空行（与前文内容隔开） |
| 长度限制 | 标题不超过一行：用 `页面可用宽度 ÷ 字号` 计算最大字数 |

**判断篇首/篇中：** 检查标题前一段是否为空段 — 空段说明是篇首，非空段说明是篇中。

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

## Step 2 — 文本清洗与切分

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

---

## Step 3 — 模板法排版

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

## Step 4 — 段落优化

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

## Step 5 — 字库字符检查

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

## Step 6 — 繁简转换

```bash
python <skill-dir>/scripts/s2twp.py <输入文件> [--mode auto]
```

- 将简体中文转为繁体（使用 opencc s2t 模式）
- 正确处理一对多映射（发→發/髮、干→幹/乾、面→面/麵 等）
- 古典字形保留（裏、喫、穀）
- 当原文为繁体时，只改一对多部分，智能判断语义

---

## Step 7 — 竖排标点

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

## Step 8 — 智能目录

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
| 竖排标点误转目录占位 | 括号被转为︵︶ | 标点处理前标记目录段落 |
| 弯引号未被替换 | QUOTE_MAP 写了 ASCII 直引号 | 用 `\u201c`/`\u201d` 转义 |

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
