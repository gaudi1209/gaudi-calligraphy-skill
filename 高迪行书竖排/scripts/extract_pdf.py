"""
从 PDF 提取古籍原文和译文。
适用于"原文+白话文"对照排版的古籍 PDF。
按章节分组，区分原文和现代译文。

用法:
    python extract_pdf.py <PDF路径> [输出JSON路径] [--start PAGE] [--end PAGE]
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import re
import json
import argparse

try:
    import fitz  # PyMuPDF
except ImportError:
    print('需要安装 PyMuPDF: pip install pymupdf')
    sys.exit(1)


# 默认章节标题模式（含"篇第X"）
CHAPTER_PATTERN = re.compile(r'(.{2,10}篇第[一二三四五六七八九十百零]+)')

# 主要对话起始标记
DIALOG_START = re.compile(
    r'^(帝曰|黃帝曰|黄帝曰|岐伯曰|岐伯對曰|雷公曰|'
    r'黃帝問曰|黃帝問)'
)


def replace_decomposed(text, decomposed_map=None, skip_content=None):
    """替换括号中的拆分部首为完整 Unicode 字符。

    许多古籍 PDF 中，生僻字被拆分为部首+部件的形式，如(月真)→䐜。
    """
    if decomposed_map is None:
        decomposed_map = {}
    if skip_content is None:
        skip_content = set()

    unknown = []

    def repl(m):
        c = m.group(1)
        if c in skip_content:
            return m.group(0)
        if c in decomposed_map:
            return decomposed_map[c]
        if len(c) <= 2 and all(0x4e00 <= ord(x) <= 0x9fff for x in c):
            unknown.append(c)
        return m.group(0)

    result = re.sub(r'[(\uff08]([^)\uff09]{1,6})[)\uff09]', repl, text)
    return result, unknown


def is_title_line(line, pattern=None):
    """检测章节标题行"""
    if pattern is None:
        pattern = CHAPTER_PATTERN
    m = pattern.search(line)
    return m and len(line.strip()) < 30


def is_likely_modern(text):
    """检测文本是否像现代白话文"""
    if not text:
        return False
    modern_words = ['从前的', '十分聪明', '也就是说', '意思是', '指的是',
                    '相当于', '怎么', '什么原因', '这样做', '的', '了',
                    '黄帝说', '歧伯说', '回答', '问到', '这时候',
                    '这些人', '那些', '主要是因为', '这是', '为什么']
    count = sum(1 for w in modern_words if w in text)
    de_ratio = text.count('的') / max(len(text), 1)
    return count >= 2 or de_ratio > 0.04


def merge_lines(lines, mode='original'):
    """将 PDF 提取的短行合并为段落。

    mode='original': 在对话切换处分段（帝曰/岐伯曰等）
    mode='translation': 空行处分段
    """
    paragraphs = []
    current = ''
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if current:
                paragraphs.append(current)
                current = ''
            continue

        if not current:
            current = stripped
        elif mode == 'original' and DIALOG_START.match(stripped):
            paragraphs.append(current)
            current = stripped
        else:
            current += stripped

    if current:
        paragraphs.append(current)
    return paragraphs


def extract(pdf_path, start_page=0, end_page=None,
            decomposed_map=None, skip_content=None):
    """从 PDF 提取章节，返回 [{title, original, translation}, ...]"""
    doc = fitz.open(pdf_path)
    if end_page is None:
        end_page = len(doc)

    all_lines = []
    all_unknown = []

    for pn in range(start_page, end_page):
        page = doc[pn]
        text = page.get_text()

        if decomposed_map:
            text, unk = replace_decomposed(text, decomposed_map, skip_content)
            all_unknown.extend(unk)

        lines = text.split('\n')
        page_lines = []

        for line in lines:
            stripped = line.strip()
            if not stripped:
                page_lines.append('')
                continue
            if re.match(r'^\d{1,3}$', stripped):
                continue
            if page_lines and stripped == page_lines[-1]:
                continue
            page_lines.append(stripped)

        all_lines.extend(page_lines)

    doc.close()

    # 按章节分组
    raw_chapters = []
    current_title = None
    current_original = []
    current_translation = []
    in_translation = False

    i = 0
    while i < len(all_lines):
        line = all_lines[i]

        if not line:
            empty_count = 0
            while i < len(all_lines) and not all_lines[i]:
                empty_count += 1
                i += 1
            if empty_count >= 2 and current_original and not in_translation:
                in_translation = True
            elif (empty_count == 1 and current_original
                  and i < len(all_lines)
                  and is_likely_modern(all_lines[i][:50])):
                in_translation = True
            continue

        if is_title_line(line):
            m = CHAPTER_PATTERN.search(line)
            new_title = m.group(1)

            def chapter_num(title):
                m2 = re.search(r'第([一二三四五六七八九十百零]+)', title)
                return m2.group(1) if m2 else ''

            new_num = chapter_num(new_title)
            if current_title and new_num == chapter_num(current_title):
                if len(new_title) > len(current_title):
                    raw_chapters[-1]['title'] = new_title
                    current_title = new_title
                i += 1
                continue

            if current_title:
                raw_chapters.append({
                    'title': current_title,
                    'original_lines': current_original,
                    'translation_lines': current_translation,
                })
            current_title = new_title
            current_original = []
            current_translation = []
            in_translation = False
            rest = line[m.end():].strip()
            if rest and not CHAPTER_PATTERN.search(rest):
                current_original.append(rest)
            i += 1
            continue

        if in_translation or is_likely_modern(line[:60]):
            if not in_translation and current_original:
                in_translation = True
            current_translation.append(line)
        else:
            current_original.append(line)

        i += 1

    if current_title:
        raw_chapters.append({
            'title': current_title,
            'original_lines': current_original,
            'translation_lines': current_translation,
        })

    # 合并行→段落
    chapters = []
    for ch in raw_chapters:
        chapters.append({
            'title': ch['title'],
            'original': merge_lines(ch['original_lines'], mode='original'),
            'translation': merge_lines(ch['translation_lines'], mode='translation'),
        })

    return chapters, all_unknown


def main():
    parser = argparse.ArgumentParser(description='从 PDF 提取古籍原文和译文')
    parser.add_argument('pdf_path', help='PDF 文件路径')
    parser.add_argument('output', nargs='?', default=None, help='输出 JSON 路径（默认同名 .json）')
    parser.add_argument('--start', type=int, default=0, help='起始页码（0-based）')
    parser.add_argument('--end', type=int, default=None, help='结束页码（不含）')
    args = parser.parse_args()

    chapters, unknowns = extract(args.pdf_path, args.start, args.end)

    print(f'提取了 {len(chapters)} 章')
    for ch in chapters[:3]:
        o_len = sum(len(p) for p in ch['original'])
        t_len = sum(len(p) for p in ch['translation'])
        print(f'  {ch["title"]}: 原文 {len(ch["original"])}段 {o_len}字, '
              f'译文 {len(ch["translation"])}段 {t_len}字')

    # 保存
    out_path = args.output or args.pdf_path.rsplit('.', 1)[0] + '_extracted.json'
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(chapters, f, ensure_ascii=False, indent=2)
    print(f'\n保存: {out_path}')

    if unknowns:
        from collections import Counter
        uk = Counter(unknowns)
        print(f'\n未识别拆分部首 ({len(uk)} 种):')
        for p, c in uk.most_common(20):
            print(f'  ({p}): {c}处')


if __name__ == '__main__':
    main()
