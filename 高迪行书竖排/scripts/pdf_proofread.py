"""
PDF vs DOCX 校对工具。
提取 PDF 原文转繁体，逐段与 docx 对比，找出文字差异。

用法:
    python pdf_proofread.py <PDF路径> <DOCX路径> [--skip N] [--no-punct]
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import re
import argparse
import difflib

try:
    import fitz
except ImportError:
    print('需要安装 PyMuPDF: pip install pymupdf')
    sys.exit(1)

try:
    import opencc
except ImportError:
    print('需要安装 opencc: pip install opencc')
    sys.exit(1)

from docx import Document


CHAPTER_PATTERN = re.compile(r'(.{2,10}篇第[一二三四五六七八九十百零]+)')
converter = opencc.OpenCC('s2t')


def extract_pdf_original(pdf_path, start_page=0, end_page=None):
    """提取 PDF 中所有原文段落（排除现代白话文）"""
    doc = fitz.open(pdf_path)
    if end_page is None:
        end_page = len(doc)

    sections = []
    current_title = ''
    current_text = []
    in_yuanwen = False

    for pn in range(start_page, end_page):
        page = doc[pn]
        text = page.get_text()
        lines = text.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue
            if re.match(r'^\d{1,3}$', line):
                continue

            if CHAPTER_PATTERN.search(line) and len(line) < 25:
                if current_title and current_text:
                    sections.append((current_title, ''.join(current_text)))
                current_title = line
                current_text = []
                in_yuanwen = True
                continue

            if '【原文】' in line:
                in_yuanwen = True
                continue

            if line.startswith('【') and '】' in line:
                marker = line[:line.index('】') + 1]
                if marker != '【原文】':
                    in_yuanwen = False
                    continue

            if in_yuanwen:
                modern_phrases = ['也就是说', '比如说', '这是指', '意思是说',
                                  '指的是', '相当于', '也可以说', '简单来说',
                                  '从前的', '十分聪明', '活到天赋']
                if not any(phrase in line for phrase in modern_phrases):
                    current_text.append(line)

    if current_title and current_text:
        sections.append((current_title, ''.join(current_text)))
    doc.close()
    return sections


def extract_docx_sections(docx_path, skip=0):
    """从 docx 提取正文段落（排除目录和译文）"""
    doc = Document(docx_path)
    sections = []
    current_title = ''
    current_paras = []
    in_translation = False

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if i < skip:
            continue

        if (CHAPTER_PATTERN.search(text) and len(text) < 25
                and '參考' not in text and '参考' not in text):
            if current_title and current_paras:
                sections.append((current_title, current_paras))
            current_title = text.split('\t')[0].strip()
            current_paras = []
            in_translation = False
            continue

        if '參考譯文' in text or '参考译文' in text:
            in_translation = True
            if current_title and current_paras:
                sections.append((current_title, current_paras))
            current_title = ''
            current_paras = []
            continue

        if in_translation:
            continue

        current_paras.append((i, text))

    if current_title and current_paras:
        sections.append((current_title, current_paras))

    return sections


def compare_sections(pdf_sections, docx_sections, ignore_punct=False):
    """逐章对比 PDF 和 docx，找出文字差异"""
    def title_key(t):
        m = re.search(r'(.+篇第.+)', t)
        return m.group(1) if m else t

    pdf_map = {title_key(converter.convert(t)): text for t, text in pdf_sections}
    docx_map = {title_key(t): paras for t, paras in docx_sections}

    common = set(pdf_map.keys()) & set(docx_map.keys())
    print(f'PDF: {len(pdf_map)} 章, DOCX: {len(docx_map)} 章, 共同: {len(common)}')

    all_diffs = []
    punct_pat = r'[，。！？；：、\u201c\u201d\u2018\u2019（）《》〈〉【】…\u2014\-–—]'

    for key in sorted(common):
        pdf_text = converter.convert(pdf_map[key])
        docx_text = ''.join(t for _, t in docx_map[key])

        pdf_clean = re.sub(r'\s', '', pdf_text)
        docx_clean = re.sub(r'\s', '', docx_text)

        if ignore_punct:
            pdf_clean = re.sub(punct_pat, '', pdf_clean)
            docx_clean = re.sub(punct_pat, '', docx_clean)

        if pdf_clean == docx_clean:
            continue

        matcher = difflib.SequenceMatcher(None, pdf_clean, docx_clean)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                continue
            pdf_seg = pdf_clean[i1:i2]
            docx_seg = docx_clean[j1:j2]

            if ignore_punct:
                pdf_no = re.sub(punct_pat, '', pdf_seg)
                docx_no = re.sub(punct_pat, '', docx_seg)
                if not pdf_no and not docx_no:
                    continue

            ctx_before = docx_clean[max(0, j1 - 8):j1]
            ctx_after = docx_clean[j2:j2 + 8]

            all_diffs.append({
                'chapter': key,
                'type': tag,
                'pdf': pdf_seg[:50],
                'docx': docx_seg[:50],
                'context_before': ctx_before,
                'context_after': ctx_after,
            })

    return all_diffs


def main():
    parser = argparse.ArgumentParser(description='PDF vs DOCX 校对')
    parser.add_argument('pdf_path', help='PDF 文件路径')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('--skip', type=int, default=86, help='跳过 docx 前 N 段（默认 86，跳过目录）')
    parser.add_argument('--no-punct', action='store_true', help='忽略标点差异')
    parser.add_argument('--pdf-start', type=int, default=0, help='PDF 起始页')
    parser.add_argument('--pdf-end', type=int, default=None, help='PDF 结束页')
    args = parser.parse_args()

    print('=== 提取 PDF 原文 ===')
    pdf_sections = extract_pdf_original(args.pdf_path, args.pdf_start, args.pdf_end)
    print(f'PDF 章节: {len(pdf_sections)}')

    print('\n=== 提取 DOCX 正文 ===')
    docx_sections = extract_docx_sections(args.docx_path, args.skip)
    print(f'DOCX 章节: {len(docx_sections)}')

    print('\n=== 对比差异 ===')
    diffs = compare_sections(pdf_sections, docx_sections, args.no_punct)
    print(f'文字差异总数: {len(diffs)}')
    for d in diffs[:50]:
        print(f'  [{d["chapter"][:10]}] {d["type"]}: '
              f'PDF="{d["pdf"]}" vs DOCX="{d["docx"]}"')
        print(f'    上下文: ...{d["context_before"]}【{d["docx"]}】'
              f'{d["context_after"]}...')


if __name__ == '__main__':
    main()
