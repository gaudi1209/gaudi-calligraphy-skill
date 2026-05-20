"""
修改 Word 文档中所有标点符号的格式和字形（竖排书法排版）。
用法: python fix_punctuation.py <输入文件> [输出文件]

处理逻辑：
  1. 书名号《》：删除标记，对书名文字添加波浪下划线（跨 run 处理）
  2. 字形替换：引号→角括号，括号→直式
  3. 格式统一：深红(#8B0000)、楷体、小四(12pt)
  4. 位置调整：上标/下标/不调整
  5. 黑括号【】︻︼：深红、楷体，但保持原字号、不调上下标
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import copy
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.text.run import Run
from lxml import etree

# ── 字形替换映射 ──
CHAR_REPLACE = {
    # 引号 → 角括号
    '\u201c': '\u300e',  # " → 『
    '\u201d': '\u300f',  # " → 』
    '\u2018': '\u300c',  # ' → 「
    '\u2019': '\u300d',  # ' → 」
    '\uff02': '\u300e',  # ＂ → 『（全角双引号）
    # 括号 → 直式
    '\uff08': '\ufe35',  # （ → ︵
    '\uff09': '\ufe36',  # ） → ︶
    '(': '\ufe35',       # ( → ︵
    ')': '\ufe36',       # ) → ︶
}

# ── 书名号标记 ──
BOOK_OPEN = set('\u300a\u3008')   # 《 〈
BOOK_CLOSE = set('\u300b\u3009')  # 》 〉
BOOK_MARKERS = BOOK_OPEN | BOOK_CLOSE

# ── 黑括号（保持原字号，不调上下标）──
LENTICULAR_BRACKETS = set('\u3010\u3011\uFE3B\uFE3C')  # 【 】 ︻ ︼

# ── 位置规则 ──
# 不调整位置（只改格式，不改上下标）
NO_SCRIPT = set('-\u2013\u2014')  # - – —
NO_SCRIPT.add('\ufe4f')           # ﹏ 浪线不调位置
NO_SCRIPT.add('\ufe35')           # ︵ 不调位置
NO_SCRIPT.add('\ufe36')           # ︶ 不调位置

# 下标
SUBSCRIPT = set('\u300d\u300f')   # 」』

# 特殊符号范围（圈号数字等，只改格式不调位置）
def is_special_symbol(char):
    return 0x2460 <= ord(char) <= 0x24FF


def is_punctuation(char):
    # 黑括号也是标点（需要改颜色和字体）
    if char in LENTICULAR_BRACKETS:
        return True
    if is_special_symbol(char):
        return True
    cp = ord(char)
    if 0x3000 <= cp <= 0x303F:
        return True
    if 0xFE30 <= cp <= 0xFE4F:  # CJK兼容形式（︵︶﹏等）
        return True
    if 0xFF00 <= cp <= 0xFFEF:
        if char in '\uff0c\u3002\uff01\uff1f\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u300a\u300b\u3008\u3009\u3010\u3011\u3001\u2026\u2014\u00b7':
            return True
    if cp == 0x00B7:
        return True
    if char in '\uff0c\u3002\uff01\uff1f\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u300a\u300b\u3008\u3009\u3010\u3011\u3001\u2026\u2014\u00b7\uff5e\u3001,.!?;:\'"()[]{}<>@#$%^&*+=|/\\~`_\u2026':
        return True
    if 0x2000 <= cp <= 0x206F:
        if char in '\u2026\u00b7\u2018\u2019\u201c\u201d':
            return True
    return False


def set_run_font(run, char):
    """设置格式：深红、楷体，根据字符类型调整位置和大小"""
    first_char = char[0] if char else ''

    # 黑括号：只改颜色和字体，保持原字号和位置
    if len(char) == 1 and first_char in LENTICULAR_BRACKETS:
        run.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
        run.font.name = '楷体'
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
        rFonts.set(qn('w:eastAsia'), '楷体')
        return

    # 常规标点格式
    run.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
    run.font.size = Pt(12)
    run.font.name = '楷体'

    # 特殊符号和浪线不调位置
    if len(char) == 1 and is_special_symbol(char):
        pass
    elif any(c in NO_SCRIPT for c in char):
        pass
    elif first_char in SUBSCRIPT:
        run.font.subscript = True
    else:
        run.font.superscript = True

    # 东亚字体
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), '楷体')


def add_wavy_underline_to_run(run):
    """对 Run 对象添加深红波浪下划线"""
    rPr = run._element.get_or_add_rPr()
    u = rPr.find(qn('w:u'))
    if u is None:
        u = etree.SubElement(rPr, qn('w:u'))
    u.set(qn('w:val'), 'wave')
    u.set(qn('w:color'), '8B0000')


def add_wavy_underline_to_element(elem):
    """对 XML run 元素添加深红波浪下划线"""
    rPr = elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(elem, qn('w:rPr'))
        elem.insert(0, rPr)
    u = rPr.find(qn('w:u'))
    if u is None:
        u = etree.SubElement(rPr, qn('w:u'))
    u.set(qn('w:val'), 'wave')
    u.set(qn('w:color'), '8B0000')


def handle_book_name_markers(para):
    """处理书名号《》：删除标记字符，对书名文字添加波浪下划线。

    必须在 process_run 之前调用。
    跨 run 处理：支持《》分布在不同 run 中的情况。
    """
    runs = list(para.runs)
    if not runs:
        return 0

    # Phase 1: 遍历所有 run，处理包含《》的 run
    # 拆分含标记的 run，删除标记字符，记录哪些 run 在书名内
    in_book = False
    modified_count = 0

    for run in runs:
        text = run.text
        if not text:
            continue

        has_marker = any(c in text for c in BOOK_MARKERS)

        if not has_marker:
            if in_book:
                add_wavy_underline_to_run(run)
            continue

        # 含标记的 run → 按《》拆分
        segments = []  # [(text, is_inside_book)]
        current = ''

        for ch in text:
            if ch in BOOK_OPEN:
                if current:
                    segments.append((current, in_book))
                    current = ''
                in_book = True
            elif ch in BOOK_CLOSE:
                if current:
                    segments.append((current, True))  # 《》内的文字
                    current = ''
                in_book = False
            else:
                current += ch

        if current:
            segments.append((current, in_book))

        modified_count += 1

        if not segments:
            # 整个 run 只有标记，删除
            run._element.getparent().remove(run._element)
            continue

        # 用拆分后的段替换原 run
        parent = run._element.getparent()
        run_index = list(parent).index(run._element)
        parent.remove(run._element)

        for j, (seg_text, seg_in_book) in enumerate(segments):
            new_elem = copy.deepcopy(run._element)
            t = new_elem.find(qn('w:t'))
            if t is not None:
                t.text = seg_text
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            if seg_in_book:
                add_wavy_underline_to_element(new_elem)
            parent.insert(run_index + j, new_elem)

    return modified_count


def process_run(run):
    text = run.text
    if not text:
        return

    # 先做字形替换（不含书名号，书名号由 handle_book_name_markers 处理）
    replaced = ''.join(CHAR_REPLACE.get(c, c) for c in text)

    if len(replaced) == 1:
        if is_punctuation(replaced):
            # 更新文本并设置格式
            t = run._element.find(qn('w:t'))
            if t is not None:
                t.text = replaced
            run.text = replaced
            set_run_font(run, replaced)
        elif replaced != text:
            # 替换了字符但不是标点（不应发生）
            t = run._element.find(qn('w:t'))
            if t is not None:
                t.text = replaced
        return

    has_punct = any(is_punctuation(c) for c in replaced)
    if not has_punct and replaced == text:
        return

    parent = run._element.getparent()
    run_index = list(parent).index(run._element)

    # 按标点/非标点分组
    chars = [(c, is_punctuation(c)) for c in replaced]
    groups = []
    cur_text = chars[0][0]
    cur_flag = chars[0][1]
    for ch, flag in chars[1:]:
        if flag == cur_flag:
            cur_text += ch
        else:
            groups.append((cur_text, cur_flag))
            cur_text = ch
            cur_flag = flag
    groups.append((cur_text, cur_flag))

    if len(groups) == 1:
        t = run._element.find(qn('w:t'))
        if t is not None:
            t.text = replaced
        if groups[0][1]:
            set_run_font(run, replaced)
        return

    new_elements = []
    for g_text, g_flag in groups:
        new_run = copy.deepcopy(run._element)
        t_elem = new_run.find(qn('w:t'))
        if t_elem is None:
            t_elem = etree.SubElement(new_run, qn('w:t'))
        t_elem.text = g_text
        if g_text.startswith(' ') or g_text.endswith(' '):
            t_elem.set(qn('xml:space'), 'preserve')
        if g_flag:
            run_obj = Run(new_run, parent)
            set_run_font(run_obj, g_text)
        new_elements.append(new_run)

    parent.remove(run._element)
    for i, elem in enumerate(new_elements):
        parent.insert(run_index + i, elem)


def main():
    if len(sys.argv) < 2:
        print('用法: python fix_punctuation.py <输入文件> [输出文件]')
        sys.exit(1)

    input_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base, ext = input_path.rsplit('.', 1)
        output_path = f'{base}_标点修改.{ext}'

    print(f'读取文件: {input_path}')
    doc = Document(input_path)

    punct_count = 0
    replace_count = 0
    book_marker_count = 0
    total_runs = 0

    for para in doc.paragraphs:
        # 先处理书名号（跨 run）
        book_marker_count += handle_book_name_markers(para)

        # 再处理其他标点（逐 run）
        runs = list(para.runs)
        for run in runs:
            total_runs += 1
            text = run.text
            if text:
                # 统计替换数
                for c in text:
                    if c in CHAR_REPLACE:
                        replace_count += 1
                # 统计标点数（替换后的）
                replaced = ''.join(CHAR_REPLACE.get(c, c) for c in text)
                punct_count += sum(1 for c in replaced if is_punctuation(c))
            process_run(run)

    print(f'处理完成: {total_runs} 个 run，书名号处理 {book_marker_count} 处，字形替换 {replace_count} 个，格式修改 {punct_count} 个标点')
    doc.save(output_path)
    print(f'保存文件: {output_path}')


if __name__ == '__main__':
    main()
