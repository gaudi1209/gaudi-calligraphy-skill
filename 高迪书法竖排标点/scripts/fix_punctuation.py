"""
修改 Word 文档中所有标点符号的格式和字形（竖排书法排版）。
用法: python fix_punctuation.py <输入文件> [输出文件]

处理逻辑：
  1. 字形替换：引号→角括号，书名号→浪线，括号→直式
  2. 格式统一：深红(#8B0000)、楷体、小四(12pt)
  3. 位置调整：上标/下标/不调整
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
    # 书名号 → 浪线
    '\u300a': '\ufe4f',  # 《 → ﹏
    '\u300b': '\ufe4f',  # 》 → ﹏
    '\u3008': '\ufe4f',  # 〈 → ﹏
    '\u3009': '\ufe4f',  # 〉 → ﹏
    # 括号 → 直式
    '\uff08': '\ufe35',  # （ → ︵
    '\uff09': '\ufe36',  # ） → ︶
    '(': '\ufe35',       # ( → ︵
    ')': '\ufe36',       # ) → ︶
}

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
    # 替换后的字符（︵︶等）也需要设置格式，不跳过
    # NO_SCRIPT 仅用于跳过位置调整，不影响格式设置
    if is_special_symbol(char):
        return True
    cp = ord(char)
    if 0x3000 <= cp <= 0x303F:
        return True
    if 0xFE30 <= cp <= 0xFE4F:  # CJK兼容形式（︵︶﹏等）
        return True
    if 0xFF00 <= cp <= 0xFFEF:
        return char in '\uff0c\u3002\uff01\uff1f\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u300a\u300b\u3008\u3009\u3010\u3011\u3001\u2026\u2014\u00b7'
    if cp == 0x00B7:
        return True
    if char in '\uff0c\u3002\uff01\uff1f\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u300a\u300b\u3008\u3009\u3010\u3011\u3001\u2026\u2014\u00b7\uff5e\u3001,.!?;:\'"()[]{}<>@#$%^&*+=|/\\~`_\u2026':
        return True
    if 0x2000 <= cp <= 0x206F:
        if char in '\u2026\u00b7\u2018\u2019\u201c\u201d':
            return True
    return False


def set_run_font(run, char):
    """设置格式：深红、楷体、小四，根据字符类型调整位置"""
    run.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
    run.font.size = Pt(12)
    run.font.name = '楷体'

    # 特殊符号和浪线不调位置
    first_char = char[0] if char else ''
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


def process_run(run):
    text = run.text
    if not text:
        return

    # 先做字形替换
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
    total_runs = 0

    for para in doc.paragraphs:
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

    print(f'处理完成: {total_runs} 个 run，字形替换 {replace_count} 个，格式修改 {punct_count} 个标点')
    doc.save(output_path)
    print(f'保存文件: {output_path}')


if __name__ == '__main__':
    main()
