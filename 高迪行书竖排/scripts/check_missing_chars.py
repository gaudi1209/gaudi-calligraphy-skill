"""
字库缺失字符检测。
比较 docx 文档中的汉字与字体文件（.ttf/.otf）的覆盖范围，
列出缺失字符。

用法:
    python check_missing_chars.py <DOCX路径> <字体文件路径> [输出目录] [--chunk N]
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import re
import argparse
from pathlib import Path
from collections import Counter

try:
    from docx import Document
except ImportError:
    print('需要安装 python-docx: pip install python-docx')
    sys.exit(1)

try:
    from fontTools.ttLib import TTFont
except ImportError:
    print('需要安装 fontTools: pip install fonttools')
    sys.exit(1)


def extract_chars_from_docx(docx_path):
    """从 docx 文件中提取所有汉字及出现频次"""
    doc = Document(docx_path)
    chars = Counter()
    for para in doc.paragraphs:
        for ch in para.text:
            # CJK 基本区 + 扩展A + 扩展B-F + 兼容区
            if ('\u4e00' <= ch <= '\u9fff'
                    or '\u3400' <= ch <= '\u4dbf'
                    or '\U00020000' <= ch <= '\U0002a6df'
                    or '\U0002a700' <= ch <= '\U0002b73f'
                    or '\U0002b740' <= ch <= '\U0002b81f'
                    or '\U0002b820' <= ch <= '\U0002ceaf'
                    or '\U0002ceb0' <= ch <= '\U0002ebef'
                    or '\U000f900' <= ch <= '\U000faff'):
                chars[ch] += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for ch in cell.text:
                    if '\u4e00' <= ch <= '\u9fff':
                        chars[ch] += 1
    return chars


def get_font_cmap(font_path):
    """获取字体文件中包含的字符"""
    font = TTFont(font_path)
    cmap = font.getBestCmap()
    font.close()
    return set(cmap.keys())


def main():
    parser = argparse.ArgumentParser(description='字库缺失字符检测')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('font_path', help='字体文件路径 (.ttf/.otf)')
    parser.add_argument('output_dir', nargs='?', default=None, help='输出目录')
    parser.add_argument('--chunk', type=int, default=200, help='每组字符数（默认 200）')
    args = parser.parse_args()

    # 提取文档字符
    print(f'读取文档: {args.docx_path}')
    doc_chars = extract_chars_from_docx(args.docx_path)
    print(f'文档中共有 {len(doc_chars)} 个不重复汉字')

    # 获取字库字符
    print(f'读取字体: {args.font_path}')
    font_chars = get_font_cmap(args.font_path)
    print(f'字体中共有 {len(font_chars)} 个字符')

    # 找缺失
    missing = sorted([c for c in doc_chars if ord(c) not in font_chars],
                     key=lambda c: doc_chars[c], reverse=True)
    print(f'缺失 {len(missing)} 个汉字')

    if not missing:
        print('没有缺失的汉字！')
        return

    # 输出
    output_dir = Path(args.output_dir) if args.output_dir else Path(args.docx_path).parent / '缺失字符'
    output_dir.mkdir(parents=True, exist_ok=True)

    # 前50预览
    print(f'\n缺失字符预览（前50个，按出现频率排序）:')
    for ch in missing[:50]:
        print(f'  {ch} (U+{ord(ch):04X}, 出现 {doc_chars[ch]} 次)')

    # 分组写入
    chunk_size = args.chunk
    for i in range(0, len(missing), chunk_size):
        chunk = missing[i:i + chunk_size]
        chunk_num = i // chunk_size + 1
        output_file = output_dir / f'缺失字符_{chunk_num}.txt'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(chunk))
        print(f'已写入: {output_file} ({len(chunk)} 个字符)')

    # 全部
    all_file = output_dir / '缺失字符_全部.txt'
    with open(all_file, 'w', encoding='utf-8') as f:
        f.write(''.join(missing))
    print(f'已写入: {all_file}')

    # Unicode 编码
    unicode_file = output_dir / '缺失字符_Unicode.txt'
    with open(unicode_file, 'w', encoding='utf-8') as f:
        for c in missing:
            f.write(f'U+{ord(c):04X}  {c}  (出现 {doc_chars[c]} 次)\n')
    print(f'已写入: {unicode_file}')


if __name__ == '__main__':
    main()
