"""
Word 文档简转繁（词组级语义判断，书法古典字形）。
用法: python s2twp.py <输入文件> [输出文件] [--skip N] [--mode auto|full|semantic]

模式:
  auto     自动检测（默认）。繁体原文仅修正一对多语义错误，简体原文完整转换。
  full     完整转换。所有字符经 s2t 处理（含字形变体）。
  semantic 仅修正一对多语义映射。

使用 opencc s2t 模式，保留古典字形（裏/喫/穀 等），
适合书法、古籍排版。正确处理一对多语义映射。
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import argparse
import os
import opencc
from docx import Document
from docx.oxml.ns import qn


# 一对多映射：简体字合并了多个繁体字，需要根据上下文语义判断
# 例如: 里→裏(里面)/里(里程), 发→發(发展)/髮(头发)
ONE_TO_MANY = set([
    # 高频一对多
    '发', '干', '面', '里', '只', '后', '历', '钟', '系', '复',
    '脏', '志', '尽', '征', '余', '仆', '托', '采', '欲', '辟',
    '谷', '注', '于', '凶', '并', '云', '尸', '斗', '游', '朴',
    '几', '舍', '布', '卷', '占', '才', '回', '咸', '向', '累',
    '了', '秋', '松', '台', '坛', '涂', '团', '须', '御', '郁',
    '叹', '升', '签', '曲', '沈', '合', '胡', '汇', '伙', '获',
    '借', '克', '夸', '困', '冬', '当', '冲', '丑', '出', '尝',
    '表', '别', '卜', '蔑', '栗', '戚', '幸', '梁', '周',
    '吁', '愿', '据', '腊', '蜡', '苏', '凭', '弦', '肴',
])


def detect_traditional(paragraphs, skip, trans_indices, converter):
    """检测文档是否主要为繁体中文"""
    sample_text = ''
    count = 0
    for i, para in enumerate(paragraphs):
        if i < skip or i in trans_indices:
            continue
        if para.text.strip():
            sample_text += para.text
            count += 1
            if count >= 100:
                break
    if not sample_text:
        return False
    sample = sample_text[:5000]
    converted = converter.convert(sample)
    changes = sum(1 for a, b in zip(sample, converted) if a != b)
    ratio = changes / len(sample)
    return ratio < 0.10  # 变化率低于10%视为繁体原文


def find_translation_ranges(paragraphs, skip=0):
    """找出所有参考译文区域的段落索引范围"""
    ranges = []
    in_trans = False
    start = None
    for i, para in enumerate(paragraphs):
        if i < skip:
            continue
        text = para.text.strip()
        if text and '参考译文' in text and len(text) < 30:
            if in_trans and start is not None:
                ranges.append((start, i - 1))
            start = i
            in_trans = True
            continue
        if in_trans and text and '篇' in text and len(text) < 30 and '---' not in text and '译文' not in text:
            ranges.append((start, i - 1))
            in_trans = False
    if in_trans and start is not None:
        ranges.append((start, len(paragraphs) - 1))
    return set(i for start, end in ranges for i in range(start, end + 1))


def main():
    parser = argparse.ArgumentParser(description='Word 简转繁（书法古典字形）')
    parser.add_argument('input', help='输入 .docx 文件')
    parser.add_argument('output', nargs='?', help='输出文件')
    parser.add_argument('--skip', type=int, default=0, help='跳过前 N 个段落')
    parser.add_argument('--mode', choices=['auto', 'full', 'semantic'],
                        default='auto', help='auto=自动检测, full=完整转换, semantic=仅一对多')
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output) if args.output else input_path.rsplit('.', 1)[0] + '_繁体.docx'

    print(f'读取: {input_path}')
    doc = Document(input_path)

    converter = opencc.OpenCC('s2t')

    # 找出参考译文区域
    trans_indices = find_translation_ranges(doc.paragraphs, args.skip)
    print(f'参考译文区域: {len(trans_indices)} 个段落将跳过转换')

    # 确定模式
    if args.mode == 'auto':
        semantic_only = detect_traditional(doc.paragraphs, args.skip, trans_indices, converter)
        mode_name = '仅修正一对多语义映射' if semantic_only else '完整 s2t 转换'
        print(f'自动检测: 繁体原文 → {mode_name}')
    elif args.mode == 'semantic':
        semantic_only = True
        print(f'模式: 仅修正一对多语义映射')
    else:
        semantic_only = False
        print(f'模式: 完整 s2t 转换')

    changed_runs = 0
    total_chars = 0
    skipped_runs = 0
    change_details = {}

    for i, para in enumerate(doc.paragraphs):
        if i < args.skip:
            continue
        if i in trans_indices:
            skipped_runs += len(para.runs)
            continue
        for run in para.runs:
            text = run.text
            if not text:
                continue
            converted = converter.convert(text)
            if converted == text:
                continue

            if semantic_only:
                # 仅修正一对多映射的字符，保留字形变体
                new_chars = []
                any_change = False
                for orig_char, conv_char in zip(text, converted):
                    if orig_char != conv_char and orig_char in ONE_TO_MANY:
                        new_chars.append(conv_char)
                        any_change = True
                        total_chars += 1
                        key = f'{orig_char}→{conv_char}'
                        change_details[key] = change_details.get(key, 0) + 1
                    else:
                        new_chars.append(orig_char)
                if not any_change:
                    continue
                new_text = ''.join(new_chars)
            else:
                new_text = converted
                for a, b in zip(text, converted):
                    if a != b:
                        total_chars += 1
                        key = f'{a}→{b}'
                        change_details[key] = change_details.get(key, 0) + 1

            t_elem = run._element.find(qn('w:t'))
            if t_elem is not None:
                t_elem.text = new_text
            run.text = new_text
            changed_runs += 1

    doc.save(output_path)
    print(f'转换完成: {changed_runs} 个 run，{total_chars} 处字符转换')
    if skipped_runs:
        print(f'跳过译文: {skipped_runs} 个 run')
    if change_details:
        top = sorted(change_details.items(), key=lambda x: -x[1])[:10]
        for k, v in top:
            print(f'  {k} : {v}处')
    print(f'保存: {output_path}')


if __name__ == '__main__':
    main()
