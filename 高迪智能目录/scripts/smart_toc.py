"""
将 Word 文档中的手动目录转换为智能 TOC 域。
支持点击跳转、排版后 F9 刷新自动更新页码。

用法: python smart_toc.py <输入文件> [输出文件] [--toc-start N] [--toc-end N] [--heading N]

参数:
  输入文件          .docx 文件路径
  输出文件          输出路径（默认：原名_智能目录.docx）
  --toc-start N    目录起始段落索引（0-based），默认自动检测
  --toc-end N      目录结束段落索引（0-based），默认自动检测
  --heading N      正文章节标题起始段落索引（0-based），默认自动检测

依赖: python-docx, pywin32 (WPS/Word 需已安装)
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import re
import os
import time
import argparse
from docx import Document
import win32com.client


def find_best_match(target, candidates):
    """基于最长公共子序列的模糊匹配"""
    best = None
    best_score = 0
    for cand in candidates:
        t_len, c_len = len(target), len(cand)
        ti = ci = 0
        common = 0
        while ti < t_len and ci < c_len:
            if target[ti] == cand[ci]:
                common += 1
                ti += 1
                ci += 1
            elif t_len - ti > c_len - ci:
                ti += 1
            else:
                ci += 1
        score = common / max(t_len, c_len)
        if score > best_score:
            best_score = score
            best = cand
    return best if best_score > 0.6 else None


def detect_toc_range(doc):
    """自动检测目录区域：找到连续含虚线+页码或 tab+页码的段落"""
    toc_start = None
    toc_end = None
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        is_toc_line = False
        if text and re.search(r'[-–—]{3,}.{0,10}$', text):
            is_toc_line = True
        elif text and '\t' in text and re.search(r'\t\d{1,5}$', text):
            is_toc_line = True
        if is_toc_line:
            if toc_start is None:
                toc_start = i
            toc_end = i
        elif toc_start is not None and toc_end is not None:
            break
    return toc_start, toc_end


def detect_content_start(doc, toc_end):
    """自动检测正文起始位置：目录之后，第一个包含"篇"字的短段落"""
    for i in range(toc_end + 1 if toc_end else 0, len(doc.paragraphs)):
        text = doc.paragraphs[i].text.strip()
        if text and len(text) < 30 and '篇' in text and '---' not in text:
            return i
    return toc_end + 1 if toc_end else 0


def main():
    parser = argparse.ArgumentParser(description='将 Word 手动目录转为智能 TOC 域')
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('output', nargs='?', help='输出文件路径')
    parser.add_argument('--toc-start', type=int, help='目录起始段落索引 (0-based)')
    parser.add_argument('--toc-end', type=int, help='目录结束段落索引 (0-based)')
    parser.add_argument('--heading', type=int, help='正文标题起始段落索引 (0-based)')
    args = parser.parse_args()

    doc_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output) if args.output else doc_path.rsplit('.', 1)[0] + '_智能目录.docx'

    # === 分析文档结构 ===
    print('分析文档结构...')
    doc = Document(doc_path)

    # 检测目录范围
    toc_start, toc_end = args.toc_start, args.toc_end
    if toc_start is None or toc_end is None:
        auto_start, auto_end = detect_toc_range(doc)
        if toc_start is None:
            toc_start = auto_start
        if toc_end is None:
            toc_end = auto_end

    if toc_start is None:
        print('未检测到目录区域（含虚线的段落），请用 --toc-start/--toc-end 手动指定')
        sys.exit(1)

    print(f'目录区域: 段落 {toc_start} - {toc_end}')

    # 收集目录条目（支持虚线或 tab 分隔格式）
    toc_entries = []
    for i in range(toc_start, toc_end + 1):
        text = doc.paragraphs[i].text.strip()
        if not text:
            continue
        chapter = None
        if re.search(r'[-–—]{3,}', text):
            chapter = re.split(r'[-–—]{3,}', text)[0].strip()
        elif '\t' in text:
            chapter = text.split('\t')[0].strip()
        if chapter:
            toc_entries.append((i + 1, chapter))  # COM 1-indexed

    if not toc_entries:
        print('目录区域中未找到有效条目')
        sys.exit(1)

    # 检测正文起始
    content_start = args.heading
    if content_start is None:
        content_start = detect_content_start(doc, toc_end)
    print(f'正文起始: 段落 {content_start}')

    # 收集正文标题（含"篇"字的短文本，排除含"译文"和虚线的）
    content_titles = {}
    for i in range(content_start, len(doc.paragraphs)):
        text = doc.paragraphs[i].text.strip()
        if text and '篇' in text and len(text) < 30 and '---' not in text and '译文' not in text:
            if text not in content_titles:
                content_titles[text] = i + 1  # COM 1-indexed

    print(f'正文标题: {len(content_titles)} 个')

    # 匹配
    match_map = {}
    for toc_idx, chapter in toc_entries:
        if chapter in content_titles:
            match_map[toc_idx] = content_titles[chapter]
        else:
            best = find_best_match(chapter, content_titles.keys())
            if best:
                match_map[toc_idx] = content_titles[best]
                print(f'  模糊: {chapter} -> {best}')

    # 去重
    seen = set()
    unique = {}
    for toc_idx, content_idx in match_map.items():
        if content_idx not in seen:
            seen.add(content_idx)
            unique[toc_idx] = content_idx

    print(f'目录: {len(toc_entries)}, 匹配: {len(match_map)}, 去重: {len(unique)}')

    if len(match_map) == 0:
        print('无匹配项，退出')
        sys.exit(1)

    # === Word COM ===
    print('启动 Word...')
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    try:
        wdoc = word.Documents.Open(doc_path)
        word.ActiveWindow.View.Type = 3
        wdoc.Repaginate()
        print(f'段落数: {wdoc.Paragraphs.Count}')

        # 插入 TC 域（从后往前）
        print('插入 TC 域...')
        tc_count = 0
        for toc_idx, content_idx in sorted(unique.items(), key=lambda x: x[1], reverse=True):
            para = wdoc.Paragraphs(content_idx)
            entry_text = para.Range.Text.strip().split('\r')[0].split('\n')[0][:30]
            rng = wdoc.Range(para.Range.End - 1, para.Range.End - 1)
            try:
                wdoc.Fields.Add(rng, -1, f'TC "{entry_text}" \\f \\l 1', True)
                tc_count += 1
            except Exception as e:
                print(f'  失败 para {content_idx}: {e}')

        print(f'TC 域: {tc_count}')

        # 删除旧目录条目（从后往前）
        print('替换目录...')
        first_toc = toc_entries[0][0]
        last_toc = toc_entries[-1][0]
        for i in range(last_toc, first_toc - 1, -1):
            para = wdoc.Paragraphs(i)
            text = para.Range.Text.strip()
            if text and text != '目\u3000录' and text != '目录':
                rng = para.Range
                del_rng = wdoc.Range(rng.Start, rng.End - 1)
                del_rng.Delete()

        # 插入 TOC 域
        toc_para = wdoc.Paragraphs(first_toc)
        toc_rng = wdoc.Range(toc_para.Range.Start, toc_para.Range.Start)
        wdoc.Fields.Add(toc_rng, -1, 'TOC \\f \\h \\z', True)
        print('TOC 域已插入')

        # 保存并刷新
        wdoc.Save()
        print('刷新域...')
        wdoc.Repaginate()
        time.sleep(3)
        try:
            word.Selection.WholeStory()
            word.Selection.Fields.Update()
        except:
            pass
        time.sleep(5)
        try:
            wdoc.Fields.Update()
        except:
            pass
        time.sleep(3)

        wdoc.SaveAs2(output_path)
        print(f'保存: {output_path}')

        # 验证
        print('验证:')
        for i in range(first_toc - 1, min(first_toc + 10, wdoc.Paragraphs.Count)):
            para = wdoc.Paragraphs(i)
            text = para.Range.Text.strip()[:70]
            if text:
                print(f'  {text}')

        wdoc.Close(False)

    except Exception as e:
        print(f'错误: {e}')
        import traceback
        traceback.print_exc()
        try:
            wdoc.Close(False)
        except:
            pass

    try:
        word.Quit()
    except:
        pass
    print('完成')


if __name__ == '__main__':
    main()
