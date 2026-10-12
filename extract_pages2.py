# -*- coding: utf-8 -*-
"""提取各章节标题所在PDF页，基于摘要页位置推算论文页码（摘要页=1）"""
import json, os
import pdfplumber

BASE = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文'
pdf = os.path.join(BASE, 'work_tmp', 'TARGET.pdf')

TITLES = [
    '1 绪论', '1.1 课题背景与研究意义', '1.2 国内外研究现状', '1.3 主要研究内容与论文结构',
    '2 相关技术介绍', '2.1 微信小程序技术', '2.2 微信小程序云开发技术', '2.3 其他相关技术',
    '3 系统需求分析', '3.1 系统可行性分析', '3.2 系统功能需求分析', '3.3 系统非功能需求分析',
    '4 系统设计', '4.1 系统总体架构设计', '4.2 系统功能模块设计', '4.3 数据库设计',
    '5 系统实现', '5.1 开发环境与工具', '5.2 学生端主要功能实现', '5.3 管理端主要功能实现',
    '5.4 消息通知与进度查询实现',
    '6 系统测试', '6.1 测试环境与方法', '6.2 功能测试用例与结果', '6.3 测试结论',
    '7 总结与展望', '7.1 工作总结', '7.2 不足与展望',
    '参考文献',
]

with pdfplumber.open(pdf) as pdfdoc:
    pages = pdfdoc.pages
    total = len(pages)
    print('PDF总页数:', total)

    page_texts = [pg.extract_text() or '' for pg in pages]

    # 找摘要页（标题"摘 要"独立成行，位于目录之后的第一个）
    abs_page = None
    for pi, t in enumerate(page_texts):
        lines = [ln.strip() for ln in t.split('\n')]
        if '摘 要' in lines or '摘  要' in lines or any('摘' in ln and '要' in ln and len(ln) <= 5 for ln in lines):
            abs_page = pi
            break
    print('摘要页 PDF页:', abs_page + 1 if abs_page is not None else None)
    if abs_page is None:
        raise SystemExit('未找到摘要页')

    # 标题匹配（只在摘要页及其之后匹配）
    def match_line(ln, title):
        return ln == title or ln == title + ' ' or ln.startswith(title + ' ')

    title_pdf = {}
    for title in TITLES:
        found = None
        for pi in range(abs_page, total):
            lines = [ln.strip() for ln in page_texts[pi].split('\n')]
            if any(match_line(ln, title) for ln in lines):
                found = pi
                break
        if found is None:
            print(f'  WARN: 未找到 "{title}"')
        title_pdf[title] = found

    # 论文页码 = PDF页 - 摘要页 + 1
    result = {}
    for title, pi in title_pdf.items():
        if pi is not None:
            result[title] = pi - abs_page + 1
            print(f'  {title} -> PDF页{pi+1} -> 论文页码{result[title]}')

    out = os.path.join(BASE, 'work_tmp', 'page_numbers.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print('已保存:', out)
