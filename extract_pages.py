# -*- coding: utf-8 -*-
"""从PDF中提取各章节标题所在页（物理页）并计算论文页码"""
import pdfplumber, json, os, re

pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.pdf'
headings = [
    "1  绪论", "1.1  研究背景与意义", "1.2  国内外研究现状", "1.3  主要研究内容与组织结构",
    "2  相关技术介绍", "2.1  微信小程序开发技术", "2.2  微信云开发技术", "2.3  数据库技术",
    "3  系统需求分析", "3.1  可行性分析", "3.2  功能需求分析", "3.3  非功能需求分析",
    "4  系统设计", "4.1  系统总体架构设计", "4.2  系统功能模块设计", "4.3  数据库设计",
    "5  系统实现", "5.1  开发环境与工具", "5.2  用户端主要功能实现", "5.3  管理端主要功能实现", "5.4  消息通知与进度查询实现",
    "6  系统测试", "6.1  测试环境与测试方法", "6.2  功能测试用例与结果", "6.3  测试结果分析",
    "7  总结与展望", "7.1  工作总结", "7.2  不足与展望",
    "参考文献",
]

# 压缩标题中的连续空格用于匹配
def norm(s):
    return re.sub(r'\s+', ' ', s).strip()

with pdfplumber.open(pdf) as pdfdoc:
    pages_text = []
    for pi, page in enumerate(pdfdoc.pages):
        text = page.extract_text() or ''
        pages_text.append(norm(text))

# 找出摘要页（含"摘 要"或"摘要"），作为论文页码1
abstract_page = None
for i, t in enumerate(pages_text):
    if '摘 要' in t or '摘要' in t:
        abstract_page = i + 1  # 1-based physical page
        break
print('摘要物理页:', abstract_page)
if abstract_page is None:
    print('ERROR: 未找到摘要页')
    raise SystemExit(1)

result = {}
for h in headings:
    hn = norm(h)
    found_page = None
    for pi, t in enumerate(pages_text):
        # 匹配标题行（标题应出现在行首附近）
        if hn in t.replace(' ', ''):
            # 确认它在页面中出现且不是目录页（目录页在摘要页之前）
            if pi + 1 >= abstract_page:
                found_page = pi + 1
                break
    if found_page is None:
        # 二次尝试：模糊匹配（去掉空格后包含）
        h_compact = hn.replace(' ', '')
        for pi, t in enumerate(pages_text):
            t_compact = t.replace(' ', '')
            if h_compact in t_compact and pi + 1 >= abstract_page:
                found_page = pi + 1
                break
    if found_page is None:
        print(f'WARNING: 未找到标题 "{h}"')
        result[h] = 0
    else:
        thesis_page = found_page - abstract_page + 1
        result[h] = thesis_page
        print(f'{h} -> 物理页{found_page} = 论文页{thesis_page}')

# 参考文献页（最后出现参考文献的位置，且不是目录）
# 目录中的"参考文献"在摘要页之前，正文的参考文献标题在摘要页之后
ref_page = None
for pi, t in enumerate(pages_text):
    if pi + 1 >= abstract_page:
        # 参考文献标题行
        lines = t.split('\n')
        for ln in lines:
            if norm(ln).replace(' ', '') == '参考文献':
                ref_page = pi + 1
                break
    if ref_page:
        break
if ref_page:
    result['参考文献'] = ref_page - abstract_page + 1
    print('参考文献 -> 物理页%d = 论文页%d' % (ref_page, ref_page - abstract_page + 1))
else:
    # 兜底：找最后一个含参考文献的物理页
    for pi in range(len(pages_text)-1, abstract_page-2, -1):
        if '参考文献' in pages_text[pi]:
            result['参考文献'] = pi + 1 - abstract_page + 1
            print('参考文献(兜底) -> 论文页%d' % result['参考文献'])
            break

out = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\page_numbers.json'
with open(out, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print('已保存:', out)
