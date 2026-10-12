# -*- coding: utf-8 -*-
"""核对正文引用编号与参考文献对应"""
import re
from docx import Document

d = Document(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.docx')
ps = d.paragraphs

# 收集正文引用（摘要之后到参考文献之前）
refs_used = set()
ref_section = False
for p in ps:
    t = p.text.strip()
    if t == '参考文献':
        ref_section = True
        continue
    if ref_section:
        continue
    # 跳过目录
    if '....' in t or '此处为' in t:
        continue
    for m in re.finditer(r'\[(\d+)\]', t):
        refs_used.add(int(m.group(1)))

print('正文中引用的编号:', sorted(refs_used))
print('缺失未引用编号:', sorted(set(range(1, 11)) - refs_used))

# 列出参考文献
print()
print('参考文献:')
ref_section = False
for p in ps:
    t = p.text.strip()
    if t == '参考文献':
        ref_section = True
        continue
    if ref_section:
        if t.startswith('['):
            print(' ', t[:80])
        elif t == '':
            continue
        else:
            break
