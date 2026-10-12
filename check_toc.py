# -*- coding: utf-8 -*-
from docx import Document
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def qn(tag): return W + tag
def localname(el): return etree.QName(el).localname

d = Document(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.docx')
body = d.element.body

def elem_text(e):
    return ''.join(t.text or '' for t in e.iter(qn('t')))

# 找所有含"请更新目录"的段落
for i, ch in enumerate(list(body)):
    if localname(ch) == 'p' and '请更新目录' in elem_text(ch):
        print(f'body[{i}] 含"请更新目录": "{elem_text(ch)[:50]}"')
        # 打印pPr
        pPr = ch.find(qn('pPr'))
        print('  pPr:', etree.tostring(pPr).decode()[:500] if pPr is not None else 'None')

# 目录SDT内容检查
for ch in body:
    if localname(ch) == 'sdt':
        for sc in ch:
            if localname(sc) == 'sdtContent':
                items = list(sc)
                print(f'SDT sdtContent 子元素数: {len(items)}')
                for ii, child in enumerate(items):
                    print(f'  [{ii}] {elem_text(child)[:45]}')
        break

# 找目录区尾部（目录SDT之后到摘要之前）的段落
sdt_found = False
for i, ch in enumerate(list(body)):
    if localname(ch) == 'sdt':
        sdt_found = True
        continue
    if sdt_found:
        if localname(ch) == 'p':
            t = elem_text(ch)
            print(f'  SDT后段落 body[{i}]: "{t[:40]}"')
