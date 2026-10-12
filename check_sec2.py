# -*- coding: utf-8 -*-
from docx import Document
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def qn(tag): return W + tag

d = Document(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\初稿1.docx')
ps = d.paragraphs
for idx in [52, 53, 54, 55, 56]:
    p = ps[idx]
    pPr = p._p.find(qn('pPr'))
    has_pbb = pPr is not None and pPr.find(qn('pageBreakBefore')) is not None
    has_sect = pPr is not None and pPr.find(qn('sectPr')) is not None
    # 检查run中的br type=page
    brs = []
    for br in p._p.iter(qn('br')):
        brs.append(br.get(qn('type')))
    print(f'段落[{idx}] "{p.text[:15]}": pageBreakBefore={has_pbb}, sectPr={has_sect}, br={brs}')
# 检查摘要段[51]和[52]是否有分页
for idx in [50, 51]:
    p = ps[idx]
    pPr = p._p.find(qn('pPr'))
    has_pbb = pPr is not None and pPr.find(qn('pageBreakBefore')) is not None
    has_sect = pPr is not None and pPr.find(qn('sectPr')) is not None
    brs = [br.get(qn('type')) for br in p._p.iter(qn('br'))]
    print(f'段落[{idx}] "{p.text[:15]}": pageBreakBefore={has_pbb}, sectPr={has_sect}, br={brs}')
