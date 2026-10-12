# -*- coding: utf-8 -*-
from docx import Document
from lxml import etree

d = Document(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\初稿1.docx')
ps = d.paragraphs
for idx in [53, 54, 55, 56, 57]:
    p = ps[idx]
    pPr = p._p.pPr
    xml = etree.tostring(pPr, pretty_print=True).decode() if pPr is not None else 'None'
    print(f'===== 段落[{idx}] "{p.text[:20]}" pPr =====')
    print(xml[:1000])
    print()
