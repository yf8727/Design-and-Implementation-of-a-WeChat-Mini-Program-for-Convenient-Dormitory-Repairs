# -*- coding: utf-8 -*-
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
doc = docx.Document(path)
body = doc.element.body

idx = 0
sec = 0
for child in body.iterchildren():
    tag = child.tag.split('}')[1]
    if tag == 'p':
        p = Paragraph(child, doc)
        # 查找 drawing 中的 rId
        blips = child.findall('.//' + qn('a:blip'))
        for b in blips:
            rid = b.get(qn('r:embed'))
            print(f'  [P{idx}] 节{sec} 含图片 rId={rid}')
        # 检查 sectPr
        if child.find(qn('w:pPr')) is not None and child.find(qn('w:pPr')).find(qn('w:sectPr')) is not None:
            sec += 1
        idx += 1
    elif tag == 'tbl':
        # 表格内图片
        blips = child.findall('.//' + qn('a:blip'))
        for b in blips:
            rid = b.get(qn('r:embed'))
            print(f'  [TABLE 节{sec}] 含图片 rId={rid}')
    elif tag == 'sectPr':
        sec += 1

# 页眉里的图片
for i, sec_obj in enumerate(doc.sections):
    hdr = sec_obj.header
    if hdr is not None and hdr.is_linked_to_previous is False:
        for b in hdr._element.findall('.//' + qn('a:blip')):
            print(f'  [Section{i} HEADER] 图片 rId={b.get(qn("r:embed"))}')
