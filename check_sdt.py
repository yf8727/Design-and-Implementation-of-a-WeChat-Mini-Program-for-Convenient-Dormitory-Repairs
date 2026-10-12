# -*- coding: utf-8 -*-
import docx
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.docx'
doc = docx.Document(path)
body = doc.element.body
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

for ch in body:
    if etree.QName(ch).localname == 'sdt':
        for p in ch.iter(W + 'p'):
            texts = ''.join(t.text or '' for t in p.iter(W + 't'))
            if texts.strip():
                print(repr(texts))
