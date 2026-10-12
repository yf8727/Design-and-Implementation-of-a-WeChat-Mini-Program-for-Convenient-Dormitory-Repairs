# -*- coding: utf-8 -*-
import zipfile
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    docxml = z.read('word/document.xml').decode('utf-8')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root = etree.fromstring(docxml.encode('utf-8'))
body = root.find(W + 'body')

print('=== body直接子元素序列 ===')
for i, child in enumerate(body):
    tag = etree.QName(child).localname
    if tag == 'p':
        texts = ''.join(t.text or '' for t in child.iter(W + 't'))
        has_sect = child.find(W + 'pPr/' + W + 'sectPr') is not None
        print(f'[{i}] <p> sectPr={has_sect} "{texts[:40]}"')
    elif tag == 'tbl':
        print(f'[{i}] <tbl>')
    elif tag == 'sdt':
        # 内容控件
        texts = ''.join(t.text or '' for t in child.iter(W + 't'))
        print(f'[{i}] <sdt> "{texts[:40]}"')
    else:
        print(f'[{i}] <{tag}>')
