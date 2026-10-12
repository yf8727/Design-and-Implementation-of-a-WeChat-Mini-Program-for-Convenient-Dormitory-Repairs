# -*- coding: utf-8 -*-
import zipfile
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    docxml = z.read('word/document.xml').decode('utf-8')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root = etree.fromstring(docxml.encode('utf-8'))
body = root.find(W + 'body')

# 找到 sdt
sdt = None
for child in body:
    if etree.QName(child).localname == 'sdt':
        sdt = child
        break

print('=== SDT结构 ===')
sdtContent = sdt.find(W + 'sdtContent')
print('sdtContent子元素数量:', len(sdtContent))
for i, child in enumerate(sdtContent):
    tag = etree.QName(child).localname
    if tag == 'p':
        texts = ''.join(t.text or '' for t in child.iter(W + 't'))
        print(f'  [{i}] <p> "{texts}"')
        # 打印pPr
        pPr = child.find(W + 'pPr')
        if pPr is not None:
            print('     pPr XML:', etree.tostring(pPr, encoding='unicode')[:500])
        # 打印第一个和最后一个run
        runs = child.findall(W + 'r')
        if runs:
            r0 = runs[0]
            print('     run0 XML:', etree.tostring(r0, encoding='unicode')[:600])
    else:
        print(f'  [{i}] <{tag}>')
