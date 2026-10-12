# -*- coding: utf-8 -*-
import zipfile
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    docxml = z.read('word/document.xml').decode('utf-8')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root = etree.fromstring(docxml.encode('utf-8'))
body = root.find(W + 'body')
sdt = None
for child in body:
    if etree.QName(child).localname == 'sdt':
        sdt = child
        break
sdtContent = sdt.find(W + 'sdtContent')

for i in [0, 1]:
    p = sdtContent[i]
    full = etree.tostring(p, encoding='unicode')
    # 去掉所有命名空间声明以便阅读
    import re
    clean = re.sub(r'\sxmlns(:\w+)?="[^"]*"', '', full)
    print(f'===== sdtContent[{i}] 完整XML =====')
    print(clean)
    print()
