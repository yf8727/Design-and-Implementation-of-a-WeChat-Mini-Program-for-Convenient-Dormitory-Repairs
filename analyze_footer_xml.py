# -*- coding: utf-8 -*-
import zipfile
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    for fn in ['word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml']:
        xml = z.read(fn).decode('utf-8')
        print(f'===== {fn} =====')
        print(xml)
        print()
