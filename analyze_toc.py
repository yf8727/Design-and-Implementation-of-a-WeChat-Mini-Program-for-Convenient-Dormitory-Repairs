# -*- coding: utf-8 -*-
import zipfile

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    docxml = z.read('word/document.xml').decode('utf-8')

for keyword in ['一级标题', '研究思路']:
    print(f'===== 搜索 "{keyword}" =====')
    idx = 0
    while True:
        i = docxml.find(keyword, idx)
        if i == -1:
            break
        start = max(0, i-600)
        end = min(len(docxml), i+200)
        print(f'--- 位置 {i} ---')
        print(docxml[start:end])
        print()
        idx = i + 1
