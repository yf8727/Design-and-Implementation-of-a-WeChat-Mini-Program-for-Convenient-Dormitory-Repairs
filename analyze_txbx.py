# -*- coding: utf-8 -*-
import zipfile, re, os

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
with zipfile.ZipFile(path) as z:
    names = z.namelist()
    print('=== zip内的xml文件 ===')
    for n in names:
        if n.endswith('.xml'):
            print(' ', n)
    docxml = z.read('word/document.xml').decode('utf-8')

print()
print('document.xml长度:', len(docxml))
print('含 txbxContent 次数:', docxml.count('txbxContent'))
print('含 "一级标题" 次数:', docxml.count('一级标题'))
print('含 "研究思路" 次数:', docxml.count('研究思路'))
print('含 "说明" 次数:', docxml.count('说明'))

# 找到txbxContent的位置
print()
idx = 0
count = 0
while True:
    i = docxml.find('txbxContent', idx)
    if i == -1:
        break
    count += 1
    start = max(0, i-100)
    end = min(len(docxml), i+500)
    print(f'--- txbxContent #{count} at {i} ---')
    print(docxml[start:end])
    print()
    idx = i + 1
    if count > 10:
        break
