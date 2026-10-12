# -*- coding: utf-8 -*-
import docx
from lxml import etree

path1 = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\初稿1.docx'
path2 = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\初稿2.docx'

for name, path in [('初稿1', path1), ('初稿2', path2)]:
    d = docx.Document(path)
    print(f'========== {name} 结构分析 ==========')
    body = d.element.body
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    # 检查图片
    imgs = body.findall('.//' + W + 'drawing')
    print(f'  图片数量(drawing): {len(imgs)}')
    # 检查分节
    sect_count = 0
    for ch in body:
        if etree.QName(ch).localname == 'p':
            sp = ch.find(W + 'pPr/' + W + 'sectPr')
            if sp is not None:
                sect_count += 1
    print(f'  分节数: {sect_count}')
    # 检查目录位置
    for i, p in enumerate(d.paragraphs):
        if '目录' in p.text or '打开文档' in p.text:
            print(f'  目录相关段落[{i}]: "{p.text[:30]}"')
    # 检查是否包含sdt
    sdts = body.findall('.//' + W + 'sdt')
    print(f'  sdt数量: {len(sdts)}')
    print()
