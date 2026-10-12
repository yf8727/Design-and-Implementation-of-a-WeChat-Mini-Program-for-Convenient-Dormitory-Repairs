# -*- coding: utf-8 -*-
import zipfile, re
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
z = zipfile.ZipFile(path)
xml = z.read('word/document.xml').decode('utf-8')

print("===== 原始 sectPr 块 =====")
for m in re.finditer(r'<w:sectPr[^>]*>.*?</w:sectPr>|<w:sectPr[^>]*/>', xml, re.S):
    print(m.group(0))
    print('---')

print()
print("===== 图片引用位置（drawing 前后的文本上下文） =====")
# 简化：找所有 blip 的 rId
for m in re.finditer(r'r:embed="(rId\d+)"', xml):
    print(f'  image ref: {m.group(1)} at pos {m.start()}')
# 打印含drawing的段落文本
docxml = z.read('word/document.xml').decode('utf-8')
# 找到每个 w:drawing 附近的文本
for m in re.finditer(r'<w:drawing>.*?</w:drawing>', docxml, re.S):
    before = docxml[max(0,m.start()-300):m.start()]
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', before)
    print(f'  drawing前文本: {texts[-3:]}')
