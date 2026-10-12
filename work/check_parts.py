# -*- coding: utf-8 -*-
import zipfile
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
z = zipfile.ZipFile(path)
print("===== footer1.xml =====")
print(z.read('word/footer1.xml').decode('utf-8'))
print()
print("===== header1.xml =====")
print(z.read('word/header1.xml').decode('utf-8'))
print()
print("===== styles.xml 中 Heading/TOC 相关样式 =====")
styles = z.read('word/styles.xml').decode('utf-8')
import re
for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', styles, re.S):
    sid, body = m.group(1), m.group(2)
    name = re.search(r'<w:name w:val="([^"]+)"', body)
    if any(k in (sid.lower() + ' ' + (name.group(1) if name else '')) for k in ['heading', 'toc', 'title']):
        print(f'  styleId={sid} name={name.group(1) if name else None}')
        # 关键属性
        fonts = re.findall(r'<w:rFonts[^/]*/>', body)
        sz = re.findall(r'<w:sz w:val="([^"]+)"', body)
        color = re.findall(r'<w:color w:val="([^"]+)"', body)
        print(f'    fonts={fonts} sz={sz} color={color}')
print()
print("===== settings.xml 开头 =====")
settings = z.read('word/settings.xml').decode('utf-8')
print(settings[:2000])
