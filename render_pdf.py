# -*- coding: utf-8 -*-
import pypdfium2 as pdfium
import os

pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.pdf'
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\pages'
os.makedirs(outdir, exist_ok=True)

doc = pdfium.PdfDocument(pdf)
n = len(doc)
print('总页数:', n)
for i in range(n):
    page = doc[i]
    bitmap = page.render(scale=1.5)
    img = bitmap.to_pil()
    out = os.path.join(outdir, f'page_{i+1:02d}.png')
    img.save(out)
    print('已保存:', out, img.size)
