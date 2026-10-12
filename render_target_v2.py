# -*- coding: utf-8 -*-
import pypdfium2 as pdfium
import os, sys, shutil

pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.pdf'
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\target_pages_v2'
if os.path.exists(outdir):
    shutil.rmtree(outdir)
os.makedirs(outdir, exist_ok=True)

doc = pdfium.PdfDocument(pdf)
n = len(doc)
print('总页数:', n)
for i in range(n):
    page = doc[i]
    bitmap = page.render(scale=1.4)
    img = bitmap.to_pil()
    out = os.path.join(outdir, f'page_{i+1:02d}.png')
    img.save(out)
print('渲染完成:', outdir)
