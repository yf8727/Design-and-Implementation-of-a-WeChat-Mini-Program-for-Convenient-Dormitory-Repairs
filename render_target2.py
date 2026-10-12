# -*- coding: utf-8 -*-
import pypdfium2 as pdfium
import os
pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.pdf'
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\integrated_pages'
os.makedirs(outdir, exist_ok=True)
doc = pdfium.PdfDocument(pdf)
print('总页数', len(doc))
for i in range(len(doc)):
    doc[i].render(scale=1.4).to_pil().save(os.path.join(outdir, f'page_{i+1:02d}.png'))
print('渲染完成')
