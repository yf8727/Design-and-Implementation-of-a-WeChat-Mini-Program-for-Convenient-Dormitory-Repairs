# -*- coding: utf-8 -*-
import pypdfium2 as pdfium
import os, sys

def render(pdf, tag):
    outdir = os.path.join(os.path.dirname(pdf), tag)
    os.makedirs(outdir, exist_ok=True)
    doc = pdfium.PdfDocument(pdf)
    n = len(doc)
    print(f'{tag}: 总页数 {n}')
    for i in range(n):
        page = doc[i]
        bitmap = page.render(scale=1.4)
        img = bitmap.to_pil()
        out = os.path.join(outdir, f'page_{i+1:02d}.png')
        img.save(out)
    print(f'{tag} 渲染完成 -> {outdir}')

render(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\初稿1.pdf', 'draft1_pages')
render(r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\初稿2.pdf', 'draft2_pages')
