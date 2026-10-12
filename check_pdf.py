# -*- coding: utf-8 -*-
import pdfplumber
pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.pdf'
with pdfplumber.open(pdf) as pdfdoc:
    print('总页数:', len(pdfdoc.pages))
    for pi in [1, 2]:
        t = pdfdoc.pages[pi].extract_text() or ''
        print(f'===== 物理页{pi+1} =====')
        print(t[:800])
        print()
