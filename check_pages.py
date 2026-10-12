# -*- coding: utf-8 -*-
import pdfplumber
pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.pdf'
with pdfplumber.open(pdf) as d:
    print('总页数', len(d.pages))
    for i, pg in enumerate(d.pages):
        t = pg.extract_text() or ''
        lines = [l.strip() for l in t.split('\n') if l.strip()]
        foot = lines[-1] if lines else ''
        first = lines[0][:35] if lines else ''
        has_h1 = [ln for ln in lines if ln.startswith(('1 绪论','2 相关技术','3 系统需求','4 系统设计','5 系统实现','6 系统测试','7 总结'))]
        print(f'PDF页{i+1}: 页脚={foot!r} | 首行={first!r} | H1={has_h1}')
