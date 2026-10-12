# -*- coding: utf-8 -*-
import docx
from docx.oxml.ns import qn
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
doc = docx.Document(path)

print("===== 封面 题目 段落 (P14) runs 详情 =====")
for i, p in enumerate(doc.paragraphs):
    if '题' in p.text[:3]:
        print(f'P{i}: text={p.text!r}')
        for j, r in enumerate(p.runs):
            rPr = r._element.rPr
            ea = rPr.rFonts.get(qn('w:eastAsia')) if (rPr is not None and rPr.rFonts is not None) else None
            print(f'  run{j}: {r.text!r} font={r.font.name}/{ea} sz={r.font.size}')
        # 段落属性
        pPr = p._element.pPr
        if pPr is not None:
            import lxml.etree as etree
            print('  pPr xml:', etree.tostring(pPr, pretty_print=True).decode()[:800])
        break

print()
print("===== 评阅表 (doc.tables[1]) 单元格结构 =====")
tbl = doc.tables[1]
for r_i, row in enumerate(tbl.rows):
    seen = set()
    for c in row.cells:
        if id(c._tc) in seen: continue
        seen.add(id(c._tc))
        print(f'  Row{r_i} cell: {c.text[:40]!r}')
        # 打印该单元格首个段落 runs
        for p in c.paragraphs[:2]:
            for r in p.runs:
                print(f'     run: {r.text!r}')
