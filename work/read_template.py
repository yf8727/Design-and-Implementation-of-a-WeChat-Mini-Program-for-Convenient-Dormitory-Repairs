# -*- coding: utf-8 -*-
"""读取模板docx的完整内容与结构"""
import docx
from docx.document import Document as _Doc
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
doc = docx.Document(path)

def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

print('===== 文档内容（按文档顺序，含表格）=====')
idx_p = 0
for block in iter_block_items(doc):
    if isinstance(block, Paragraph):
        t = block.text
        style = block.style.name
        # 检查段落格式
        pf = block.paragraph_format
        align = pf.alignment
        print(f'[P] style={style} align={align} | {t}')
        idx_p += 1
    else:
        print(f'===== [TABLE] rows={len(block.rows)} cols={len(block.columns)} =====')
        for r_i, row in enumerate(block.rows):
            cells = []
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(c.text.replace('\n', '⏎'))
            print(f'  R{r_i}: ' + ' || '.join(cells))

print()
print('===== 节/页面设置 =====')
for i, sec in enumerate(doc.sections):
    print(f'Section {i}: page={sec.page_width}x{sec.page_height} margin L={sec.left_margin} R={sec.right_margin} T={sec.top_margin} B={sec.bottom_margin}')

print()
print('===== 样式定义 =====')
for s in doc.styles:
    try:
        if s.type is not None and s.type in (1, 2, 3, 4, 5):
            f = s.font
            print(f'style={s.name} type={s.type} font={f.name} size={f.size}')
    except Exception:
        pass
