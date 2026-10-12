# -*- coding: utf-8 -*-
import docx
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\TARGET.docx'
doc = docx.Document(path)
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
print('=== 行长>40且非440行距的段落 ===')
for i, p in enumerate(doc.paragraphs):
    t = p.text.strip()
    if t and len(t) > 40 and not t.startswith('['):
        sp = p._element.find(W + 'pPr/' + W + 'spacing')
        line = sp.get(W + 'line') if sp is not None else None
        rule = sp.get(W + 'lineRule') if sp is not None else None
        if not (line == '440' and rule == 'exact'):
            print(f'  段落[{i}]: 行距={line}/{rule} 文本前30: {t[:30]}')
