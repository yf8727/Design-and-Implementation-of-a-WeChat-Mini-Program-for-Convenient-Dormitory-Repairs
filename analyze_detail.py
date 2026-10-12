# -*- coding: utf-8 -*-
import docx
from lxml import etree

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.docx'
doc = docx.Document(path)
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 需要详细查看的段落索引
targets = [14, 29, 37, 38, 39, 40, 41, 42, 43, 44, 52, 53, 56, 59, 63, 64, 65, 66, 67, 68, 74, 75, 89, 90, 91, 92, 93, 94, 101]

for i in targets:
    p = doc.paragraphs[i]
    pPr = p._element.find(W + 'pPr')
    style = ''
    spacing = ''
    ind = ''
    if pPr is not None:
        st = pPr.find(W + 'pStyle')
        if st is not None:
            style = st.get(W + 'val')
        sp = pPr.find(W + 'spacing')
        if sp is not None:
            spacing = f'line={sp.get(W+"line")} lineRule={sp.get(W+"lineRule")} before={sp.get(W+"before")} after={sp.get(W+"after")}'
        in_ = pPr.find(W + 'ind')
        if in_ is not None:
            ind = f'firstLine={in_.get(W+"firstLine")} firstLineChars={in_.get(W+"firstLineChars")} left={in_.get(W+"left")}'
    # 第一个run的rPr
    rpr_info = ''
    runs = p._element.findall(W + 'r')
    if runs:
        r0 = runs[0]
        rPr = r0.find(W + 'rPr')
        if rPr is not None:
            fonts = rPr.find(W + 'rFonts')
            sz = rPr.find(W + 'sz')
            b = rPr.find(W + 'b')
            if fonts is not None:
                rpr_info += f'fonts(ea={fonts.get(W+"eastAsia")} ascii={fonts.get(W+"ascii")}) '
            if sz is not None:
                rpr_info += f'sz={sz.get(W+"val")} '
            if b is not None:
                rpr_info += 'bold '
    # 是否有jc
    jc = ''
    if pPr is not None:
        j = pPr.find(W + 'jc')
        if j is not None:
            jc = f'jc={j.get(W+"val")}'
    print(f'[{i}] style={style} | {spacing} | {ind} | {jc} | run: {rpr_info} | text="{p.text[:40]}"')
