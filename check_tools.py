# -*- coding: utf-8 -*-
import subprocess, os, glob

# 尝试用 pdftoppm / mutool 渲染；若都没有，用 LibreOffice 转 png 不现实，改用 pdf2image 检查
pdf = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\SOURCE.pdf'
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work_tmp\pages'
os.makedirs(outdir, exist_ok=True)

def find_tool(name):
    for p in os.environ.get('PATH','').split(';'):
        cand = os.path.join(p, name)
        if os.path.exists(cand):
            return cand
    return None

# 检查可能工具
for t in ['pdftoppm.exe','mutool.exe','magick.exe','gs.exe']:
    print(t, '->', find_tool(t))
