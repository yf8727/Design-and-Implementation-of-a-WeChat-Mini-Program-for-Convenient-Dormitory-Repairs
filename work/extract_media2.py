# -*- coding: utf-8 -*-
import zipfile, os
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
z = zipfile.ZipFile(path)
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work\media'
names = [n for n in z.namelist() if n.startswith('word/media/')]
print('names:', names)
for n in names:
    data = z.read(n)
    base = os.path.basename(n)
    print('base:', repr(base))
    fn = os.path.join(outdir, base)
    print('fn:', repr(fn))
    with open(fn, 'wb') as f:
        f.write(data)
print('done')
