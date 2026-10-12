# -*- coding: utf-8 -*-
import zipfile, os
path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\专科毕业论文（设计）模板及要求(1).docx'
z = zipfile.ZipFile(path)
outdir = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work\media'
for n in z.namelist():
    if n.startswith('word/media/') and not n.endswith('/'):
        data = z.read(n)
        fn = os.path.join(outdir, os.path.basename(n))
        with open(fn, 'wb') as f:
            f.write(data)
        print(os.path.basename(n), len(data))
print('done')
