# -*- coding: utf-8 -*-
import docx

path = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\初稿1.docx'
d = docx.Document(path)
ps = d.paragraphs

targets = {
    51: '两类用户',
    62: '陈淼',
    70: '[8]',
    71: '[9]',
    76: '贾志杰',
    92: '（1）性能需求',
    93: '（2）安全需求',
    94: '（3）易用性需求',
    95: '（4）可维护性',
    133: '报修数据统计',
    158: '陈淼',
}
for idx, sub in targets.items():
    p = ps[idx]
    runs = p.runs
    print(f'--- 段落[{idx}] runs={len(runs)} 文本前40: "{p.text[:40]}"')
    for ri, r in enumerate(runs):
        if sub in r.text:
            print(f'    run[{ri}] 含目标: "{r.text[:80]}"')
    # 检查子串是否跨run
    full = p.text
    pos = full.find(sub)
    if pos >= 0:
        # 找到pos对应的run
        cum = 0
        for ri, r in enumerate(runs):
            cum += len(r.text)
            if cum > pos:
                # 检查子串是否完全在一个run内
                run_start = cum - len(r.text)
                if run_start + len(r.text) >= pos + len(sub):
                    pass
                break
    # 打印所有run文本（短）
    print('    全部runs:', [r.text for r in runs][:12])
