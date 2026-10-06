#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_adjudicated.py —— 复核外部 AI 的内容报告前，先按《内容裁定台账》过滤重复报（2026-10-05）
用法：python3 check_adjudicated.py <报告.md> [<报告2.md> …]
对台账每条，从「地址」和「理由」里取关键词（节号、条目号、数值、专名），在报告里找命中的行，
列出「疑似重复报已裁定」的候选供人工确认。只读，不改文件。"""
import sys, os, re, json

LEDGER = os.path.expanduser('~/flight-repos/gh-private/内容裁定台账.json')   # 正本在私有库根目录；交接包与噜噜镜像里是副本

def keys_of(e):
    t = e['地址']
    ks = set(re.findall(r'\d\.\d+\s*[A-H]?-?\d*', t)) | set(re.findall(r'M1[4-9]-[A-Z]?\d+', t + e.get('轮', '')))
    ks |= set(re.findall(r'[A-Z]{2,}[  ]?[A-Z/]*', t)) - {'PDF', 'FCOM', 'QRH', 'AFM', 'SOP', 'MEL'}
    ks |= set(re.findall(r'\d+(?:±\d+)?(?:ft|kt|psi|℃|m|s|min|V|Hz)', t))
    for w in re.findall(r'[一-鿿]{4,8}', t): ks.add(w)
    return {k.strip() for k in ks if len(k.strip()) >= 3}

def main():
    if len(sys.argv) < 2: print(__doc__); return
    ledger = json.load(open(LEDGER, encoding='utf-8'))
    for rp in sys.argv[1:]:
        lines = open(rp, encoding='utf-8').read().split('\n')
        print('== %s（台账 %d 条）' % (os.path.basename(rp), len(ledger)))
        hits = 0
        for e in ledger:
            ks = keys_of(e)
            for n, l in enumerate(lines, 1):
                got = [k for k in ks if k in l]
                if len(got) >= 2 or (len(got) == 1 and re.search(r'\d\.\d', got[0])):
                    print('  ? 第 %d 行命中〔%s｜%s（%s）〕：%s' % (n, e['地址'][:24], e['结论'], e['轮'], l.strip()[:60]))
                    hits += 1; break
        print('  共 %d 条疑似重复报已裁定（人工确认后按台账驳回）' % hits)

if __name__ == '__main__': main()
