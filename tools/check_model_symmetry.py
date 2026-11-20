#!/usr/bin/env python3
"""K-S37（原称 K-S23，11-19 改号避开 check_src 的 K-S23）：对生成器逐格导出检查两型文字对称，以及两型对比表（含共有格）样式。
11-20 用户（SD-268，取代 SD-255 旧口径）：两型对比表（含单型列与跨两列共有格）按格内内容统一判定——
单独的单一数值 / 范围 /「—」占位格居中无点，其余文字描述格靠左加点。与 build_docx.js isSingleVal 同口径。"""
import argparse,collections,json,re
from pathlib import Path

def load(path):
    if path.is_dir():
        path, = path.glob('cells_*.jsonl')
    return [t for line in path.read_text().splitlines() if 'cells' in (t:=json.loads(line))]

def model(label):
    label = re.sub(r'[\s\u200b\u2060\u00ad]', '', label).upper()
    ng=bool(re.search(r'737[-‐‑–—]?NG',label))
    eight=bool(re.search(r'737[-‐‑–—]?(?:8(?![0-9])|MAX)',label))
    return 'ng' if ng and not eight else '8' if eight and not ng else None

def is_single_val(cell):
    """与 build_docx.js isSingleVal 同口径：单段、无中文与分号，纯「—」占位或单独单一数值/范围。"""
    if len([p for p in cell.get('items', []) if p.get('html', '').strip()]) != 1:
        return False
    t = re.sub(r'[\s\u200b\u2060\u00ad]', '', cell['t'])
    if not t:
        return False
    if re.fullmatch(r'[—－\-–/／]+', t):
        return True
    return bool(re.search(r'\d', t) and re.fullmatch(r'[+\-−–~～≤≥<>±×·.,/°%℃℉a-zA-Z0-9()（）]+', t))

def is_value(cell):
    return is_single_val(cell)

def signature(cell):
    return sorted({p['align'] for p in cell['items']}), any(p['bullet']=='•' for p in cell['items'])

def scan(tables):
    total=pairs=skipped=values=0
    value_errors=[]
    misses=[]
    for table in tables:
        columns=collections.defaultdict(set)
        rows=collections.defaultdict(dict)
        for c in table['cells']:
            if c['h'] and c['cs']==1 and (m:=model(c['t'])):
                columns[m].add(c['c'])
            if c['kind']=='data' and not c['ghost']:
                rows[c['r']][c['c']]=c
        if not columns['ng'] or not columns['8']:
            continue
        total+=1
        model_columns = columns['ng'] | columns['8']
        for row_cells in rows.values():
            for col, cell in row_cells.items():
                hits_model = (col in model_columns and cell['cs'] == 1) or (cell['cs'] >= 2 and any(j in model_columns for j in range(col, col + cell['cs'])))
                if hits_model and is_single_val(cell):
                    if signature(cell) != (['center'], False):   # SD-268：单独单一数值 / 范围 /「—」居中无点
                        value_errors.append({'item':table['k'], 'source':table['src'], 'row':cell['r'],
                                             'column':col,'text':cell['t'],'style':signature(cell)})
        for row, cells in rows.items():
            for ng in sorted(columns['ng']):
                for eight in sorted(columns['8']):
                    a,b=cells.get(ng),cells.get(eight)
                    if not a or not b or a['cs']!=1 or b['cs']!=1 or not a['t'].strip() or not b['t'].strip():
                        skipped+=1
                        continue
                    if is_value(a) or is_value(b):
                        values+=1
                        continue
                    pairs+=1
                    if signature(a)!=signature(b):
                        misses.append({'item':table['k'],'table':table['ti'],'source':table['src'],'row':row,'ng':a['t'],'8':b['t'],
                                       'ngStyle':signature(a),'8Style':signature(b)})
    return {'tables':total,'pairs':pairs,'sharedOrUnpaired':skipped,'excludedValues':values,'valueErrors':value_errors,'mismatches':misses}

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('dump',type=Path)
    p.add_argument('--report',type=Path)
    args=p.parse_args()
    report=scan(load(args.dump))
    if args.report:
        args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print('K-S37', {k:len(v) if isinstance(v,list) else v for k,v in report.items()})
    raise SystemExit(bool(report['mismatches'] or report['valueErrors']))
