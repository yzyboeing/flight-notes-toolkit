#!/usr/bin/env python3
"""K-S23：对生成器逐格导出检查两型文字对称，以及取值／占位格居中无点。"""
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

def is_value(cell):
    text = re.sub(r'[\s\u200b\u2060\u00ad]', '', cell['t'])
    text = re.sub(r'（[^）]{0,14}）|\([^)]{0,14}\)', '', text)
    if re.fullmatch(r'[—－\-–]', text):
        return True
    parts = [x for x in re.split(r'[;；]', text) if x]
    for value in parts:
        numeric = re.sub(r'^(约|至少|不少于|不超过|最多|最少|小于|大于)', '', value)
        numeric = re.sub(r'千米|公里|英尺|夸脱|分钟|小时|米|磅|升|秒|天|度|个', 'u', numeric)
        if not re.search(r'\d', numeric) or not re.fullmatch(r'[\d.,:：~～+\-−–≤≥<>≦≧≈=%/°×±℃℉a-zA-Z·→]*', numeric):
            return False
    return bool(parts)

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
                if col in model_columns and cell['cs'] == 1 and is_value(cell):
                    if signature(cell) != (['center'], False):
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
    print('K-S23', {k:len(v) if isinstance(v,list) else v for k,v in report.items()})
    raise SystemExit(bool(report['mismatches'] or report['valueErrors']))
