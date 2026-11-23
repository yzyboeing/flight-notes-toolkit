# -*- coding: utf-8 -*-
"""punct_rule.py —— 表格单元格句末标点规则 A（SD-203，2026-10-17 用户选定「句末标点用 A」）：
   表格单元格（首列、表头、序号列除外）里的每一项（以 <br> 分隔）：
   · 完整句子 → 句末「。」；多项时项间「；」、末项「。」
   · 短语（名词性 / 数值 / 术语 / 条件短语）→ 不加句末标点；多项短语项间「；」、末项不加
   用法：python3 punct_rule.py [--repo 笔记库] [--apply] [--report 文件]
   不加 --apply 只报：「高把握改」= 违反规则 A、可机改；「存疑」= 机器判不准，人工看（不算错误）"""
import io, os, re, glob, sys, json
G = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '~/flight-repos/gh-private'))
STRONG = re.compile(r'(没有|得到|失去|是|为|应|须|必须|不得|不要|禁止|可以|会|将|使|进行|提供|显示|选择|转换|转到|断开|接通|脱开|打开|关闭|亮|灭|出现|消失|保持|执行|完成|检查|核实|确认|报告|申请|使用|调定|设置|置于|放出|收上|恢复|增加|减小|降低|升高|开始|结束|要求|需要|能|应当|导致|造成|影响|取代|代替|移到|移至|产生|允许|继续|停止|改平|截获|预位|生效|失效|计算|判断|避免|防止|确保|实施|处置|操纵|控制|耗|流经|指示|告警|警告|触发|供压|供电|跳开|有效|无效|不可用|可用|不受|受|取决|属于|包括|对应|超过|低于|高于|大于|小于|达到|减至|加至|回到|处于|位于|在 [A-Z]|视为|按照|根据)')
WEAK = re.compile(r'(时|可|需|按|由|向|对|自动|人工|工作|飞行|起飞|着陆|复飞|上升|下降|注意|通过)')
def plain(s): return re.sub(r'<[^>]+>', '', s).replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').strip()
STATE = {'可用', '不可用', '禁止', '允许', '正常', '适用', '不适用', '无', '有', '是', '否', '空白', '接通', '断开', '预位', '开', '关', '亮', '灭', '工作', '不工作', '无效', '有效', '相同', '不同', '—', '-'}
def classify(item):
    kd, p = _classify(item)
    core0 = re.sub(r'^【[^】]*】', '', re.sub(r'[。；;，,]$', '', p)).strip()
    if kd != 'skip' and (core0 in STATE or re.fullmatch(r'[（(][^）)]*[）)]', core0)): return 'P!', p   # SD-272 状态词、整格只有括注：不加句号
    if kd in ('P', 'P?'):
        core = re.sub(r'^【[^】]*】', '', re.sub(r'[。；;，,]$', '', p))
        if len(re.findall(r'[\u4e00-\u9fff]', core)) <= 4 and not STRONG.search(core): return 'P!', p   # SD-275：单个词、状态词、颜色、数值（「空白」「F/D 目标速度」「【737-NG】白色」）永不加句号，也不参与同列多数
    return kd, p
def _classify(item):
    p = plain(item)
    if not p or re.fullmatch(r'[①-⑳\d\s、，,]+|「[^」]*」', p) or p in ('—', '-', '无', '有', '是', '否') or re.search(r'([：:且或及和与、，,]|——)$', p) or re.search(r'[：:][；;。]?$', p): return 'skip', p
    core = re.sub(r'[。；;，,]$', '', p)
    m = re.match(r'^[^，。；：:]{1,14}[：:](.+)$', core)      # 「标签：内容」按内容判断
    body = m.group(1) if m else core
    bv = re.sub(r'（[^）]*）|\([^)]*\)', '', body)                       # 括注里的词不算谓语
    if '→' in bv: return 'S', p          # 11-12：箭头链（A → B → C）是动作序列，按句子加「。」（全书现状 233 : 68）
    bv = re.sub(r'(得到|受到|失去|恢复)(控制|指示|显示)', r'\1', bv)   # 11-12 用户截图 5.1.2：「飞行状态得到控制」是句子，「控制」不是名词
    bv = re.sub(r'(显示|控制|警告|指示|告警|计算|供电|工作|选择|起动|关闭|打开|保持|调定|设置|使用|检查|增压|释压|截获|预位|转换|供压)(组件|器|面板|灯|系统|位置|源|计算机|选择器|信息|方式|电门|汇流条|手柄|页|活门|信号牌|喇叭|程序|单|限制|功能|模式|逻辑|高度|速度|时间|压力|装置|条件|标准|阶段)', '名', bv)   # 「显示组件」「控制面板」「检查单」等是名词
    bv = re.sub(r'(控制|告警|警告|指示|显示|检查)(?=、|$)', '名', bv)       # 「人工增压控制、」列举里的名词
    bv = bv.replace('停止端', '名')
    # 2026-10-29（SD-244，用户看 App 5.4 第 1 条）：故障 / 状态名称（「转换汇流条断开（…）」「驱动装置故障 + BTB 断开」）与「《…》检查单」是短语，不加句号
    _nm = re.sub(r'（[^）]*）|\([^)]*\)', '', body).strip()
    if re.fullmatch(r'(《[^》]+》\s*(检查单)?[、，\s]*)+', _nm) or (re.fullmatch(r'[^，；。,;]{1,24}(断开|故障|失效)', _nm) and not re.search(r'(的|了|会|将|使|后|时)', _nm) and not re.match(r'(出现|发生|有|无)', _nm) and '——' not in _nm): return 'P', p
    # 条件短语：去掉比较词后再无谓语（「RA ≥ 2000ft，且襟翼未收」「低于 200ft RA 至接地 30s 后」）→ 短语
    if not STRONG.search(re.sub(r'(低于|高于|超过|大于|小于|达到|未收|未放|在 [A-Z])', '', bv)) and re.search(r'[≥≤<>]|(时|后|前|期间)$', bv): return 'P', p
    cjk = len(re.findall(r'[\u4e00-\u9fff]', body))
    if re.fullmatch(r'(《[^》]+》[、，\s]*)+', body) or cjk <= 2 or cjk / max(len(body), 1) < 0.15: return 'P', p   # 书名号清单、数值 / 代号为主
    if re.fullmatch(r'[^，。]*?\d[^，。]*?(（[^）]*）)?([；;][^，。]*?\d[^，。]*?(（[^）]*）)?)+', body) and cjk <= 12: return 'P', p   # 「2.0%（在地面）；2.5%（在空中）」
    if re.search(r'[，；。]', bv) and len(bv) >= 8 and STRONG.search(bv): return 'S', p
    if STRONG.search(bv): return ('S' if len(body) >= 10 else 'S?'), p
    if WEAK.search(bv) or re.search(r'[，；。]', bv): return '?', p
    if len(body) <= 8 or re.search(r'(的情况|的项目|等|者|的)$', body): return 'P', p
    return 'P?', p
def pseudo_head(attrs, body):
    """表身里的小表头格（如「失效时机」「目标速度」）：居中且全格加粗、无句末标点，不参与同列统一"""
    return 'center' in attrs and re.fullmatch(r'\s*<strong>[^<]{1,12}</strong>\s*', body) is not None
def fix_item(raw, want_end):
    """raw 是带标签的一项；want_end：'。' / '；' / ''"""
    m = re.search(r'([。；;，,])((?:\s|</[a-z]+>)*)$', raw)
    if m: raw = raw[:m.start(1)] + raw[m.start(1) + 1:]
    if want_end:
        m2 = re.search(r'((?:\s|</[a-z]+>)*)$', raw)
        # 句号放在闭合标签之后（与原文多数写法一致：「……</strong>。」）
        raw = raw.rstrip() + want_end
    return raw
def main():
    apply = '--apply' in sys.argv
    rep = {'改': [], '疑': []}
    global DEC
    dp = os.path.join(G, '句末标点裁定.json')
    DEC = json.load(open(dp, encoding='utf-8')) if os.path.exists(dp) else {}; files = glob.glob(G + '/notes_src/[1-6]*/[0-9]*.md') + ([G + '/速查/速查源.md'] if os.environ.get('BUILD_QR') else [])   # SD-242：速查源只读溯源
    for f in files:
        t = io.open(f, encoding='utf-8').read(); out = t
        def last_kind(body):
            parts = re.split(r'<br\s*/?>\s*\n?', body)
            idx = [x for x in parts if plain(x)]
            if not idx: return None
            kd, p = classify(idx[-1])
            key = re.sub(r'[。；;，,]$', '', p)
            if key in DEC: kd = 'S' if DEC[key] == '。' else 'P'
            return kd
        def do_table(mtab):
            tb = mtab.group(0)
            if re.search(r'<th[^>]*>[^<]*原文', tb): return tb      # 11-12：「QRH 原文」列照手册原文，标点不动
            rows = [r for r in re.finditer(r'<tr([^>]*)>.*?</tr>', tb, re.S) if not re.search(r'class="[^"]*(hdr|premise|note)', r.group(1) or '')]
            # 11-12（用户截图 5.1.2）：同一列句末标点统一。先看整列各格末项的判定
            col = {}; seq_tab = False
            for r in rows:
                cells = re.findall(r'<td[^>]*>(.*?)</td>', r.group(0), re.S)
                if cells and re.fullmatch(r'[①-⑳]|\d+', plain(cells[0])): seq_tab = True
                for k, c in enumerate(cells):
                    if k == 0: continue
                    kd = last_kind(c)
                    if kd and kd != 'skip': col.setdefault(k, []).append(kd)
            # SD-275（2026-11-22 用户：「为什么现在还会有同一列表格中有的有句号有的没有呢？」）：
            # 按网格列（计入 colspan / rowspan 占位）分组，同组（起始列 + 跨列数相同）各格末项句末标点统一
            allrows = list(re.finditer(r'<tr([^>]*)>.*?</tr>', tb, re.S)); occ = set(); POS = []
            for ri, r in enumerate(allrows):
                c_ = 0; pos = []
                for cm in re.finditer(r'<(td|th)([^>]*)>', r.group(0)):
                    while (ri, c_) in occ: c_ += 1
                    cs = int((re.search(r'colspan="(\d+)"', cm.group(2)) or [0, 1])[1]); rs = int((re.search(r'rowspan="(\d+)"', cm.group(2)) or [0, 1])[1])
                    for i_ in range(rs):
                        for j_ in range(cs): occ.add((ri + i_, c_ + j_))
                    if cm.group(1) == 'td': pos.append((c_, cs))
                    c_ += cs
                POS.append(pos)
            GK = {}
            for ri, r in enumerate(allrows):
                if re.search(r'class="[^"]*(hdr|premise|note)', r.group(1) or ''): continue
                if re.match(r'(QRH 状况|目的|QRH 书名)', plain((re.findall(r'<td[^>]*>(.*?)</td>', r.group(0), re.S) or [''])[0])): continue
                for k_, ((gc, cs), (ca, c)) in enumerate(zip(POS[ri], re.findall(r'<td([^>]*)>(.*?)</td>', r.group(0), re.S))):
                    if gc == 0 or pseudo_head(ca, c) or (k_ == 0 and cs < 2): continue
                    kd = last_kind(c)
                    if kd and kd != 'skip': GK.setdefault((gc, cs), []).append(kd)
            def group_want(g):
                ks = GK.get(g, [])
                ks = [x for x in ks if x != 'P!']
                if len(ks) < 2: return None
                nS = sum(1 for x in ks if x in ('S', 'S?')); nP = sum(1 for x in ks if x in ('P', 'P?'))
                if nP == 0: return '。'
                if nS == 0: return ''
                return '。' if nS > nP else ''
            def col_want(k, kd, g=None):
                if kd == 'P!': return ''
                gw = group_want(g) if g else None
                if gw is not None: return gw
                ks = [x for x in col.get(k, []) if x in ('S', 'P')]
                if seq_tab and ks and ks.count('S') != ks.count('P'):
                    return '。' if ks.count('S') > ks.count('P') else ''   # 序号表的同列各项是并列条目：按多数统一（多数是句子 → 整列「。」；多数是短语 → 整列不加）
                if kd in ('S', 'P'): return '。' if kd == 'S' else ''
                return None
            RI = [-1]
            def do_tr(mt):
                RI[0] += 1; pos = POS[RI[0]] if RI[0] < len(POS) else []
                tr = mt.group(0)
                if re.search(r'class="[^"]*(hdr|premise|note)', mt.group(1) or ''): return tr
                cells = list(re.finditer(r'(<td[^>]*>)(.*?)(</td>)', tr, re.S))
                if not cells: return tr
                first_plain = plain(cells[0].group(2))
                if re.match(r'(QRH 状况|目的|QRH 书名)', first_plain): return tr   # 11-12：QRH 状况 / 目的是手册原文照录，标点照原文
                new = tr
                for k, c in enumerate(cells):
                    g = pos[k] if k < len(pos) else None
                    if (g[0] if g else k) == 0: continue   # SD-275：按网格列判首列
                    if k == 0 and not (g and g[1] >= 2): continue   # 上方 rowspan 占位时本行第一个 td 多是小标签（如「复飞前发动机失效」），只有跨列内容格才参与
                    if g and pseudo_head(c.group(1), c.group(2)): continue
                    body = c.group(2)
                    items = re.split(r'(<br\s*/?>\s*\n?)', body)
                    parts = items[0::2]; seps = items[1::2]
                    idx = [i for i, x in enumerate(parts) if plain(x)]
                    if not idx: continue
                    kinds = [classify(parts[i]) for i in idx]
                    if any(k_ == 'skip' for k_, _ in kinds) and len(idx) == 1: continue
                    last_i = idx[-1]
                    for j, i in enumerate(idx):
                        kd, p = kinds[j]
                        if kd == 'skip': continue
                        is_last = (i == last_i)
                        if not is_last and re.match(r'[（(]', plain(parts[idx[j + 1]])): continue   # 1019R6：下一行是括注——续行，不算分项
                        cur_end = p[-1] if p[-1] in '。；;，,' else ''
                        key = re.sub(r'[。；;，,]$', '', p)
                        if is_last:
                            if key in DEC: kd = 'S' if DEC[key] == '。' else 'P'   # 人工裁定优先
                            want = col_want(k, kd, g if g and g[0] > 0 else None)
                        else:
                            want = '；'; kd = 'S' if kd != 'P' else 'P'          # 多项格中间项一律「；」（规则 A）
                        if want is None:
                            k0 = kinds[j][0]
                            if (k0 == 'S?' and cur_end == '。') or (k0 == 'P?' and cur_end == '') or (k0 == '?' and cur_end in ('。', '')): continue
                            rep['疑'].append((os.path.basename(f), first_plain[:16], key, k0, cur_end))
                            continue
                        if cur_end == want or (want == '；' and cur_end == ';'): continue
                        rep['改'].append((os.path.basename(f), first_plain[:16], p[-40:], kd, repr(cur_end), repr(want)))
                        parts[i] = fix_item(parts[i], want)
                    body2 = ''.join(x for pair in zip(parts, seps + ['']) for x in pair)
                    if body2 != body:
                        new = new.replace(c.group(0), c.group(1) + body2 + c.group(3), 1)
                return new
            return re.sub(r'<tr([^>]*)>.*?</tr>', do_tr, tb, flags=re.S)
        # 11-12：第五章 <!-- callouts:… --> 旁注表由叠加层生成，不在这里改（改叠加层的 short / full）
        segs = re.split(r'(<!-- callouts:.*?<!-- /callouts:[^>]*-->)', t, flags=re.S)
        out = ''.join(sg if sg.startswith('<!-- callouts:') else re.sub(r'<table.*?</table>', do_table, sg, flags=re.S) for sg in segs)
        if apply and out != t: io.open(f, 'w', encoding='utf-8').write(out)
    print('高把握改', len(rep['改']), '存疑', len(rep['疑']))
    from collections import Counter
    print(Counter((r[4], r[5]) for r in rep['改']).most_common(8))
    if '--report' in sys.argv:
        json.dump(rep, open(sys.argv[sys.argv.index('--report') + 1], 'w'), ensure_ascii=False, indent=0)

if __name__ == '__main__':
    main()
