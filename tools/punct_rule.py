# -*- coding: utf-8 -*-
"""punct_rule.py —— 表格单元格句末标点规则 A（SD-203，2026-10-17 用户选定「句末标点用 A」）：
   表格单元格（首列、表头、序号列除外）里的每一项（以 <br> 分隔）：
   · 完整句子 → 句末「。」；多项时项间「；」、末项「。」
   · 短语（名词性 / 数值 / 术语 / 条件短语）→ 不加句末标点；多项短语项间「；」、末项不加
   用法：python3 punct_rule.py [--repo 笔记库] [--apply] [--report 文件]
   不加 --apply 只报：「高把握改」= 违反规则 A、可机改；「存疑」= 机器判不准，人工看（不算错误）"""
import io, os, re, glob, sys, json
G = os.path.abspath(os.path.expanduser(sys.argv[sys.argv.index('--repo') + 1] if '--repo' in sys.argv else '~/flight-repos/gh-private'))
STRONG = re.compile(r'(是|为|应|须|必须|不得|不要|禁止|可以|会|将|使|进行|提供|显示|选择|转换|转到|断开|接通|脱开|打开|关闭|亮|灭|出现|消失|保持|执行|完成|检查|核实|确认|报告|申请|使用|调定|设置|置于|放出|收上|恢复|增加|减小|降低|升高|开始|结束|要求|需要|能|应当|导致|造成|影响|取代|代替|移到|移至|产生|允许|继续|停止|改平|截获|预位|生效|失效|计算|判断|避免|防止|确保|实施|处置|操纵|控制|耗|流经|指示|告警|警告|触发|供压|供电|跳开|有效|无效|不可用|可用|不受|受|取决|属于|包括|对应|超过|低于|高于|大于|小于|达到|减至|加至|回到|处于|位于|在 [A-Z]|视为|按照|根据)')
WEAK = re.compile(r'(时|可|需|按|由|向|对|自动|人工|工作|飞行|起飞|着陆|复飞|上升|下降|注意|通过)')
def plain(s): return re.sub(r'<[^>]+>', '', s).replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').strip()
def classify(item):
    p = plain(item)
    if not p or p in ('—', '-', '无', '有', '是', '否') or re.search(r'[：:且或及和与、，,]$', p) or re.search(r'[：:][；;。]?$', p): return 'skip', p
    core = re.sub(r'[。；;，,]$', '', p)
    m = re.match(r'^[^，。；：:]{1,14}[：:](.+)$', core)      # 「标签：内容」按内容判断
    body = m.group(1) if m else core
    bv = re.sub(r'（[^）]*）|\([^)]*\)', '', body)                       # 括注里的词不算谓语
    bv = re.sub(r'(显示|控制|警告|指示|告警|计算|供电|工作|选择|起动|关闭|打开|保持|调定|设置|使用|检查|增压|释压|截获|预位|转换|供压)(组件|器|面板|灯|系统|位置|源|计算机|选择器|信息|方式|电门|汇流条|手柄|页|活门|信号牌|喇叭|程序|单|限制|功能|模式|逻辑|高度|速度|时间|压力|装置|条件|标准|阶段)', '名', bv)   # 「显示组件」「控制面板」「检查单」等是名词
    bv = re.sub(r'(控制|告警|警告|指示|显示|检查)(?=、|$)', '名', bv)       # 「人工增压控制、」列举里的名词
    bv = bv.replace('停止端', '名')
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
    DEC = json.load(open(dp, encoding='utf-8')) if os.path.exists(dp) else {}; files = glob.glob(G + '/notes_src/[1-5]*/[0-9]*.md') + [G + '/速查/速查源.md']
    for f in files:
        t = io.open(f, encoding='utf-8').read(); out = t
        def do_tr(mt):
            tr = mt.group(0)
            if re.search(r'class="[^"]*(hdr|premise|note)', mt.group(1) or ''): return tr
            cells = list(re.finditer(r'(<td[^>]*>)(.*?)(</td>)', tr, re.S))
            if not cells: return tr
            first_plain = plain(cells[0].group(2))
            seq = bool(re.fullmatch(r'[①-⑳]|\d+', first_plain))
            new = tr; shift = 0
            for k, c in enumerate(cells):
                if k == 0: continue
                if 'colspan' in c.group(1) and k == 1 and len(cells) == 2 and False: pass
                body = c.group(2)
                items = re.split(r'(<br\s*/?>\s*\n?)', body)
                parts = items[0::2]; seps = items[1::2]
                idx = [i for i, x in enumerate(parts) if plain(x)]
                if not idx: continue
                kinds = [classify(parts[i]) for i in idx]
                if any(k_ == 'skip' for k_, _ in kinds) and len(idx) == 1: continue
                # 有分项引导语（「…：」开头的项）或含「–」子项的格，结构复杂，只处理末项
                last_i = idx[-1]
                for j, i in enumerate(idx):
                    kd, p = kinds[j]
                    if kd == 'skip': continue
                    is_last = (i == last_i)
                    cur_end = p[-1] if p[-1] in '。；;，,' else ''
                    key = re.sub(r'[。；;，,]$', '', p)
                    if is_last:
                        want = '。' if kd in ('S',) else ('' if kd in ('P',) else None)
                        if want is None and key in DEC: want = '。' if DEC[key] == '。' else ''; kd = 'S' if want else 'P'
                    else:
                        want = '；'; kd = 'S' if kd != 'P' else 'P'          # 多项格中间项一律「；」（规则 A）
                    if want is None:
                        k0 = kinds[j][0]
                        # 现状已与倾向一致（S? 有句号、P? 无标点）或无倾向（?）且不是「；」结尾的，保持现状；只有冲突的进存疑
                        if (k0 == 'S?' and cur_end == '。') or (k0 == 'P?' and cur_end == '') or (k0 == '?' and cur_end in ('。', '')): continue
                        rep['疑'].append((os.path.basename(f), first_plain[:16], key, k0, cur_end))
                        continue
                    if cur_end == want or (want == '；' and cur_end == ';'): continue
                    conf = True
                    rec = (os.path.basename(f), first_plain[:16], p[-40:], kd, repr(cur_end), repr(want))
                    if not conf: rep['疑'].append(rec[:4] + ('建议 %s' % (want or '不加'),)); continue
                    rep['改'].append(rec)
                    parts[i] = fix_item(parts[i], want)
                body2 = ''.join(x for pair in zip(parts, seps + ['']) for x in pair)
                if body2 != body:
                    new = new.replace(c.group(0), c.group(1) + body2 + c.group(3), 1)
            return new
        out = re.sub(r'<tr([^>]*)>.*?</tr>', do_tr, t, flags=re.S)
        if apply and out != t: io.open(f, 'w', encoding='utf-8').write(out)
    print('高把握改', len(rep['改']), '存疑', len(rep['疑']))
    from collections import Counter
    print(Counter((r[4], r[5]) for r in rep['改']).most_common(8))
    if '--report' in sys.argv:
        json.dump(rep, open(sys.argv[sys.argv.index('--report') + 1], 'w'), ensure_ascii=False, indent=0)

if __name__ == '__main__':
    main()
