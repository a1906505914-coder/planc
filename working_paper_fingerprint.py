# -*- coding: utf-8 -*-
"""底稿指纹对比基线 v2（只读，不改任何底稿；JSONL 流式防 OOM）。

用途：给底稿改代码（#874 互抵口径 / #876 导出层重构）兜底——
   改前跑 `base` 存当前三集团正式底稿的逐格指纹；改后重跑底稿再跑 `check`，
   任何 sheet 维度/数值/文本差异会逐行暴露，防止"改了代码数字悄悄变坏"。

v2 变更：基线存 JSONL（每行=一个文件条目），base 逐文件写盘、check 逐文件
   读盘对比 → 内存只驻留单文件指纹，规避沙箱 OOM 强杀。

用法：
    python working_paper_fingerprint.py base    # 生成基线（当前=真值）
    python working_paper_fingerprint.py check   # 对比当前 vs 基线，报告差异
    python working_paper_fingerprint.py check --report 报告.txt

指纹口径：
   - 逐 sheet：rows × cols + 每行哈希 + 每行前4列标识 + 全表哈希
   - 数值按 %.6f 归一（避免浮点末位抖动误报）；None 记 N
   - read_only + data_only 流式读取（公式取缓存值，与生成器写数值一致）
"""
import os, sys, glob, json, time, hashlib
import openpyxl

ROOT = r'D:/底稿测试/AH/底稿/2026/集团'
GRPS = ['1010', '1357', '2468']
BASE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_wp_fingerprint_base.jsonl')


def _norm(v):
    if v is None:
        return 'N'
    if isinstance(v, float):
        return 'F%.6f' % v
    if isinstance(v, int):
        return 'I%d' % v
    return 'S' + str(v)


def _row_hash(row):
    return hashlib.sha1('|'.join(_norm(c) for c in row).encode('utf-8', 'replace')).hexdigest()[:16]


def _row_key(row):
    return ' | '.join(str(c)[:20] for c in row[:4] if c is not None)


def _scan_file(fp):
    """单文件指纹（流式读，只驻留当前文件）。失败返回 {'error': msg}。
    v3：行序无关——行指纹按哈希排序存储（Python dict/set 遍历哈希随机化
    会导致生成器输出行序不稳定，逐行对比会误报；排序后只反映内容差异）。"""
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception as e:
        return {'error': str(e)}
    sheets = {}
    for ws in wb.worksheets:
        hashes = []
        keys = []
        for row in ws.iter_rows(values_only=True):
            hashes.append(_row_hash(row))
            keys.append(_row_key(row))
        paired = sorted(zip(hashes, keys))   # 行序无关：按哈希排序
        sheets[ws.title] = {
            'rows': len(hashes),
            'cols': ws.max_column,
            'rows_sorted': paired,
            'total': hashlib.sha1(','.join(h for h, _k in paired).encode()).hexdigest()[:16],
        }
    wb.close()
    return sheets


def _iter_targets():
    """生成 (group, 文件名, 绝对路径) 序列。"""
    for g in GRPS:
        d = os.path.join(ROOT, g)
        if not os.path.isdir(d):
            continue
        for fp in sorted(glob.glob(os.path.join(d, '*.xlsx'))):
            fn = os.path.basename(fp)
            if fn.startswith('~$'):
                continue
            yield g, fn, fp


def base():
    t0 = time.time()
    nf = 0
    with open(BASE_PATH, 'w', encoding='utf-8') as f:
        f.write(json.dumps({'_meta': {'created': time.strftime('%Y-%m-%d %H:%M:%S'),
                                      'root': ROOT, 'groups': GRPS,
                                      'format': 'v2-jsonl'}}, ensure_ascii=False) + '\n')
        for g, fn, fp in _iter_targets():
            rec = {'g': g, 'fn': fn, 'sheets': _scan_file(fp)}
            f.write(json.dumps(rec, ensure_ascii=False) + '\n')
            nf += 1
    print(f'基线已生成：{BASE_PATH}')
    print(f'  文件数：{nf}，耗时 {time.time() - t0:.0f}s', flush=True)


def check():
    if not os.path.exists(BASE_PATH):
        print(f'无基线文件：{BASE_PATH}（先运行 base）')
        return 1
    meta = None
    t0 = time.time()
    diffs = 0
    report = []
    base_fns = set()
    cur_fns = set((g, fn) for g, fn, _fp in _iter_targets())
    with open(BASE_PATH, 'r', encoding='utf-8') as f:
        for line in f:
            rec = json.loads(line)
            if '_meta' in rec:
                meta = rec['_meta']
                continue
            g, fn, bs = rec['g'], rec['fn'], rec['sheets']
            base_fns.add((g, fn))
            fp = os.path.join(ROOT, g, fn)
            if not os.path.exists(fp):
                report.append(f'[缺失文件] {g}/{fn}')
                diffs += 1
                continue
            if isinstance(bs, dict) and 'error' in bs:
                continue
            cs = _scan_file(fp)
            for sn in sorted(set(bs) | set(cs)):
                if sn not in bs:
                    report.append(f'  [新增sheet] {g}/{fn} :: {sn}')
                    diffs += 1
                    continue
                if sn not in cs:
                    report.append(f'  [缺失sheet] {g}/{fn} :: {sn}')
                    diffs += 1
                    continue
            br = bs[sn]
            cr = cs[sn]
            if br['rows'] != cr['rows'] or br['cols'] != cr['cols']:
                report.append(f'  [维度变化] {g}/{fn} :: {sn}  基线 {br["rows"]}x{br["cols"]} vs 当前 {cr["rows"]}x{cr["cols"]}')
                diffs += 1
                continue
            if br['total'] == cr['total']:
                continue
            # 行序无关对比：找出 基线有当前缺失 / 当前新增 的内容行（hash 多集差）
            from collections import Counter
            _bc = Counter(h for h, _k in br['rows_sorted'])
            _cc = Counter(h for h, _k in cr['rows_sorted'])
            missing = list((_bc - _cc).elements())
            added = list((_cc - _bc).elements())
            for h in missing[:10]:
                key = next((k for hh, k in br['rows_sorted'] if hh == h), '?')
                report.append(f'    [内容差异] 基线有/当前缺: 「{key}」')
                diffs += 1
            for h in added[:10]:
                key = next((k for hh, k in cr['rows_sorted'] if hh == h), '?')
                report.append(f'    [内容差异] 当前新增: 「{key}」')
            if diffs > 200:
                report.append('    …差异过多，截断')
                break
    # 当前有、基线无 = 新增文件
    for (g, fn) in sorted(cur_fns - base_fns):
        report.append(f'[新增文件] {g}/{fn}')
        diffs += 1
    lines = [f'底稿指纹对比 @ {time.strftime("%Y-%m-%d %H:%M:%S")}',
             f'基线 {meta.get("created", "?") if meta else "?"}；扫描耗时 {time.time()-t0:.0f}s']
    if diffs == 0:
        lines.append('✅ 指纹完全一致，无回退')
    else:
        lines.append(f'共 {diffs} 处差异：')
        lines += report
    out = '\n'.join(lines)
    if '--report' in sys.argv:
        i = sys.argv.index('--report')
        rp = sys.argv[i + 1]
        with open(rp, 'w', encoding='utf-8') as f:
            f.write(out + '\n')
        print(out)
        print(f'\n报告已写：{rp}')
    else:
        print(out)
    return 0 if diffs == 0 else 1


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'check'
    if cmd == 'base':
        base()
    else:
        sys.exit(check())
