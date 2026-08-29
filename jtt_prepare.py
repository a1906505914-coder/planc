# -*- coding: utf-8 -*-
"""JTt 方案 A 适配器（2026-08-05）：把 ga/jj 集团的『用友 U8 多账簿导出』
拆分成 FY 标准结构（每主体一年一文件：科目余额表 / 综合查询明细表 / 辅助核算余额表），
13 个 builder 零改动直接跑。

输入：<JTt 目录>（含 ga2025分主体科目余额表.xlsx、JJ2025分主体科目余额表.xlsx、序时账*、*辅助余额表.xlsx）
输出：<输出目录>/ga/ 与 <输出目录>/jj/ 下每主体三个标准文件
  NN{主体名}2025年科目余额表.xlsx        （表头 r2：科目编码|科目名称|期初方向|期初金额|本期借方|本期贷方|期末方向|期末金额|等级）
  NN{主体名}2025年综合查询明细表.xlsx    （表头 r2：科目名称|日期|凭证字|凭证号|对方科目|摘要|借方金额|贷方金额）
  NN{主体名}2025年辅助核算余额表.xlsx    （表头 r2：科目名称|往来单位名称|期初方向|期初金额|本期借方|本期贷方|期末方向|期末金额|往来类型）

清洗规则：
- 主体名：去『-基准账簿/-基准账/-基准』后缀、去换行符，排序后补 NN 编号
- 科目名：去『编码\』前缀（10020101\银行存款\结算中心 → 银行存款\结算中心），保留中文层级
- 辅助表：剔『核算账簿累计』汇总行；客商名『【客商：XXX】』→ XXX
- GL：ga 12 月序时账（双层表头，取本币列）与 jj 2 文件（单层表头）各自按主体合并为全年
"""
import paths as P
import openpyxl, os, re, sys, io, glob

# 2026-08-05：stdout 包装移入 __main__（模块级包装在 import 场景会关闭底层 buffer，
# 引发 'I/O operation on closed file'；与 currency_annotate 同一教训）。

YEAR = '2025'


def clean_ent(raw):
    """主体名清洗：去换行、去『-基准账簿/-基准账/-基准』后缀。"""
    s = str(raw or '').strip().replace('\n', '').replace('\r', '').replace('\u3000', '')
    s = re.sub(r'[-－]\s*(基准账簿|基准账|基准)$', '', s)
    return s.strip()


def clean_subj(raw):
    """科目名清洗：去『编码\』前缀，保留中文层级（10020101\银行存款\结算中心 → 银行存款\结算中心）。"""
    s = str(raw or '').strip()
    s = re.sub(r'^\d+\\', '', s)
    return s.strip()


def strip_kehu(raw):
    """客商名：『【客商：XXX】』 → XXX。"""
    s = str(raw or '').strip()
    m = re.search(r'【客商：(.+?)】', s)
    return m.group(1).strip() if m else s


def _num(v):
    """金额清洗：None/空/'None' 字符串 → 0.0（U8 导出空单元格可能是字符串 'None'）。"""
    if v is None:
        return 0.0
    if isinstance(v, str):
        s = v.strip()
        if s == '' or s.lower() == 'none':
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def load_rows(path, min_row):
    """读全部数据行。read_only 优先；部分文件（JJ TB / ga 辅助表）read_only dims 异常
    只读到 1 行/0 行 → 自动回退非只读重读。返回 (rows 列表, 实际起始行)。"""
    for ro in (True, False):
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=ro)
            ws = wb[wb.sheetnames[0]]
            rows = [r for r in ws.iter_rows(min_row=min_row, values_only=True)]
            wb.close()
        except Exception:
            continue
        # 数据行判定：非全空行数
        n = sum(1 for r in rows if r and any(x is not None and str(x).strip() for x in r))
        if n >= 10:
            return rows, ro
    # 最后尝试非只读
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws = wb[wb.sheetnames[0]]
    rows = [r for r in ws.iter_rows(min_row=min_row, values_only=True)]
    wb.close()
    return rows, False


# ===================== TB 拆分 =====================
def split_tb(src, out_dir, prefix):
    """分主体科目余额表 → 每主体一个标准科目余额表。返回 {主体: 输出文件路径}。"""
    rows, _ = load_rows(src, 1)
    groups = {}
    hdr_seen = False
    for r in rows:
        if not r or r[0] is None or not str(r[0]).strip():
            continue
        # 2026-08-05 修复：剔除『核算账簿累计』汇总行（源 TB 每账簿末行，金额=该账簿全科目
        # 合计，混入会导致每主体 TB 总额翻倍——曾致 GL=TB/2 假象）。aux 拆分已有同规则。
        if '核算账簿累计' in str(r[0]):
            continue
        if not hdr_seen:
            if '科目编码' in str(r[0]):
                hdr_seen = True
            continue
        ent = clean_ent(r[2] if len(r) > 2 else '')
        if not ent:
            continue
        groups.setdefault(ent, []).append(r)
    print(f'  TB 主体数：{len(groups)}')
    paths = {}
    for i, ent in enumerate(sorted(groups), 1):
        nn = f'{i:02d}'
        outp = os.path.join(out_dir, f'{nn}{ent}{YEAR}年科目余额表.xlsx')
        wout = openpyxl.Workbook()
        wso = wout.active
        wso.title = '科目余额表'
        wso.append([f'{ent} {YEAR}年科目余额表'])
        wso.append(['科目编码', '科目名称', '期初方向', '期初金额', '本期借方',
                    '本期贷方', '期末方向', '期末金额', '等级'])
        for r in groups[ent]:
            # 源列：0编码 1名称 2账簿 3期初方向 4期初余额 5本期借方 6本期贷方 7借累 8贷累 9期末方向 10期末余额
            wso.append([
                str(r[0]).strip() if r[0] is not None else '',
                clean_subj(r[1]) if len(r) > 1 else '',
                str(r[3]).strip() if len(r) > 3 and r[3] is not None else '',
                _num(r[4]) if len(r) > 4 else 0.0,
                _num(r[5]) if len(r) > 5 else 0.0,
                _num(r[6]) if len(r) > 6 else 0.0,
                str(r[9]).strip() if len(r) > 9 and r[9] is not None else '',
                _num(r[10]) if len(r) > 10 else 0.0,
                '',
            ])
        wout.save(outp)
        paths[ent] = outp
    return paths


# ===================== GL 拆分 =====================
def _detect_jj_header(rows):
    """探测 JJ 序时账表头类型并返回 (start_row, colmap)：
    - 新结构（2026-08-05 按月导出，双层表头 r1/r2、数据 r3 起）：
        r1: 年|年|核算账簿名称|业务单元|凭证号|分录号|摘要|科目编码|科目名称|辅助项|币种|借方|借方|贷方|贷方|核销信息
        r2: 月|日|…|原币|本币|原币|本币|…
      colmap: 月0 日1 主体2 业务3 凭证号4 分录5 摘要6 编码7 科目8 辅助9 币种10 借本币12 贷本币14
    - 旧结构（2026-08-04 序时账1/2，单层表头 r1、数据 r2 起）：
        月0 日1 主体2 业务3 凭证号4 摘要8 编码9 科目10 借方13 贷方14
      colmap: 月0 日1 主体2 业务3 凭证号4 摘要8 编码9 科目10 借本币13 贷本币14
    返回 (start_row_index, colmap)；无法识别返回 (None, None)。"""
    for i in range(min(4, len(rows))):
        r = rows[i]
        if not r:
            continue
        txt = [str(c) if c is not None else '' for c in r]
        if '核算账簿名称' in txt and '凭证号' in txt and '摘要' in txt:
            # 判断是否双层：下一行含 月/日 且含 原币/本币
            if i + 1 < len(rows) and rows[i + 1] is not None:
                nxt = [str(c) if c is not None else '' for c in rows[i + 1]]
                if '月' in nxt and '日' in nxt and ('原币' in nxt or '本币' in nxt):
                    return i + 2, {'month': 0, 'day': 1, 'ent': 2, 'biz': 3, 'vno': 4,
                                   'summ': 6, 'code': 7, 'name': 8, 'jf': 12, 'df': 14}
            return i + 1, {'month': 0, 'day': 1, 'ent': 2, 'biz': 3, 'vno': 4,
                           'summ': 8, 'code': 9, 'name': 10, 'jf': 13, 'df': 14}
    return None, None


def split_gl(files, out_dir, ents, mode, paths_out):
    """序时账 → 每主体全年综合查询明细表。
    mode='ga'：ga 12 月序时账，表头 r2/r3、数据 r4 起；
               c0=月 c1=日 c2=主体 c3=凭证号 c5=摘要 c6=编码 c7=科目
               c10=借原币 c11=借本币 c12=贷原币 c13=贷本币
    mode='jj'：JJ 序时账（2026-08-05 更新为按月 12 个文件、双层表头；兼容旧单层表头），
               自动探测表头类型（_detect_jj_header），本币列取数（借本币/贷本币）。
    """
    per_ent = {e: [] for e in ents}
    other_ents = set()
    for fp in sorted(files):
        if mode == 'jj':
            rows, ro = load_rows(fp, 1)
            _start, cm = _detect_jj_header(rows)
            if _start is None:
                print(f'  ⚠️ {os.path.basename(fp)} 表头未识别，跳过')
                continue
            data_rows = rows[_start:]
        else:
            rows, ro = load_rows(fp, 4)
            cm = {'month': 0, 'day': 1, 'ent': 2, 'biz': 3, 'vno': 4,
                  'summ': 5, 'code': 6, 'name': 7, 'jf': 11, 'df': 13}
            data_rows = rows
        for r in data_rows:
            if not r or len(r) <= max(cm.get('jf', 0), cm.get('df', 0), 10):
                continue
            ent = clean_ent(r[cm['ent']] if len(r) > cm['ent'] else '')
            if not ent:
                continue
            if ent not in per_ent:
                other_ents.add(ent)
                continue
            try:
                m = int(float(r[cm['month']])) if r[cm['month']] is not None else 0
                d = int(float(r[cm['day']])) if r[cm['day']] is not None else 0
            except (TypeError, ValueError):
                continue
            vno = str(r[cm['vno']] if len(r) > cm['vno'] else '' or '')
            vt = vno.split('-')[0].strip() if '-' in vno else vno.strip()
            vn = vno.split('-')[1].strip() if '-' in vno else ''
            summ = str(r[cm['summ']] if len(r) > cm['summ'] else '' or '')
            code = str(r[cm['code']] if len(r) > cm['code'] else '' or '')
            nm = r[cm['name']] if len(r) > cm['name'] else None
            jf = r[cm['jf']] if len(r) > cm['jf'] else None
            df = r[cm['df']] if len(r) > cm['df'] else None
            jf = _num(jf); df = _num(df)
            if jf == 0 and df == 0:
                continue
            per_ent[ent].append((m, d, clean_subj(nm), f'{YEAR}-{m:02d}-{d:02d}', vt, vn,
                                 '', summ, jf, df))
    if other_ents:
        try:
            from mask_dict import mask_names as _mn
            _ents = _mn(sorted(other_ents), 5)
        except Exception:
            _ents = sorted(other_ents)[:5]
        print(f'  ⚠️ GL 含 TB 无主体 {len(other_ents)} 个（已忽略）：' + '、'.join(_ents))
    for ent in sorted(per_ent):
        rows = per_ent[ent]
        outp = paths_out.get(ent)
        if not outp:
            continue
        wout = openpyxl.Workbook()
        wso = wout.active
        wso.title = '综合查询明细表'
        wso.append([f'{ent} {YEAR}年综合查询明细表'])
        wso.append(['科目名称', '日期', '凭证字', '凭证号', '对方科目', '摘要', '借方金额', '贷方金额'])
        for m, d, nm, dt, vt, vn, opp, summ, jf, df in rows:
            wso.append([nm, dt, vt, vn, opp, summ, jf, df])
        wout.save(outp)
        print(f'  GL {ent}: {len(rows)} 行 → {os.path.basename(outp)}')


# ===================== aux 拆分 =====================
def split_aux(files, out_dir, ents, paths_out):
    """辅助余额表（ga 7 个 / jj 6 个）→ 每主体一个辅助核算余额表（多科目合并）。
    源表头 r2：摘要|核算账簿名称|科目名称|客商名称|方向|期初余额|本期借方|本期贷方|借累|贷累|方向|期末余额
    数据 r3 起；『核算账簿累计』汇总行剔除。"""
    per_ent = {e: [] for e in ents}
    for fp in sorted(files):
        rows, _ = load_rows(fp, 3)
        for r in rows:
            if not r or len(r) < 11:
                continue
            summ_mark = str(r[0] or '').strip()
            if summ_mark == '核算账簿累计':
                continue
            ent = clean_ent(r[1] if len(r) > 1 else '')
            if not ent or ent not in per_ent:
                continue
            km = clean_subj(r[2]) if len(r) > 2 else ''
            kehu = strip_kehu(r[3]) if len(r) > 3 else ''
            if not km or not kehu:
                continue
            per_ent[ent].append([
                km,
                kehu,
                str(r[4]).strip() if len(r) > 4 and r[4] is not None else '',   # 期初方向
                _num(r[5]) if len(r) > 5 else 0.0,                              # 期初金额
                _num(r[6]) if len(r) > 6 else 0.0,                              # 本期借方
                _num(r[7]) if len(r) > 7 else 0.0,                              # 本期贷方
                str(r[10]).strip() if len(r) > 10 and r[10] is not None else '',  # 期末方向
                _num(r[11]) if len(r) > 11 else 0.0,                            # 期末金额
                '',                                                              # 往来类型
            ])
    for ent in sorted(per_ent):
        rows = per_ent[ent]
        outp = paths_out.get(ent)
        if not outp:
            continue
        wout = openpyxl.Workbook()
        wso = wout.active
        wso.title = '辅助核算余额表'
        wso.append([f'{ent} {YEAR}年辅助核算余额表'])
        # 2026-08-05 表头对齐 FY 样式（纯『方向/金额』列，_resolve_aux_columns 按顺序识别：
        # 第一个方向/金额=期初、第二个=期末）：科目名称|往来单位名称|方向|金额|本期借方|本期贷方|方向|金额|往来类型
        wso.append(['科目名称', '往来单位名称', '方向', '金额', '本期借方',
                    '本期贷方', '方向', '金额', '往来类型'])
        for row in rows:
            wso.append(row)
        wout.save(outp)
        print(f'  aux {ent}: {len(rows)} 行 → {os.path.basename(outp)}')


def process(src_dir, out_root):
    os.makedirs(out_root, exist_ok=True)
    for tag, tb_file, gl_files, aux_files in [
        ('ga',
         os.path.join(src_dir, 'ga2025分主体科目余额表.xlsx'),
         sorted(glob.glob(os.path.join(src_dir, 'ga2025序时账*月.xlsx'))),
         sorted(glob.glob(os.path.join(src_dir, 'ga2025*辅助余额表.xlsx')))),
        ('jj',
         os.path.join(src_dir, 'JJ2025分主体科目余额表.xlsx'),
         # 2026-08-05：JJ 序时账更新为按月 12 个文件（JJ2025序时账1月~12月.xlsx，双层表头），
         # 全部纳入；旧版 序时账1/2（单层表头，序时账2 只含借方不完整）不再使用。
         sorted(glob.glob(os.path.join(src_dir, 'JJ2025序时账*月.xlsx'))),
         sorted(glob.glob(os.path.join(src_dir, 'JJ2025*辅助余额表.xlsx')))),
    ]:
        print(f'\n===== {tag} =====')
        out_dir = os.path.join(out_root, tag)
        os.makedirs(out_dir, exist_ok=True)
        if not os.path.isfile(tb_file):
            print(f'  ⚠️ 缺 TB：{tb_file}')
            continue
        tb_paths = split_tb(tb_file, out_dir, tag)
        ents = sorted(tb_paths)
        print(f'  主体 {len(ents)} 个')
        # GL 输出路径复用 TB 同名主体
        gl_out = {e: os.path.join(out_dir, os.path.basename(p).replace('科目余额表', '综合查询明细表'))
                  for e, p in tb_paths.items()}
        if gl_files:
            split_gl(gl_files, out_dir, ents, 'ga' if tag == 'ga' else 'jj', gl_out)
        else:
            print('  ⚠️ 无序时账文件，GL 跳过')
        if aux_files:
            aux_out = {e: os.path.join(out_dir, os.path.basename(p).replace('科目余额表', '辅助核算余额表'))
                       for e, p in tb_paths.items()}
            split_aux(aux_files, out_dir, ents, aux_out)
        else:
            print('  ⚠️ 无辅助余额表，aux 跳过')
    print(f'\n✅ 拆分完成：{out_root}')


def main():
    argv = sys.argv[1:]
    src = argv[0] if argv else P.JTT
    out = argv[1] if len(argv) > 1 else os.path.join(src, 'prepared')
    if not os.path.isdir(src):
        print(f'⚠️ 源目录不存在：{src}')
        return 1
    process(src, out)
    print(f'\n后续：python current_account_detail.py "{out}" （或 regen_all.py "{out}"）')


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    main()
