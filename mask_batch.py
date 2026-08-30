# -*- coding: utf-8 -*-
"""mask_batch.py —— 批量统一脱敏（2026-08-19）。

扫描 D:/底稿测试 下所有【专项输出底稿】（文件名含 复核/核对/清单/分析/函证/穿透 等关键词），
对每个文件用【启发式列分类】自动脱敏 → 生成 <同名>_脱敏.xlsx。

启发式分类（读表头按列名关键词）：
  S1     主体列（我方/对方/公司/往来单位/客户/供应商/存款人/户名…）→ A 代码
  S1sub  长文本列（说明/备注/口径/结论/建议…）→ 文本内主体名子串替换
  S3     业务文本列（摘要/文本/内容/用途/品名…）→ 类型化
  S2     账号列（账号/卡号/税号…）→ 掩码
  其余   科目/代码/金额/日期等 → 不动（金额保留，核对必需）

已有显式配置的（mask_config 中 active 专项，如 ah_intra 三件套）走显式配置；
其余走启发式。版本文件（_dup/_v\\d+/_新/-反馈）跳过。

用法：python mask_batch.py [--scan] [--out-dir <脱敏输出目录>]
"""
import paths as P
import argparse
import os
import re
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import openpyxl  # noqa: E402

import mask_config as MC  # noqa: E402
import mask_dict as MD  # noqa: E402
import mask_engine as ME  # noqa: E402

DATA_ROOT = os.environ.get('AUDIT_DATA_ROOT', r'D://底稿测试')

# 专项输出文件名关键词（进模型分析的核心底稿）
PAT_FILE = re.compile(
    r'(复核|回函|网银核对|开户账户核对|发生额异常|账龄|试算表核对|关联交易和余额|'
    r'合同摘要|合同全覆盖|差异穿透|序时账_|内部往来核对|毛利率|截止性测试|监盘支持表|'
    r'抽凭清单|折算附注|差异对照|差异清单|调整分录|合并工作底稿|科目间勾稽|'
    r'附注审定数|未覆盖科目|关联方清单|账户汇总表|租赁合同汇总)'
)
# 排除：已脱敏/历史版本/模板/常规科目审计底稿
EXC_FILE = re.compile(
    r'(_脱敏|_dup|模板|勿删|_外壳|_v\d+|_生成_v|_新|反馈|审计底稿_\d{4}_)'
)

# 启发式列分类关键词
S1_KEYS = ('我方主体', '对方主体', '核算主体', '往来单位', '客户', '供应商',
           '存款人', '承租方', '出租方', '甲方', '乙方', '单位名称', '账户名称',
           '户名', '主体', '公司', '账套', '我方名称', '对方名称', '对方单位',
           '对方户名', '付款方', '收款方', '客商', '被审计单位', '单位', '银行',
           '我方', '对方')
S2_KEYS = ('账号', '卡号', '票据号', '税号', '信用代码')
S1SUB_KEYS = ('说明', '备注', '口径', '结论', '建议', '反馈', '文件')
S3_KEYS = ('摘要', '文本', '内容', '用途', '品名', '标的', '条款', '业务描述', '地址')

# 通用 S3 类型化关键词（启发式列用）
DEFAULT_KW = {
    '材料/采购': ['材料', '采购', '购入', '集采购'],
    '销售/收入': ['销售', '收入', '销'],
    '费用': ['费用', '报销', '手续费', '服务费'],
    '收付款': ['付款', '支付', '收款', '收到', '转账'],
    '资金': ['资金池', '归集', '上收', '下拨'],
    '利息': ['利息', '结息'],
    '结转调整': ['结转', '暂估', '冲销', '调整', '重分类'],
    '票据': ['承兑', '托收', '背书'],
    '租赁': ['租金', '租赁', '押金', '保证金'],
    '固定资产': ['折旧', '资产', '设备', '厂房', '车辆'],
}


def _scan():
    """扫描候选文件（含路径）。"""
    out = []
    for root, dirs, files in os.walk(DATA_ROOT):
        if not any(k in root for k in ('底稿', 'prepared', '中间产物')):
            continue
        for f in files:
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            if EXC_FILE.search(f):
                continue
            if PAT_FILE.search(f):
                out.append(os.path.join(root, f))
    return sorted(out)


def auto_sheet_cfg(hdr):
    """按表头启发式生成 {列名: 规则}。"""
    cfg = {}
    for h in hdr:
        hs = str(h or '').strip()
        if not hs:
            continue
        # 仅排除【科目代码/编码】列；科目名称列不排除——主体小计行可能把主体名放在科目名称列
        if '科目代码' in hs or '科目编码' in hs or hs.endswith('代码'):
            continue
        if '对方科目' in hs:          # ⚡ 对方科目列实为摘要/往来描述（银行底稿），类型化
            cfg[hs] = 'S3'
            continue
        if '银行账户' in hs:          # ⚡ 银行账户列→账号掩码（S2 特判，防被 S1 的'银行'抢）
            cfg[hs] = 'S2'
            continue
        if '摘要' in hs and '对方' in hs:   # ⚡ 网银对方/摘要 等混合列 → 摘要优先（S3）
            cfg[hs] = 'S3'
            continue
        if any(k in hs for k in S1_KEYS):
            cfg[hs] = 'S1'
        elif any(k in hs for k in S2_KEYS):
            cfg[hs] = 'S2'
        elif any(k in hs for k in S1SUB_KEYS):
            cfg[hs] = 'S1sub'
        elif any(k in hs for k in S3_KEYS):
            cfg[hs] = 'S3'
    return cfg


# 通用表头词（辅助定位表头行——很多底稿表头在 3-5 行，前几行是标题）
HDR_WORDS = ('科目', '金额', '期末', '发生', '日期', '数量', '余额', '差异', '合计',
             '笔数', '名称', '期初', '银行', '核对', '结论', '明细', '应收', '应付',
             '收入', '成本', '主体', '公司')


def find_header(ws, max_scan=12):
    """多行扫描定位表头行。返回 (row_idx, header_tuple)；找不到返回 (None, None)。

    ⚡ 稳健策略：以【非空列数】为主（表头行几乎每列有值），关键词为辅。
       ⚠️ 只对短值（≤15字）计关键词——数据行含公司名（"XX有限公司…"）会伪造高分，
          长值不计，避免数据行冒充表头（2026-08-19 实测 3500 网银 sheet 误定位）。"""
    best_idx, best_row, best_score = None, None, -1
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i >= max_scan:
            break
        if not r:
            continue
        cells = [str(c or '').strip() for c in r]
        nonempty = sum(1 for c in cells if c and len(c) <= 30)
        kw = 0
        for c in cells:
            if not c or len(c) > 15:      # 长值=数据（客户名/摘要），不计关键词
                continue
            if any(k in c for k in S1_KEYS):
                kw += 3
            if any(k in c for k in S2_KEYS):
                kw += 2
            if any(k in c for k in S3_KEYS):
                kw += 1
            if any(k in c for k in HDR_WORDS):
                kw += 1
        score = nonempty * 2 + kw
        if score > best_score:
            best_score, best_idx, best_row = score, i, r
    if best_score <= 0:
        return None, None
    return best_idx, best_row


def inspect(fp):
    """逐 sheet 定位表头 + 生成列规则。
    返回 {sheet: {'header': row_idx, 'cols': {列名: 规则}} 或 None(无表头，整 sheet 原样复制)}。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    out = {}
    try:
        for sn in wb.sheetnames:
            ws = wb[sn]
            idx, hdr = find_header(ws)
            if idx is None:
                out[sn] = None
            else:
                out[sn] = {'header': idx, 'cols': auto_sheet_cfg(hdr)}
    finally:
        wb.close()
    return out


def _build_alias(data):
    """构建简称别名映射（复用 MC.ALIAS）。返回 (alias_map, alias_sorted)。"""
    alias_map = {}
    for full, aliases in MC.ALIAS.items():
        code = data.get(full)
        if code:
            for a in aliases:
                alias_map[a] = code
    return alias_map, sorted(alias_map, key=len, reverse=True)


def _run(fp, insp, out):
    """read_only → write_only 读写循环：保留标题行/表头行，数据行按列规则脱敏。"""
    data = MD.names()
    sorted_names = sorted(data, key=len, reverse=True)
    alias_map, alias_sorted = _build_alias(data)
    cfg = {'text_kw': DEFAULT_KW}
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    wout = openpyxl.Workbook(write_only=True)
    try:
        for sn in wb.sheetnames:
            ws = wb[sn]
            wso = wout.create_sheet(sn)
            info = insp.get(sn)
            hdr_idx = info['header'] if info else None
            cols = info['cols'] if info else {}
            hmap = None
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if not r:
                    continue
                if hdr_idx is None or i < hdr_idx:
                    wso.append(list(r))
                    continue
                if i == hdr_idx:
                    hmap = {str(h): k for k, h in enumerate(r) if h}
                    wso.append(list(r))
                    continue
                out_row = list(r)
                for col, rule in cols.items():
                    idx = hmap.get(col)
                    if idx is None or idx >= len(out_row):
                        continue
                    if out_row[idx] in (None, ''):
                        continue
                    out_row[idx] = ME._apply(rule, out_row[idx], cfg, data,
                                             sorted_names, alias_map, alias_sorted)
                wso.append(out_row)
    finally:
        wb.close()
    wout.save(out)


def process(fp, out_dir=None):
    """对单个文件脱敏。返回 (basename, 是否已处理, 说明)。"""
    base = os.path.basename(fp)
    # 已有显式配置（ah_intra 等）→ 直接走引擎
    for name, cfg in MC.CONFIG.items():
        if base in (cfg.get('files') or {}):
            out = os.path.join(out_dir, base.replace('.xlsx', '_脱敏.xlsx')) if out_dir else None
            if not out:
                out = os.path.splitext(fp)[0] + '_脱敏.xlsx'
            _, stats = ME.transform(fp, out)
            return base, True, 'config:' + name
    # 启发式：定位表头 + 列规则
    insp = inspect(fp)
    has_cols = any(v and v['cols'] for v in insp.values())
    if not has_cols:
        return base, False, '无敏感列可识别，跳过'
    out = os.path.join(out_dir, base.replace('.xlsx', '_脱敏.xlsx')) if out_dir else None
    if not out:
        out = os.path.splitext(fp)[0] + '_脱敏.xlsx'
    _run(fp, insp, out)
    n_sheets = sum(1 for v in insp.values() if v and v['cols'])
    return base, True, 'auto:%d sheets' % n_sheets


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--scan', action='store_true', help='只扫描清单，不脱敏')
    ap.add_argument('--out-dir', help='脱敏输出目录（默认与源同目录）')
    a = ap.parse_args(argv)

    files = _scan()
    print('扫描到专项输出文件:', len(files))
    if a.scan:
        for fp in files:
            print(' ', fp)
        return 0

    ok, skip, fail = [], [], []
    for i, fp in enumerate(files, 1):
        try:
            base, done, note = process(fp, a.out_dir)
            status = '✓' if done else '·'
            print('[%d/%d] %s %s | %s' % (i, len(files), status, base, note))
            (ok if done else skip).append(base)
        except Exception as e:
            fail.append((base, str(e)[:60]))
            print('[%d/%d] ✗ %s | %s' % (i, len(files), os.path.basename(fp), str(e)[:60]))
    print()
    print('处理完成：脱敏 %d | 跳过 %d | 失败 %d' % (len(ok), len(skip), len(fail)))
    if fail:
        print('失败清单:')
        for b, e in fail:
            print('  ', b, e)
    return 0


if __name__ == '__main__':
    sys.exit(main())
