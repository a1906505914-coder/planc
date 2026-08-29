# -*- coding: utf-8 -*-
"""fa_ledger_mask.py —— 固定资产台账 本地先行脱敏（2026-08-22）。

按台账脱敏三件套，为 ADF 固定资产台账（SAP 资产主数据导出）生成脱敏版：
  输入：7月固定资产台账.xlsx + 12月固定资产台账/12月资产台账/*.XLSX
  输出：7月固定资产台账_脱敏.xlsx + 12月资产台账_脱敏/*.XLSX
  残留校验 0 才放行。

规则（SAP 资产主数据列序）：
  col0  资产号码          → 保留（跨期配对核对 key）
  col1  资产名称          → 主体/公司特征子串泛化
  col2  资产主号文本      → 同 col1
  col3  公司代码          → 借X（contract_mask borrower_dir 映射；未收录保留代码）
  col6  规格型号          → 同 col1
  col11 资产地点 / col12 资产地点描述 → 同 col1（防地点含主体）
  col13 WBS / col14 WBS描述           → 同 col1（防项目名含主体）
  其余（金额/日期/分类/状态）→ 保留（核对必需）

用法：
  python fa_ledger_mask.py            # 生成脱敏版 + 残留校验
  python fa_ledger_mask.py --dry      # 只打印列级替换统计（不写文件）
"""
import glob
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA = r'd:/底稿测试/ADF/数据/2026'
END_FP = os.path.join(DATA, '7月固定资产台账.xlsx')
BEG_DIR = os.path.join(DATA, '12月固定资产台账', '12月资产台账')
OUT_END = os.path.join(DATA, '7月固定资产台账_脱敏.xlsx')
OUT_BEG_DIR = os.path.join(DATA, '12月固定资产台账', '12月资产台账_脱敏')

_COMPANY_PAT = re.compile(
    r'[\u4e00-\u9fff]{2,}(有限公司|股份有限公司|有限责任公司|集团有限公司|集团|股份公司|'
    r'银行|分行|支行|控股|有限|公司|工厂|厂|合作社|事务所|分公司|总厂)')

# 公司代码 → 借X（复用 contract_mask borrower_dir 映射）
_ENTITY_MAP = {}
try:
    from loan_detail import _load_entity_map as _lem
    _ENTITY_MAP = _lem()
except Exception:
    pass


def _mask_comp(v):
    """公司代码 → 借X（如 3100→借01）；未收录保留。统一走 Desensitizer（contract 语境）。"""
    from desensitizer import get_contract
    dz = get_contract()
    dz.register('entity', _ENTITY_MAP)
    return dz.mask(v, 'entity')


def _mask_text(v):
    """长文本内主体/公司特征子串泛化（保留其余业务内容）。统一走 Desensitizer。"""
    from desensitizer import get_contract
    return get_contract().mask(v, 'text')


# 列规则：索引 → 处理函数（None=保留）
COL_RULES = {
    1: _mask_text,    # 资产名称
    2: _mask_text,    # 资产主号文本
    3: _mask_comp,    # 公司代码
    6: _mask_text,    # 规格型号
    9: _mask_text,   # 工厂
    10: _mask_text,  # 工厂描述（含主体名，如 XX创新中心工厂）
    11: _mask_text,   # 资产地点
    12: _mask_text,   # 资产地点描述
    13: _mask_text,   # WBS
    14: _mask_text,   # WBS描述
}

_FA_HEADER = ['资产号码', '资产名称', '资产主号文本', '公司代码', '资产分类', '资产分类描述',
              '规格型号', '序列号/老系统资产编号', '存货号', '工厂', '工厂描述', '资产地点',
              '资产地点描述', 'WBS', 'WBS描述', '数量', '计量单位', '库存位置', '资本化日期',
              '报废日期']


def _collect_real_names():
    """收集原始台账中的候选真实名（含公司特征，供残留校验）。"""
    names = set()
    fps = [END_FP] + sorted(glob.glob(os.path.join(BEG_DIR, '*.XLSX')))
    import openpyxl
    for fp in fps:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(values_only=True, max_row=3000):
                if not row:
                    continue
                for v in row[:20]:
                    if isinstance(v, str) and _COMPANY_PAT.search(v):
                        for m in _COMPANY_PAT.finditer(v):
                            start = max(0, m.start() - 2)
                            end = min(len(v), m.end() + 2)
                            names.add(v[start:end])
            wb.close()
        except Exception:
            continue
    return names


def _process_fp(fp, out_fp, stats):
    import openpyxl
    wb = openpyxl.load_workbook(fp, read_only=True)
    ws = wb.worksheets[0]
    wout = openpyxl.Workbook()
    wso = wout.active
    wso.title = 'Sheet1'
    header = None
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] in (None, ''):
            continue
        if header is None:
            header = row
            wso.append(header)
            continue
        out = list(row)
        for j, fn in COL_RULES.items():
            if j < len(out) and out[j] not in (None, ''):
                old = str(out[j])
                new = fn(old)
                if new != old:
                    stats[j] = stats.get(j, 0) + 1
                out[j] = new
        wso.append(out)
    wb.close()
    os.makedirs(os.path.dirname(out_fp), exist_ok=True)
    wout.save(out_fp)
    return stats


def main():
    dry = '--dry' in sys.argv
    if dry:
        print('[DRY] 预计处理：%s + 12月资产台账/*.XLSX' % os.path.basename(END_FP))
        print('[DRY] 列规则: %s' % {k: (v.__name__ if hasattr(v, '__name__') else 'lambda') for k, v in COL_RULES.items()})
        print('[DRY] 公司代码映射 %d 条（借X）' % len(_ENTITY_MAP))
        return 0

    stats = {}
    _process_fp(END_FP, OUT_END, stats)
    beg_files = sorted(glob.glob(os.path.join(BEG_DIR, '*.XLSX')))
    for fp in beg_files:
        rel = os.path.basename(fp)
        _process_fp(fp, os.path.join(OUT_BEG_DIR, rel), stats)
    print('已生成脱敏版：')
    print('  %s' % OUT_END)
    print('  %s（%d 个文件）' % (OUT_BEG_DIR, len(beg_files)))
    print('替换统计（按列）: %s' % {k: v for k, v in sorted(stats.items())})

    # 残留校验
    real_names = _collect_real_names()
    residue = []
    import openpyxl
    for fp in [OUT_END] + sorted(glob.glob(os.path.join(OUT_BEG_DIR, '*.XLSX'))):
        wb = openpyxl.load_workbook(fp, read_only=True)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(values_only=True):
            for v in row[:20]:
                if isinstance(v, str):
                    for nm in real_names:
                        if nm and nm in v:
                            residue.append((os.path.basename(fp), nm))
                            break
        wb.close()
    if residue:
        print('[FAIL] 残留 %d 处（前 20）:' % len(residue))
        for f, nm in residue[:20]:
            print('  %s : %s' % (f, nm))
        return 2
    print('[OK] 残留 0（候选真实名 %d 个全部覆盖）' % len(real_names))
    return 0


if __name__ == '__main__':
    sys.exit(main())
