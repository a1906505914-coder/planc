# -*- coding: utf-8 -*-
"""ledger_registry.py —— 台账注册表（2026-08-22）。

统一登记【台账类数据】的三件套：manifest(数据源) + mask脚本(脱敏) + 读取器(消费)，
配合 mask_gate（入口体检）/ mask_safe（脱敏版优先读取）。新增台账 = 登记一行，
不再每类各写各的脱敏/读取入口。

用法：
  python ledger_registry.py --list                 # 列出台账及脱敏状态
  python ledger_registry.py --check                # 检查全部台账脱敏版是否齐备
  python ledger_registry.py --check financing      # 只查某台账
  from ledger_registry import masked_path, check
"""
import glob
import os
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT = os.environ.get('AUDIT_DATA_ROOT', r'D:\底稿测试')


def _p(*parts):
    return os.path.join(DATA_ROOT, *parts)


# ============================================================
# 台账注册表（新增台账 = 在此加一行）
#   src      原始文件（相对 DATA_ROOT；支持 * 通配）
#   mask     脱敏脚本（相对 APP_DIR）
#   masked   脱敏版路径（相对 DATA_ROOT；支持 * 通配）
#   reader   消费/读取模块
# ============================================================
LEDGER = {
    'financing': {
        'name': '融资明细（借款/应付票据/表外）',
        'accts': ['ADF'],
        'src': [_p('ADF', '数据', '2026', '融资明细.xls')],
        'mask': 'financing_mask.py',
        'masked': [_p('ADF', '数据', '2026', '融资明细_脱敏.xlsx')],
        'reader': 'financing_ledger',
        'note': '贷款类 ST39/LT70；loan_detail 只读脱敏版',
    },
    'fa': {
        'name': '固定资产台账（SAP 资产主数据）',
        'accts': ['ADF'],
        'src': [_p('ADF', '数据', '2026', '7月固定资产台账.xlsx'),
                _p('ADF', '数据', '2026', '12月固定资产台账', '12月资产台账', '*.XLSX')],
        'mask': 'fa_ledger_mask.py',
        'masked': [_p('ADF', '数据', '2026', '7月固定资产台账_脱敏.xlsx'),
                   _p('ADF', '数据', '2026', '12月固定资产台账', '12月资产台账_脱敏', '*.XLSX')],
        'reader': 'fa_adf / fa_review',
        'note': '约 40 万行/9 主体；公司代码→借X',
    },
    'contract': {
        'name': '借款合同 OCR 识别文本',
        'accts': ['ADF'],
        'src': [_p('ADF', '数据', '2026', '借款合同_ocr', '识别文本', '**', '*.txt')],
        'mask': 'contract_mask_local.py（借款合同脱敏.bat）',
        'masked': [_p('ADF', '数据', '2026', '借款合同_ocr', '识别文本_脱敏', '**', '*.txt')],
        'reader': 'contract_reader / loan_contract_reader',
        'note': '借X/贷X 映射；92 份残留 0',
    },
    'confirm': {
        'name': '银行/往来函证回函 OCR',
        'accts': ['AZ'],
        'src': [_p('AZ', '数据', '2025', '回函', '_ocr_识别结果', '**', '*.txt')],
        'mask': 'confirm_mask_local.py（回函处理.bat）',
        'masked': [_p('AZ', '数据', '2025', '回函', '回函_脱敏版', '**', '*.txt')],
        'reader': 'confirm_parse_all',
        'note': 'A编号/银行单字/C编号；91 份残留 0',
    },
    'lease': {
        'name': '租赁合同台账',
        'accts': ['AJ', 'GFY'],
        'src': [],            # 台账路径随账套配置（lease_manifest）
        'mask': 'mask_batch.py（账套底稿脱敏.bat）',
        'masked': [],
        'reader': 'lease_review / lease_ledger',
        'note': 'contracts 已在 lease_manifest/lease_ledger，随账套配置',
    },
}


def _resolve(p):
    """路径展开：通配/不存在返回 None。"""
    hits = glob.glob(p, recursive=True)
    return hits if hits else None


def list_ledgers():
    print('=== 台账注册表 ===')
    for key, cfg in LEDGER.items():
        print(f'  {key:10s} {cfg["name"]}')
        print(f'    mask={cfg["mask"]}  reader={cfg["reader"]}')
        print(f'    脱敏版: {"、".join(os.path.basename(m) for m in cfg["masked"]) or "（随账套配置）"}')
    print()


def check(key=None):
    """检查台账脱敏版是否齐备。返回缺失台账列表。

    ⚠️ 目录级脱敏（合同/回函：脱敏版子目录名也变 3100国望→借01）无法逐文件路径配对，
    按【数量校验】：脱敏版展开文件数 >= 原始文件数 即视为齐备。"""
    keys = [key] if key else list(LEDGER)
    miss = []
    for k in keys:
        cfg = LEDGER.get(k)
        if not cfg:
            print(f'  ⚠️ 未登记台账: {k}')
            continue
        raw_files = [f for pat in cfg['src'] for f in glob.glob(pat, recursive=True)]
        if not raw_files:
            print(f'  ⚠️ [{k}] 无原始文件（{cfg["name"]}，账套未提供？）')
            continue
        masked_files = [f for pat in cfg['masked'] for f in glob.glob(pat, recursive=True)]
        # 个别脱敏版文件逐个配对检查（文件级：融资/固定资产）
        pair_miss = 0
        for fp in raw_files:
            from mask_safe import prefer_masked
            if prefer_masked(fp) == fp:   # 未找到配对脱敏版
                pair_miss += 1
        if len(masked_files) >= len(raw_files):
            print(f'  ✅ [{k}] {cfg["name"]} 脱敏版齐备（原始{len(raw_files)} / 脱敏{len(masked_files)}）')
        else:
            miss.append(k)
            print(f'  ⚠️ [{k}] {cfg["name"]} 脱敏版不全（原始{len(raw_files)} / 脱敏{len(masked_files)}）'
                  f' → 运行: python {cfg["mask"]}')
    return miss


def masked_path(raw_abs):
    """原始台账路径 → 脱敏版路径（复用 mask_safe）。"""
    from mask_safe import prefer_masked
    return prefer_masked(raw_abs)


def main():
    import argparse
    ap = argparse.ArgumentParser(description='台账注册表')
    ap.add_argument('--list', action='store_true', help='列出台账')
    ap.add_argument('--check', nargs='?', const='', default=None, help='检查脱敏版齐备（可选指定台账）')
    a = ap.parse_args()
    if a.list:
        list_ledgers()
        return 0
    if a.check is not None:
        miss = check(a.check or None)
        if miss:
            print('\n缺失脱敏版的台账：%s → 先跑对应 mask 脚本再分析' % miss)
            return 2
        return 0
    list_ledgers()
    check()
    return 0


if __name__ == '__main__':
    sys.exit(main())
