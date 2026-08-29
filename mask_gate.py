# -*- coding: utf-8 -*-
"""mask_gate.py —— 脱敏入口闸门 / 数据体检（2026-08-22）。

铁律：**今后任何提供给我的数据，第一步都是脱敏（特别是账套主体名称）。**
本工具把"第一步脱敏"变成可执行的检查：扫描数据目录，识别含真实主体名
（账套主体/公司/银行/集团）的【敏感文件】，并报告其脱敏状态。

用法：
  python mask_gate.py                  # 扫描全部账套（D:/底稿测试）
  python mask_gate.py --dir ADF/数据/2026    # 只扫指定目录（相对 DATA_ROOT 或绝对）
  python mask_gate.py --list                  # 列出敏感未脱敏文件（默认同 --dir 全部）
  python mask_gate.py --gate                  # 门禁模式：存在敏感未脱敏 → 退出码 2

输出约定：只打印【文件名 + 状态】，不打印文件内容（防把敏感内容带进对话）。
脱敏状态判定：
  - 已有脱敏版（同文件 *_脱敏.xlsx / 文件在 脱敏/或_脱敏/ 目录） → 已脱敏
  - 内容含真实名/公司特征且无脱敏版                          → 敏感·未脱敏 ⚠️
  - 内容不含敏感特征                                        → 无敏感

真实名来源（本地对照表，仅本机）：mask_dict.json + contract_mask.json + confirm_mask.json。
"""
import argparse
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT = os.environ.get('AUDIT_DATA_ROOT', r'D:\底稿测试')

# 公司/银行特征（无字典命中的兜底判定）
_COMPANY_PAT = re.compile(
    r'[\u4e00-\u9fff]{2,}(有限公司|股份有限公司|有限责任公司|集团有限公司|集团|'
    r'股份公司|银行|分行|支行|金租|租赁|财务公司|信托|证券|保险|控股|有限|'
    r'公司|工厂|厂|合作社|事务所|分公司|总厂)')

# 扫描排除的目录/文件
_EXCLUDE_DIR = ('__pycache__', '.git', '.gl_cache', '备份', '_old', 'archive',
                '脱敏', '_脱敏', '识别文本_脱敏', '回函_脱敏版')
_EXCLUDE_EXT = ('.py', '.bat', '.pkl', '.pklz', '.pyc', '.log', '.ini', '.exe',
                '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.gif')
_TEXT_EXT = ('.json', '.txt', '.md', '.csv', '.htm', '.html')

# 原始账套数据特征（文件名命中 → 本机程序专用，需真实名做勾稽，不直接进对话）
_RAW_DATA_KW = ('科目余额表', '序时账', '综合查询', '供应商', '客户', '余额表',
                '往来明细', '明细账', '对账单', '卡片', '辅助核算', '辅助余额', '项目余额',
                '发生额', '试算表', '利润表', '资产负债表', '现金流量表', '网银')
# 本机中间产物目录（OCR 识别文本等，脱敏版在 *_脱敏 目录，程序已只读脱敏版）
_MID_DIR_KW = ('_ocr_识别结果', '识别文本')


def _load_real_names():
    """收集本地对照表中的真实名（mask_dict + contract_mask + confirm_mask）。"""
    names = set()
    for fn in ('mask_dict.json', 'contract_mask.json', 'confirm_mask.json'):
        fp = os.path.join(APP_DIR, fn)
        try:
            with open(fp, encoding='utf-8') as f:
                d = json.load(f)
        except Exception:
            continue
        # mask_dict: {真实名: A代码}
        if fn == 'mask_dict.json' and isinstance(d, dict):
            names |= set(d.keys())
            continue
        # contract_mask / confirm_mask：递归收集字符串键
        def collect(x):
            if isinstance(x, dict):
                for k, v in x.items():
                    if isinstance(k, str) and len(k) >= 2:
                        names.add(k)
                    collect(v)
            elif isinstance(x, list):
                for i in x:
                    collect(i)
        collect(d)
    # 过滤明显非名称的键
    names = {n for n in names if len(n) >= 2 and re.search(r'[\u4e00-\u9fff]', n)}
    return names


_REAL_NAMES = _load_real_names()


def _is_mask_candidate(rel):
    """文件名/目录名含脱敏标记 → 该文件本身就是脱敏版。"""
    base = os.path.basename(rel).lower()
    if '_脱敏' in base or '脱敏版' in base or base.startswith('脱敏'):
        return True
    return False


def _has_paired_mask(root, f):
    """存在配对脱敏版（复用 mask_safe 统一逻辑：同目录/同级/父级 _脱敏 或 _脱敏版）。"""
    from mask_safe import _paired_mask
    return _paired_mask(root, f) is not None


def _scan_xlsx(fp):
    """读 xlsx 前 40 行前 12 列，返回字符串集合（read_only 高效）。"""
    vals = []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for ri, row in enumerate(ws.iter_rows(values_only=True, max_row=40, max_col=12)):
            for v in row:
                if isinstance(v, str) and v.strip() and len(v.strip()) >= 2:
                    vals.append(v.strip())
            if ri > 40:
                break
        wb.close()
    except Exception:
        pass
    return vals


def _scan_xls(fp):
    vals = []
    try:
        import xlrd
        wb = xlrd.open_workbook(fp)
        ws = wb.sheet_by_index(0)
        for r in range(min(40, ws.nrows)):
            for c in range(min(12, ws.ncols)):
                v = ws.cell_value(r, c)
                if isinstance(v, str) and v.strip() and len(v.strip()) >= 2:
                    vals.append(v.strip())
    except Exception:
        pass
    return vals


def _scan_text(fp):
    try:
        with open(fp, encoding='utf-8', errors='ignore') as f:
            return f.read(200000).split('\n')[:200]
    except Exception:
        return []


def analyze_file(fp, rel):
    """返回 ('masked'|'sensitive'|'clean', 命中样本数)。"""
    if _is_mask_candidate(rel):
        return 'masked', 0
    ext = os.path.splitext(fp)[1].lower()
    if ext == '.xlsx':
        vals = _scan_xlsx(fp)
    elif ext == '.xls':
        vals = _scan_xls(fp)
    elif ext in _TEXT_EXT:
        vals = _scan_text(fp)
    else:
        return 'clean', 0
    hit = 0
    for v in vals:
        if any(nm and nm in v for nm in _REAL_NAMES):
            hit += 1
        elif _COMPANY_PAT.search(v):
            hit += 1
        if hit >= 3:
            break
    if hit:
        return 'sensitive', hit
    return 'clean', 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='脱敏入口闸门：扫描敏感文件与脱敏状态')
    ap.add_argument('--dir', default=None, help='扫描目录（相对 DATA_ROOT 或绝对）')
    ap.add_argument('--gate', action='store_true', help='门禁模式：敏感未脱敏则退出码 2')
    ap.add_argument('--list', action='store_true', help='只列敏感未脱敏（默认）')
    a = ap.parse_args(argv)

    if a.dir:
        scan_root = a.dir if os.path.isabs(a.dir) else os.path.join(DATA_ROOT, a.dir)
        if not os.path.isdir(scan_root):
            print('[ERROR] 目录不存在：%s' % scan_root)
            return 1
    else:
        scan_root = DATA_ROOT

    print('=== 脱敏入口闸门 ===')
    print('扫描根：%s' % scan_root)
    print('真实名对照表条数：%d' % len(_REAL_NAMES))
    print()

    sensitive_miss = []
    sensitive_raw = []
    sensitive_ok = []
    clean_n = masked_n = 0
    scanned = 0
    for root, dirs, files in os.walk(scan_root):
        dirs[:] = [d for d in dirs if not any(x in d for x in _EXCLUDE_DIR)]
        for f in files:
            if f.startswith('~$'):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext in _EXCLUDE_EXT:
                continue
            fp = os.path.join(root, f)
            rel = os.path.relpath(fp, DATA_ROOT)
            scanned += 1
            status, hit = analyze_file(fp, rel)
            if status == 'masked':
                masked_n += 1
            elif status == 'sensitive':
                # 敏感文件本身是脱敏版？重复检查目录
                if _is_mask_candidate(rel):
                    masked_n += 1
                else:
                    # 有同文件名脱敏版（原始本机用 + 脱敏版进对话）→ 视为已脱敏
                    paired = _has_paired_mask(root, f)
                    # 本机中间产物（OCR 识别文本等，脱敏版在 *_脱敏 目录，程序已只读脱敏版）
                    mid = any(k in rel for k in _MID_DIR_KW)
                    # 原始账套数据（科目余额表/序时账/网银/往来明细…）→ 本机程序专用
                    raw = any(k in rel for k in _RAW_DATA_KW)
                    if paired or mid:
                        masked_n += 1
                    elif raw:
                        sensitive_raw.append(rel)
                    else:
                        sensitive_miss.append((rel, hit))
            else:
                clean_n += 1

    print('扫描文件数：%d（已脱敏 %d / 无敏感 %d / 敏感 %d）' % (scanned, masked_n, clean_n,
                                                      len(sensitive_miss) + len(sensitive_raw)))
    print()
    print('--- ① 产物类·敏感未脱敏（%d 个，进对话/分析前须先脱敏）---' % len(sensitive_miss))
    for rel, hit in sorted(sensitive_miss)[:60]:
        print('  ⚠️ %s (命中%d)' % (rel, hit))
    if len(sensitive_miss) > 60:
        print('  …共 %d 个' % len(sensitive_miss))
    print()
    print('--- ② 原始账套数据（%d 个，本机程序专用需真实名，不直接进对话）---' % len(sensitive_raw))
    for rel in sorted(sensitive_raw)[:15]:
        print('  · %s' % rel)
    if len(sensitive_raw) > 15:
        print('  …共 %d 个' % len(sensitive_raw))

    if a.gate and sensitive_miss:
        print('\n[GATE] 存在 %d 个产物类敏感未脱敏文件 → 拒绝进入分析' % len(sensitive_miss))
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
