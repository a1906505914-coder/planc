# -*- coding: utf-8 -*-
"""doc_router —— 外部证据识别 → 路由归集 → 注入科目底稿（2026-08-16 铁律131i）。

架构（用户 11:38 洞察：外部资料先有共同性、最终落实到具体科目底稿）：
  外部资料(合同/函证/银行/保函) → 统一识别层(OCR+解析) → 结构化对象(doc_type)
    → 路由表(doc_type → 科目/处理模块) → 生成核对底稿 → 注入科目底稿("前缀-"sheets)

路由表 ROUTES（新增识别类型只加一行）：
  doc_type        subject      label         module          prefix
  bank_list       银行存款     开户账户核对    bank_acct_recon  开户核对-
  bank_statement  银行存款     对账单核对      (待接入)         对账单-
  confirm_bank    银行存款     银行函证回函    (待接入)         银行函证-
  confirm_ar      应收账款     往来函证回函    (待接入)         往来函证-
  contract_sale   营业收入     销售合同核对    (待接入)         合同核对-
  contract_buy    存货/应付    采购合同核对    (待接入)         采购合同-

用法：
  python doc_router.py <资料目录> --acct AYL --year 2025            # 全自动：识别→路由→注入
  python doc_router.py --route bank_list --acct AYL --year 2025      # 只跑指定类型
  python doc_router.py --list                                        # 列路由表
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import argparse
import glob
import json
import os
import re
import shutil
import sys

DATA_ROOT = paths.DATA_ROOT

# 路由表：doc_type → 归集科目/处理模块/sheet 前缀
ROUTES = {
    'bank_list': {
        'subject': '银行存款', 'label': '开户账户核对', 'module': 'bank_acct_recon',
        'func': 'run_recon', 'prefix': '开户核对-',
        'target': '银行存款审计底稿_{year}_生成.xlsx',
    },
    'bank_statement': {
        'subject': '银行存款', 'label': '银行对账单核对', 'module': None,
        'prefix': '对账单核对-', 'target': '银行存款审计底稿_{year}_生成.xlsx',
    },
    'confirm_bank': {
        'subject': '银行存款', 'label': '银行函证回函', 'module': None,
        'prefix': '银行函证-', 'target': '银行存款审计底稿_{year}_生成.xlsx',
    },
    'confirm_ar': {
        'subject': '应收账款', 'label': '往来函证回函', 'module': None,
        'prefix': '往来函证-', 'target': '应收账款审计底稿_{year}_生成.xlsx',
    },
    'contract_sale': {
        'subject': '营业收入', 'label': '销售合同核对', 'module': None,
        'prefix': '合同核对-', 'target': '营业收入审计底稿_{year}_生成.xlsx',
    },
    'contract_buy': {
        'subject': '存货', 'label': '采购合同核对', 'module': None,
        'prefix': '采购合同-', 'target': '存货审计底稿_{year}_生成.xlsx',
    },
}

# 识别类型判定关键词（统一识别层：内容特征 → doc_type）
_CLASSIFY_RULES = [
    ('bank_list', ['已开立银行结算账户清单', '开立银行账户清单']),
    ('bank_statement', ['月结单', '对账单', 'Statement', '承前余额', '承前轉結', 'B/F']),
    ('confirm', ['询证函', '回函']),
    ('contract', ['合同', '工程名称', '发包人', '乙方']),
]


def classify(text, fname=''):
    """内容特征 → doc_type（合同/函证细分为销售/采购/往来/银行）。"""
    if not text:
        return ''
    if any(k in text for k in _CLASSIFY_RULES[0][1]):
        return 'bank_list'
    if any(k in text for k in _CLASSIFY_RULES[1][1]):
        return 'bank_statement'
    if '询证函' in text or '回函' in text:
        if '银行' in text or '存款' in text:
            return 'confirm_bank'
        return 'confirm_ar'
    if '合同' in text:
        if any(k in text for k in ('采购', '购买', '供应商', '供货')):
            return 'contract_buy'
        return 'contract_sale'
    return ''


def recognize(src_dir, acct='', year=''):
    """统一识别层：扫描目录 → 每文件 OCR/解析 → {file, doc_type, text} 清单。"""
    import contract_reader as CR
    exts = ('.pdf', '.docx', '.doc', '.xlsx', '.txt', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff')
    files = []
    for root, _, fs in os.walk(src_dir):
        for f in fs:
            if f.lower().endswith(exts) and not f.startswith('~$'):
                files.append(os.path.join(root, f))
    files.sort()
    out = []
    for f in files:
        try:
            text, fmt = CR.extract_text(f)
            dt = classify(text, os.path.basename(f))
            out.append({'file': os.path.basename(f), 'path': f, 'fmt': fmt,
                        'doc_type': dt, 'chars': len(text)})
            if dt:
                print(f'  ✅ {os.path.basename(f)[:40]} → {dt}')
            else:
                print(f'  ⚪ {os.path.basename(f)[:40]} 未识别类型')
        except Exception as ex:
            print(f'  ⚠️ {os.path.basename(f)[:40]} 识别失败: {ex}')
    return out


def _copy_sheet(src_ws, dst_wb, name):
    """跨工作簿复制 sheet（值+轻量样式+列宽+冻结）。"""
    import openpyxl
    ws = dst_wb.create_sheet(title=name)
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                ws.cell(row=cell.row, column=cell.column, value=cell.value)
    if not src_ws.parent.read_only:
        for row in src_ws.iter_rows():
            for cell in row:
                if cell.has_style:
                    nc = ws.cell(row=cell.row, column=cell.column)
                    nc.font = openpyxl.styles.Font(name=cell.font.name, size=cell.font.size,
                                                   bold=cell.font.bold)
                    nc.number_format = cell.number_format
        for col, dim in src_ws.column_dimensions.items():
            if dim.width:
                ws.column_dimensions[col].width = dim.width
    ws.freeze_panes = src_ws.freeze_panes
    return ws


def _append_recon_summary(dst_wb, src_wb, prefix, src_xlsx, target_xlsx):
    """注入后附加「核对摘要」sheet（2026-08-22 注入升级：复制 sheet → 附带勾稽提示）。
    摘要内容：来源文件、注入时间、各 sheet 数值合计（供人工与目标底稿审定数勾稽）。
    保守策略：只列来源侧数值合计 + 提示，不强行取目标审定数（各底稿结构不同，避免错判）。"""
    import datetime
    ws = dst_wb.create_sheet(f'核对摘要-{prefix}')
    ws.append(['注入核对摘要'])
    ws.append(['目标底稿', os.path.basename(target_xlsx)])
    ws.append(['来源核对表', os.path.basename(src_xlsx)])
    ws.append(['注入时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    ws.append([])
    ws.append(['来源 sheet', '数值列合计', '说明'])
    for s in src_wb.sheetnames:
        src_ws = src_wb[s]
        tot = 0.0
        has_num = False
        for row in src_ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, (int, float)):
                    tot += float(v)
                    has_num = True
        ws.append([s, round(tot, 2) if has_num else '', '数值合计，人工与目标审定数勾稽'])
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 34


def inject(target_xlsx, src_xlsx, prefix):
    """把核对底稿全部 sheets 注入科目底稿（前缀命名，跳过已存在同名 sheet）。
    2026-08-22 升级：注入后附加「核对摘要」sheet（勾稽提示，复制→核对）。"""
    import openpyxl
    if not os.path.exists(target_xlsx):
        print(f'  ⚠️ 目标科目底稿不存在: {os.path.basename(target_xlsx)}')
        return 0
    # 备份
    bak = target_xlsx.replace('.xlsx', '_bak_before_inject.xlsx')
    if not os.path.exists(bak):
        shutil.copy2(target_xlsx, bak)
    src_wb = openpyxl.load_workbook(src_xlsx)
    dst_wb = openpyxl.load_workbook(target_xlsx)
    added = 0
    for s in src_wb.sheetnames:
        name = f'{prefix}{s}'
        if name in dst_wb.sheetnames:
            del dst_wb[name]
        _copy_sheet(src_wb[s], dst_wb, name)
        added += 1
    try:
        _append_recon_summary(dst_wb, src_wb, prefix, src_xlsx, target_xlsx)
    except Exception as ex:
        print(f'  ⚠️ 核对摘要生成失败：{ex}')
    dst_wb.save(target_xlsx)
    print(f'  ✅ 注入 {added} 个 sheets + 核对摘要（{prefix}…）→ {os.path.basename(target_xlsx)}')
    return added


def route(doc_type, acct, year, src_dir=None):
    """路由执行：doc_type → 处理模块 → 核对底稿 → 注入科目底稿。"""
    r = ROUTES.get(doc_type)
    if not r:
        print(f'⚠️ 未知 doc_type: {doc_type}')
        return None
    year = str(year)
    out_root = os.path.join(DATA_ROOT, acct, '底稿', year)
    os.makedirs(out_root, exist_ok=True)
    if r['module'] == 'bank_acct_recon':
        import bank_acct_recon as BR
        # 参数：ocr_dir 指向 银行资料/_ocr_识别结果（或资料目录本身）
        ocr_dir = src_dir
        if src_dir and os.path.exists(os.path.join(src_dir, '_ocr_识别结果')):
            ocr_dir = os.path.join(src_dir, '_ocr_识别结果')
        recon_fp = BR.run_recon(acct, year, ocr_dir=ocr_dir, out_dir=out_root)
        target = os.path.join(out_root, r['target'].format(year=year))
        inject(target, recon_fp, r['prefix'])
        return recon_fp
    print(f'⚠️ {doc_type} 处理模块未接入（ROUTES 表已预留 {r["label"]}）')
    return None


def main():
    ap = argparse.ArgumentParser(description='外部证据识别 → 路由归集 → 注入科目底稿')
    ap.add_argument('src', nargs='?', default=None, help='资料目录（自动识别全类型）')
    ap.add_argument('--acct', default='AYL', help='账套（默认 AYL）')
    ap.add_argument('--year', default='2025', help='年度（默认 2025）')
    ap.add_argument('--route', default=None, help='只跑指定 doc_type（如 bank_list）')
    ap.add_argument('--list', action='store_true', help='列路由表')
    a = ap.parse_args()

    if a.list:
        print('doc_type        → 归集科目 / 标签 / 前缀')
        for k, v in ROUTES.items():
            print(f'  {k:16} → {v["subject"]} / {v["label"]} / {v["prefix"]}')
        return

    if a.route:
        route(a.route, a.acct, a.year)
        return

    if not a.src:
        ap.print_help()
        return
    # 全自动：识别 → 逐类型路由
    print(f'=== 统一识别 {a.src} ===')
    objs = recognize(a.src, a.acct, a.year)
    doc_types = sorted({o['doc_type'] for o in objs if o['doc_type']})
    print(f'识别到类型: {doc_types}')
    for dt in doc_types:
        print(f'\n=== 路由 {dt} → {ROUTES.get(dt, {}).get("subject", "?")} ===')
        route(dt, a.acct, a.year, src_dir=a.src)


if __name__ == '__main__':
    main()
