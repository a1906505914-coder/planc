# -*- coding: utf-8 -*-
"""固定资产复核 统一专项（2026-08-18 用户定：同类合并为一个专项，配置适配账套）。

合并对象：
  · XBJ  → depreciation_review（计提复核）
  · GJX  → fa_movement（增减变动·卡片xls+GL）
  · ADF  → fa_adf（增减变动·SAP台账多主体）

机制：fa_manifest.jobs()[acct] 取配置，按 kind 分派到对应模块；
     模块级常量由本文件按配置 setattr 覆盖（原脚本逻辑零改动，结果与旧版一致）。
新增账套：fa_manifest.jobs() 加一行即可，不新建脚本。

用法：
  python fa_review.py XBJ           # 单账套
  python fa_review.py GJX --out .   # 指定输出目录
"""
import os
import sys

import paths as P
import fa_manifest as FM

DATA_ROOT = P.DATA_ROOT


def gen(acct, out_dir=None):
    cfg = FM.job_for(acct)
    if not cfg:
        raise KeyError(f'fa_manifest 无配置: {acct}')
    kind = cfg['kind']

    # ---- 计提复核：depreciation_review.build(cfg, out_path) ----
    if kind == 'accrual':
        import depreciation_review as DR
        import depreciation_manifest as DM
        dcfg = DM.job_for(acct)
        if not dcfg:
            raise KeyError(f'depreciation_manifest 无配置: {acct}')
        year = dcfg.get('year', 2025)
        out_path = out_dir or os.path.join(DATA_ROOT, acct, '底稿', str(year),
                                           f'固定资产折旧计提复核_{year}.xlsx')
        return DR.build(dcfg, out_path)

    # ---- 增减变动·卡片xls+GL：fa_movement.run_period ----
    if kind == 'movement':
        import fa_movement as FMV
        FMV.ACCT = acct
        FMV.YEAR = cfg['year']
        FMV.CARD_DIR = cfg['card_dir']
        FMV.GL25 = cfg['gl']['2025']
        FMV.GL26 = cfg['gl']['2026']
        FMV.TB25 = cfg['tb']['2025']
        FMV.TB26 = cfg['tb']['2026']
        FMV.CARD_GL_CLASS = cfg['card_gl_class']
        FMV.GL_CLASSES = cfg['gl_classes']
        outs = []
        for p in cfg['periods']:
            outs.append(FMV.run_period(p, out_dir))
        return outs

    # ---- 增减变动·SAP台账多主体：fa_adf.main ----
    if kind == 'movement_sap':
        import fa_adf as FA
        FA.ACCT = acct
        FA.YEAR = str(cfg['year'])
        FA.COMPS = cfg['comps']
        FA.END_YM = cfg['end_ym']
        FA.LEDGER_END = cfg['ledger_end']
        FA.LEDGER_BEG_DIR = cfg['ledger_beg_dir']
        FA.TB_DIR = cfg['tb_dir']
        FA.CLS_MAP = cfg['cls_map']
        FA.GL_CLASSES = cfg['gl_classes']
        FA.TB_CLASS = cfg['tb_class']
        FA.TB_CLASS_DEP = cfg['tb_class_dep']
        return FA.main()

    raise KeyError(f'未知 kind: {kind}')


def main():
    import argparse
    ap = argparse.ArgumentParser(description='固定资产复核统一专项')
    ap.add_argument('acct', help='账套（XBJ/GJX/ADF，见 fa_manifest）')
    ap.add_argument('--out', default=None, help='输出目录（默认 底稿/{year}）')
    a = ap.parse_args()
    outs = gen(a.acct, a.out)
    print(f'{a.acct}: 生成 {len(outs) if isinstance(outs, list) else 1} 个文件')
    return 0


if __name__ == '__main__':
    sys.exit(main())
