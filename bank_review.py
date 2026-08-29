# -*- coding: utf-8 -*-
"""银行专项统一入口（2026-08-18 用户定：合并 bank_reconcile_detail / bank_acct_recon / bank_abnormal 为一个专项）。

机制：bank_manifest.jobs()[acct] 取 forms，逐个路由到对应模块（复用原逻辑，结果与旧版一致）。
新增账套：bank_manifest 加一行即可，不新建脚本。

用法：
  python bank_review.py ADF           # 双向核对（按 bank_reconcile_manifest 主体逐个）
  python bank_review.py AYL           # 开户清单核对 + 发生额异常分析
"""
import os
import sys

import paths as P
import bank_manifest as BM

ROOT = P.DATA_ROOT


def gen(acct, out_dir=None):
    cfg = BM.job_for(acct)
    if not cfg:
        raise KeyError(f'bank_manifest 无配置: {acct}')
    outs = []
    for form in cfg['forms']:
        if form == 'reconcile':
            import bank_reconcile_detail as BRD
            import bank_reconcile_manifest as BRM
            # ⚡ 用 jobs() 完整任务（含 stmt/journal 相对路径、排除大主体）。
            #   修正原 gen_one 遍历 subjects + stmt=None 的缺陷（ADF/GJX 均找不到网银目录）。
            for comp, data_dir, stmt, jf, year in BRM.jobs(acct=acct):
                out = BRD.build_bank_reconcile(data_dir, stmt, comp=comp, year=str(year),
                                               journal_file=jf)
                if out:
                    outs.append(out)
        elif form == 'acct':
            import bank_acct_recon as BAC
            year = cfg.get('year', '2025')
            out = BAC.run_recon(acct, year, ocr_dir=cfg.get('ocr_dir'),
                                out_dir=out_dir or cfg.get('out_dir'))
            if out:
                outs.append(out)
        elif form == 'abnormal':
            import bank_abnormal as BA
            data_dir = cfg.get('data_dir') or P.G
            out_path = cfg.get('out_path') or os.path.join(P.G, '银行存款发生额异常分析_生成.xlsx')
            out = BA.build(data_dir, out_path)
            if out:
                outs.append(out)
        else:
            raise KeyError(f'未知 form: {form}')
    return outs


def main():
    import argparse
    ap = argparse.ArgumentParser(description='银行专项统一入口')
    ap.add_argument('acct', help='账套（ADF/GJX/AYL，见 bank_manifest）')
    ap.add_argument('--out', default=None, help='输出目录（覆盖默认）')
    a = ap.parse_args()
    outs = gen(a.acct, a.out)
    print(f'{a.acct}: 生成 {len(outs)} 个文件')
    return 0


if __name__ == '__main__':
    sys.exit(main())
