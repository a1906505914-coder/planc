# -*- coding: utf-8 -*-
"""银行专项统一 manifest（2026-08-18 用户定：银行三个程序合并为一个专项，配置适配账套）。

此前 bank_reconcile_detail（双向核对）/bank_acct_recon（开户清单核对）/bank_abnormal（发生额异常）
三个程序独立、形态重叠。现统一为一个专项（bank_review.py），本文件管配置。
新增账套 = 加一行 forms，不新建脚本。

形态 forms（一个账套可多个）：
  reconcile  双向核对（bank_reconcile_detail.build_bank_reconcile）——网银流水 vs 银行日记账
             ⚡ 账套内主体/网银/日记账配置仍以 bank_reconcile_manifest.ACCTS 为唯一事实来源
  acct       开户清单核对（bank_acct_recon.run_recon）——银行开户清单(OCR) vs 账面账户
  abnormal   银行存款发生额异常分析（bank_abnormal.build）——序时账大额/异常发生筛查
"""
import bank_reconcile_manifest as BRM  # 双向核对账套配置唯一来源


def jobs():
    jobs_ = {}
    for acct in BRM.ACCTS:
        jobs_[acct] = {'forms': ['reconcile']}
    # ⚡ AYL：开户清单核对 + 发生额异常分析（paths.G == AYL 目录）
    jobs_['AYL'] = {'forms': ['acct', 'abnormal'], 'year': '2025'}
    return jobs_


def job_for(acct):
    return jobs().get(acct)


if __name__ == '__main__':
    for acct, cfg in jobs().items():
        print(f'{acct:6s} forms={cfg["forms"]}')
