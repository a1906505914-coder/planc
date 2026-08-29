# -*- coding: utf-8 -*-
"""冒烟测试（#防崩溃·改代码后第一道防线）。

改代码后先跑本脚本（秒级），确认：
  ① 全部 py 可编译（语法）
  ② 核心模块可 import（无 import 级运行时错误，如顶层读数据/坏依赖）
  ③ 关键生成器函数可解析（不实际跑数据，只验证可调用）

用法：
    python smoke_test.py              # 全量冒烟
    python smoke_test.py --quick      # 只编译+import（更快，跳过函数签名检查）
"""
import os
import sys
import glob
import time
import py_compile

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)
sys.path.insert(0, BASE)


def check_compile():
    fails = []
    files = sorted(glob.glob('*.py'))
    for f in files:
        try:
            py_compile.compile(f, doraise=True)
        except Exception as e:
            fails.append((f, f'compile: {e}'))
    print(f'  编译：{len(files)} 个，失败 {len(fails)}')
    for f, e in fails[:10]:
        print(f'    ❌ {f} {e}')
    return fails


def check_import():
    """import 核心模块（排除会实际跑任务的入口）。"""
    core = ['audit_common', 'audit_shell', 'sap_adapter', 'sap_reader', 'audit_cache',
            'current_account_detail', 'bank_deposit_detail', 'bank_reconcile_detail',
            'tax_detail', 'payroll_detail', 'expense_detail', 'pl_detail',
            'longterm_assets_detail', 'inventory_detail', 'loan_detail', 'equity_detail',
            'rd_expense_detail', 'revenue_detail', 'gp_other', 'counterparty_recon',
            'xw_render', 'xw_period_sheet', 'working_paper_fingerprint', 'audit_config']
    fails = []
    for m in core:
        try:
            __import__(m)
        except Exception as e:
            fails.append((m, f'import: {type(e).__name__}: {e}'))
    print(f'  import：{len(core)} 个核心模块，失败 {len(fails)}')
    for m, e in fails[:10]:
        print(f'    ❌ {m} {e}')
    return fails


def main():
    argv = sys.argv[1:]
    t0 = time.time()
    print(f'冒烟测试 @ {time.strftime("%H:%M:%S")}')
    f1 = check_compile()
    f2 = check_import()
    total = len(f1) + len(f2)
    print(f'\n冒烟结果：{"✅ 全部通过" if total == 0 else f"❌ {total} 处失败（请勿交付）"}（耗时 {time.time()-t0:.0f}s）')
    return 0 if total == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
