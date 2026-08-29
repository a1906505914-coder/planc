# -*- coding: utf-8 -*-
"""launcher.py —— 统一启动器（2026-08-04，替代 13 个 *_launcher.py）。

用法：
    python launcher.py <模块> <数据文件夹> [--skip-finalize]
    模块：revenue / expense / pl / tax / payroll / ca / bank / equity /
          inventory / longterm / loan / rd / gp
          （或直接 builder 文件名，如 current_account_detail）

与 regen_all 共用同一调用机制（_run_module 兼容各 builder 入口），单模块跑完
统一做一次终审（audit_finalize：审定表注入/终审清理/对方科目补全/自动检查/公式浅色）。
--skip-finalize：开发模式，跳过终审链以提速（仅刷新底稿数据，不跑 recalc/checker/公式浅色）。
"""
import os
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

MODS = {
    'revenue': 'revenue_detail',
    'expense': 'expense_detail',
    'pl': 'pl_detail',
    'tax': 'tax_detail',
    'payroll': 'payroll_detail',
    'ca': 'current_account_detail',
    'bank': 'bank_deposit_detail',
    'equity': 'equity_detail',
    'inventory': 'inventory_detail',
    'longterm': 'longterm_assets_detail',
    'loan': 'loan_detail',
    'rd': 'rd_expense_detail',
    'gp': 'gp_other',
}


def _crash_log(module, exc):
    try:
        import datetime
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        logp = os.path.join(HERE, 'crash_%s_%s.log' % (module, ts))
        with open(logp, 'w', encoding='utf-8') as f:
            f.write(''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        print('[FATAL] 已写入崩溃日志：%s' % logp)
    except Exception:
        pass


def _run_module(module_name, data_dir, extra_args=None):
    """在进程内调用某 builder 的处理函数（与 regen_all._run_module 同机制）。
    extra_args：2026-08-06 透传给 builder 的额外参数（如 ['--only-subj', 'AR'] 单科目重跑）。"""
    saved_argv = sys.argv
    sys.argv = [module_name + '.py', data_dir] + list(extra_args or [])
    try:
        if module_name in sys.modules:
            del sys.modules[module_name]
        mod = __import__(module_name)
        if hasattr(mod, 'build_all'):
            mod.build_all(data_dir)
        elif hasattr(mod, 'main'):
            mod.main()
        else:
            return False, '无 main()/build_all()'
        return True, 'OK'
    except SystemExit as se:
        return (True, 'OK(exit)') if (se.code in (None, 0)) else (False, 'exit=%s' % se.code)
    except Exception as ex:
        traceback.print_exc()
        return False, 'ERROR: %s' % ex
    finally:
        sys.argv = saved_argv


def main():
    argv = sys.argv[1:]
    if len(argv) < 2 or argv[0] in ('-h', '--help'):
        print(__doc__)
        return
    skip_finalize = '--skip-finalize' in argv
    argv = [a for a in argv if a != '--skip-finalize']
    # 2026-08-06 单科目重跑透传：python launcher.py ca <文件夹> --only-subj AR
    only_subj = None
    if '--only-subj' in argv:
        _i = argv.index('--only-subj')
        if _i + 1 < len(argv):
            only_subj = argv[_i + 1]
        argv = [a for a in argv if a != '--only-subj' and a != only_subj]
    key, data_dir = argv[0], argv[1]
    module = MODS.get(key, key.replace('.py', ''))
    if not os.path.isdir(data_dir):
        print('ERROR: 无效文件夹：%s' % data_dir)
        return
    print('>>> launcher: 模块=%s 文件夹=%s%s%s' % (module, data_dir,
          '（开发模式：跳过终审）' if skip_finalize else '',
          f'（单科目：{only_subj}）' if only_subj else ''))
    if skip_finalize:
        os.environ['SKIP_FINALIZE'] = '1'
    extra = ['--only-subj', only_subj] if only_subj else None
    ok, msg = _run_module(module, data_dir, extra)


if __name__ == '__main__':
    main()
