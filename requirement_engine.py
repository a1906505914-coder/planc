# -*- coding: utf-8 -*-
"""审计要求清单引擎（2026-08-14 架构试点）。

工作流（用户方法论：给数据 → 先问要求 → 要求明确后按需生成）：
    1. 准备一份要求清单 JSON（科目+要求勾选 + 参数）
    2. 引擎读清单 → 按序调用已注册的要求知识点 → 汇总输出文件

清单格式（requirement_manifest.json）：
    {
      "comp": "3300",                      # 主体代码
      "year": "2026",                      # 年度
      "data_dir": "<账套根>/3300网银明细2026年1-7月",   # 数据目录
      "subjects": ["bank", "current_account"],   # 可选：科目层（13 个科目生成器，先生成底稿）
      "requirements": [                    # 要求层勾选清单（后执行，按序）
        {
          "name": "bank_reconcile",        # 知识点名称（注册表键）
          "params": {                      # 该知识点的输入参数（可选）
            "stmt_folder": "...",
            "journal_file": "..."
          }
        }
      ]
    }

用法：
    python requirement_engine.py requirement_manifest.json
    python requirement_engine.py --list          # 列出全部已注册知识点
    python requirement_engine.py --make-example  # 生成示例清单
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import paths as P  # noqa: E402
import requirements  # noqa: E402  （触发注册表 import，加载全部知识点）
import requirements.bank_reconcile  # noqa: E402,F401  注册银行核对知识点


def _load_all():
    """显式 import 各知识点模块（注册）。新增知识点在此登记。"""
    import requirements.bank_reconcile  # noqa: E402,F401
    import requirements.cutoff_test  # noqa: E402,F401
    import requirements.gross_margin  # noqa: E402,F401
    import requirements.confirmation_list  # noqa: E402,F401
    import requirements.audit_sampling  # noqa: E402,F401
    import requirements.cross_recon  # noqa: E402,F401
    import requirements.inventory_count  # noqa: E402,F401
    import requirements.bank_abnormal  # noqa: E402,F401
    import requirements.diff_intelligence  # noqa: E402,F401


def run_manifest(path):
    """按清单文件执行（自动注册全部知识点——编程调用不依赖外部先 _load_all）。"""
    _load_all()
    with open(path, 'r', encoding='utf-8') as f:
        man = json.load(f)
    return run_manifest_from(man)


def run_manifest_from(man):
    """按清单 dict 执行（regen_all 要求层挂接/编程调用用）。自动注册知识点。

    ⚡ 2026-08-14 路径约定：清单中所有路径字段（data_dir 及 params 里的
    stmt_folder/journal_file 等）一律写【相对 DATA_ROOT 的路径】，
    由 P.resolve_rel() 统一展开（绝对路径也容忍）。"""
    _load_all()
    ctx = {
        'data_dir': P.resolve_rel(man.get('data_dir', '') or '') or '',
        'comp': man.get('comp'),
        'year': man.get('year', '2026'),
        'verbose': man.get('verbose', True),
    }
    if not ctx['data_dir']:
        print('[ERROR] 清单缺少 data_dir')
        return 1
    outputs = []
    # ── 阶段 1：科目层（可选）——先生成科目底稿（13 个科目生成器） ──
    subjects = man.get('subjects') or []
    if subjects:
        print(f'▶ 阶段1 科目底稿生成：{", ".join(subjects)}')
        # ⚡⚡ 2026-08-14 用【子进程】跑科目层：run_u8_on_sap.main 的 patch_audit_common()
        #   是全局副作用（把 audit_common 数据接口切到 SAP 适配版），同一进程内会污染
        #   后续要求知识点（U8 账套 confirmation_list 解析 TB 列错位 float('借') 实证）——
        #   子进程隔离保证主进程状态干净。也符合铁律115（独立进程）。
        import subprocess
        _py = sys.executable
        _eng = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'run_u8_on_sap.py')
        cmd = [_py, _eng, ctx['data_dir']]
        if ctx['comp']:
            cmd += ['--only', ctx['comp']]
        cmd += ['--subj', ','.join(subjects)]
        print('  CMD:', ' '.join(cmd))
        rc = subprocess.call(cmd)
        if rc != 0:
            print(f'  ✗ 科目生成子进程退出码 {rc}')
    # ── 阶段 2：要求层——逐个执行要求知识点 ──
    for req in man.get('requirements', []):
        name = req.get('name')
        if name not in requirements.REQUIREMENTS:
            print(f'[ERROR] 未注册的要求知识点：{name}（可用 --list 查看）')
            continue
        meta = requirements.REQUIREMENTS[name]
        rctx = dict(ctx)
        for k, v in (req.get('params') or {}).items():
            rctx[k] = v
        print(f'▶ 执行要求知识点：{name} —— {meta["desc"]}')
        try:
            outs = meta['fn'](rctx)
            if outs:
                outputs.extend(outs)
                print(f'  ✓ 产出 {len(outs)} 个文件')
        except Exception as e:
            import traceback
            print(f'  ✗ 执行失败：{e}')
            traceback.print_exc()
    print()
    print('===== 执行完成 =====')
    for p in outputs:
        print('  输出：', p)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='审计要求清单引擎')
    ap.add_argument('manifest', nargs='?', help='要求清单 JSON 路径')
    ap.add_argument('--list', action='store_true', help='列出全部已注册知识点')
    ap.add_argument('--for-subject', help='列出适用于某科目的知识点（如 银行存款）')
    ap.add_argument('--make-example', action='store_true', help='生成示例清单')
    args = ap.parse_args(argv)
    _load_all()
    if args.list:
        print('已注册要求知识点：')
        for r in requirements.list_requirements():
            print(f'  {r["name"]:<18} 依赖科目: {", ".join(r["subjects"]) or "-"}')
            print(f'    {"":<18} {r["desc"]}')
            if r['params']:
                print(f'    {"":<18} 参数: {", ".join(r["params"])}')
        return 0
    if args.for_subject:
        subj = args.for_subject
        print(f'适用于科目「{subj}」的要求知识点：')
        hits = [r for r in requirements.list_requirements() if subj in r['subjects']]
        if not hits:
            print('  （无）')
        for r in hits:
            print(f'  {r["name"]:<18} {r["desc"]}')
        return 0
    if args.make_example:
        import bank_reconcile_manifest as M
        j3300 = M.job_for('3300')  # ⚡ 2026-08-14 示例路径从 manifest 动态生成（取消硬编码）
        dd, stmt, jf = (j3300[1], j3300[2], j3300[3]) if j3300 else ('', '', '')
        # ⚡ 2026-08-14 示例一律写【相对 DATA_ROOT 路径】——账套拷贝/改名零改动
        dd = os.path.relpath(dd, P.DATA_ROOT) if dd else ''
        stmt = os.path.relpath(stmt, P.DATA_ROOT) if stmt else ''
        jf = os.path.relpath(jf, P.DATA_ROOT) if jf else ''
        ex = {
            'comp': '3300', 'year': '2026',
            'data_dir': dd,
            'requirements': [
                {'name': 'bank_reconcile',
                 'params': {
                     'stmt_folder': stmt,
                     'journal_file': jf,
                 }},
                {'name': 'cutoff_test', 'params': {}},
                {'name': 'gross_margin', 'params': {}},
                {'name': 'confirmation_list', 'params': {}},
                {'name': 'audit_sampling', 'params': {}},
            ],
        }
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'requirement_manifest.example.json')
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(ex, f, ensure_ascii=False, indent=2)
        print(f'示例清单已生成（含全部 5 个知识点）：{out}')
        print('  路径字段为相对 <DATA_ROOT> 路径（引擎自动展开为绝对路径）')
        print('  用法：复制后改 data_dir/comp/year/params，即可按需生成')
        return 0
    if not args.manifest:
        ap.print_help()
        return 1
    return run_manifest(args.manifest)


if __name__ == '__main__':
    sys.exit(main())
