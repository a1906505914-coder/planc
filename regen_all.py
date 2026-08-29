# -*- coding: utf-8 -*-
"""regen_all.py —— 6 阶段流水线总控（2026-08-06 架构升级）。

用户方法论（2026-08-06）：拿到账套 → ① 程序自建试算表（摸结构/勾稽）→ ② 生成全套底稿 →
③ 底稿 vs 自建试算表核对（同源，差异=取数 bug）→ ④ 重分类调整分录 → ⑤ 报告生成（母分汇总/
合并报表，程序只算口径数据）→ ⑥ 外部试算表核对（自建 vs 用户试算表=审计调整/合并抵消口径差）。

各阶段：
  ① self_tb_gen     自建试算表（单体 TB 直出，叶子过滤，平衡校验+报表勾稽）
  ② 13 builder      生成全套底稿（原有能力）
  ③ recon_self      底稿审定数 vs 自建试算表（同源核对，差异=取数 bug）
  ④ reclass_gen     重分类调整分录生成器（从明细表负余额生成分录）——待接入
  ⑤ report_gen      报告生成（形态判定：单体/母分汇总/合并；程序只算口径数据）——待接入
  ⑥ recon_ext       外部试算表核对（自建 vs 用户试算表）——待接入（现 tb_recon_detail 独立跑）

用法：
  python regen_all.py <文件夹> [<另一文件夹> ...] [--skip-finalize] [--phase 1,3]
  例：python regen_all.py FY --phase 1,3   # 只跑自建试算表 + 同源核对
      python regen_all.py FY               # 默认全流程（①②③ + 已实现的后续阶段）
      python regen_all.py FY --skip-finalize   # 开发模式：跳过终审
"""
import os
import sys
import traceback

import paths as P

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# 13 个 builder 模块 + 通用勾稽
BUILDERS = [
    ("revenue_detail", "收入"),
    ("expense_detail", "费用"),
    ("pl_detail", "利润表"),
    ("tax_detail", "应交税费"),
    ("payroll_detail", "职工薪酬"),
    ("current_account_detail", "往来款"),
    ("bank_deposit_detail", "银行存款"),
    ("equity_detail", "所有者权益"),
    ("inventory_detail", "存货"),
    ("longterm_assets_detail", "长期资产"),
    ("loan_detail", "借款"),
    ("rd_expense_detail", "研发费用"),
    ("gp_other", "其他类"),
]

# 6 阶段流水线（key → (模块名, 标签, 是否已实现)）
PHASES = [
    ("self_tb", "self_tb_gen", "① 自建试算表", True),
    ("build",   None,           "② 13 builder 底稿", True),
    ("recon_self", "recon_self", "③ 底稿 vs 自建试算表核对", True),
    # ⚡ 2026-08-24 ④⑥ 接骨架（reclass_gen/recon_ext，设计见 docs/audit_pipeline_design.md）：
    #   安全边界——只产建议/差异，不改写审定数；⑤ report_gen 仍预留未实现。
    ("reclass", "reclass_gen", "④ 重分类调整分录", True),
    ("report",  "report_gen",  "⑤ 报告生成（母分汇总/合并）", False),
    ("recon_ext", "recon_ext", "⑥ 外部数据核对", True),
]


def _run_module(module_name, label, data_dir):
    """在进程内调用某 builder 的处理函数。finalize 改为统一在全部 builder 跑完后只做一次，
    避免每 builder 都跑一遍 finalize_folder 造成的重复处理（控制台看到同一个文件名多次注入 审定表/补全 对方科目）。"""
    saved_argv = sys.argv
    sys.argv = [module_name + ".py", data_dir]
    try:
        # 强制重新加载模块，避免 __import__ 缓存旧代码（如 _entity_name 修复）
        if module_name in sys.modules:
            del sys.modules[module_name]
        mod = __import__(module_name)
        # 不同 builder 入口函数不统一：gp_other 只有 build_all；其余均有 main()
        if hasattr(mod, "build_all"):
            mod.build_all(data_dir)
        elif hasattr(mod, "main"):
            mod.main()
        else:
            return False, "无 main()/build_all()"
        return True, "OK"
    except SystemExit as se:
        return (True, "OK(exit)") if (se.code in (None, 0)) else (False, "exit=%s" % se.code)
    except Exception as ex:
        traceback.print_exc()
        return False, "ERROR: %s" % ex
    finally:
        sys.argv = saved_argv


def _finalize_once(data_dir):
    """全部 builder 跑完后，统一做一次终审：注入审定表 + 终审清理 + fill_voucher_counterparty 补全对方科目。"""
    try:
        import audit_finalize
        audit_finalize.finalize_folder(data_dir)
    except Exception as ex:
        print(f"  ⚠️ finalize 异常：{ex}")


def main():
    # 编排模式标记：builder 的 finalize_after_build 检测到该变量则跳过单跑终审，
    # 由本编排器在全部 builder 跑完后统一 finalize 一次（避免每 builder 重复扫描全文件夹）。
    os.environ['REGEN_ALL_MODE'] = '1'
    argv = sys.argv[1:]
    # ⚡ 2026-08-12 修复：全量重跑前清磁盘缓存（.cache/*.pkl）——本次 DQ 2024 期权益/
    #   往来/专项应付款全缺的根因之一：旧缓存存小写 'dq' 实体，discover 是 'DQ'，
    #   生成器读缓存匹配不到 → 误判"无数据"跳过。全量重跑=数据重建，缓存必须失效。
    #   --keep-cache 可跳过（增量单科目跑时保留缓存提速）。
    keep_cache = '--keep-cache' in argv
    if keep_cache:
        argv = [a for a in argv if a != '--keep-cache']
    else:
        # ⚡ 2026-08-12：磁盘缓存 os.remove 在 Windows 沙箱下被拦截（回收站不可用）→
        #   改用 AUDIT_NO_CACHE=1 环境变量禁用缓存（audit_cache.load 检测该变量）。
        #   全量重跑=数据重建，缓存必须失效，否则旧实体名/旧结构会污染新生成
        #   （本次 DQ 2024 期权益/往来/专项应付款全缺的根因之一）。
        os.environ['AUDIT_NO_CACHE'] = '1'
        try:
            _cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.cache')
            if os.path.isdir(_cache_dir):
                _n = 0
                for _fn in os.listdir(_cache_dir):
                    if _fn.endswith('.pkl'):
                        try:
                            os.remove(os.path.join(_cache_dir, _fn)); _n += 1
                        except Exception:
                            pass
                if _n:
                    print(f"[缓存清理] 已清除 {_n} 个磁盘缓存 + AUDIT_NO_CACHE=1（全量重跑强制失效）")
                else:
                    print("[缓存清理] 无残留缓存；AUDIT_NO_CACHE=1 已设")
        except Exception:
            pass
    skip_finalize = '--skip-finalize' in argv
    if skip_finalize:
        argv = [a for a in argv if a != '--skip-finalize']
        # 开发模式：整条终审链（公式浅色/合计重算/检查器）全部跳过，只刷新底稿数据
        os.environ['SKIP_FINALIZE'] = '1'
        print("[开发模式] --skip-finalize：跳过终审（公式浅色/合计重算/检查器），仅刷新底稿数据")
    # --phase 1,3 分段跑（2026-08-06：只跑指定阶段，1-based）
    phase_sel = None
    if '--phase' in argv:
        _i = argv.index('--phase')
        if _i + 1 < len(argv):
            phase_sel = {int(x) for x in argv[_i + 1].split(',') if x.strip().isdigit()}
            argv = argv[:_i] + argv[_i + 2:]
        else:
            argv.remove('--phase')
    dirs = argv
    if not dirs:
        print("用法：python regen_all.py <文件夹> [<另一文件夹> ...] [--skip-finalize] [--phase 1,3]")
        print("  例：python regen_all.py FY            # 全流程（①②③+已实现阶段）")
        print("      python regen_all.py FY --phase 1,3   # 只跑 自建试算表 + 同源核对")
        print("      python regen_all.py g J          # 刷新 g、J 两文件夹的全部 13 程序")
        return
    for d in dirs:
        data_dir = d if os.path.isabs(d) else os.path.join(os.path.expanduser("~"), "Desktop", d)
        if not os.path.isdir(data_dir):
            print("\n[跳过] 文件夹不存在：%s" % data_dir)
            continue
        print("\n" + "=" * 70)
        print("文件夹 %s (%s)" % (d, data_dir))
        print("=" * 70)
        # ⚡⚡ 2026-08-24 架构预检：科目代码归属冲突检测（防『修了又犯』——配置盲区提前暴露）
        try:
            import code_conflict_check as _ccc
            _owners = _ccc.scan_all_generators()
            _cf = _ccc.find_conflicts(_owners)
            if _cf:
                print('  ⚠️ [科目代码冲突检测] %d 处代码被多科目组配置（生成时按名称互斥裁决，需确认归属）：' % len(_cf))
                for _cd, _own in _cf:
                    print('      %s: ' % _cd + ' 与  '.join(_own))
        except Exception as _ex:
            print(f'  ⚠️ 科目代码冲突检测执行异常: {_ex}')
        # —— 阶段① 自建试算表（摸结构/勾稽）——
        if phase_sel is None or 1 in phase_sel:
            print("\n----- [阶段①] 自建试算表 -----")
            try:
                import self_tb_gen
                self_tb_gen.main_stub(data_dir) if hasattr(self_tb_gen, 'main_stub') else _run_stage_module('self_tb_gen', data_dir)
            except Exception as _ex:
                print(f"  ⚠️ 自建试算表失败：{_ex}")
        # —— 阶段② 13 builder 底稿 ——
        if phase_sel is None or 2 in phase_sel:
            # ⚡ 阶段三 3.2 摄取层 TB 控制数 Fail Fast（治本：数据问题在摄取时拦截，
            #   不静默生成坏底稿）。FAIL=硬性问题（提示并继续，审计员确认数据后重跑）；
            #   --tb-fail-fast 时 FAIL 直接中止该账套。
            try:
                import tb_validate
                _tb_ok, _tb_issues = tb_validate.validate_tb(data_dir, fail_fast='--tb-fail-fast' in argv)
            except tb_validate.TBValidationError as _ex:
                print(f'  ⛔ [阶段三 3.2] {_ex}')
                print(f'  [跳过] {d}：TB 控制数不平衡（Fail Fast），修复数据后重跑')
                continue
            except Exception as _ex:
                print(f'  ⚠️ TB 校验异常（继续）：{_ex}')
            n_ok = _run_builders(data_dir, skip_finalize)
            if not skip_finalize:
                print("\n----- [终审] audit_finalize (注入审定表 + 补全对方科目) -----")
                _finalize_once(data_dir)
            else:
                print("\n[开发模式] 跳过终审（SKIP_FINALIZE=1）—— 产出未跑 recalc/checker/公式浅色，仅供开发核对")
            # ⚡⚡ 2026-08-14 要求层挂接：科目层跑完后执行要求知识点（清单驱动/自动发现）
            _run_requirements(data_dir)
            # 收尾：审计结果中间层 + 产物完整性自检（随阶段②跑）
            _post_build(data_dir)
            print("[完成] 文件夹 %s：%d/%d builder 成功" % (d, n_ok, len(BUILDERS)))
        # —— 阶段③ 底稿 vs 自建试算表核对 ——
        if phase_sel is None or 3 in phase_sel:
            print("\n----- [阶段③] 底稿 vs 自建试算表核对 -----")
            try:
                import recon_self
                recon_self.recon_self(data_dir)
            except Exception as _ex:
                print(f"  ⚠️ 同源核对失败：{_ex}")
            # —— 阶段③.5 合并工作底稿（集团模式多主体自动分表，2.3）——
            print("\n----- [阶段③.5] 合并工作底稿（集团模式自动分表） -----")
            try:
                import audit_merge
                audit_merge.build_merge_workbook(data_dir)
            except Exception as _ex:
                print(f"  ⚠️ 合并工作底稿失败：{_ex}")
        # —— 阶段④⑤⑥ 后续阶段（待接入）——
        # ⚡ 2026-08-24 修复：phase_sel 是数字集合（--phase 1,4），而 PHASES[3:] 的
        #   key 是字符串（reclass/report/recon_ext）——原按 key 匹配永远 continue，
        #   ④⑤⑥ 之前全 False 未暴露。改为按【流水线序号】匹配（序号 = PHASES 下标+1）。
        for _idx, (_pid, _mod, _label, _impl) in enumerate(PHASES[3:], 4):
            if phase_sel is not None and (_idx not in phase_sel):
                continue
            if not _impl:
                print(f"\n----- [{_label}] 未接入（{_mod} 待实现，跳过） -----")
                continue
            try:
                __import__(_mod)
                _run_stage_module(_mod, data_dir)
            except ImportError:
                print(f"\n----- [{_label}] {_mod} 未实现，跳过 -----")
            except Exception as _ex:
                print(f"  ⚠️ [{_label}] 执行异常：{_ex}")


def _run_builders(data_dir, skip_finalize):
    """阶段②：依次跑 13 个 builder + 每 builder 后规范检查。返回成功数。"""
    n_ok = 0
    for mod, label in BUILDERS:
        print("\n----- [%s] %s -----" % (label, mod))
        ok, msg = _run_module(mod, label, data_dir)
        print("  -> %s %s" % ("✓" if ok else "✗", msg))
        n_ok += 1 if ok else 0
        # 2026-08-04：每跑完一个 builder 立即做规范检查（不等全部跑完）——
        # 结构不规范当场暴露在对应 builder 的日志里，而非最后才发现。
        if ok and not skip_finalize:
            try:
                import audit_checker
                _te, _tw = audit_checker.process_folder(data_dir, quiet=True)
                if _te or _tw:
                    print("  ⚠️ [规范检查] 该 builder 产物有 %d ERROR / %d WARN —— 请查看上方明细" % (_te, _tw))
                else:
                    print("  ✅ [规范检查] 该 builder 产物结构规范（0 ERROR / 0 WARN）")
            except Exception as _ex:
                print("  ⚠️ [规范检查] 执行异常：%s" % _ex)
    return n_ok


def _run_stage_module(module_name, data_dir):
    """阶段模块统一入口：模块内须有 main() 或 build_all(data_dir) 或 recon_self(data_dir)。"""
    saved_argv = sys.argv
    sys.argv = [module_name + ".py", data_dir]
    try:
        if module_name in sys.modules:
            del sys.modules[module_name]
        mod = __import__(module_name)
        if hasattr(mod, "build_all"):
            mod.build_all(data_dir)
        elif hasattr(mod, "recon_self"):
            mod.recon_self(data_dir)
        elif hasattr(mod, "main"):
            # 2026-08-06 修复：main 签名不一——self_tb_gen.main(argv) 是必选参数，
            # 其余 builder 为 main() 或 main(argv=None)（argv=None 时内部自读
            # sys.argv，而 sys.argv 已被本函数注入 [module, data_dir]，必须无参调用）。
            # 判定标准：存在无默认值的必选参数才传参（传 sys.argv[1:]，
            # 与 main(argv=None) 自读 sys.argv[1:] 行为一致）。
            import inspect
            _need_argv = any(p.default is inspect.Parameter.empty
                             for p in inspect.signature(mod.main).parameters.values())
            if _need_argv:
                mod.main(sys.argv[1:])
            else:
                mod.main()
        else:
            print(f"  ⚠️ {module_name} 无 build_all/recon_self/main 入口")
    except SystemExit:
        pass
    except Exception as ex:
        traceback.print_exc()
        print(f"  ⚠️ {module_name} 异常：{ex}")
    finally:
        sys.argv = saved_argv


def _run_requirements(data_dir):
    """⚡⚡ 2026-08-14 要求层挂接：科目层跑完后，自动执行该账套的要求层知识点。

    清单发现（约定）：
      ① 显式：账套目录下 manifest_*.json（如 manifest_3300_combined.json）——
         data_dir 字段匹配本账套则执行；
      ② 隐式：无清单时，对已注册要求知识点中"能自动定参"的（无 params 依赖），
         直接按 data_dir 执行（如 cutoff_test/gross_margin/confirmation_list）。
    任何知识点失败不影响其余（引擎内部已隔离）。"""
    try:
        import requirement_engine as RE
        import glob
        # ① 显式清单：data_dir 匹配本账套
        cands = [f for f in glob.glob(os.path.join(data_dir, 'manifest_*.json'))
                 + glob.glob(os.path.join(HERE, 'manifest_*.json'))]
        hit = None
        for f in cands:
            try:
                import json
                man = json.load(open(f, encoding='utf-8'))
                # ⚡ 2026-08-14 manifest 路径字段支持相对 DATA_ROOT（<DATA_ROOT> 占位亦可）
                _md = str(P.resolve_rel(str(man.get('data_dir', '')) or '')).lower()
                if _md and (os.path.normpath(_md) == os.path.normpath(data_dir).lower()
                            or os.path.basename(os.path.normpath(_md)) == os.path.basename(os.path.normpath(data_dir)).lower()):
                    hit = f
                    break
            except Exception:
                continue
        if hit:
            print(f"[要求层] 发现清单 {os.path.basename(hit)} → 按清单执行知识点")
            RE.run_manifest(hit)
            return
        # ② 隐式：无 params 依赖的知识点直接跑
        _auto = []
        for r in RE.requirements.list_requirements():
            if not r['params']:
                _auto.append(r['name'])
        if _auto:
            print(f"[要求层] 无匹配清单，自动执行无参知识点：{', '.join(_auto)}")
            man = {'data_dir': data_dir, 'requirements': [{'name': n} for n in _auto]}
            RE.run_manifest_from(man)
        else:
            print('[要求层] 无匹配清单且无无参知识点，跳过（可放 manifest 到账套目录启用）')
    except Exception as _ex:
        print(f"  ⚠️ 要求层挂接异常（不影响科目层结果）：{_ex}")


def _post_build(data_dir):
    """阶段②收尾：审计结果中间层 + 产物完整性自检。"""
    # 2026-08-05 P0 挂接：审计结果中间层（全科目审定数 JSON + 汇总表）
    try:
        import audit_result_export
        _n = len(audit_result_export.export_folder(data_dir, quiet=True))
        _x = audit_result_export.export_summary_xlsx(data_dir)
        print(f"[审计结果中间层] 导出 {_n} 个科目 JSON -> {os.path.join(data_dir, '审计结果')}")
        if _x:
            print(f"  ✓ 全科目审定数汇总表：{_x}")
    except Exception as _ex:
        print(f"  ⚠️ 审计结果中间层导出失败：{_ex}")
    # 2026-08-06 P0 挂接：产物完整性自检
    try:
        import audit_result_export as _are
        _are.run_coverage(data_dir)
    except Exception as _ex:
        print(f"  ⚠️ 完整性自检执行异常：{_ex}")
    # 2026-08-07 P0 挂接：未覆盖科目清单（TB 有数据但注册表/builder 未覆盖的一级科目）
    try:
        import uncovered_subjects as _us
        _rows = _us.uncovered_subjects(data_dir)
        if _rows:
            _out = _us.write_uncovered_xlsx(data_dir, _rows)
            print(f"  ⚠️ [未覆盖科目清单] 共 {len(_rows)} 个一级科目未生成底稿（详见 {_out}）")
        else:
            print('  ✅ [未覆盖科目清单] 全部有数据一级科目均已覆盖')
    except Exception as _ex:
        print(f"  ⚠️ 未覆盖科目清单执行异常：{_ex}")
    # ⚡⚡ 2026-08-24 架构自检：空白底稿检测（TB 有数据但整份空=取数bug，自动接入）
    try:
        import self_check
        _ok, _lines = self_check.self_check(data_dir, quiet=True)
        _bug = [l for l in _lines if '❌' in l]
        if _bug:
            print(f"  ⛔ [空白底稿自检] {len(_bug)} 处空白底稿（TB 有数据，需人工处理）：")
            for l in _bug[:8]:
                print('      ' + l.strip())
        else:
            print('  ✅ [空白底稿自检] 无空白底稿')
    except Exception as _ex:
        print(f"  ⚠️ 空白底稿自检执行异常：{_ex}")
    # ⚡⚡ 2026-08-24 架构自检：先查 TB 自身勾稽（qc+jf-df=qm 带符号），再查审定表 期初+增-减=期末
    try:
        import wp_recon_check
        _n_tb = wp_recon_check.check_tb_self(data_dir, open(os.devnull, 'w'))
        if _n_tb:
            print(f"  ⛔ [TB自检] {_n_tb} 行 TB 不平（qc+jf-df≠qm，源科目余额表需客户核实）")
        else:
            print('  ✅ [TB自检] TB 自身勾稽全部平衡')
        _cnt = wp_recon_check.check_folder(data_dir, open(os.devnull, 'w'))
        if _cnt:
            print(f"  ⛔ [勾稽核对] {_cnt} 处审定表勾稽不平（期初+增-减≠期末，需查源TB或生成器）")
        else:
            print('  ✅ [勾稽核对] 审定表勾稽全部平衡')
    except Exception as _ex:
        print(f"  ⚠️ 勾稽核对执行异常：{_ex}")


if __name__ == "__main__":
    main()
