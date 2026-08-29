# -*- coding: utf-8 -*-
"""recalc_workbook.py —— 2026-08-07 新增（任务②底稿重算器）。

背景：程序产物全部数值化（无公式）→ 审计人员在 Excel 里把合计/小计改回公式后保存，
Excel 自动写入计算缓存值 v → 程序 data_only 可读回 ✓。但【程序用 openpyxl 二次保存】
会把 v 缓存清掉（只写公式 f 不写 v）→ 收回读不回。

本脚本：对审计返回/程序二次保存过的底稿，用【本机 Excel COM】打开 → 全表重算 →
保存（Excel 保存自动写回缓存值）→ 关闭。跑完后 data_only=True 即可读回所有公式结果。

用法：
  python recalc_workbook.py <文件夹>                # 扫描该文件夹 *_生成.xlsx
  python recalc_workbook.py <文件1> [文件2 ...]     # 指定文件
  python recalc_workbook.py <文件夹> --all          # 含非 _生成 后缀的 xlsx
  python recalc_workbook.py <文件夹> --dry          # 只列出含公式无缓存的文件，不重算

依赖：本机安装 Microsoft Excel（Office16 已验证）；pip install pywin32。
无 Excel 时打印提示并跳过（程序 _resolve_value 已能解析 SUM/引用 兜底）。
"""
import glob
import os
import sys


def _find_xlsx(paths, all_xlsx=False):
    files = []
    for p in paths:
        if os.path.isfile(p):
            files.append(p)
        elif os.path.isdir(p):
            pat = '*' if all_xlsx else '*_生成.xlsx'
            files += sorted(glob.glob(os.path.join(p, pat)))
    # 去重保序
    return list(dict.fromkeys(files))


def _has_formula_no_cache(fp):
    """该文件是否有【公式但无缓存值】的单元格（需要重算的对象）。
    data_only=False 有公式 / data_only=True 读 None。"""
    try:
        import openpyxl
        wbf = openpyxl.load_workbook(fp, read_only=True, data_only=False)
        wbv = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        n = 0
        for ws_f, ws_v in zip(wbf.worksheets, wbv.worksheets):
            for r_f, r_v in zip(ws_f.iter_rows(), ws_v.iter_rows()):
                for c_f, c_v in zip(r_f, r_v):
                    if isinstance(c_f.value, str) and c_f.value.startswith('='):
                        if c_v.value is None:
                            n += 1
                        if n > 0:
                            break
                if n > 0:
                    break
            if n > 0:
                break
        wbf.close(); wbv.close()
        return n > 0
    except Exception:
        return False


def recalc(files, dry=False):
    """用 Excel COM 重算并保存（写回公式缓存值）。返回 (成功, 失败, 跳过)。"""
    if not files:
        print('无匹配文件')
        return 0, 0, 0
    need = [f for f in files if _has_formula_no_cache(f)] if not dry else files
    if not need and not dry:
        print(f'扫描 {len(files)} 个文件：全部为数值（无公式无缓存问题），无需重算')
        return 0, 0, len(files)
    if dry:
        print(f'需要重算（含公式无缓存）{len(need)} 个：')
        for f in need:
            print('  ·', os.path.basename(f))
        return len(need), 0, len(files) - len(need)
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print('⚠️ 未安装 pywin32（pip install pywin32），跳过 Excel COM 重算。'
              '程序 _resolve_value 可兜底解析 SUM/单元格引用。')
        return 0, 0, len(files)
    app = None
    ok = fail = 0
    try:
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        for fp in need:
            try:
                wb = app.Workbooks.Open(os.path.abspath(fp), UpdateLinks=0,
                                        ReadOnly=False, CorruptLoad=1)
                try:
                    wb.Application.CalculateFull()
                except Exception:
                    wb.Application.Calculate()
                wb.Save()
                wb.Close(False)
                ok += 1
                print(f'  ✓ 重算：{os.path.basename(fp)}')
            except Exception as ex:
                fail += 1
                print(f'  ✗ 失败：{os.path.basename(fp)}：{ex}')
                try:
                    app.Workbooks.Close()
                except Exception:
                    pass
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            app = None
        pythoncom.CoUninitialize()
    print(f'重算完成：成功 {ok} / 失败 {fail} / 无需 {len(files) - len(need)}')
    return ok, fail, len(files) - len(need)


def main():
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        return
    dry = '--dry' in argv
    all_x = '--all' in argv
    paths = [a for a in argv if not a.startswith('--')]
    files = _find_xlsx(paths, all_x)
    recalc(files, dry=dry)


if __name__ == '__main__':
    main()
