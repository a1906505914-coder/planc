# -*- coding: utf-8 -*-
"""bank_reconcile_manifest.py —— 银行双向核对数据源清单（2026-08-14 取消硬编码）。

⚡⚡⚡ 设计目标：顶级目录改名/移动时【只改 paths.py 一处】，本文件与所有脚本零改动。
  - 账套名：唯一事实来源 = paths.py ACCT_NAMES（DATA_ROOT + 账套名 = 账套根）
  - 本文件：只声明【账套内子结构】，全部相对路径（相对账套根）
  - 脚本：统一 `import bank_reconcile_manifest as M; M.jobs()` 生成任务，不写死任何路径

结构：
  ACCTS = {账套名: {
      year: 核对年度,
      subjects: {主体代码: {
          dir:     底稿输出目录（相对账套根；缺省=账套根）
          stmt:    网银流水文件夹（相对账套根）
          journal: 日记账文件（相对账套根；None=从 GL 抽取/adapter）
      }}
  }}
"""
import os
import paths as P

ACCTS = {
    'ADF': {
        'year': '2026',
        # ⚡ 2026-08-15 用户定：ADF 双向核对底稿统一输出到账套根（不再分主体子目录）。
        #   子目录（3100/3200/...）仍为源数据目录（科目余额表/序时账等），不输出底稿。
        'subjects': {
            '3000': {'stmt': '数据/2026/网银/3000网银明细2026年1-7月',
                     'journal': '3000网银明细2026年1-7月/3000 银行日记账.XLSX'},
            '3100': {'stmt': '数据/2026/网银/3100网银明细2026年1-7月',
                     'journal': '3100网银明细2026年1-7月/3100-3400银行日记账.XLSX'},
            '3200': {'stmt': '数据/2026/网银/3200网银明细2026年1-7月',
                     'journal': '3100网银明细2026年1-7月/3100-3400银行日记账.XLSX'},
            '3300': {'stmt': '数据/2026/网银/3300网银明细2026年1-7月',
                     'journal': '3100网银明细2026年1-7月/3100-3400银行日记账.XLSX'},
            '3400': {'stmt': '数据/2026/网银/3400网银明细2026年1-7月',
                     'journal': '3100网银明细2026年1-7月/3100-3400银行日记账.XLSX'},
            # ⚡ 3500 大主体（392MB）：单独进程跑（_run_3500.py），不在此批量（jobs 默认排除）
            '3500': {'stmt': '数据/2026/网银/3500网银明细2026年1-7月',
                     'journal': '3500网银明细2026年1-7月/3500-3900.XLSX'},
            '3700': {'stmt': '数据/2026/网银/3700网银明细2026年1-7月',
                     'journal': '3500网银明细2026年1-7月/3500-3900.XLSX'},
            '3800': {'stmt': '数据/2026/网银/3800网银明细2026年1-7月',
                     'journal': '3500网银明细2026年1-7月/3500-3900.XLSX'},
            '3900': {'stmt': '数据/2026/网银/3900网银明细2026年1-7月',
                     'journal': '3500网银明细2026年1-7月/3500-3900.XLSX'},
            '6100': {'stmt': '数据/2026/网银/6100网银明细2026年1-7月', 'journal': None},
            '6900': {'stmt': '数据/2026/网银/6900网银明细2026年1-7月', 'journal': None},
        },
    },
    'GJX': {
        'year': '2025',
        'subjects': {
            'DQ': {'stmt': '数据/2025/网银/DQ网银流水2025', 'journal': None},
        },
    },
}


def _root(acct):
    """账套根（DATA_ROOT + 账套名——账套名唯一事实来源=paths.py）。"""
    return os.path.join(P.DATA_ROOT, acct)


# ⚡ 大主体单独跑（不在批量 jobs，job_for 可查）
_BIG_ONLY = {'3500'}


def jobs(acct=None, include_big=False):
    """生成核对任务 [(comp, data_dir, stmt_folder, journal_file, year)]（绝对路径）。
    include_big=False（默认）排除大主体（3500）——大主体单独进程跑。"""
    out = []
    for a, conf in ACCTS.items():
        if acct and a != acct:
            continue
        root = _root(a)
        for comp, s in conf['subjects'].items():
            if not include_big and comp in _BIG_ONLY:
                continue
            # ⚡ 2026-08-18 修正：默认 data_dir=账套根 discover 为空（ADF/GJX 均无法识别结构）；
            #    默认 数据/{year} 层 discover 正确识别主体（ADF→9主体、GJX→DQ）
            data_dir = os.path.join(root, s['dir']) if s.get('dir') else os.path.join(root, '数据', conf['year'])
            stmt = os.path.join(root, s['stmt'])
            jf = os.path.join(root, s['journal']) if s.get('journal') else None
            out.append((comp, data_dir, stmt, jf, conf['year']))
    return out


def job_for(comp):
    """按主体代码查任务（含 3500 等大主体）。"""
    for j in jobs(include_big=True):
        if j[0] == comp:
            return j
    return None


def resolve(acct, rel):
    """相对路径 → 绝对路径（rel 为 None 返回 None）。"""
    return os.path.join(_root(acct), rel) if rel else None


if __name__ == '__main__':
    for comp, dd, stmt, jf, y in jobs():
        print(f'{comp:<6} year={y}  out={dd}')
        print(f'       stmt={stmt}')
        print(f'       jf  ={jf}')
    j = job_for('3500')
    print('3500:', j)
