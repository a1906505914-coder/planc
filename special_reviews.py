# -*- coding: utf-8 -*-
"""专项底稿统一框架（2026-08-16 用户定：一个框架 + 多个专项插件）。

架构（铁律131e）：
  · 一个框架（本文件）：统一 CLI / 注册表 / 三层路由 / 公共基建
  · 多个专项插件：每个专项 = 生成器模块 + manifest + 拆分整合逻辑，
    在 REGISTRY 注册一行即可接入三层路由
  · 三层路由（铁律131c）：
    ① 集团底稿（覆盖全集团，统一复核）
    ② 单体拆分（按主体）
    ③ 整合进对应科目审计底稿（如 银行核对-* → 银行存款审计底稿）

专项插件约定接口（见 REGISTRY 字段）：
  module    生成器模块（须有 build/main 或 run_one）
  manifest  数据源清单模块（须有 jobs()/resolve()）
  kind      输出文件名模板；subject 整合目标科目；split 拆分方式
  gen_call  生成调用（统一包装，因各专项入口签名不同）

用法：
  python special_reviews.py                          # 全部专项全部账套
  python special_reviews.py --review=bank            # 仅银行核对
  python special_reviews.py --review=dep,lease       # 折旧+租赁
  python special_reviews.py --acct=ADF --review=bank # 指定账套
  python special_reviews.py --split                  # 仅做拆分整合（不重新生成）
  python special_reviews.py --list                   # 列出已注册专项
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import argparse
import importlib
import os
import sys

ROOT = paths.DATA_ROOT

# ⚡ 专项注册表：未来新增专项只需在此加一行（模块 + manifest + 拆分路由）
REGISTRY = {
    'bank': {
        'module': 'bank_review',      # ⚡ 2026-08-18 合并：双向核对+开户核对+异常分析 → 统一 bank_review（bank_manifest 配置适配）
        'manifest': 'bank_reconcile_manifest',
        'name': '银行专项（双向核对/开户核对/异常分析）',
        'subject': '银行存款',            # 整合目标科目（科目底稿名关键词）
        'split': 'bank',                  # 拆分路由：bank=按主体整合+集团汇总
        'year_attr': 'year',              # manifest 中年份字段
        'accts': ['ADF', 'GJX', 'AYL'],   # 适用账套
    },
    'dep': {
        'module': 'fa_review',        # ⚡ 2026-08-18 合并：XBJ/GJX/ADF 三个固定资产复核 → 统一 fa_review（fa_manifest 配置适配）
        'manifest': 'depreciation_manifest',
        'name': '固定资产复核',
        'subject': '固定资产',
        'split': 'dep',
        'year_attr': 'year',
        'accts': ['XBJ', 'GJX', 'ADF'],
    },
    'lease': {
        'module': 'lease_review',
        'manifest': 'lease_manifest',
        'name': '使用权资产/租赁负债复核',
        'subject': '使用权资产',
        'split': 'lease',
        'year_attr': 'year',
        'accts': ['AJ'],
    },
    'aging': {
        'module': 'aging_review',
        'manifest': 'aging_manifest',
        'name': '应收+合同资产联动账龄',
        'subject': '应收账款',
        'split': 'aging',
        'year_attr': 'year',
        'accts': ['XBJ', 'AJ'],   # AH 待迁移后接入（SAP 专用读取器）
    },
    'project': {
        'module': 'project_review',
        'manifest': 'project_manifest',
        'name': '项目制核算收入底稿',
        'subject': '营业收入',
        'split': None,            # 暂不做拆分整合（框架先行）
        'year_attr': 'year',
        'accts': ['XBJ'],         # AH SAP 迁移后接入（WBS/合同号项目键）
    },
    'ah_intra': {
        'module': 'ah_intra_group_recon',
        'manifest': 'ah_intra_group_recon',
        'name': 'AH88家关联往来核对',
        'subject': '往来',
        'split': None,            # 独立核对表，不整合进科目底稿
        'year_attr': 'year',
        'accts': ['AH'],          # ⚠️ 账套专属（数据源=AH/prepared 三集团底稿）
    },
    'fx': {
        'module': 'fx_review',    # ⚡ 2026-08-18 提炼：fx_consolidate/fx_rates/thai_mapping → 境外报表折算专项
        'manifest': 'fx_review',
        'name': '境外报表折算（外币折算合并）',
        'subject': None,          # 折算底稿，无整合目标科目
        'split': None,
        'year_attr': 'year',
        'accts': ['AZ'],
    },
    'financing': {
        'module': 'financing_review',   # ⚡ 2026-08-22 统一融资专项：借款+表外+融资全景
        'manifest': 'financing_ledger',
        'name': '统一融资专项（借款/表外/融资全景）',
        'subject': '短期借款',
        'split': None,            # 独立专项，不整合进科目底稿（loan_detail 已按科目出底稿）
        'year_attr': 'year',
        'accts': ['ADF'],
    },
}


def _mod(name):
    return importlib.import_module(name)


def ASE_PREP():
    """AH/prepared 路径（关联往来专项输出目录）。"""
    return os.path.join(paths.DATA_DIRS['AH'], '中间产物', 'prepared')


def list_reviews():
    print('已注册专项底稿：')
    for key, cfg in REGISTRY.items():
        print(f'  {key:8s} {cfg["name"]:16s} 整合→{cfg["subject"]} 账套={cfg["accts"]}')


# ---------------------------------------------------------------- 生成
def gen_one(key, acct):
    """调用专项生成器（统一包装，兼容不同入口签名）。返回产物文件路径列表。"""
    cfg = REGISTRY[key]
    m = _mod(cfg['module'])
    mm = _mod(cfg['manifest'])
    outs = []
    if key == 'ah_intra':
        # ⚡ 2026-08-18 关联往来专项三步：①抽序时账 ②88家核对 ③差异穿透底稿（每对差异→具体分录）
        #    各自独立进程执行（同进程串行 66k 行序时账 + 24.7k 行穿透会 OOM）
        import subprocess
        _d = os.path.dirname(os.path.abspath(__file__))
        for s in ['ah_intra_seq_extract.py', 'ah_intra_group_recon.py', 'ah_intra_diff_trace.py']:
            rc = subprocess.run([sys.executable, os.path.join(_d, s)])
            if rc.returncode != 0:
                print(f'  ⚠️ {s} 退出码 {rc.returncode}')
        outs.append(os.path.join(ASE_PREP(), '关联往来序时账_2026.xlsx'))
        outs.append(os.path.join(ASE_PREP(), '88家内部往来核对_2026.xlsx'))
        outs.append(os.path.join(ASE_PREP(), '关联往来差异穿透底稿_2026.xlsx'))
    elif key == 'bank':
        # ⚡ 2026-08-18 统一专项：bank_review 按 bank_manifest.forms 路由
        #    reconcile(双向核对)/acct(开户清单核对)/abnormal(发生额异常)
        import bank_review as BR
        outs = BR.gen(acct)
        if isinstance(outs, str):
            outs = [outs]
    elif key == 'dep':
        # ⚡ 2026-08-18 统一专项：fa_review 按 fa_manifest 路由 XBJ/GJX/ADF（原三个脚本逻辑复用，结果一致）
        import fa_review as FR
        outs = FR.gen(acct)
        if isinstance(outs, str):
            outs = [outs]
        outs = [o for o in outs if o]
    elif key == 'lease':
        # ⚡ 2026-08-18：专项底稿交付形态=集团合并文件。先生成分主体（输出中间产物），
        #   再由 lease_group_merge 合并为集团底稿（未来拆分从集团底稿开始，不再保留分主体交付）。
        acct_cfg, subs = mm.jobs()[acct], list(mm.jobs()[acct]['subs'])
        for sub in subs:
            m.run_one(acct, sub)
        import lease_group_merge as LGM
        _year = mm.jobs()[acct]['subs'][subs[0]].get('year', 2025)
        rc = LGM.main(acct, _year)
        if rc == 0:
            outs.append(os.path.join(ROOT, acct, '底稿', str(_year),
                                     f'使用权资产租赁复核_{_year}.xlsx'))
    elif key == 'fx':
        # ⚡ 2026-08-18 境外报表折算专项：fx_review 按 fx_manifest 路由
        import fx_review as FX
        rc = FX.gen(acct)
        if isinstance(rc, int):       # process 返回码（0/1）
            outs = []
        elif isinstance(rc, str):
            outs = [rc]
        else:
            outs = rc
    elif key == 'financing':
        # ⚡ 2026-08-22 统一融资专项：借款底稿 + 表外事项 + 融资全景（ADF）
        import financing_review as FR
        outs = FR.gen(acct)
        if isinstance(outs, str):
            outs = [outs]
    elif key == 'project':
        # project_review：run_one(acct, sub)，按 manifest 账套×子集团
        for sub in mm.jobs()[acct]['subs']:
            rc = m.run_one(acct, sub)
            if rc == 0:
                year = mm.jobs()[acct]['subs'][sub].get('year', 2025)
                sub_dir = sub if sub != '_root' else ''
                fp = os.path.join(ROOT, acct, '底稿', str(year))
                if sub_dir:
                    fp = os.path.join(fp, sub_dir)
                outs.append(os.path.join(fp, f'项目制收入底稿_{year}.xlsx'))
    return outs


# ---------------------------------------------------------------- 拆分整合
def split_one(key, acct):
    """按专项拆分路由整合（复用 _reconcile_split 能力）。

    ⚡ 2026-08-19 修复：_reconcile_split.py 在整理文件夹时被误删且无备份——
    加缺失降级，框架主流程不崩；整合功能待按原语义恢复（bank_integrate/
    dep_integrate 把复核表数值整合进科目审计底稿）。"""
    try:
        import _reconcile_split as RS
    except ModuleNotFoundError:
        print(f'  ⚠️ {acct}: 拆分整合模块 _reconcile_split 缺失（整理时误删，待恢复），跳过整合')
        return 0
    cfg = REGISTRY[key]
    mm = _mod(cfg['manifest'])
    if key == 'bank':
        n = RS.bank_integrate(acct)
        g = RS.bank_group_summary(acct)
        print(f'  {acct}: 银行核对 单体整合 {n} 主体, 集团底稿 {g}')
        return n
    if key == 'dep':
        if acct != 'XBJ':
            print(f'  ⏭️ {acct}: 折旧复核暂仅 XBJ 有配置')
            return 0
        dep_fp = os.path.join(ROOT, acct, '底稿', '2025', '固定资产折旧计提复核_2025.xlsx')
        subj_fp = os.path.join(ROOT, acct, '底稿', '2025', '固定资产审计底稿_2025_生成.xlsx')
        return RS.dep_integrate(acct, '2025', dep_fp, subj_fp)
    if key == 'lease':
        # 租赁复核 → 使用权资产审计底稿（业务在 ga 子集团 03 公司）
        subs = mm.jobs()[acct]['subs']
        year = subs[sorted(subs)[0]].get('year', 2025)
        lease_fp = os.path.join(ROOT, acct, '底稿', str(year), f'使用权资产租赁复核_{year}.xlsx')
        # 只在有租赁业务的子集团找科目底稿（ga 有 03 公司，jj 无）
        import glob
        subj_cands = []
        for sub in subs:
            for fp in glob.glob(os.path.join(ROOT, acct, '底稿', str(year), sub,
                                             '*使用权资产*审计底稿*.xlsx')):
                subj_cands.append(fp)
        if not subj_cands:
            print(f'  ⚠️ {acct}: 未找到 使用权资产审计底稿')
            return 0
    if key == 'aging':
        # 应收+合同资产账龄 → 应收账款审计底稿（整合 复核 sheets）
        year = mm.jobs()[acct]['subs'][sorted(mm.jobs()[acct]['subs'])[0]].get('year', 2025)
        import glob
        # 账龄表：可能在 底稿/{year}/ 根（_root）或 底稿/{year}/{sub}/（子集团）
        aging_cands = glob.glob(os.path.join(ROOT, acct, '底稿', str(year), '**',
                                             '应收合同资产账龄复核_*.xlsx'), recursive=True)
        subj_cands = glob.glob(os.path.join(ROOT, acct, '底稿', str(year), '**',
                                            '*应收账款*审计底稿*.xlsx'), recursive=True)
        if not aging_cands or not subj_cands:
            print(f'  ⚠️ {acct}: 账龄复核表或应收账款审计底稿缺失')
            return 0
        total = 0
        # 按目录配对：账龄表目录 → 同目录应收账款底稿
        for aging_fp in aging_cands:
            adir = os.path.dirname(aging_fp)
            for subj_fp in subj_cands:
                if os.path.dirname(subj_fp) == adir:
                    total += RS.dep_integrate(acct, str(year), aging_fp, subj_fp,
                                              prefix='账龄复核-', label='账龄复核')
        return total
    return 0


# ---------------------------------------------------------------- 主入口
def main():
    ap = argparse.ArgumentParser(description='专项底稿统一框架')
    ap.add_argument('--review', default=None, help='专项（逗号分隔，默认全部）')
    ap.add_argument('--acct', default=None, help='账套（默认全部）')
    ap.add_argument('--split', action='store_true', help='仅拆分整合（不重新生成）')
    ap.add_argument('--list', action='store_true', help='列出已注册专项')
    a = ap.parse_args()

    if a.list:
        list_reviews()
        return 0

    keys = [k.strip() for k in a.review.split(',')] if a.review else list(REGISTRY)
    rc = 0
    for key in keys:
        if key not in REGISTRY:
            print(f'  ⚠️ 未注册专项: {key}（--list 查看）')
            rc = 1
            continue
        cfg = REGISTRY[key]
        accts = [a.acct] if a.acct else cfg['accts']
        print(f'\n===== 专项: {cfg["name"]} =====')
        for acct in accts:
            if not a.split:
                outs = gen_one(key, acct)
                print(f'  {acct}: 生成 {len(outs)} 个文件')
            split_one(key, acct)
    print('\n完成')
    return rc


if __name__ == '__main__':
    sys.exit(main())
