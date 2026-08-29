# -*- coding: utf-8 -*-
"""recon_registry.py —— 外部数据核对注册表（2026-08-24 架构预留 R1）。

设计背景（见 docs/audit_pipeline_design.md §六）：
    账套→底稿只是审计流程的前段（约 30%），剩余工作量靠外部数据获取→分析→
    核对→调整→报告。41 个外部数据/核对程序目前独立 __main__ 运行、无总控。
    本模块把零散核对程序【登记】为统一契约，供 recon_ext 总控调度。

安全边界：
    - 核对只【发现差异】，不自动改写底稿审定数（审计判断留给人工）
    - 差异统一写为 JSON 契约（§6.1），供后续人工确认状态流转

用法：
    from recon_registry import REGISTRY, list_recons
    list_recons()                          # 列出已登记核对
    entry = REGISTRY['bank']               # 取某专项配置
    recon = load_recon('bank')             # 动态导入专项模块

契约字段（每专项一行）：
    module    差异产出模块（含 recon_entry() 或 main()）
    manifest  配置清单模块（读取该专项的账套/年份/文件配置）
    name      专项名称
    subject   整合目标科目（科目底稿名关键词）
    accts     适用账套列表
    diff_json 差异 JSON 输出相对路径（data_dir 下）
    entry_fn  统一入口函数名（默认 recon_entry；缺失则回退 main）
"""
import importlib
import os
import sys


def _acct_from_dir(data_dir):
    """从数据目录尾段推导账套代码：d:/底稿测试/AZ → AZ。"""
    if not data_dir:
        return None
    head, tail = os.path.split(os.path.normpath(data_dir.rstrip('/\\')))
    return tail or os.path.basename(head)

# ⚡ 核对注册表：未来新增核对程序只需在此加一行（模块 + manifest + 差异输出契约）。
#   第一阶段只登记【已验证成熟】的专项；其余零散程序随收编逐步补登记。
REGISTRY = {
    'bank': {
        'module': 'bank_review',          # 银行专项（双向核对/开户核对/异常分析）
        'manifest': 'bank_reconcile_manifest',
        'name': '银行专项核对（对账/开户/异常）',
        'subject': '银行存款',
        'accts': ['ADF', 'GJX', 'AYL'],
        'diff_json': '差异_银行核对.json',
        'entry_fn': 'gen',             # bank_review.gen(acct, out_dir=None)
        # bank_review.gen(acct)：acct 从 data_dir 推导（d:/底稿测试/AYL → AYL）
        'entry_args': {'acct': 'from_dir'},
        # 差异抽取规则：从开户核对底稿『逐账户明细』抽差异行
        #   data_dir = 账套根（如 d:/底稿测试/AYL），file 相对账套根
        'diff_sheets': [
            {
                'file': '底稿/*/银行开户账户核对_*.xlsx',
                'sheet': '逐账户明细',
                'marks': [7],           # 匹配结果列（⚠️ 清单独有/账面独有）
                'cols': [0, 6, 7],      # 主体/账面科目/匹配结果
                'colmap': {'entity': '0', 'code': '6', 'reason': '7'},
            },
            {
                # 网银双向核对『差异清单』：来源列含『有账无/有网银无』即差异
                'file': '底稿/*/银行网银核对底稿_*.xlsx',
                'sheet': '差异清单',
                'marks': [1],           # 来源列（账有网银无/网银有账无）
                'cols': [0, 1, 3, 4, 6, 7],  # 银行/来源/收入/支出/凭证号/原因
                'colmap': {'entity': '0', 'reason': '7', 'code': '6'},
                'skip_rows': 3,
            },
        ],
    },
    'aging': {
        'module': 'aging_review',         # 应收+合同资产账龄复核
        'manifest': 'aging_manifest',
        'name': '应收合同资产账龄复核',
        'subject': '应收账款',
        'accts': [],                      # 空 = 适用全部集团账套
        'diff_json': '差异_账龄复核.json',
        'entry_fn': 'recon_entry',
        # 差异抽取规则：账龄复核『与TB核对』sheet，差异列非 0 即差异行
        #   （表头在 r3；差异=col3，说明=col4）
        'diff_sheets': [{
            'file': '底稿/*/应收合同资产账龄复核_*.xlsx',
            'sheet': '与TB核对',
            'marks': [3],           # 差异列（数值非 0）
            'cols': [0, 1, 2, 3, 4],  # 项目/账龄表合计/TB期末/差异/说明
            'colmap': {'name': '0', 'gl_amt': '1', 'ext_amt': '2', 'diff': '3', 'reason': '4'},
            'num_mark': True,       # 数值型差异标记（≠0 即差异）
            'skip_rows': 3,         # 表头前 3 行跳过
        }],
    },
    'cross': {
        'module': 'cross_recon',          # 科目间勾稽核对
        'manifest': None,
        'name': '科目间勾稽核对',
        'subject': None,                  # 跨科目，不整合单一科目
        'accts': [],
        'diff_json': '差异_科目勾稽.json',
        'entry_fn': 'recon_entry',
        # 差异抽取规则：勾稽汇总『差异』列非 0 即差异（表头在 r4）
        'diff_sheets': [{
            'file': '中间产物/勾稽中间文件/科目间勾稽核对_生成.xlsx',
            'sheet': '科目间勾稽汇总',
            'marks': [6],           # 差异列（数值≠0）
            'cols': [0, 2, 3, 4, 5, 6, 7, 8],  # 主体/关系/科目/TB/GL/差异/状态/说明
            'colmap': {'entity': '0', 'name': '3', 'gl_amt': '5', 'ext_amt': '4',
                       'diff': '6', 'reason': '8'},
            'num_mark': True,
            'skip_rows': 4,
        }],
    },
    'confirm': {
        'module': 'confirm_reconcile',    # 函证回函核对（往来/银行）
        'manifest': None,
        'name': '函证回函核对',
        'subject': '应收账款',            # 以往来科目为主
        'accts': ['AZ'],
        'diff_json': '差异_函证核对.json',
        'entry_fn': 'recon_entry',
        # confirm_reconcile.main(acct='AZ')：acct 从 data_dir 尾段推导（d:/底稿测试/AZ → AZ）
        'entry_args': {'acct': 'from_dir'},
        # 差异抽取规则：往来函证核对『差异列≠0 或 结论=不符』；银行函证『回函金额缺失』
        'diff_sheets': [
            {
                'file': '底稿/*/函证回函_核对表_*.xlsx',
                'sheet': '往来函证核对',
                'marks': [8, 3],    # 差异列(回函-底稿) 数值≠0；文件名结论列含『不符』
                'cols': [0, 1, 2, 3, 4, 5, 6, 7, 8],  # 编号/主体/对方/结论/科目/发函/回函/底稿/差异
                'colmap': {'entity': '1', 'name': '2', 'reason': '3',
                           'gl_amt': '7', 'ext_amt': '6', 'diff': '8'},
                'num_mark': True,
                'mark_text_cols': [3],   # 附加文本标记列（含『不符』即差异）
                'skip_rows': 1,
            },
        ],
    },
}


def list_recons():
    """列出已登记核对专项。"""
    print('已登记外部核对专项：')
    for key, cfg in REGISTRY.items():
        accts = '全部' if not cfg.get('accts') else '/'.join(cfg['accts'])
        print(f'  {key:10s} {cfg["name"]:24s} 账套={accts}')


def load_recon(key):
    """动态导入专项模块（返回 module 对象）。"""
    cfg = REGISTRY.get(key)
    if not cfg:
        raise KeyError(f'未登记核对专项: {key}')
    return importlib.import_module(cfg['module'])


def recon_entry(key, data_dir, **kw):
    """统一调用入口：
    1. 调专项的 entry_fn（默认 recon_entry；缺失回退 main）生成底稿；
    2. 从产出的 xlsx 抽取差异 → 统一差异 JSON（recon_diff_extract）。
    返回差异 JSON 输出路径。"""
    cfg = REGISTRY[key]
    m = load_recon(key)
    fn_name = cfg.get('entry_fn', 'recon_entry')
    fn = getattr(m, fn_name, None) or getattr(m, 'main', None)
    if fn is None:
        raise AttributeError(f'{cfg["module"]} 无 {fn_name}()/main()')
    # ① 跑专项生成（若配置了 diff_sheets 且数据目录可访问）
    from recon_diff_extract import write_diffs, extract_all, find_diff_files
    has_src = bool(cfg.get('diff_sheets')) and os.path.isdir(data_dir)
    if has_src and not kw.get('extract_only'):
        try:
            import inspect
            try:
                sig = inspect.signature(fn)
                _params = list(sig.parameters.values())
                _required = [p for p in _params if p.default is inspect.Parameter.empty
                             and p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
            except (TypeError, ValueError):
                _required = []
            # 签名适配：main() 无参 → 模拟 argv；main(acct) → 按 entry_args 映射；否则传 data_dir
            if not _required:
                _old = sys.argv
                sys.argv = [cfg['module'] + '.py', data_dir]
                try:
                    fn()
                finally:
                    sys.argv = _old
            elif 'entry_args' in cfg:
                _a = cfg['entry_args']
                _vals = []
                for p in _params:
                    if p.kind not in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD):
                        continue
                    if p.name in _a:
                        _av = _a[p.name]
                        _vals.append(_acct_from_dir(data_dir) if _av == 'from_dir' else _av)
                    elif p.name == 'data_dir':
                        _vals.append(data_dir)
                    else:
                        _vals.append(None)
                fn(*_vals)
            else:
                fn(data_dir)
        except Exception as ex:
            print(f'  ⚠️ {cfg["module"]} 生成异常（继续抽取已有差异）：{ex}')
    # ② 抽取差异 → 统一 JSON
    diff_out = kw.get('diff_out') or (
        os.path.join(data_dir, cfg['diff_json']) if cfg.get('diff_json') else None)
    if has_src and diff_out:
        return write_diffs(data_dir, cfg, diff_out=diff_out)
    return diff_out


if __name__ == '__main__':
    import sys
    list_recons()
