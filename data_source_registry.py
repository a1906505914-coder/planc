# -*- coding: utf-8 -*-
"""data_source_registry.py —— 外部数据源获取/预处理注册表（2026-08-24 架构 R4）。

设计背景（docs/audit_pipeline_design.md §6.2）：41 个外部程序分两类——
    · 【核对差异类】→ recon_registry（已收编 bank/aging/cross/confirm）
    · 【数据获取/预处理/专项分析类】→ 本注册表（合同 OCR/函证 OCR/银行流水解析/
      多币种拆分等）

本注册表把这些「上游数据源程序」统一登记，提供：
    · list_sources()  列出已登记数据源
    · run_source()    按注册配置调用（签名适配，同 recon_entry）
    · 产出登记（out_dir/out_file）便于后续核对程序定位数据

kind 语义：
    'preprocess'  源数据清洗/拆分（如 currency_split：多币种源按币种拆目录）
    'extract'     外部文档 OCR/解析（合同/函证/银行流水 → 结构化 JSON/xlsx）
    'analyze'     专项分析（异常识别/关注清单，输出 md/json）

用法：
    from data_source_registry import run_source, list_sources
    list_sources()
    run_source('currency_split', src_dir)
"""
import importlib
import os
import sys

# ⚡ 数据源注册表：未来新增数据源程序只需在此加一行。
#   kind / module / entry / inputs（源目录相对账套根或绝对路径）
REGISTRY = {
    'currency_split': {
        'kind': 'preprocess',
        'module': 'currency_split',
        'entry': 'process',
        'name': '多币种源数据拆分（CNY/THB/USD）',
        'accts': ['AZ', 'ADF'],
        'inputs': ['数据'],            # src_dir
        'outputs': '数据/{币种}/',      # 拆分后子目录
    },
    'contract_ocr': {
        'kind': 'extract',
        'module': 'contract_ocr_batch',
        'entry': 'main',
        'name': '借款合同 PDF OCR（批量）',
        'accts': ['ADF'],
        'inputs': ['数据/借款合同'],
        'outputs': '数据/借款合同_ocr/识别文本/',
    },
    'contract_summary': {
        'kind': 'extract',
        'module': 'contract_summary',
        'entry': 'main',
        'name': '合同 OCR 结构化摘要 + 汇总',
        'accts': ['ADF'],
        'inputs': ['数据/借款合同_ocr/识别文本'],
        'outputs': '合同摘要/',
    },
    'confirm_ocr': {
        'kind': 'extract',
        'module': 'confirm_batch',
        'entry': 'main',
        'name': '函证回函照片/PDF OCR',
        'accts': ['AZ'],
        'inputs': ['数据/回函'],
        'outputs': '数据/回函/_ocr_识别结果/',
    },
    'confirm_parse': {
        'kind': 'extract',
        'module': 'confirm_parser',
        'entry': 'main',
        'name': '回函 OCR 文本 → 结构化信息',
        'accts': ['AZ'],
        'inputs': ['数据/回函/_ocr_识别结果'],
        'outputs': '数据/回函/回函_解析结果_全量.json',
    },
    'bank_stmt_parse': {
        'kind': 'extract',
        'module': 'bank_statement_parser',
        'entry': 'main',
        'name': '银行流水解析（网银明细）',
        'accts': ['GJX', 'AYL'],
        'inputs': ['*网银*'],
        'outputs': '银行流水解析/',
    },
    'contract_anomaly': {
        'kind': 'analyze',
        'module': 'contract_anomaly',
        'entry': 'main',
        'name': '借款合同异常条款识别',
        'accts': ['ADF'],
        'inputs': ['数据/借款合同_ocr/识别文本'],
        'outputs': '借款合同_异常条款清单.md',
    },
}


def _acct_from_dir(data_dir):
    head, tail = os.path.split(os.path.normpath(data_dir.rstrip('/\\')))
    return tail or os.path.basename(head)


def _resolve_input(data_dir, inp):
    """输入相对账套根解析；含 * 通配 → 返回匹配的第一个目录。"""
    p = os.path.join(data_dir, inp) if not os.path.isabs(inp) else inp
    if '*' in p:
        import glob
        g = sorted(glob.glob(p))
        return g[0] if g else p
    return p


def list_sources():
    print('已登记外部数据源程序：')
    for key, cfg in REGISTRY.items():
        print(f'  {key:16s} [{cfg["kind"]:10s}] {cfg["name"]}')


def run_source(key, data_dir, **kw):
    """按注册配置调用数据源程序（签名适配，同 recon_entry）。返回输出提示。"""
    cfg = REGISTRY[key]
    m = importlib.import_module(cfg['module'])
    fn_name = cfg.get('entry', 'main')
    fn = getattr(m, fn_name, None) or getattr(m, 'main', None)
    if fn is None:
        raise AttributeError(f'{cfg["module"]} 无 {fn_name}()/main()')
    import inspect
    try:
        sig = inspect.signature(fn)
        params = [p for p in sig.parameters.values()
                  if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
        required = [p for p in params if p.default is inspect.Parameter.empty]
    except (TypeError, ValueError):
        required, params = [], []
    # 输入目录优先从 inputs 解析；无配置用 data_dir
    src = None
    if cfg.get('inputs'):
        src = _resolve_input(data_dir, cfg['inputs'][0])
    if not required:
        # 无参 main → 模拟 argv（传数据源路径）
        _old = sys.argv
        sys.argv = [cfg['module'] + '.py', src or data_dir] + list(kw.get('extra_args', []))
        try:
            fn()
        finally:
            sys.argv = _old
    else:
        p0 = required[0].name
        args = {p.name: (src or data_dir) if p.name == p0 else None for p in params}
        args.update({k: v for k, v in kw.items() if k in args})
        fn(*[args[p.name] for p in params])
    out = cfg.get('outputs', '')
    return os.path.join(data_dir, out) if out else None


if __name__ == '__main__':
    list_sources()
