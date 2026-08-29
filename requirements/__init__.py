# -*- coding: utf-8 -*-
"""审计要求知识点注册表（2026-08-14 架构试点：解决核对逻辑"四分五裂"——跨科目核对
无处安放，散落生成器/公共层。将"审计要求/核对"独立成知识点，声明数据依赖，引擎按
清单组装执行，科目生成器不再需要为核对逻辑改动）。

知识点 = 一个可独立执行的审计要求/核对程序（如 银行核对/函证/截止测试/毛利率/存货盘点）。
接口约定：
    @register_requirement(name='bank_reconcile', desc='...', subjects=('银行存款',),
                          params=('stmt_folder', 'journal_file'))
    def bank_reconcile(ctx):        # ctx: dict
        # ctx 含引擎注入的全局字段 + 本知识点 params
        ...
        return [out_path, ...]      # 输出文件列表（可空）

全局字段：data_dir / comp / year / verbose
本知识点字段：params 里声明的键，从清单 requirements[].params 注入。
"""
REQUIREMENTS = {}  # name -> {name, desc, subjects, params, fn}


def register_requirement(name, desc='', subjects=(), params=()):
    """注册一个审计要求知识点。name 全局唯一。"""
    def deco(fn):
        if name in REQUIREMENTS:
            raise ValueError(f'要求知识点重复注册：{name}')
        REQUIREMENTS[name] = {
            'name': name,
            'desc': desc,
            'subjects': tuple(subjects),
            'params': tuple(params),
            'fn': fn,
        }
        return fn
    return deco


def list_requirements():
    """列出全部已注册知识点。"""
    return sorted(REQUIREMENTS.values(), key=lambda r: r['name'])
