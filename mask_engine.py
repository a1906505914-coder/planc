# -*- coding: utf-8 -*-
"""mask_engine.py —— 通用脱敏引擎（2026-08-19）。

按 mask_config 配置 + mask_dict 字典，把原始 xlsx 转为脱敏版。
原始版留在本地审计用；脱敏版用于进模型对话/外发协作。

用法：
    python mask_engine.py <src.xlsx> [--out <dst.xlsx>]
    python mask_engine.py --list                    # 列出已接入配置的专项
    python mask_engine.py --test <src.xlsx>         # 预演：只打印将替换的列与样本，不写文件

规则实现：
    S1    主体整格替换（mask_dict.lookup）
    S1sub 文本内主体名子串替换（字典按名称长度降序，避免短名抢先）
    S2    账号/编号掩码（前4后4；不足8位前3后3）
    S3    业务文本类型化（专项 text_kw 关键词→业务大类；未命中截断）
    S4    金额保留
    S6    不动（默认）
"""
import argparse
import os
import re
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

import openpyxl  # noqa: E402

import mask_config as MC  # noqa: E402
import mask_dict as MD  # noqa: E402


# ---------------------------------------------------------------- 规则实现
# 公司名特征后缀：S1 兜底判定用（外部客商名不在字典时打码）
_COMPANY_PAT = re.compile(
    r'(有限公司|有限责任公司|股份有限公司|集团有限公司|集团公司|股份公司|'
    r'公司|集团|工厂|厂|合伙企业|分公司|有限|总厂|'
    r'局|委|府|旗|县|镇|乡|村|'
    r'中心|研究院|研究所|事务所|学校|学院|医院|银行|联社|合作社|商行|饭店|酒店|'
    r'科技|能源|化工|钢铁|石化|煤业|矿业|铝业|电力|电气|机械|建设|建筑|'
    r'设备|材料|电子|信息|软件|网络|通信|物流|运输|贸易|商贸|实业|控股|投资|'
    r'ltd|corp|corporation|inc|gmbh|pte|llc|co\b)',
    re.IGNORECASE)


def _mask_num(v):
    """账号/编号掩码：≥8位 前4后4；≥6位 前3后3；否则全掩。"""
    s = str(v)
    if len(s) >= 8:
        return s[:4] + '*' * (len(s) - 8) + s[-4:]
    if len(s) >= 6:
        return s[:3] + '*' * (len(s) - 6) + s[-3:]
    return '*' * len(s)


# ⚡ S3 fallback 动作词概括：未命中业务大类时，只保留【动作+性质】，彻底去掉客户/项目专名
#    （外部客商"浙石化/中海壳牌"等不可穷举；客户身份由 S1 列承载，文本冗余信息打码）
_ACT = (
    ('销售', ['销售', '销']),
    ('采购', ['采购', '购入', '购']),
    ('付款', ['付款', '支付', '付']),
    ('收款', ['收款', '收到', '收']),
    ('结转调整', ['结转', '暂估', '冲销', '调整', '重分类']),
    ('费用', ['费用', '报销', '手续费', '服务费']),
    ('结算', ['结算', '清账', '核销']),
)


def _fallback(t):
    for cat, kws in _ACT:
        for k in kws:
            if k in t:
                return cat
    return '往来业务'


def _classify(txt, text_kw, cut=16):
    """业务文本类型化：命中专项关键词→业务大类；否则→动作词概括（无客户专名）。"""
    t = str(txt or '').strip()
    if not t:
        return t
    for cat, kws in text_kw.items():
        for k in kws:
            if k in t:
                return cat
    return _fallback(t)[:cut]


def _sub_names(txt, data, sorted_names, alias_map=None, alias_sorted=None):
    """长文本内主体名子串替换（sorted_names 按名称长度降序）。
    alias_map/alias_sorted：简称别名（如 杭氧股份→集团代码），全称替换后再替换。
    ⚡ 最后做外部组织名兜底：剩余含公司特征词的片段 → 【客商】（覆盖未入字典的外部客商/文件名）。"""
    t = str(txt or '')
    if not t:
        return t
    for nm in sorted_names:
        if nm and nm in t:
            t = t.replace(nm, data[nm])
    if alias_map and alias_sorted:
        for a in alias_sorted:
            if a and a in t:
                t = t.replace(a, alias_map[a])
    t = _COMPANY_PAT.sub('【客商】', t)
    return t


# ---------------------------------------------------------------- 引擎
def _apply(rule, value, cfg, data, sorted_names, alias_map=None, alias_sorted=None):
    """按规则处理单值。rule: S1/S1sub/S2/S3/S4/S6/None。"""
    if rule in (None, 'S4', 'S6'):
        return value
    if rule == 'S1':
        hit = data.get(str(value))
        if hit:
            return hit
        # 外部客商名兜底：不在字典但含组织名特征 → 【客商】（集团内主体已由字典覆盖）
        v = str(value)
        if len(v) >= 4 and _COMPANY_PAT.search(v):
            return '【客商】'
        return value          # 无公司特征的（代码/人名/短词）保留原样
    if rule == 'S1sub':
        return _sub_names(value, data, sorted_names, alias_map, alias_sorted)
    if rule == 'S2':
        return _mask_num(value)
    if rule == 'S3':
        # 先剔除文本内主体名（全称+简称），再类型化——防"销杭氧集团…"仍暴露主体
        return _classify(_sub_names(value, data, sorted_names, alias_map, alias_sorted),
                         cfg.get('text_kw', {}))
    return value


def _find_config(src_fp):
    """按文件名匹配专项配置 → (special_name, cfg)。"""
    base = os.path.basename(src_fp)
    for name, cfg in MC.CONFIG.items():
        files = cfg.get('files') or {}
        if base in files:
            return name, cfg
    return None, None


def _header_map(hdr):
    """表头行 → {列名: 索引}。"""
    return {str(h): i for i, h in enumerate(hdr) if h}


def transform(src_fp, out_fp=None, dry=False):
    """主转换。dry=True 只打印样本不写文件。
    返回 (专项名, 替换统计 dict)。"""
    name, cfg = _find_config(src_fp)
    if cfg is None:
        print('⚠️ 未找到该文件的脱敏配置：%s' % os.path.basename(src_fp))
        print('   已接入配置的专项：%s' % MC.active())
        return None, {}
    data = MD.names()
    if not data:
        print('⚠️ 脱敏字典为空，请先运行 python mask_dict.py 构建')
        return name, {}
    sorted_names = sorted(data, key=len, reverse=True)   # ⚡ 预排序一次（S1sub/S3 复用，防每行重排）
    # 简称别名 → 代码（杭氧股份/杭氧集团/杭氧 → 集团 A 代码）
    alias_map, alias_sorted = {}, []
    for full, aliases in MC.ALIAS.items():
        code = data.get(full)
        if code:
            for a in aliases:
                alias_map[a] = code
    alias_sorted = sorted(alias_map, key=len, reverse=True)

    if dry:
        print('=== 预演（不写文件）：%s / %s ===' % (name, os.path.basename(src_fp)))
    else:
        print('=== 脱敏：%s / %s ===' % (name, os.path.basename(src_fp)))

    wb = openpyxl.load_workbook(src_fp, read_only=True, data_only=True)
    wout = openpyxl.Workbook(write_only=True) if not dry else None
    stats = {}

    files_cfg = cfg['files'][os.path.basename(src_fp)]
    sheet_cfg = files_cfg.get('sheets', {})

    for sn in wb.sheetnames:
        ws = wb[sn]
        wso = wout.create_sheet(sn) if not dry else None
        col_cfg = sheet_cfg.get(sn, {})
        no_hdr = col_cfg.get('_no_header', False)   # ⚡ 无表头 sheet（如口径说明）：整表按 _idxN 处理
        stats[sn] = {'rows': 0, 'S1': 0, 'S1sub': 0, 'S2': 0, 'S3': 0}
        first = True
        hmap = {}
        for r in ws.iter_rows(values_only=True):
            if not r:
                continue
            out = list(r)
            if first and not no_hdr:
                hmap = _header_map(r)
                first = False
                if not dry:
                    wso.append(out)
                continue
            first = False
            for col, rule in col_cfg.items():
                if col.startswith('_idx'):
                    idx = int(col[4:])               # 按列索引（无表头 sheet 用）
                elif no_hdr:
                    continue
                else:
                    idx = hmap.get(col)              # 按列名
                if idx is None or idx >= len(out):
                    continue
                if out[idx] in (None, ''):
                    continue
                old = out[idx]
                new = _apply(rule, out[idx], cfg, data, sorted_names, alias_map, alias_sorted)
                if isinstance(new, str) and new != old:
                    stats[sn][rule] += 1
                out[idx] = new
                if dry and stats[sn]['rows'] < 3:
                    print('  [%s] %s=%s → %s' % (sn, col, str(old)[:24], new))
            if not dry:
                wso.append(out)
            stats[sn]['rows'] += 1
    wb.close()

    if not dry:
        wout.save(out_fp)
        print('已生成脱敏版：%s' % out_fp)
    return name, stats


def main(argv=None):
    ap = argparse.ArgumentParser(description='通用脱敏引擎')
    ap.add_argument('src', nargs='?', help='原始 xlsx')
    ap.add_argument('--out', help='脱敏版输出路径（默认 <同名>_脱敏.xlsx）')
    ap.add_argument('--dry', action='store_true', help='预演不写文件')
    ap.add_argument('--list', action='store_true', help='列出已接入专项')
    a = ap.parse_args(argv)

    if a.list:
        print('已接入(active):', MC.active())
        print('待接入(pending) %d 个:' % len(MC.pending()))
        for k in MC.pending():
            print('  - %s：%s' % (k, MC.CONFIG[k]['name']))
        return 0
    if not a.src:
        print('用法：python mask_engine.py <src.xlsx> [--out <dst.xlsx>] [--dry] [--list]')
        return 1

    if not a.out:
        base, ext = os.path.splitext(a.src)
        a.out = base + '_脱敏' + ext
    name, stats = transform(a.src, a.out, dry=a.dry)
    if name:
        for sn, s in stats.items():
            print('  %s: %d 行 | S1=%d S1sub=%d S3=%d' % (sn, s['rows'], s['S1'], s['S1sub'], s['S3']))
    return 0


if __name__ == '__main__':
    sys.exit(main())
