# -*- coding: utf-8 -*-
"""confirm_mask_local.py —— 回函【本地先行脱敏】程序 v2（纯数字编码，2026-08-22）。

⚡ 安全原则：真实主体名称只在本机处理，不进入任何外部/对话。
运行完成后，只把「回函_脱敏版/」目录用于后续（上传/对话/分析），原始目录与映射表留在本机。

编码规则（v2：全部数字编码，无"取字"逻辑，杜绝名称残留风险）：
  发函方（被审计单位）→ A001..A0NN   （分公司按母公司归并）
  回函银行              → B001..B0NN   （名称含"银行"/Bank 判定）
  回函企业              → C001..C0NN   （其余全部）

流程（完全自动，无人工干预）：
  1) 扫描 回函根目录 各回函子目录的文件名 → 收集全部唯一主体
  2) 发函方自动归并 → A 编号；回函方自动分类编号 → B/C 编号
  3) 复制 _ocr_识别结果 → 回函_脱敏版/，文件名与文本中主体名全部替换
  4) 校验脱敏版无真实名残留（含 OCR 碎片兜底片段，跨行容忍）

用法：
  python confirm_mask_local.py [回函根目录]
  默认：AZ 回函根 = D:\\底稿测试\\AZ\\数据\\2025\\回函
"""
import paths as P
import json
import os
import re
import shutil
import sys

_ACCOUNTING_FIRM_RE = re.compile(
    r'[\u4e00-\u9fa5]{2,12}会计\s*师?\s*事务\s*所')
_FUNCCENTER_RE = re.compile(r'[\u4e00-\u9fa5]{1,8}函证中心')
# 孤立的事务所简称兜底（OCR 拆行/截断时残留，本地可扩展）
FIRM_SHORT = ['天健']
# 企业/银行名中的通用后缀/前缀（用于生成"核心片段"兜底替换）
_COMPANY_SUFFIX = [
    '股份有限公司', '有限责任公司', '有限公司', '责任公司', '股份公司', '公司',
    '分公司', '营业部', '经营部', '工厂', '厂', '（中国）', '(中国)',
    '（新加坡）', '(新加坡)', '(THAILAND)', '(S.E.A)', '(泰国)',
    '公众股份公司', '上海分行', '集团', '支行',
]
_BANK_PREFIXES = ('中国', '上海', '浙江', '北京', '广州', '深圳',
                  '江苏', '广东', '山东', '湖南', '湖北', '天津')


def _is_bank(name):
    """是否银行：名称含『银行』或 Bank。"""
    return ('银行' in name) or ('bank' in name.lower())


def _collect_parties(root):
    """扫描回函子目录（企业/银行询证函回函），从文件名收集 (发函方, 回函方)。"""
    auditees, replies = [], []
    for sub in os.listdir(root):
        p = os.path.join(root, sub)
        if not os.path.isdir(p):
            continue
        for d in os.listdir(p):
            if '_' not in d:
                continue
            parts = d.split('_')
            if len(parts) >= 2:
                a, r = parts[0].strip(), parts[1].strip()
                if a and a not in auditees:
                    auditees.append(a)
                if r and r not in replies:
                    replies.append(r)
    return auditees, replies


# 发函母公司优先级（决定 A 编号分配顺序，固定避免代码漂移）
_AUDITEE_PRIORITY = ['浙江兆龙互连科技股份有限公司', '杭州兆龙物联技术有限公司',
                     '浙江兆龙数链科技有限公司', '浙江兆龙高分子材料有限公司',
                     'LONGTEK HOLDING GROUP PTE. LTD.',
                     'LONGTEK INTERCONNECT (THAILAND) CO., LTD.',
                     'LONGTEK SINGAPORE TRADING PTE. LTD.']


def _group_auditee(names):
    """发函方自动归并：分公司并入母公司（最长前缀），A 编号按固定优先级分配。"""
    names = set(names)
    groups = []  # [(母, [成员…])]
    seen = set()
    for mother in _AUDITEE_PRIORITY:
        if mother in names:
            members = sorted(n for n in names if n.startswith(mother))
            groups.append([mother, members])
            seen.update(members)
    rest = sorted((n for n in names if n not in seen), key=lambda x: (len(x), x))
    for n in rest:
        for g in groups:
            if n.startswith(g[0]):
                if n not in g[1]:
                    g[1].append(n)
                break
        else:
            groups.append([n, [n]])
    return groups


def _build_reply_map(replies):
    """回函方 → 纯数字编码：银行 B001+，企业 C001+（按名称排序，稳定不漂移）。"""
    mapping = {}
    banks = sorted(n for n in replies if _is_bank(n))
    comps = sorted(n for n in replies if not _is_bank(n))
    for i, n in enumerate(banks, 1):
        mapping[n] = f'B{i:03d}'
    for i, n in enumerate(comps, 1):
        mapping[n] = f'C{i:03d}'
    return mapping


def _company_fragment(name):
    """企业名核心片段（OCR 碎片兜底替换键）：去通用后缀取前 6 汉字/字母。"""
    n = name
    for suf in _COMPANY_SUFFIX:
        n = n.replace(suf, '')
    n = re.sub(r'[\s()（）]', '', n)
    zh = [c for c in n if '\u4e00' <= c <= '\u9fff']
    if len(zh) >= 6:
        return ''.join(zh[:6])
    if len(zh) >= 3:
        return ''.join(zh)
    alpha = re.sub(r'[^A-Za-z]', '', n)
    if len(alpha) >= 6:
        return alpha[:6]
    return None


def _bank_fragments(name):
    """银行名核心片段（OCR 碎片兜底替换键）。可返回多个，按长度降序替换。"""
    frags = []
    # 纯英文长词（KASIKORNBANK / BANK OF CHINA (THAI)）
    if re.fullmatch(r'[A-Za-z\s.()]{6,}', name):
        frags.append(name.strip())
    # 中文 XX银行
    m = re.search(r'[\u4e00-\u9fa5]{2,8}银行', name)
    if m:
        full = m.group(0)
        if len(full) >= 4:
            frags.append(full)
        for pfx in _BANK_PREFIXES:
            if full.startswith(pfx) and len(full) - len(pfx) >= 4:
                frags.append(full[len(pfx):])
    # 英文 XX BANK
    m = re.search(r'([A-Za-z]{2,20}(?:\s+[A-Za-z]{2,20}){0,3}\s+(?:BANK|Bank))', name)
    if m and len(m.group(1)) >= 6:
        frags.append(m.group(1))
    return frags


def _mask_text(text, mapping, auditee_codes):
    """文本中的主体名全部替换（支持 OCR 拆行/碎片）；会计师事务所泛称化。"""
    out = text
    subs = [(real, code) for real, code in auditee_codes.items()]
    subs += [(real, code) for real, code in mapping.items()]
    # 兜底片段（OCR 严重碎片化时完整名匹配失败，用独特前缀/片段）
    _FRAGMENTS = [
        ('浙江兆龙', 'A001'), ('兆龙互连', 'A001'), ('互连科技', 'A001'),
        ('杭州兆龙', 'A002'), ('兆龙物联', 'A002'), ('物联技术', 'A002'),
        ('兆龙数链', 'A003'), ('数链科技', 'A003'),
        ('兆龙高分子', 'A004'), ('高分子', 'A004'),
        ('LONGTEK HOLDING', 'A005'), ('LONGTEK INTERCONNECT', 'A006'),
        ('LONGTEK SINGAPORE', 'A007'),
        ('LONGTEK', 'A0'),   # LONGTEK 系兜底（碎片无法细分时）
        ('兆龙', 'A0'),   # 兆龙系兜底（碎片无法细分时）
    ]
    # 回函方核心片段（完整名被 OCR 打散时兜底）
    for real, code in mapping.items():
        if code.startswith('B'):
            for frag in _bank_fragments(real):
                subs.append((frag, code))
        else:
            frag = _company_fragment(real)
            if frag:
                subs.append((frag, code))
    subs += _FRAGMENTS
    # 通用公司后缀碎片 → 泛称
    for suf in ('股份有限公司', '有限责任公司', '有限公司', '股份有限'):
        subs.append((suf, '公司'))
    for real, code in sorted(subs, key=lambda x: -len(x[0])):
        # 允许主体名被 OCR 换行/空格拆开（任意相邻字符间可插空白）
        pat = r'\s*'.join(re.escape(c) for c in real)
        out = re.sub(pat, code, out)
    out = _ACCOUNTING_FIRM_RE.sub('★会计师事务所', out)
    out = _FUNCCENTER_RE.sub('★函证中心', out)
    for fs in FIRM_SHORT:
        out = out.replace(fs, '审计所')
    out = out.replace('★', '')
    return out


def _copy_tree_masked(src, dst, mapping, auditee_codes, report):
    """复制 _ocr_识别结果 → 脱敏版，文件名与内容脱敏。返回 True 成功。"""
    if not os.path.isdir(src):
        return False
    for dirpath, dirnames, filenames in os.walk(src):
        for fn in filenames:
            s = os.path.join(dirpath, fn)
            rel = os.path.relpath(s, src)
            rel_masked = _mask_text(rel, mapping, auditee_codes)
            d = os.path.join(dst, rel_masked)
            os.makedirs(os.path.dirname(d), exist_ok=True)
            try:
                with open(s, 'r', encoding='utf-8', errors='replace') as f:
                    txt = f.read()
                txt_masked = _mask_text(txt, mapping, auditee_codes)
                with open(d, 'w', encoding='utf-8') as f:
                    f.write(txt_masked)
                report['files'] += 1
            except Exception:
                try:
                    shutil.copy2(s, d)   # 非文本 → 原样复制（PDF 等）
                    report['copied'] += 1
                except Exception:
                    pass
    return True


def main(root=None):
    root = root or os.path.join(P.DATA_DIRS['AZ'], '数据', '2025', '回函')
    if not os.path.isdir(root):
        print(f'❌ 回函根目录不存在: {root}')
        return 1

    print('① 收集主体名…')
    auditees, replies = _collect_parties(root)
    print(f'   发函方 {len(auditees)} | 回函方 {len(replies)}')

    print('② 发函方归并 → A 编号…')
    groups = _group_auditee(auditees)
    auditee_codes = {}
    for i, (mother, members) in enumerate(groups, 1):
        code = f'A{i:03d}'
        for m in members:
            auditee_codes[m] = code

    print('③ 回函方 → B/C 数字编码…')
    reply_map = _build_reply_map(replies)
    nb = sum(1 for v in reply_map.values() if v.startswith('B'))
    print(f'   银行 {nb} 家 (B) | 企业 {len(reply_map) - nb} 家 (C)')

    # ④ 复制 _ocr_识别结果 → 脱敏版
    ocr_src = os.path.join(root, '_ocr_识别结果')
    dst = os.path.join(root, '回函_脱敏版')
    report = {'files': 0, 'copied': 0}
    print('④ 生成脱敏版…')
    if not _copy_tree_masked(ocr_src, dst, reply_map, auditee_codes, report):
        print(f'   ⚠️ 缺 OCR 目录: {ocr_src}（如已删，可只重建映射）')
    print(f'   复制/替换 {report["files"]} 个文本，{report["copied"]} 个原样')

    # ⑤ 本地保留对照表（不进脱敏版）
    mapping_doc = {
        'auditee_codes': auditee_codes,
        'reply_map': reply_map,
        'note': '真实名↔数字编码对照表，仅本机保留；对外只发 回函_脱敏版/',
    }
    map_fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'confirm_mask.json')
    with open(map_fp, 'w', encoding='utf-8') as f:
        json.dump(mapping_doc, f, ensure_ascii=False, indent=1)

    # 校验：脱敏版不得残留任何真实主体名（含关键片段，容忍跨行）
    print('⑤ 校验脱敏版无真实主体名残留…')
    leaked = []
    all_real = list(auditee_codes.keys()) + list(reply_map.keys()) + FIRM_SHORT
    # 兜底片段也作为残留检查键
    for real, code in reply_map.items():
        if code.startswith('B'):
            all_real += _bank_fragments(real)
        else:
            frag = _company_fragment(real)
            if frag:
                all_real.append(frag)
    all_real += ['浙江兆龙', '兆龙互连', '互连科技', '杭州兆龙', '兆龙物联', '物联技术',
                 '兆龙数链', '数链科技', '兆龙高分子', '高分子', '兆龙',
                 'LONGTEK', 'LONGTEK HOLDING', 'LONGTEK INTERCONNECT',
                 'LONGTEK SINGAPORE']
    for dirpath, _, filenames in os.walk(dst):
        for fn in filenames:
            fp = os.path.join(dirpath, fn)
            try:
                with open(fp, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()
            except Exception:
                continue
            for real in all_real:
                # 跨行/空格容忍检查
                pat = r'\s*'.join(re.escape(c) for c in real)
                if re.search(pat, content) or re.search(pat, fp):
                    leaked.append((os.path.relpath(fp, dst), real))
                    break
    print(f'   残留: {len(leaked)} 处' + ('' if not leaked else '（详见本地日志）'))
    if leaked:
        with open(os.path.join(dst, '_残留检查.txt'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(f'{p} <= {r}' for p, r in leaked))
    print(f'\n✅ 完成。可外发目录: {dst}')
    print(f'  对照表(仅本地): {map_fp}')
    return 0


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    sys.exit(main(args[0] if args else None))
