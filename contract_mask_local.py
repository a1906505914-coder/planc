# -*- coding: utf-8 -*-
"""contract_mask_local.py —— 借款合同【本地先行脱敏】v2（纯数字编码，2026-08-22）

⚡ 安全原则：借款公司/银行真实名只在本机处理，不进入任何外部/对话。
运行后只把「识别文本_脱敏/」用于后续分析，原始识别文本与对照表留在本机。

编码规则（v2：纯数字，无简称字）：
  借款公司 → 借01..借06   （目录名 3100国望 → 借01）
  银行     → 贷01..贷07   （关键词匹配）

本模块同时导出 mask_path/mask_text，供 loan_contract_reader 读取脱敏版时复用。

用法：
  python contract_mask_local.py [识别文本目录] [脱敏版输出目录]
"""
import paths as P
import json
import os
import re
import shutil
import sys

# 借款公司目录 → 借X（与 loan_contract_reader.BORROWER_MAP 一致）
BORROWER_DIR = [('3100国望', '借01'), ('3200苏州纤维', '借02'), ('3300中鲈', '借03'),
                ('3400港虹', '借04'), ('3900苏震', '借05'), ('6100新视界', '借06')]
# 银行关键词 → 贷X（与 loan_contract_reader.LENDER_MAP 一致）
BANK_KW = [('中国银行', '贷01'), ('中行', '贷01'),
           ('农业银行', '贷02'), ('农行', '贷02'),
           ('工商银行', '贷03'), ('工行', '贷03'),
           ('建设银行', '贷04'), ('建行', '贷04'),
           ('江苏银行', '贷05'), ('招银金租', '贷06'),
           ('委托贷款', '贷07')]
# 文本内借款公司关键词 → 借X
COMPANY_KW = [('江苏东方盛虹', '借01'), ('东方盛虹', '借01'), ('国望', '借01'),
              ('盛虹纤维', '借02'), ('苏州纤维', '借02'),
              ('中鲈', '借03'), ('港虹', '借04'),
              ('苏震', '借05'), ('新视界', '借06')]
_FIRM_RE = re.compile(r'[\u4e00-\u9fa5]{2,12}会计\s*师?\s*事务\s*所')

BORROWER_DIR_MAP = {d: code for d, code in BORROWER_DIR}
BANK_KW_LIST = [(kw, code) for kw, code in BANK_KW]


def mask_text(text):
    """文本主体名替换（借款公司/银行/事务所泛称）。
    ⚡ 2026-08-24 委托统一 desensitizer（contract 语境：借X/贷X）：
    实体+银行部分替换 + 公司名整体泛化，不误伤业务词。"""
    if not text:
        return text
    from desensitizer import get_contract
    dz = get_contract()
    # 本模块关键词映射注册（与 contract_mask.json 同源，补丁保险）
    dz.register('entity', {kw: code for kw, code in COMPANY_KW})
    dz.register('bank', {kw: code for kw, code in BANK_KW_LIST})
    out = dz.mask(text, 'text')
    out = _FIRM_RE.sub('会计师事务所', out)
    return out


def mask_path(rel):
    """相对路径脱敏：借款公司目录→借X；银行目录→贷X+残段（保留区分，防同名合并）；
    文件名内主体词→文本替换保留其余。"""
    parts = rel.replace('/', os.sep).replace('\\', os.sep).split(os.sep)
    out_parts = []
    for i, p in enumerate(parts):
        if p in BORROWER_DIR_MAP:
            out_parts.append(BORROWER_DIR_MAP[p])
        elif i == len(parts) - 1:
            # 文件名：只替换其中的主体词，保留编号/金额/扩展名
            out_parts.append(mask_text(p))
        else:
            hit = next((code for kw, code in BANK_KW_LIST if kw in p), None)
            if hit:
                # 保留目录残段（借款公司名也脱敏），避免同名文件被合并
                rest = mask_text(p)
                for _, code in BANK_KW_LIST:
                    rest = rest.replace(code, '')
                rest = rest.strip('（）() -_ ')
                out_parts.append(hit + (('-' + rest) if rest else ''))
            else:
                out_parts.append(p)
    return os.sep.join(out_parts)


def main(src=None, dst=None):
    src = src or os.path.join(P.DATA_DIRS['ADF'], '数据', '2026', '借款合同_ocr', '识别文本')
    dst = dst or os.path.join(P.DATA_DIRS['ADF'], '数据', '2026', '借款合同_ocr', '识别文本_脱敏')
    if not os.path.isdir(src):
        print(f'❌ 识别文本目录不存在: {src}')
        return 1

    # ① 映射表（仅本机）
    mapping = {'borrower_dir': BORROWER_DIR_MAP, 'bank_dir': {kw: code for kw, code in BANK_KW_LIST},
               'note': '借款合同真实名↔数字代号对照（借X/贷X），仅本机保留'}
    map_fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'contract_mask.json')
    with open(map_fp, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=1)

    # ② 复制 + 替换
    report = {'files': 0}
    for dirpath, _, filenames in os.walk(src):
        for fn in filenames:
            s = os.path.join(dirpath, fn)
            rel = os.path.relpath(s, src)
            rel_m = mask_path(rel)
            d = os.path.join(dst, rel_m)
            os.makedirs(os.path.dirname(d), exist_ok=True)
            try:
                with open(s, 'r', encoding='utf-8', errors='replace') as f:
                    txt = f.read()
                with open(d, 'w', encoding='utf-8') as f:
                    f.write(mask_text(txt))
                report['files'] += 1
            except Exception:
                shutil.copy2(s, d)
    print(f'复制/替换 {report["files"]} 个文本')

    # ③ 校验无真实名残留
    leaked = []
    real_kw = [kw for kw, _ in BANK_KW_LIST] + [kw for kw, _ in COMPANY_KW]
    for dirpath, _, filenames in os.walk(dst):
        for fn in filenames:
            fp = os.path.join(dirpath, fn)
            try:
                content = open(fp, 'r', encoding='utf-8', errors='replace').read()
            except Exception:
                continue
            for kw in real_kw:
                if kw in content:
                    leaked.append((os.path.relpath(fp, dst), kw))
                    break
    print(f'校验残留: {len(leaked)} 处')
    if leaked:
        with open(os.path.join(dst, '_残留检查.txt'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(f'{p} <= {k}' for p, k in leaked))
    print(f'✅ 完成。可外发目录: {dst}')
    return 0


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    sys.exit(main(args[0] if len(args) > 0 else None,
                  args[1] if len(args) > 1 else None))
