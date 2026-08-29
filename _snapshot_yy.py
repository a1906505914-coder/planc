# -*- coding: utf-8 -*-
"""_snapshot_yy.py —— 底稿指纹快照（run_u8_on_sap 收口后自动对比回归）。

⚠️ 2026-08-19 整理文件夹时被误删，按 run_u8_on_sap 调用契约重建：
    _S.DIR        快照根目录（底稿目录）
    _S.snap(json) 拍当前快照（相对路径 → {size, mtime}）
    _S.diff(prev, curr, ignore) 对比，返回变化文件数（消失/新增/大小或时间变化）
用法见 run_u8_on_sap.py:482-499（[指纹] 收口后自动对比 + 滚动基准）。
"""
import json
import os


DIR = ''


def _walk():
    out = {}
    if not DIR or not os.path.isdir(DIR):
        return out
    for root, dirs, files in os.walk(DIR):
        for f in files:
            if f.startswith('~$'):
                continue
            fp = os.path.join(root, f)
            rel = os.path.relpath(fp, DIR).replace('\\', '/')
            try:
                st = os.stat(fp)
            except OSError:
                continue
            out[rel] = {'size': st.st_size, 'mtime': int(st.st_mtime)}
    return out


def snap(out_path):
    """拍当前快照 → json 文件。返回文件数。"""
    data = _walk()
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=0)
    return len(data)


def diff(prev_path, curr_path, ignore):
    """对比两次快照，返回变化文件数（消失/新增/大小或修改时间变化）。"""
    try:
        with open(prev_path, encoding='utf-8') as f:
            prev = json.load(f)
    except Exception:
        return 0
    try:
        with open(curr_path, encoding='utf-8') as f:
            curr = json.load(f)
    except Exception:
        return 0
    ig = set(ignore or [])
    bad = 0
    for k in prev:
        if k in ig or k not in curr:
            if k not in ig:
                bad += 1
            continue
        p, c = prev[k], curr[k]
        if p.get('size') != c.get('size') or p.get('mtime') != c.get('mtime'):
            bad += 1
    for k in curr:
        if k not in ig and k not in prev:
            bad += 1
    return bad


if __name__ == '__main__':
    import sys
    DIR = sys.argv[1] if len(sys.argv) > 1 else '.'
    out = sys.argv[2] if len(sys.argv) > 2 else '_snapshot.json'
    print('快照文件数:', snap(out), '→', out)
