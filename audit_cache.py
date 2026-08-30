# -*- coding: utf-8 -*-
"""TB/GL 解析结果磁盘缓存（2026-08-05 效率优化）。
背景：每次改 bug 后重跑 builder，都要重新解析 FY 21.8MB/几十万行 GL（单次 3-8 分钟）。
方案：read_tb_full/read_gl_rows 的解析结果按【目录源文件签名】pickle 到 .cache/，
源文件（km/gl）mtime/size 未变 → 直接加载，跳过重解析。
- 签名 = 目录下全部 km/gl 文件 (名, mtime, size) 的 hash → 任何源文件变化自动失效。
- 开关：环境变量 AUDIT_NO_CACHE=1 强制禁用（改代码调试/首次构建后清理用）。
- 清理：保留最近 30 个缓存文件，超出按 mtime 删除最旧。
"""
import os
import hashlib
import pickle

CACHE_DIR = os.environ.get(
    'AUDIT_CACHE_DIR',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '.cache'))
# ⚡ 2026-08-30 多用户隔离：env AUDIT_CACHE_DIR 可指向各机本地缓存目录，
#   避免多人共享同一程序目录时并发写 .cache/*.pkl 损坏；不设则保持现状。
KEEP_MAX = 30


def _enabled():
    return os.environ.get('AUDIT_NO_CACHE', '') not in ('1', 'true', 'TRUE')


# 解析器版本：任何改变解析结果的代码修改（表头兼容/实体名清洗/列映射等）必须递增此号，
# 否则缓存签名不变 → 加载旧解析结果（曾致 jtt 实体名修复后 read_tb_full 键仍带『年』，
# 与 discover_entities 键不一致 → 往来审定表全 0，2026-08-06）。
PARSER_VERSION = '2026-08-15-v2'   # read_tb_full 表头行位置兼容（XBJ/SSS 表头在第 1 行；原固定 r2 探测+r3 读 → 1002 银行存款被当表头跳过）


def _file_sig(data_dir, entities, kind):
    """目录下 km/gl 文件签名（名+mtime+size + 解析器版本）→ md5。
    ⚡⚡ 2026-08-29 P0 修复：AH/SAP 账套 entities 的 gl 是【文件列表】（1010-1月.xlsx...），
    原只当单文件路径 → os.path.isfile(列表) 抛 TypeError → load 恒返回 None → 缓存永远 miss
    → 每次运行重解析大 GL（1357 247万行 17 分钟/2468 150万行 18 分钟）。支持单/多文件。"""
    files = []
    for e, yd in entities.items():
        for y, p in yd.items():
            for k in ('km', 'gl'):
                fp = p.get(k) if p else None
                fps = fp if isinstance(fp, (list, tuple)) else ([fp] if fp else [])
                for f in fps:
                    if f and os.path.isfile(f):
                        try:
                            st = os.stat(f)
                            files.append((os.path.normcase(f), int(st.st_mtime), st.st_size))
                        except OSError:
                            continue
    files.sort()
    h = hashlib.md5()
    h.update(kind.encode('utf-8'))
    h.update(PARSER_VERSION.encode('utf-8'))
    # ⚡⚡ 2026-08-15 修复：签名必须包含 discover 结构（主体→年度键）——
    #   否则 _extract_year 修复后 entities 结构变了（2020/2022 空年度键消失）而源文件未变，
    #   缓存命中修复前的旧解析（含空年度）→ self_tb_gen 的 years[0] 取到空年度 → 试算表全 0。
    #   （同型坑：PARSER_VERSION 注释记录的 jtt 实体名修复后 TB 键不一致事件，2026-08-06）
    for e, yd in sorted(entities.items()):
        h.update(str(e).encode('utf-8', 'replace'))
        h.update(','.join(sorted(str(y) for y in yd)).encode('utf-8'))
    for fp, mt, sz in files:
        h.update(fp.encode('utf-8', 'replace'))
        h.update(str(mt).encode('ascii'))
        h.update(str(sz).encode('ascii'))
    return h.hexdigest()


def _path(data_dir, entities, kind):
    sig = _file_sig(data_dir, entities, kind)
    return os.path.join(CACHE_DIR, '%s_%s.pkl' % (sig, kind))


def load(data_dir, entities, kind):
    """返回 (数据, 缓存路径) 或 (None, None)（未命中/禁用）。"""
    if not _enabled():
        return None, None
    try:
        p = _path(data_dir, entities, kind)
        if not os.path.isfile(p):
            return None, p
        with open(p, 'rb') as f:
            data = pickle.load(f)
        return data, p
    except Exception:
        return None, None


def save(data_dir, entities, kind, data):
    """写缓存（失败静默，不影响主流程）。"""
    try:
        if not _enabled():
            return
        os.makedirs(CACHE_DIR, exist_ok=True)
        p = _path(data_dir, entities, kind)
        tmp = p + '.tmp'
        with open(tmp, 'wb') as f:
            pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, p)
        _cleanup()
    except Exception:
        pass


def _cleanup():
    """保留最近 KEEP_MAX 个缓存文件。"""
    try:
        if not os.path.isdir(CACHE_DIR):
            return
        items = []
        for fn in os.listdir(CACHE_DIR):
            if fn.endswith('.pkl'):
                fp = os.path.join(CACHE_DIR, fn)
                try:
                    items.append((os.path.getmtime(fp), fp))
                except OSError:
                    continue
        items.sort()
        for _, fp in items[:-KEEP_MAX]:
            try:
                os.remove(fp)
            except OSError:
                pass
    except Exception:
        pass
