# -*- coding: utf-8 -*-
"""cache_manager.py —— 缓存统一管理器（2026-08-24 架构阶段3）。

收口三套缓存体系的 key 生成/版本/清理：
- GL 磁盘缓存（sap_common .gl_cache/*.pklz + .sig）
- TB 磁盘缓存（audit_cache .cache/*.pkl）
- 进程内函数缓存（audit_common _FN_CACHE/_CACHE_TB/_CACHE_GL）

关键：
- stable_md5：跨进程稳定哈希（禁用 Python hash()，其 tuple 带随机种子 PYTHONHASHSEED）
- CACHE_VERSION：格式版本前缀，解析格式变化时 bump 自动失效旧缓存
- cleanup_stale：统一清理 .stale/.done/.trash 残留（沙箱禁 os.remove，move 到 .trash）
"""
import os

CACHE_VERSION = 'v1'   # ⚡ 格式版本：bump 使旧格式缓存自动失效


def stable_md5(*parts):
    """跨进程稳定哈希（md5），替代 Python hash()（随机种子）。"""
    import hashlib
    h = hashlib.md5()
    for p in parts:
        if isinstance(p, bytes):
            h.update(p)
        else:
            h.update(str(p).encode('utf-8', 'replace'))
    return h.hexdigest()[:16]


# ---------------- GL 磁盘缓存 key ----------------
def gl_cache_dir(data_dir):
    """GL 磁盘缓存目录（2026-08-30 多用户隔离）：
    env AUDIT_GL_CACHE_DIR 设置时 → <AUDIT_GL_CACHE_DIR>/<data_dir 哈希>
    （各机本地缓存，避免多人共享数据目录时并发写 .gl_cache 损坏、以及共享盘 I/O 慢）；
    不设则保持 <data_dir>/.gl_cache（单机现状不变）。"""
    root = os.environ.get('AUDIT_GL_CACHE_DIR', '')
    if root:
        return os.path.join(root, stable_md5(os.path.normpath(data_dir)))
    return os.path.join(data_dir, '.gl_cache')


def gl_cache_key(data_dir, comp, cols, ver=CACHE_VERSION):
    """GL 缓存文件名 key：{comp}_{ver}_{md5(cols)}.pklz。
    历史：hash() 随机化（不同进程 key 不同→永远 miss→OOM 重扫）、无版本（格式变化误用旧缓存）。"""
    h = stable_md5(str(sorted(cols)))
    return os.path.join(gl_cache_dir(data_dir), f'{comp}_{ver}_{h}.pklz')


def gl_cache_sig_key(cache_key):
    """GL 缓存签名文件：{key}.sig（源文件 mtime+size，normpath 统一分隔符）。"""
    return cache_key + '.sig'


def gl_source_sig(data_dir, comp):
    """GL 源文件签名（mtime+size）→ 缓存失效判断。路径 normpath 统一（正/反斜杠混传 bug）。"""
    sig = []
    try:
        from sap_common import _gl_files
        for fp in _gl_files(data_dir, comp):
            try:
                st = os.stat(fp)
                sig.append(f'{os.path.normpath(fp)}:{int(st.st_mtime)}:{st.st_size}')
            except OSError:
                pass
    except Exception:
        pass
    return '|'.join(sig)


# ---------------- TB 磁盘缓存 key（audit_cache 风格） ----------------
def tb_cache_key(parts, kind, ver=CACHE_VERSION):
    """TB 缓存 key：{ver}_{md5(parts)}_tb.pkl。parts 含 源文件签名+entities 结构+kind。"""
    h = stable_md5(*parts, kind)
    return f'{ver}_{h}_{kind}.pkl'


# ---------------- 清理 ----------------
def cleanup_stale(cache_dir, move_to='.trash'):
    """把 .stale/.done 残留 move 到 {cache_dir}/.trash（沙箱禁 os.remove，move 兼容）。
    返回移动数。"""
    if not os.path.isdir(cache_dir):
        return 0
    trash = os.path.join(cache_dir, move_to)
    try:
        os.makedirs(trash, exist_ok=True)
    except OSError:
        return 0
    import shutil
    n = 0
    for fn in sorted(os.listdir(cache_dir)):
        if fn.startswith(move_to):
            continue
        if '.stale' in fn or '.done' in fn:
            src = os.path.join(cache_dir, fn)
            try:
                shutil.move(src, os.path.join(trash, fn))
                n += 1
            except Exception:
                pass
    return n
