# -*- coding: utf-8 -*-
"""audit_lock.py —— 同数据包并发锁（2026-08-30 多用户化）。

场景：多用户共享同一数据目录时，两人同时全量跑同一账套 → 写同一输出/中间产物互踩。
方案：数据目录放 .audit.lock 文件，内容 = {host}:{pid}:{started_ts}。
  - acquire_run_lock(root, ttl)：独占获取。锁不存在或已过期 → 建立并返回 token；
    否则返回 None（被他人占用）。
  - release_run_lock(root, token)：释放。仅当持有者 token 匹配（防误删他人锁）。
  - check_run_lock(root, ttl)：只读检查（preflight 用）→ (occupied, info)。

注意：Windows 共享盘文件锁跨机器有限，用 TTL（默认 30 分钟）兜底——进程崩溃后
锁残留超时自动视为可接管，避免阻塞全员。
"""
import os
import time
import socket

LOCK_NAME = '.audit.lock'
DEFAULT_TTL = 1800   # 30 分钟：超过视为残留（崩溃/关机），可接管


def _lock_fp(root):
    return os.path.join(root, LOCK_NAME)


def _read(root):
    fp = _lock_fp(root)
    try:
        with open(fp, encoding='utf-8') as f:
            return f.read().strip()
    except OSError:
        return None


def check_run_lock(root, ttl=DEFAULT_TTL):
    """返回 (occupied, info)。occupied=True = 锁存在且未过期（他人正在跑）。"""
    content = _read(root)
    if not content:
        return False, None
    parts = content.split(':')
    try:
        ts = float(parts[2])
    except (IndexError, ValueError):
        ts = 0.0
    if time.time() - ts > ttl:
        return False, content   # 过期残留：可接管
    return True, content


def acquire_run_lock(root, ttl=DEFAULT_TTL):
    """独占获取同数据包锁。成功返回 token；被占用返回 None。"""
    occupied, _info = check_run_lock(root, ttl)
    if occupied:
        return None
    token = '%s:%s:%.3f' % (socket.gethostname(), os.getpid(), time.time())
    try:
        os.makedirs(root, exist_ok=True)
        with open(_lock_fp(root), 'w', encoding='utf-8') as f:
            f.write(token)
    except OSError:
        return None
    return token


def release_run_lock(root, token=None):
    """释放锁。token 非空且不匹配 → 拒绝（防误删他人锁）。返回是否释放。"""
    cur = _read(root)
    if cur is None:
        return False
    if token is not None and cur != token:
        return False
    try:
        os.remove(_lock_fp(root))
        return True
    except OSError:
        return False
