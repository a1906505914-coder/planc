# -*- coding: utf-8 -*-
"""tests/test_architecture.py —— 架构回归单测（2026-08-30 测试金字塔起步）。

固化 2026-08-30 多用户化 + DAO 收尾反复手工验证的逻辑，防止改公共层回归：
  - audit_lock 并发锁（acquire/占用/check/错误token拒绝/释放/重获/过期兜底）
  - assert_build_totals 构建时内建断言（warn 命中/正确合计不命中/fail 抛异常/env 开关）
  - _conv_sap_gl_std 行规范化（debit/credit 回退 dr/cr、voucher 拼接、指令号透传）
  - _mem_total_fix 合计校验（错误命中/正确 0/公式跳过/集团合计跳过）
  - paths 环境变量切换（AUDIT_DATA_ROOT 跟随）
  - cache_manager.gl_cache_dir 缓存隔离（env 分桶/默认不变）

运行：python tests/test_architecture.py   （纯标准库，无需 pytest）
"""
import os
import sys
import time
import tempfile

APP = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
sys.path.insert(0, os.path.abspath(APP))

PASS = FAIL = 0


def check(name, cond, extra=''):
    global PASS, FAIL
    if cond:
        PASS += 1
        print('  ✓ %s' % name)
    else:
        FAIL += 1
        print('  ✗ %s %s' % (name, extra))


# ---------------------------------------------------------------- audit_lock
def test_audit_lock():
    print('[audit_lock 并发锁]')
    from audit_lock import acquire_run_lock, check_run_lock, release_run_lock
    root = tempfile.mkdtemp()
    t1 = acquire_run_lock(root)
    check('首次 acquire 成功', t1 is not None)
    t2 = acquire_run_lock(root)
    check('二次 acquire 被拒(None)', t2 is None)
    occ, info = check_run_lock(root)
    check('check 占用', occ and info)
    check('错误 token 释放被拒', not release_run_lock(root, 'other:1:1'))
    check('正确 token 释放成功', release_run_lock(root, t1))
    check('释放后重获', acquire_run_lock(root) is not None)
    release_run_lock(root)
    with open(os.path.join(root, '.audit.lock'), 'w') as f:
        f.write('oldhost:999:%.3f' % (time.time() - 99999))
    occ, _ = check_run_lock(root)
    check('过期锁视为空闲', not occ)
    check('过期后 acquire 成功', acquire_run_lock(root) is not None)


# ---------------------------------------------------------------- assert_build_totals
def _make_wb(wrong):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '测试科目 审定表'
    # 表头 ≥3 关键词单元格（序号/项目名称/期末余额），否则 _is_header_row_mem 不识别 → 合计不检查
    ws.append(['序号', '项目名称', '借方发生额', '贷方发生额', '期末余额'])
    ws.append([1, '明细1', 100.0, 100.0, 100.0])
    ws.append([2, '明细2', 200.0, 200.0, 200.0])
    ws.append(['', '合计', 1000.0 if wrong else 300.0, 300.0, 300.0])
    return wb


def test_assert_build_totals():
    print('[assert_build_totals 内建断言]')
    from audit_shell import assert_build_totals
    issues = assert_build_totals(_make_wb(True), fail_fast=False)
    check('错误合计 warn 命中', len(issues) == 1, str(issues))
    check('正确合计不命中', not assert_build_totals(_make_wb(False), fail_fast=False))
    try:
        assert_build_totals(_make_wb(True), fail_fast=True)
        check('fail 模式抛异常', False)
    except AssertionError:
        check('fail 模式抛异常', True)
    # env 开关
    os.environ['AUDIT_FAIL_FAST'] = '1'
    try:
        assert_build_totals(_make_wb(True))
        check('env AUDIT_FAIL_FAST=1 抛异常', False)
    except AssertionError:
        check('env AUDIT_FAIL_FAST=1 抛异常', True)
    finally:
        os.environ.pop('AUDIT_FAIL_FAST', None)


# ---------------------------------------------------------------- _conv_sap_gl_std
def test_conv_sap_gl_std():
    print('[_conv_sap_gl_std 行规范化]')
    from ledger_backend import _conv_sap_gl_std
    rows = [
        {'name': '原材料', 'vtype': 'SA', 'vno': '100', 'dr': 12.5, 'cr': 0.0,
         'wbs': 'W001', 'ord': '', 'mat': 'M9', 'proj': '', 'po': ''},
        {'name': '库存商品', 'vtype': '', 'vno': '', 'debit': 3.0, 'credit': 7.5,
         'wbs': '', 'ord': '', 'mat': '', 'proj': '', 'po': ''},
    ]
    out = _conv_sap_gl_std(rows)
    check('行数一致', len(out) == 2)
    check('voucher 拼接(SA-100)', out[0]['voucher'] == 'SA-100')
    check('debit 回退 dr(12.5)', out[0]['debit'] == 12.5)
    check('debit 直读(3.0)', out[1]['debit'] == 3.0)
    check('credit 直读(7.5)', out[1]['credit'] == 7.5)
    check('指令号透传 wbs', out[0]['wbs'] == 'W001')
    check('指令号透传 mat', out[0]['mat'] == 'M9')
    check('空凭证号回退空', out[1]['voucher'] == '')


# ---------------------------------------------------------------- _mem_total_fix
def test_mem_total_fix():
    print('[_mem_total_fix 合计校验]')
    from audit_checker import _mem_total_fix
    hdr = ['序号', '科目名称', '借方发生额', '贷方发生额', '期末余额']
    ok_rows = [hdr, [1, '明细1', 50.0, 20.0, 30.0], [2, '明细2', 70.0, 40.0, 30.0],
               ['', '合计', 120.0, 60.0, 60.0]]
    check('正确合计 0 命中', _mem_total_fix(ok_rows) == 0)
    bad_rows = [r[:] for r in ok_rows]
    bad_rows[3][2] = 999.0
    check('错误合计 命中', _mem_total_fix(bad_rows) > 0)
    # 公式单元格跳过
    fml_rows = [r[:] for r in ok_rows]
    fml_rows[3][2] = '=SUM(C2:C3)'
    check('公式单元格跳过', _mem_total_fix(fml_rows) == 0)
    # 集团合计跳过（builder 精确汇总，重算双计）
    grp_rows = [r[:] for r in ok_rows]
    grp_rows[3][0] = '集团合计'
    check('集团合计跳过', _mem_total_fix(grp_rows) == 0)


# ---------------------------------------------------------------- paths env
def test_paths_env():
    print('[paths 环境变量切换]')
    import paths as P
    default_root = P.DATA_ROOT
    os.environ['AUDIT_DATA_ROOT'] = 'E:/审计数据'
    import importlib
    importlib.reload(P)
    try:
        check('env 切换后 DATA_ROOT', P.DATA_ROOT == 'E:/审计数据', P.DATA_ROOT)
        check('env 切换后 DATA_DIRS[AH]', P.DATA_DIRS['AH'] == os.path.join('E:/审计数据', 'AH'))
    finally:
        os.environ.pop('AUDIT_DATA_ROOT', None)
        importlib.reload(P)
    check('还原默认 DATA_ROOT', P.DATA_ROOT == default_root)


# ---------------------------------------------------------------- cache_manager.gl_cache_dir
def test_gl_cache_dir():
    print('[cache_manager.gl_cache_dir 缓存隔离]')
    import cache_manager as CM
    d = r'D:\底稿测试\AH'
    check('默认 .gl_cache', CM.gl_cache_dir(d).endswith('.gl_cache'))
    os.environ['AUDIT_GL_CACHE_DIR'] = 'C:/tmp/gl'
    import importlib
    importlib.reload(CM)
    try:
        p = CM.gl_cache_dir(d)
        check('env 分桶到本地', p.startswith('C:/tmp/gl'), p)
        check('env 分桶含数据目录 hash', len(os.path.basename(p)) == 16)
    finally:
        os.environ.pop('AUDIT_GL_CACHE_DIR', None)
        importlib.reload(CM)
    check('还原默认 .gl_cache', CM.gl_cache_dir(d).endswith('.gl_cache'))


if __name__ == '__main__':
    test_audit_lock()
    test_assert_build_totals()
    test_conv_sap_gl_std()
    test_mem_total_fix()
    test_paths_env()
    test_gl_cache_dir()
    print('\n结果：PASS %d / FAIL %d' % (PASS, FAIL))
    sys.exit(1 if FAIL else 0)
