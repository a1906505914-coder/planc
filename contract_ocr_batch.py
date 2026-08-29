# -*- coding: utf-8 -*-
"""contract_ocr_batch.py —— 批量合同 PDF OCR 管线 v2（2026-08-17 稳定版）。
v2 修复（用户反馈多次卡死后的根因修复）：
  · 不落盘临时图：fitz 渲染 → numpy 数组 → 内存 OCR（绕开 safe-delete 删除失败）
  · 每 PDF 独立子进程 + 超时 kill：单页 OCR 死锁不再阻塞整个任务（v1 卡死的根因）
  · 串行稳定（workers 参数保留但默认 1——多进程并行/GPU 均实测会卡死）
特性：
  · 断点续跑：输出 txt 已存在则跳过（--force 重跑）
  · 进度可见：实时打印 完成数/总数/失败
  · 失败不中断：单文件失败/超时记入失败清单，其余继续
用法：
  python contract_ocr_batch.py <合同目录> [--out <输出目录>] [--dpi 150] [--timeout 300] [--force]
  python contract_ocr_batch.py <file1.pdf> <file2.pdf> ...
输出：每个 PDF → {输出目录}/{相对子路径}/{原文件名}.txt（含每页分隔头）
"""
import argparse
import glob
import os
import subprocess
import sys
import time

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


def _ocr_pdf_worker(pdf, out_file, dpi, force):
    """子进程内执行：单 PDF → 逐页内存 OCR → 落盘 txt。返回 (basename, 页数, status)。"""
    try:
        if os.path.exists(out_file) and not force:
            return (os.path.basename(pdf), 0, 'skip')
        from contract_reader import ocr_image_array
        try:
            import fitz
        except ImportError:
            import pymupdf as fitz  # PyMuPDF 1.24+ 模块名改为 pymupdf
        import numpy as np
        doc = fitz.open(pdf)
        n_pages = len(doc)
        parts = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=dpi)
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            img = arr[:, :, :3] if pix.n >= 3 else arr
            t = ocr_image_array(img)          # 内存 OCR，不落盘
            parts.append(f'===== 第{i + 1}页 =====\n' + (t or ''))
        doc.close()
        os.makedirs(os.path.dirname(out_file), exist_ok=True)
        with open(out_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(parts))
        return (os.path.basename(pdf), n_pages, 'ok')
    except Exception as ex:
        return (os.path.basename(pdf), 0, f'err:{ex}')


def _run_one(pdf, out_file, dpi, force, timeout):
    """单 PDF 用独立 python 子进程隔离执行，超时 kill（Windows 沙箱禁止句柄复制，spawn+Queue 不可用）。
    子进程 = 本脚本 --worker 模式，渲染+OCR+写 txt 都在子进程完成，父进程仅等待+超时管理。"""
    env = dict(os.environ)
    env['CONTRACT_OCR_CPU'] = '1'
    if APP_DIR not in env.get('PYTHONPATH', ''):
        env['PYTHONPATH'] = APP_DIR + os.pathsep + env.get('PYTHONPATH', '')
    cmd = [sys.executable, os.path.abspath(__file__), '--worker', pdf, '--worker-out', out_file,
           '--worker-dpi', str(dpi)]
    if force:
        cmd.append('--worker-force')
    try:
        proc = subprocess.Popen(cmd, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()
            return (os.path.basename(pdf), 0, f'timeout>{timeout}s')
        rc = proc.returncode
    except Exception as ex:
        return (os.path.basename(pdf), 0, f'err:{ex}')
    if rc == 0 and os.path.exists(out_file):
        return (os.path.basename(pdf), 0, 'ok')
    return (os.path.basename(pdf), 0, f'rc={rc}')


def main(argv=None):
    ap = argparse.ArgumentParser(description='批量合同 PDF OCR（内存版，稳定串行 + 每 PDF 超时 kill）')
    ap.add_argument('paths', nargs='*', help='合同 PDF 或目录（worker 子进程模式不需要）')
    ap.add_argument('--out', default=None, help='输出目录（默认 输入目录/_ocr_识别结果）')
    ap.add_argument('--workers', type=int, default=1, help='并行进程数（保留参数，默认1=串行最稳定）')
    ap.add_argument('--dpi', type=int, default=150, help='渲染分辨率（默认150；文本清晰可降）')
    ap.add_argument('--timeout', type=int, default=600, help='单 PDF 超时秒数（默认600=10分钟）')
    ap.add_argument('--force', action='store_true', help='强制重跑（覆盖已有 txt）')
    ap.add_argument('--continuous', action='store_true',
                    help='连续模式：单进程内逐份 OCR，引擎只加载一次（适配 easyocr 等加载慢的引擎；无超时隔离）')
    # 子进程模式（父进程用 subprocess 调用自身）
    ap.add_argument('--worker', metavar='PDF', default=None, help=argparse.SUPPRESS)
    ap.add_argument('--worker-out', default=None, help=argparse.SUPPRESS)
    ap.add_argument('--worker-dpi', type=int, default=150, help=argparse.SUPPRESS)
    ap.add_argument('--worker-force', action='store_true', help=argparse.SUPPRESS)
    a = ap.parse_args(argv)
    if a.worker:
        # 子进程模式：单 PDF OCR 完成后写 txt，正常退出码 0
        name, pages, status = _ocr_pdf_worker(a.worker, a.worker_out, a.worker_dpi, a.worker_force)
        if status == 'ok':
            return 0
        print(f'WORKER-FAIL {name} {status}', file=sys.stderr)
        return 1
    os.environ['CONTRACT_OCR_CPU'] = '1'   # 一律强制 CPU（GPU 单/多进程均实测卡死）

    if not a.paths:
        ap.print_usage()
        print('错误：缺少合同 PDF 或目录参数。')
        return 1
    pdfs = []
    for p in a.paths:
        if os.path.isdir(p):
            pdfs += glob.glob(os.path.join(p, '**', '*.pdf'), recursive=True)
        elif p.lower().endswith('.pdf') and os.path.exists(p):
            pdfs.append(p)
    pdfs = sorted(set(pdfs))
    if not pdfs:
        print('未找到 PDF 文件。')
        return 1

    base_dir = a.out or (a.paths[0] if os.path.isdir(a.paths[0]) else os.path.dirname(a.paths[0]))
    root_dir = a.paths[0] if os.path.isdir(a.paths[0]) else os.path.dirname(a.paths[0])
    out_root = os.path.join(base_dir, '_ocr_识别结果')
    tasks = []
    for pdf in pdfs:
        rel = os.path.relpath(pdf, root_dir)
        out_file = os.path.join(out_root, rel + '.txt')
        tasks.append((pdf, out_file, a.dpi, a.force, a.timeout))

    print(f'共 {len(pdfs)} 份 PDF → 输出 {out_root}（dpi={a.dpi}, timeout={a.timeout}s/份）', flush=True)
    t0 = time.time()
    done = ok = skip = fail = 0
    failed = []
    for i, (pdf, out_file, dpi, force, timeout) in enumerate(tasks, 1):
        if os.path.exists(out_file) and not force:
            name, pages, status = (os.path.basename(pdf), 0, 'skip')
        elif a.continuous:
            # 连续模式：单进程内直接 OCR（引擎全局单例复用，easyocr 加载 31s 只需一次）
            name, pages, status = _ocr_pdf_worker(pdf, out_file, dpi, force)
        else:
            name, pages, status = _run_one(pdf, out_file, dpi, force, timeout)
        done += 1
        if status == 'ok':
            ok += 1
        elif status == 'skip':
            skip += 1
        else:
            fail += 1
            failed.append((name, status))
        el = time.time() - t0
        print(f'  [{done}/{len(tasks)}] {name[:40]} | {pages}页 | {status} | 已用{el:.0f}s', flush=True)
    print(f'\n完成：共{len(tasks)}份 | 成功{ok} 跳过{skip} 失败{fail} | 总耗时 {time.time()-t0:.0f}s')
    if failed:
        print('失败清单：')
        for name, status in failed:
            print(f'  {name}: {status}')
    return 1 if fail else 0


if __name__ == '__main__':
    sys.exit(main())
