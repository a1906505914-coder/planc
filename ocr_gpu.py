# -*- coding: utf-8 -*-
"""ocr_gpu.py —— 借款合同批量 OCR（DirectML GPU 加速，RTX 独显；--cpu 走 CPU 引擎）。

⚡⚡ 2026-08-25 模块规整：合并 ocr_keypages.py（CPU 旧版）——本模块加 --cpu 参数
切换 CPU 引擎（contract_reader.ocr_image_array，与旧 ocr_keypages 一致），
ocr_keypages.py 标记废弃（逻辑已并入）。比 CPU 快约 5 倍（实测单页 441ms vs CPU 2.5s）。

用法（用专用 venv 运行，自动走 DML GPU）:
    C:\\Users\\lvlh\\WorkBuddy\\2026-07-15-15-26-19\\.ocr_dml_venv\\Scripts\\python.exe ocr_gpu.py [合同目录] [输出目录]
    ocr_gpu.py [合同目录] [输出目录] --cpu   # CPU 引擎（无 GPU 环境）
    ocr_gpu.py [合同目录] [输出目录] --all    # 扫描件全页识别（默认前4页）

默认目录:
    合同: d:/底稿测试/ADF/数据/2026/借款合同/化纤本年度新增借款合同
    输出: d:/底稿测试/ADF/数据/2026/借款合同_ocr/识别文本

文本型PDF → 直接提取全页文本；扫描件 → OCR前4页。输出受控本地，断点续跑。
"""
import os
import sys
import glob

# --cpu：不使用 DML GPU（CPU 引擎走 contract_reader，与旧 ocr_keypages 一致）
USE_CPU = '--cpu' in sys.argv or os.environ.get('OCR_CPU') == '1'
if not USE_CPU:
    # 启用 DirectML GPU（device_id=1 = RTX 5060 独显）
    os.environ['OCR_DML'] = '1'
    os.environ['OCR_DML_DEVICE'] = '1'

import pymupdf as fitz
import numpy as np
if USE_CPU:
    from contract_reader import ocr_image_array as _ocr_engine
else:
    from rapidocr_onnxruntime import RapidOCR

CONTRACT_DIR = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith('-') else \
    r'd:/底稿测试/ADF/数据/2026/借款合同/化纤本年度新增借款合同'
OUT_DIR = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('-') else \
    r'd:/底稿测试/ADF/数据/2026/借款合同_ocr/识别文本'
# --all 全页识别模式；默认只 OCR 扫描件前4页（关键信息页）
ALL_PAGES = '--all' in sys.argv or os.environ.get('OCR_ALL_PAGES') == '1'
MAX_PAGES = 9999 if ALL_PAGES else 4
DPI = 150

_engine = None


def get_engine():
    """全局复用 OCR 引擎（避免每份重复加载模型，DML 会缓存算子）。
    --cpu 模式返回 None（直接走 contract_reader 单页函数）。"""
    global _engine
    if USE_CPU:
        return None
    if _engine is None:
        _engine = RapidOCR()
    return _engine


def _needs_reocr(out_file):
    """判断是否需要重新识别：全页模式且已有txt是'前N页'扫描件产物（含页标记）→ 重跑。
    文本型（无页标记）全页一致，可跳过。"""
    if not os.path.exists(out_file):
        return True
    if not ALL_PAGES:
        return False
    try:
        with open(out_file, 'r', encoding='utf-8', errors='replace') as f:
            head = f.read(500)
        return '===== 第' in head  # 扫描件才含页标记
    except Exception:
        return False


def ocr_pdf(fp):
    doc = fitz.open(fp)
    n = len(doc)
    # 文本层有效 → 全页提取
    txt = ''.join(p.get_text() for p in doc)
    if txt and len(txt.strip()) > 80:
        doc.close()
        return 'text', txt
    # 扫描件 → OCR 前 MAX_PAGES 页（GPU 走 RapidOCR / CPU 走 contract_reader）
    parts = []
    eng = get_engine()
    for i in range(min(MAX_PAGES, n)):
        try:
            pix = doc[i].get_pixmap(dpi=DPI)
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n)
            img = arr[:, :, :3] if pix.n >= 3 else arr
            if USE_CPU:
                t = _ocr_engine(img)
                if t and t.strip():
                    parts.append(f'===== 第{i + 1}页 =====\n' + t)
            else:
                result, _ = eng(img)
                if result:
                    lines = [ln[1] for ln in result if len(ln) > 1 and ln[1]]
                    parts.append(f'===== 第{i + 1}页 =====\n' + '\n'.join(lines))
        except Exception as ex:
            print(f'  ⚠️ 第{i + 1}页失败: {str(ex)[:40]}', flush=True)
    doc.close()
    return 'ocr', '\n'.join(parts)


def main():
    pdfs = sorted(glob.glob(os.path.join(CONTRACT_DIR, '**', '*.pdf'),
                            recursive=True))
    print(f'OCR 启动: {len(pdfs)} 份合同 | {"CPU(contract_reader)" if USE_CPU else "DML GPU (RTX 5060)"}'
          f' | 模式={"全页" if ALL_PAGES else "前4页"}',
          flush=True)
    done = skip = fail = 0
    for i, fp in enumerate(pdfs, 1):
        rel = os.path.relpath(fp, CONTRACT_DIR)
        out_file = os.path.join(OUT_DIR, rel + '.txt')
        if not _needs_reocr(out_file):
            skip += 1
            continue
        try:
            kind, txt = ocr_pdf(fp)
            if not txt or not txt.strip():
                raise ValueError('空文本')
            os.makedirs(os.path.dirname(out_file), exist_ok=True)
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write(txt)
            done += 1
            print(f'  [{i}/{len(pdfs)}] {kind} {rel} ({len(txt)}字)', flush=True)
        except Exception as ex:
            fail += 1
            print(f'  [{i}/{len(pdfs)}] FAIL {rel}: {str(ex)[:40]}', flush=True)
    print(f'\n完成: 新增{done} | 跳过{skip} | 失败{fail} | 总{len(pdfs)}')


if __name__ == '__main__':
    main()
