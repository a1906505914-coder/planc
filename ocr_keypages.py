# -*- coding: utf-8 -*-
"""ocr_keypages.py —— 借款合同关键页OCR（内存方式，不落盘临时图）
⚠️ 2026-08-25 模块规整：本模块已并入 ocr_gpu.py（--cpu 参数走同引擎），
保留仅作兼容引用，新调用请用 ocr_gpu.py（GPU 默认 / --cpu 同此引擎）。
文本型PDF → pymupdf直接提取全页文本
扫描件PDF → 内存OCR前4页(借款金额/利率/期限/双方 在前几页)
输出 → 借款合同_ocr/识别文本/<相对路径>.txt（受控本地，断点续跑）
"""
import os, sys, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pymupdf as fitz
import numpy as np
from contract_reader import ocr_image_array

CONTRACT_DIR = r'd:/底稿测试/ADF/数据/2026/借款合同/化纤本年度新增借款合同'
OUT_DIR = r'd:/底稿测试/ADF/数据/2026/借款合同_ocr/识别文本'
MAX_PAGES = 4          # 扫描件只OCR前4页（关键信息页）
DPI = 150


def ocr_pdf(fp, out_file):
    doc = fitz.open(fp)
    n = len(doc)
    # 文本层有效 → 全页提取
    txt = ''.join(p.get_text() for p in doc)
    if txt and len(txt.strip()) > 80:
        doc.close()
        return 'text', txt
    # 扫描件 → 内存OCR前 MAX_PAGES 页
    parts = []
    for i in range(min(MAX_PAGES, n)):
        pix = doc[i].get_pixmap(dpi=DPI)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        img = arr[:, :, :3] if pix.n >= 3 else arr
        t = ocr_image_array(img)
        if t and t.strip():
            parts.append(f'===== 第{i + 1}页 =====\n' + t)
    doc.close()
    return 'ocr', '\n'.join(parts)


def main():
    pdfs = sorted(glob.glob(os.path.join(CONTRACT_DIR, '**', '*.pdf'), recursive=True))
    done = skip = fail = 0
    for i, fp in enumerate(pdfs, 1):
        rel = os.path.relpath(fp, CONTRACT_DIR)
        out_file = os.path.join(OUT_DIR, rel + '.txt')
        if os.path.exists(out_file):
            skip += 1
            continue
        try:
            kind, txt = ocr_pdf(fp, out_file)
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
