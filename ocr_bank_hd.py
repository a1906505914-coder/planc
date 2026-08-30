# -*- coding: utf-8 -*-
"""ocr_bank_hd.py —— 银行询证函回函【高清重OCR】（GPU DirectML，DPI 300）
旧识别文本为 CPU DPI150，表格结构丢失严重；本脚本以更高分辨率重渲染重识别，
覆盖 _ocr_识别结果/银行询证函回函/ 下的 txt（原始 PDF 在 银行询证函回函/，可随时重生成）。
文本型 PDF → 直接用文本层；扫描件 → GPU OCR 全页。
"""
import paths as P
import os
import sys
import glob

os.environ['OCR_DML'] = '1'
os.environ['OCR_DML_DEVICE'] = '1'

import pymupdf as fitz
import numpy as np
from rapidocr_onnxruntime import RapidOCR

SRC = os.path.join(P.DATA_DIRS['AZ'], '数据', '2025', '回函', '银行询证函回函')
DST = os.path.join(P.DATA_DIRS['AZ'], '数据', '2025', '回函', '_ocr_识别结果', '银行询证函回函')
DPI = 300

_engine = None


def get_engine():
    global _engine
    if _engine is None:
        _engine = RapidOCR()
    return _engine


def ocr_pdf(pdf):
    doc = fitz.open(pdf)
    # 文本型 → 直接提取
    txt = '\n'.join(p.get_text() for p in doc)
    if txt and len(txt.strip()) > 80:
        doc.close()
        return txt
    # 扫描件 → GPU OCR 全页（DPI 300）
    parts = []
    eng = get_engine()
    for i in range(len(doc)):
        pix = doc[i].get_pixmap(dpi=DPI)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n)
        img = arr[:, :, :3] if pix.n >= 3 else arr
        result, _ = eng(img)
        if result:
            lines = [ln[1] for ln in result if len(ln) > 1 and ln[1]]
            parts.append(f'===== 第{i + 1}页 =====\n' + '\n'.join(lines))
    doc.close()
    return '\n'.join(parts)


def main():
    pdfs = sorted(glob.glob(os.path.join(SRC, '**', '*.pdf'), recursive=True))
    print(f'银行回函 PDF: {len(pdfs)} 份 | DML GPU | DPI={DPI}', flush=True)
    done = fail = 0
    for i, pdf in enumerate(pdfs, 1):
        rel = os.path.relpath(pdf, SRC)
        out = os.path.join(DST, rel + '.txt')
        try:
            txt = ocr_pdf(pdf)
            os.makedirs(os.path.dirname(out), exist_ok=True)
            with open(out, 'w', encoding='utf-8') as f:
                f.write(txt)
            done += 1
            print(f'  [{i}/{len(pdfs)}] {rel} ({len(txt)}字)', flush=True)
        except Exception as e:
            fail += 1
            print(f'  [{i}/{len(pdfs)}] FAIL {rel}: {str(e)[:40]}', flush=True)
    print(f'\n完成: 新增{done} | 失败{fail} | 总{len(pdfs)}')


if __name__ == '__main__':
    main()
