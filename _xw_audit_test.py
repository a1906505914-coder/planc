# -*- coding: utf-8 -*-
import os, time
os.environ['CA_XW'] = '1'
import current_account_detail as CA
t0 = time.time()
try:
    CA.build_all_in_dir(r'D:/底稿测试/XBJ/数据/2025', out_dir=r'D:/底稿测试/XBJ/_xw_out',
                        period_mode='Y', only_subj='AR')
    print(f'[XW-AUDIT] AR xw（含审定表）耗时 {time.time()-t0:.1f}s', flush=True)
except Exception:
    import traceback; traceback.print_exc()
