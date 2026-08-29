# -*- coding: utf-8 -*-
"""paths.py —— 路径配置中心（2026-08-09 建立，根治硬编码路径）。

⚡⚡⚡ 数据路径铁律（2026-08-14 用户定：程序是程序，数据是数据）：
  1. 程序代码【禁止】出现任何数据绝对路径（盘符/D:/底稿测试/账套名）；
  2. 数据路径唯一锚点 = 本文件 DATA_ROOT（或环境变量 AUDIT_DATA_ROOT）；
  3. 账套根 = DATA_ROOT + 账套名（账套名唯一事实来源 = ACCT_NAMES）；
  4. 账套【内】子结构一律写相对路径（相对账套根/DATA_ROOT），消费方用 P.resolve_rel() 展开；
  5. 文档/示例里的路径一律用占位符 <DATA_ROOT> / <APP_ROOT>；
  6. 解释器一律 sys.executable（bat 用 %USERPROFILE% 探测），不写绝对路径。
账套整体拷贝/改名：只需改 DATA_ROOT（或设 AUDIT_DATA_ROOT）→ 全部脚本零改动。

环境变量优先级：
  AUDIT_DATA_ROOT  > 本文件 DATA_ROOT 默认值
"""
import os

# ========== 数据根目录（所有账套文件夹所在父目录） ==========
# 2026-08-09：由 C:/Users/lvlh/Desktop 迁移至 D:/底稿测试（用户确认）
DATA_ROOT = os.environ.get('AUDIT_DATA_ROOT', r'D:\底稿测试')

# 小程序根目录（本文件所在目录，程序自身——随代码整体搬迁）
APP_ROOT = os.path.dirname(os.path.abspath(__file__))

# 账套文件夹名 → 完整路径（各账套数据目录）
# ⚡⚡ 2026-08-15 全账套统一改名（用户定，映射存档见 D:\底稿测试\文件夹改名同步清单.md）：
#   c→ACB  dq→GJX  F→ARF  FY→GFY  g→AYL  H→AJJ  J→ADZ  JTt→AJ
#   Q→AQ  R→AFJ  S→AS  T→ATT  Z→AZ  yy→AH  300→ADF  SSS→XBJ
#   （3300/建机电子底稿/FY2026年1-3月明细表=历史遗留，目录不存在，已移出）
ACCT_NAMES = ['ACB', 'GJX', 'ARF', 'GFY', 'AYL', 'AJJ', 'ADZ', 'AJ',
              'AQ', 'AFJ', 'AS', 'ATT', 'AZ', 'AH', 'ADF', 'XBJ']
DATA_DIRS = {n: os.path.join(DATA_ROOT, n) for n in ACCT_NAMES}

# ========== 常用账套快捷别名（新名 → 原义；变量名保持稳定，代码零改动） ==========
C = DATA_DIRS['ACB']                    # 诚本
DQ = DATA_DIRS['GJX']                   # DQ（2025+2026 1-3月）
F = DATA_DIRS['ARF']                    # 浙江日发控股集团
FY = DATA_DIRS['GFY']                   # FY 本级（12 实体双期间）
G = DATA_DIRS['AYL']                    # 上海朗炫
H = DATA_DIRS['AJJ']                    # 金洁环境
J = DATA_DIRS['ADZ']                    # 东轴
JTT = DATA_DIRS['AJ']                   # ga2025 + jj2025
Q = DATA_DIRS['AQ']                     # 上海金桥信息
R = DATA_DIRS['AFJ']                    # 山东日发 / 日发纺机
S = DATA_DIRS['AS']                     # energy
T = DATA_DIRS['ATT']                    # 北京同心 / 江西图之腾
YY = DATA_DIRS['AH']                    # SAP 主账套（88 家主体）
YY2026 = YY                             # 兼容旧引用（等价别名）
Z = DATA_DIRS['AZ']                     # 浙江兆龙互连（泰国三币种）
SSS = DATA_DIRS['XBJ']                  # 水利水电（3003 pilot）
D300 = DATA_DIRS['ADF']                 # 300 集团（3000-3900/6100/6900，网银核对主账套）

APP = APP_ROOT                          # 小程序（程序自身）


def acct(name):
    """取账套目录完整路径（不存在则返回 None）。"""
    p = os.path.join(DATA_ROOT, name)
    return p if os.path.isdir(p) else None


def resolve_rel(rel):
    """相对路径 → 绝对路径。

    约定：manifest/清单 JSON 的路径字段一律写【相对 DATA_ROOT 的路径】
    （如 '300/3300网银明细2026年1-7月'）；本函数展开为绝对路径。
    - rel 为 None/空 → 返回 None
    - rel 已是绝对路径（盘符或 UNC）→ 原样返回（容忍历史遗留）
    - 其余 → os.path.join(DATA_ROOT, rel)
    """
    if not rel:
        return None
    rel = str(rel).replace('/', os.sep).replace('\\', os.sep)
    if os.path.isabs(rel) or (len(rel) > 1 and rel[1] == ':'):
        return rel
    return os.path.join(DATA_ROOT, rel)


if __name__ == '__main__':
    print('DATA_ROOT =', DATA_ROOT)
    print('APP_ROOT  =', APP_ROOT)
    for n, p in DATA_DIRS.items():
        print(f'  {n:12s} -> {p}  {"✓" if os.path.isdir(p) else "（未找到）"}')
