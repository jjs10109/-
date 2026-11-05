# -*- coding: utf-8 -*-
"""全局参数配置。"""
ATR_N = 14
MA_WINDOW = 60
TREND_WINDOW = 60
UP_RATIO_WINDOW = 20
ADV_SHORT_WINDOW = 20
ADV_LONG_WINDOW = 60
VOL_WINDOWS = (60, 252)
DEFAULT_PRICE_ADJ = "F"
DEFAULT_ADV_MIN = 5e7

# 网格适用性标签阈值。GI 仅用于解释“积极/正常/保守”，不直接决定档数。
DENSITY_DENSE = 0.62
DENSITY_MEDIUM = 0.50

MIN_POINTS_PER_INTERVAL = 1
MAX_POINTS_PER_INTERVAL = 16

# v2-grid-count-refactor:
# GI 只做“适用性/风险提示”，不再直接决定每段拆多少个点。
# 档数由价格距离、ATR、仓位变化、流动性共同决定。
GRID_PRICE_ATR_MULT = 0.80          # 单档价格步长约等于 0.8 * ATR；ATR 越大，价格档位越宽
GRID_ADV_LOW = 5e7                  # 低流动性参考 ADV
GRID_ADV_HIGH = 1e9                 # 高流动性参考 ADV
GRID_MIN_SINGLE_POS_STEP = 0.0020   # 低流动性时，单档仓位变化尽量不超过 0.20%
GRID_MAX_SINGLE_POS_STEP = 0.0080   # 高流动性时，单档仓位变化可放大到 0.80%
MIN_STEP_PCT = 0.003
MAX_STEP_PCT = 0.05
BUY_GAMMA_DEFAULT = 1.0
SELL_GAMMA_DEFAULT = 1.0

# v2-trend-gamma-adjustment:
# gamma 用来决定网格仓位分布是前置还是后置。
# 默认启用趋势修正：交易方向趋势越强，gamma 越大，仓位越后置；
# 交易方向趋势越弱/有反转迹象，gamma 越小，仓位越前置。
GAMMA_TREND_ADJUST_ENABLED = True
GAMMA_TREND_BETA = 0.55
GAMMA_TREND_MIN = 0.65
GAMMA_TREND_MAX = 2.50
GAMMA_TREND_RET20_SCALE = 0.12
GAMMA_TREND_SCORE_SCALE = 0.12
GAMMA_TREND_STREAK_START = 2.0
GAMMA_TREND_STREAK_FULL = 10.0

# v2-anchor-split-safety:
# 自动补中间锚点时，中间锚点最多只能完成两端经理目标仓位差的 70%。
# 目的：防止类似“55 卖出 1%，60 卖完”的指令在 56 附近就被自动中点过度卖完。
AUTO_SPLIT_MAX_MID_COMPLETION = 0.70
AUTO_SPLIT_EARLY_MID_COMPLETION = 0.50
# v2-pre-terminal-progress-guard:
# 终点前保护：没到远端基金经理价格之前，不能提前完成太大比例的本段仓位变化。
# x<=0.5 时最多完成 50%；0.5<x<1.0 时从 50% 线性放宽到 85%；只有 x=1.0 才能完成 100%。
PRE_TERMINAL_CAP_FRONT = 0.50
PRE_TERMINAL_CAP_BACK = 0.85
PRE_TERMINAL_TREND_REDUCTION = 0.20  # 趋势沿交易方向越强，终点前完成比例最多再压低 20%

# 博士权重学习参数
B_LEVELS = (0.8, 1.6, 2.4)
CORE_STEP = 0.05
MAX_SUB_WEIGHT = 0.40
MAX_TOTAL_WEIGHT = 1.00
COST_RATE = 0.001
