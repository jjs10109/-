# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd

from config import (
    DEFAULT_PRICE_ADJ, BUY_GAMMA_DEFAULT, SELL_GAMMA_DEFAULT,
    MIN_STEP_PCT, MAX_STEP_PCT, MIN_POINTS_PER_INTERVAL, MAX_POINTS_PER_INTERVAL,
    GRID_PRICE_ATR_MULT, GRID_ADV_LOW, GRID_ADV_HIGH,
    GRID_MIN_SINGLE_POS_STEP, GRID_MAX_SINGLE_POS_STEP,
    GAMMA_TREND_ADJUST_ENABLED, GAMMA_TREND_BETA, GAMMA_TREND_MIN, GAMMA_TREND_MAX,
    GAMMA_TREND_RET20_SCALE, GAMMA_TREND_SCORE_SCALE, GAMMA_TREND_STREAK_START, GAMMA_TREND_STREAK_FULL,
    AUTO_SPLIT_MAX_MID_COMPLETION, AUTO_SPLIT_EARLY_MID_COMPLETION,
    PRE_TERMINAL_CAP_FRONT, PRE_TERMINAL_CAP_BACK, PRE_TERMINAL_TREND_REDUCTION,
)
from utils import safe_float, clip
from data_loader import fetch_wind_daily_ohlcva, fetch_wind_intraday_minute, normalize_intraday_df, norm_code, guess_market
from features_daily import prepare_daily_features, summarize_stock_diagnostics
from features_minute import compute_intraday_micro_features
from gi_score import compute_grid_intensity, compute_grid_components
from core_grid_weights import fit_core_grid_weights

# 检查基金经理的档位是否正确
def validate_manager_levels_soft(levels: List[Dict[str, float]], side: str) -> Tuple[bool, str, List[Dict[str, float]]]:
    try:
        # 检查交易方向和档位
        if side not in ("buy", "sell"):
            return False, "side must be buy or sell", []
        if levels is None or len(levels) == 0:
            return False, "manager levels empty", []
        cleaned = []
        # 检查每个档位的价格和目标仓位，并转换为浮点数
        for x in levels:
            if "price" not in x or "target_pos" not in x:
                return False, f"level missing price or target_pos: {x}", []
            price = safe_float(x["price"], np.nan)
            target = safe_float(x["target_pos"], np.nan)
            if np.isnan(price) or np.isnan(target):
                return False, f"level contains non-numeric value: {x}", []
            if target < -1e-12 or target > 1 + 1e-12:
                return False, f"target_pos out of [0,1]: {x}", []
            cleaned.append({"price": float(price), "target_pos": float(target)})

        # 检查买入指令的价格是否随着档位上升而下降
        tmp = {float(x["price"]): float(x["target_pos"]) for x in cleaned}
        if side == "buy":
            ordered = [{"price": p, "target_pos": tmp[p]} for p in sorted(tmp.keys(), reverse=True)]
            for i in range(len(ordered) - 1):
                if not (ordered[i]["target_pos"] <= ordered[i + 1]["target_pos"] + 1e-12):
                    return False, f"buy target_pos must increase as price falls: {ordered[i]} -> {ordered[i + 1]}", ordered
        # 检查卖出指令的价格是否随着档位上升而上升
        else:
            ordered = [{"price": p, "target_pos": tmp[p]} for p in sorted(tmp.keys())]
            for i in range(len(ordered) - 1):
                if not (ordered[i]["target_pos"] >= ordered[i + 1]["target_pos"] - 1e-12):
                    return False, f"sell target_pos must decrease as price rises: {ordered[i]} -> {ordered[i + 1]}", ordered
        return True, "OK", ordered
    except Exception as e:
        return False, str(e), []

# 检查卖出指令的第一档目标仓位是否不高于当前仓位
def validate_sell_consistency(manager_sell_levels: List[Dict[str, float]], current_pos: float) -> Tuple[bool, str]:
    if not manager_sell_levels:
        return True, "OK"
    ok, reason, lv = validate_manager_levels_soft(manager_sell_levels, "sell")
    if not ok:
        return False, reason
    first_target = float(lv[0]["target_pos"])
    if current_pos + 1e-12 < first_target:
        return False, f"current_pos={current_pos} is lower than first sell target_pos={first_target}"
    return True, "OK"

# 根据当前价格，算出基金经理指令里的目标仓位。
def interpolate_target_pos(price: float, manager_levels: List[Dict[str, float]], side: str, gamma: float = 1.0, outside: str = "clamp") -> float:
    ok, reason, lv = validate_manager_levels_soft(manager_levels, side)
    if not ok:
        raise ValueError(reason)
    p = float(price)
    gamma = float(max(gamma, 1e-6))
    ps = [x["price"] for x in lv]
    qs = [x["target_pos"] for x in lv]
    if side == "buy":
        if p >= ps[0]: return float(qs[0]) if outside == "clamp" else np.nan
        if p <= ps[-1]: return float(qs[-1]) if outside == "clamp" else np.nan
        for i in range(len(ps) - 1):
            if ps[i] >= p >= ps[i + 1]:
                t = ((ps[i] - p) / (ps[i] - ps[i + 1])) ** gamma
                return float(qs[i] + (qs[i + 1] - qs[i]) * t)
    else:
        if p <= ps[0]: return float(qs[0]) if outside == "clamp" else np.nan
        if p >= ps[-1]: return float(qs[-1]) if outside == "clamp" else np.nan
        for i in range(len(ps) - 1):
            if ps[i] <= p <= ps[i + 1]:
                t = ((p - ps[i]) / (ps[i + 1] - ps[i])) ** gamma
                return float(qs[i] + (qs[i + 1] - qs[i]) * t)
    return float(qs[-1])

def _score_liquidity_from_adv(adv20: float) -> float:
    """把近20日成交额映射到 0~1；越大表示流动性越好。"""
    adv20 = safe_float(adv20, np.nan)
    if np.isnan(adv20) or adv20 <= 0:
        return 0.30
    lo = float(GRID_ADV_LOW)
    hi = float(GRID_ADV_HIGH)
    return float(np.clip((np.log(max(adv20, 1.0)) - np.log(lo)) / (np.log(hi) - np.log(lo)), 0.0, 1.0))



def _clip_signed(x: float, lo: float = -1.0, hi: float = 1.0) -> float:
    try:
        if np.isnan(float(x)):
            return 0.0
        return float(np.clip(float(x), lo, hi))
    except Exception:
        return 0.0


def compute_trend_gamma_adjustment(
    daily_profile: pd.DataFrame,
    side: str,
    base_gamma: float,
    enabled: bool = GAMMA_TREND_ADJUST_ENABLED,
) -> Dict[str, float]:
    """根据最近趋势对 gamma 做温和修正。

    gamma 的含义：
    - gamma = 1：线性分配；
    - gamma > 1：仓位变化后置，靠近基金经理远端价格时交易更多；
    - gamma < 1：仓位变化前置，更早交易。

    趋势修正方向：
    - 卖出侧：上涨趋势强 -> gamma 变大 -> 更靠后卖；上涨弱/回落 -> gamma 变小 -> 更早卖。
    - 买入侧：下跌趋势强 -> gamma 变大 -> 更靠后买；下跌弱/反弹 -> gamma 变小 -> 更早买。
    """
    base_gamma = safe_float(base_gamma, 1.0)
    if np.isnan(base_gamma) or base_gamma <= 0:
        base_gamma = 1.0

    meta = {
        "gamma_base": float(base_gamma),
        "gamma_final": float(base_gamma),
        "gamma_multiplier": 1.0,
        "trend_raw": 0.0,
        "trend_side": 0.0,
        "trend_ret20_component": 0.0,
        "trend_score_component": 0.0,
        "trend_streak_component": 0.0,
        "gamma_trend_enabled": bool(enabled),
        "gamma_min": float(GAMMA_TREND_MIN),
        "gamma_max": float(GAMMA_TREND_MAX),
        "gamma_beta": float(GAMMA_TREND_BETA),
    }
    if not enabled or daily_profile is None or len(daily_profile) < 21 or side not in ("buy", "sell"):
        return meta

    df = daily_profile.copy().sort_values("date").reset_index(drop=True)
    latest = df.iloc[-1]
    close_now = safe_float(latest.get("close", np.nan), np.nan)
    close_20 = safe_float(df.iloc[-21].get("close", np.nan), np.nan) if len(df) >= 21 else np.nan
    if not np.isnan(close_now) and not np.isnan(close_20) and close_20 > 0:
        ret20 = close_now / close_20 - 1.0
        ret_comp = _clip_signed(ret20 / float(GAMMA_TREND_RET20_SCALE))
    else:
        ret_comp = 0.0

    trend_score = safe_float(latest.get("TrendScore_60", np.nan), np.nan)
    trend_comp = _clip_signed(trend_score / float(GAMMA_TREND_SCORE_SCALE)) if not np.isnan(trend_score) else 0.0

    streak_len = safe_float(latest.get("streak_len", np.nan), np.nan)
    log_ret = safe_float(latest.get("log_ret", np.nan), np.nan)
    if np.isnan(log_ret) and len(df) >= 2:
        prev_close = safe_float(df.iloc[-2].get("close", np.nan), np.nan)
        if not np.isnan(prev_close) and prev_close > 0 and not np.isnan(close_now):
            log_ret = np.log(close_now / prev_close)
    streak_sign = 0.0 if np.isnan(log_ret) or abs(log_ret) <= 1e-12 else float(np.sign(log_ret))
    if not np.isnan(streak_len):
        streak_strength = float(np.clip((streak_len - float(GAMMA_TREND_STREAK_START)) / float(GAMMA_TREND_STREAK_FULL), 0.0, 1.0))
        streak_comp = streak_sign * streak_strength
    else:
        streak_comp = 0.0

    trend_raw = _clip_signed(0.40 * ret_comp + 0.30 * trend_comp + 0.30 * streak_comp)
    trend_side = trend_raw if side == "sell" else -trend_raw
    trend_side = _clip_signed(trend_side)

    multiplier = float(np.exp(float(GAMMA_TREND_BETA) * trend_side))
    gamma_final = float(np.clip(base_gamma * multiplier, float(GAMMA_TREND_MIN), float(GAMMA_TREND_MAX)))

    meta.update({
        "gamma_final": gamma_final,
        "gamma_multiplier": float(gamma_final / base_gamma) if base_gamma > 0 else 1.0,
        "trend_raw": float(trend_raw),
        "trend_side": float(trend_side),
        "trend_ret20_component": float(ret_comp),
        "trend_score_component": float(trend_comp),
        "trend_streak_component": float(streak_comp),
    })
    return meta


def choose_points_for_interval(
    price_a: float,
    price_b: float,
    atr: float,
    close: float,
    base_points: int = 1,
    micro_score: float = np.nan,
    pos_a: float = np.nan,
    pos_b: float = np.nan,
    adv20: float = np.nan,
    side: str = "buy",
    latest_features: Optional[pd.Series] = None,
    grid_intensity: Optional[Dict[str, Any]] = None,
) -> int:
    """决定两个基金经理锚点之间插入多少个中间网格点。

    v2-default-grid 口径：默认做网格，指标只调节密度，不直接否定网格。

    核心方向：
    - 波动偏高：说明有触发空间，但价格步长拉宽，所以拆得略少。
    - 流动性低：单档仓位更小，所以更细；但最终受 MAX_POINTS_PER_INTERVAL 限制。
    - 价格偏离方向有利：卖出遇高估/买入遇低估时更密；反向偏离时更稀。
    - 趋势强：不取消网格，只降低密度，避免过早过度执行。

    返回 n_inner：两个经理锚点之间插入的中间点数；总价格点数 = n_inner + 2。
    """
    width = abs(float(price_a) - float(price_b))
    if width <= 0:
        return int(MIN_POINTS_PER_INTERVAL)

    close = safe_float(close, np.nan)
    atr = safe_float(atr, np.nan)
    if np.isnan(close) or close <= 0:
        close = max(abs(float(price_a)), abs(float(price_b)), 1.0)
    if np.isnan(atr) or atr <= 0:
        atr = close * 0.03

    if latest_features is not None:
        comp = compute_grid_components(latest_features, side=side)
        adv20 = safe_float(latest_features.get("ADV_20", adv20), adv20)
        atr_pct = safe_float(latest_features.get("ATR_pct", atr / close), atr / close)
    elif grid_intensity is not None and isinstance(grid_intensity, dict):
        comp = (grid_intensity.get("components", {}) or {})
        atr_pct = atr / close
    else:
        atr_pct = atr / close
        liq_good = _score_liquidity_from_adv(adv20)
        comp = {
            "liquidity_good_score": liq_good,
            "liquidity_need_score": 1.0 - liq_good,
            "volatility_high_spacing_score": clip((atr_pct - 0.040) / (0.120 - 0.040), 0.0, 1.0),
            "side_favorable_deviation": 0.0,
            "side_unfavorable_deviation": 0.0,
            "trend_risk": 0.0,
        }

    liq_good = safe_float(comp.get("liquidity_good_score", _score_liquidity_from_adv(adv20)), 0.5)
    liq_need = safe_float(comp.get("liquidity_need_score", 1.0 - liq_good), 0.5)
    vol_high = safe_float(comp.get("volatility_high_spacing_score", 0.0), 0.0)
    favorable_dev = safe_float(comp.get("side_favorable_deviation", 0.0), 0.0)
    unfavorable_dev = safe_float(comp.get("side_unfavorable_deviation", 0.0), 0.0)
    trend_risk = safe_float(comp.get("trend_risk", 0.0), 0.0)

    # 1）价格维度：ATR 越高，单档价格步长越宽，所以高波动不会机械加密。
    vol_step_widen = 1.0 + 0.35 * vol_high
    price_step = float(GRID_PRICE_ATR_MULT) * atr * vol_step_widen
    price_step = float(np.clip(price_step, close * MIN_STEP_PCT, close * MAX_STEP_PCT))
    n_price_inner = max(1, int(np.ceil(width / price_step)) - 1)

    # 2）仓位/流动性维度：低流动性 -> 单档仓位更小 -> 档位更多，但受 MAX_POINTS 限制。
    pos_delta = abs(safe_float(pos_b, np.nan) - safe_float(pos_a, np.nan))
    max_single_pos_step = float(GRID_MIN_SINGLE_POS_STEP + liq_good * (GRID_MAX_SINGLE_POS_STEP - GRID_MIN_SINGLE_POS_STEP))
    max_single_pos_step = max(max_single_pos_step, float(GRID_MIN_SINGLE_POS_STEP))
    if np.isnan(pos_delta) or pos_delta <= 1e-12:
        n_pos_inner = 1
    else:
        n_pos_inner = max(1, int(np.ceil(pos_delta / max_single_pos_step)) - 1)

    # 3）密度调节：默认做网格，只调节疏密。
    base_n = max(n_price_inner, n_pos_inner, int(base_points), int(MIN_POINTS_PER_INTERVAL))
    m_liq = 1.0 + 0.35 * liq_need
    m_vol = 1.0 - 0.25 * vol_high
    m_dev = 1.0 + 0.45 * favorable_dev - 0.25 * unfavorable_dev
    m_trend = 1.0 - 0.35 * trend_risk
    multiplier = float(np.clip(m_liq * m_vol * m_dev * m_trend, 0.55, 1.65))

    n = int(round(base_n * multiplier))
    return int(np.clip(n, MIN_POINTS_PER_INTERVAL, MAX_POINTS_PER_INTERVAL))

def generate_interval_prices(price_a: float, price_b: float, side: str, n_inner: int, method: str = "linear") -> List[float]:
    a, b = float(price_a), float(price_b)
    if method == "geometric" and a > 0 and b > 0:
        arr = np.geomspace(a, b, n_inner + 2)
    else:
        arr = np.linspace(a, b, n_inner + 2)
    prices = [float(x) for x in arr]
    return sorted(prices, reverse=(side == "buy"))


def _pre_terminal_completion_cap(x: float, trend_side: float = 0.0) -> float:
    """终点前仓位完成比例上限。

    x 是当前网格价格在本段经理锚点区间里的进度：
    - x=0：靠近本段起点；
    - x=1：到达本段远端经理锚点。

    规则：
    - x <= 0.5：最多完成 50%；
    - 0.5 < x < 1：最多从 50% 线性放宽到 85%；
    - x >= 1：允许完成 100%。

    趋势保护：如果趋势沿交易方向很强，说明价格可能继续往远端走，
    终点前完成比例再小幅压低，但不取消网格。
    """
    try:
        x = float(np.clip(float(x), 0.0, 1.0))
    except Exception:
        x = 0.0
    if x >= 1.0 - 1e-12:
        return 1.0
    front = float(PRE_TERMINAL_CAP_FRONT)
    back = float(PRE_TERMINAL_CAP_BACK)
    if x <= 0.50:
        cap = front
    else:
        cap = front + (back - front) * ((x - 0.50) / 0.50)
    try:
        ts = max(float(trend_side), 0.0)
    except Exception:
        ts = 0.0
    # 趋势判断可能有噪声，所以这里只做温和压低，不直接否定成交。
    cap = cap * (1.0 - float(PRE_TERMINAL_TREND_REDUCTION) * min(ts, 1.0))
    return float(np.clip(cap, 0.0, 1.0))


def _apply_terminal_target_guard(
    px: float,
    target: float,
    manager_levels: List[Dict[str, float]],
    side: str,
    trend_side: float = 0.0,
) -> Tuple[float, bool, float, float]:
    """保护每一段经理锚点的远端目标，避免提前完成。

    这里不是只保护整条 ladder 最后的终点，而是保护当前所在的
    两个基金经理锚点区间：每段只能消化本段仓位差，且没到本段
    远端价格之前，不能完成过高比例。

    返回：target_after_guard, guard_applied, interval_progress_x, completion_cap。
    """
    try:
        if manager_levels is None or len(manager_levels) < 2:
            return float(target), False, np.nan, np.nan
        ok, _, lv = validate_manager_levels_soft(manager_levels, side)
        if not ok or len(lv) < 2:
            return float(target), False, np.nan, np.nan
        px = float(px)
        target = float(target)
        eps = 1e-10
        for i in range(len(lv) - 1):
            left = lv[i]
            right = lv[i + 1]
            p0 = float(left["price"])
            p1 = float(right["price"])
            w0 = float(left["target_pos"])
            w1 = float(right["target_pos"])
            if abs(p1 - p0) <= eps or abs(w1 - w0) <= 1e-12:
                continue
            if side == "sell":
                # 卖出区间价格从低到高，仓位从高到低。
                if not (p0 - eps <= px <= p1 + eps):
                    continue
                x = (px - p0) / (p1 - p0)
                if x >= 1.0 - eps or abs(px - p1) <= eps:
                    return target, False, float(np.clip(x, 0.0, 1.0)), 1.0
                cap = _pre_terminal_completion_cap(x, trend_side=trend_side)
                guarded_target = w0 - (w0 - w1) * cap
                if target < guarded_target - 1e-12:
                    return float(np.clip(guarded_target, 0.0, 1.0)), True, float(np.clip(x, 0.0, 1.0)), cap
                return target, False, float(np.clip(x, 0.0, 1.0)), cap
            elif side == "buy":
                # 买入区间价格从高到低，仓位从低到高。
                if not (p1 - eps <= px <= p0 + eps):
                    continue
                x = (p0 - px) / (p0 - p1)
                if x >= 1.0 - eps or abs(px - p1) <= eps:
                    return target, False, float(np.clip(x, 0.0, 1.0)), 1.0
                cap = _pre_terminal_completion_cap(x, trend_side=trend_side)
                guarded_target = w0 + (w1 - w0) * cap
                if target > guarded_target + 1e-12:
                    return float(np.clip(guarded_target, 0.0, 1.0)), True, float(np.clip(x, 0.0, 1.0)), cap
                return target, False, float(np.clip(x, 0.0, 1.0)), cap
    except Exception:
        return float(target), False, np.nan, np.nan
    return float(target), False, np.nan, np.nan


def auto_split_anchored_buy_levels(current_pos: float, upper_price: float, upper_target: float, lower_price: float, lower_target: float, daily_profile: pd.DataFrame) -> Tuple[List[Dict[str, float]], Dict[str, Any]]:
    if not (upper_price > lower_price):
        raise ValueError("buy anchors require upper_price > lower_price")
    if not (0 <= current_pos <= upper_target <= lower_target <= 1):
        raise ValueError("buy anchors require 0 <= current_pos <= upper_target <= lower_target <= 1")
    fit = fit_core_grid_weights(daily_profile)
    w2 = float(fit["weights"][1])
    w3 = float(fit["weights"][2])
    if w2 + w3 <= 1e-12:
        w2_norm = w3_norm = 0.5
    else:
        w2_norm = w2 / (w2 + w3)
        w3_norm = w3 / (w2 + w3)
    residual = lower_target - upper_target

    anchor_span = upper_price - lower_price
    latest_atr = float(fit.get("latest_atr", np.nan))
    if np.isnan(latest_atr) or latest_atr <= 0 or anchor_span <= 0:
        frac_mid = w2_norm
        mode = "weight_fallback"
    else:
        atr_frac = clip((0.8 * latest_atr) / anchor_span, 0.20, 0.80)
        deep_bias = clip(w3_norm - w2_norm, -0.50, 0.50)
        frac_mid = clip(atr_frac + 0.20 * deep_bias, 0.20, 0.80)
        mode = "atr_anchored_weight_bias"
    price_mid = upper_price - frac_mid * (upper_price - lower_price)

    # 安全约束：自动补出的中间锚点价格是 flexible 的，但如果中间价格落在
    # “前半段”（买入侧=更靠近上方价格，还没到精确中点），仓位完成度最多 50%；
    # 到了后半段才允许最高 70%。这样避免还没接近远端经理目标价就提前买满。
    exact_mid_price = (upper_price + lower_price) / 2.0
    early_mid = bool(price_mid > exact_mid_price + 1e-12)
    max_mid_completion = float(np.clip(AUTO_SPLIT_EARLY_MID_COMPLETION if early_mid else AUTO_SPLIT_MAX_MID_COMPLETION, 0.0, 1.0))
    raw_pos_mid = upper_target + residual * w2_norm
    max_pos_mid = upper_target + residual * max_mid_completion
    pos_mid = min(raw_pos_mid, max_pos_mid)
    pos_mid = clip(pos_mid, upper_target, lower_target)
    mid_completion_frac = float((pos_mid - upper_target) / residual) if abs(residual) > 1e-12 else 0.0
    levels = [
        {"price": float(upper_price), "target_pos": float(upper_target)},
        {"price": float(price_mid), "target_pos": float(pos_mid)},
        {"price": float(lower_price), "target_pos": float(lower_target)},
    ]
    ok, reason, ordered = validate_manager_levels_soft(levels, "buy")
    if not ok:
        raise ValueError(f"auto split generated invalid buy levels: {reason}")
    meta = {
        "split_mode": mode,
        "price_mid": price_mid,
        "pos_mid": pos_mid,
        "raw_pos_mid": raw_pos_mid,
        "mid_completion_frac": mid_completion_frac,
        "max_mid_completion": max_mid_completion,
        "early_mid": early_mid,
        "exact_mid_price": exact_mid_price,
        "mid_completion_capped": bool(raw_pos_mid > max_pos_mid + 1e-12),
        "frac_mid": frac_mid,
        "w2_norm": w2_norm,
        "w3_norm": w3_norm,
        "fit": fit,
    }
    return ordered, meta




def auto_split_anchored_sell_levels(current_pos: float, lower_price: float, lower_target: float, upper_price: float, upper_target: float, daily_profile: pd.DataFrame) -> Tuple[List[Dict[str, float]], Dict[str, Any]]:
    """卖出侧两端锚点自动补中间档，让卖出侧和买入侧逻辑对称。

    卖出档位按价格从低到高排列：
    lower_price -> lower_target, upper_price -> upper_target。
    随着价格上涨，目标仓位下降，所以要求 upper_target <= lower_target <= current_pos。
    """
    if not (upper_price > lower_price):
        raise ValueError("sell anchors require upper_price > lower_price")
    if not (0 <= upper_target <= lower_target <= current_pos <= 1):
        raise ValueError("sell anchors require 0 <= upper_target <= lower_target <= current_pos <= 1")
    fit = fit_core_grid_weights(daily_profile)
    w2 = float(fit["weights"][1])
    w3 = float(fit["weights"][2])
    if w2 + w3 <= 1e-12:
        w2_norm = w3_norm = 0.5
    else:
        w2_norm = w2 / (w2 + w3)
        w3_norm = w3 / (w2 + w3)
    residual = lower_target - upper_target

    anchor_span = upper_price - lower_price
    latest_atr = float(fit.get("latest_atr", np.nan))
    if np.isnan(latest_atr) or latest_atr <= 0 or anchor_span <= 0:
        frac_mid = w2_norm
        mode = "weight_fallback"
    else:
        atr_frac = clip((0.8 * latest_atr) / anchor_span, 0.20, 0.80)
        deep_bias = clip(w3_norm - w2_norm, -0.50, 0.50)
        frac_mid = clip(atr_frac + 0.20 * deep_bias, 0.20, 0.80)
        mode = "atr_anchored_weight_bias"
    price_mid = lower_price + frac_mid * (upper_price - lower_price)

    # 安全约束：自动补出的中间锚点价格是 flexible 的，但如果中间价格落在
    # “前半段”（卖出侧=更靠近下方价格，还没到精确中点），仓位完成度最多 50%；
    # 到了后半段才允许最高 70%。这样避免还没接近远端经理目标价就提前卖完。
    exact_mid_price = (lower_price + upper_price) / 2.0
    early_mid = bool(price_mid < exact_mid_price - 1e-12)
    max_mid_completion = float(np.clip(AUTO_SPLIT_EARLY_MID_COMPLETION if early_mid else AUTO_SPLIT_MAX_MID_COMPLETION, 0.0, 1.0))
    raw_pos_mid = lower_target - residual * w2_norm
    min_pos_mid = lower_target - residual * max_mid_completion
    pos_mid = max(raw_pos_mid, min_pos_mid)
    pos_mid = clip(pos_mid, upper_target, lower_target)
    mid_completion_frac = float((lower_target - pos_mid) / residual) if abs(residual) > 1e-12 else 0.0
    levels = [
        {"price": float(lower_price), "target_pos": float(lower_target)},
        {"price": float(price_mid), "target_pos": float(pos_mid)},
        {"price": float(upper_price), "target_pos": float(upper_target)},
    ]
    ok, reason, ordered = validate_manager_levels_soft(levels, "sell")
    if not ok:
        raise ValueError(f"auto split generated invalid sell levels: {reason}")
    meta = {
        "split_mode": mode,
        "price_mid": price_mid,
        "pos_mid": pos_mid,
        "raw_pos_mid": raw_pos_mid,
        "mid_completion_frac": mid_completion_frac,
        "max_mid_completion": max_mid_completion,
        "early_mid": early_mid,
        "exact_mid_price": exact_mid_price,
        "mid_completion_capped": bool(raw_pos_mid < min_pos_mid - 1e-12),
        "frac_mid": frac_mid,
        "w2_norm": w2_norm,
        "w3_norm": w3_norm,
        "fit": fit,
    }
    return ordered, meta


def split_manager_levels(manager_levels: List[Dict[str, float]], side: str, daily_profile: pd.DataFrame, micro: Optional[Dict[str, float]] = None, grid_intensity: Optional[Dict[str, Any]] = None, method: str = "linear", gamma: Optional[float] = None) -> pd.DataFrame:
    ok, reason, lv = validate_manager_levels_soft(manager_levels, side)
    if not ok:
        raise ValueError(reason)
    if len(lv) == 1:
        gamma_base = BUY_GAMMA_DEFAULT if side == "buy" else SELL_GAMMA_DEFAULT
        gamma_meta = compute_trend_gamma_adjustment(daily_profile, side=side, base_gamma=gamma_base)
        return pd.DataFrame([{
            "side": side, "interval_id": 0, "is_manager_level": True,
            "price": lv[0]["price"], "target_pos": lv[0]["target_pos"], "add_or_reduce_pos": 0.0,
            "source_upper_price": lv[0]["price"], "source_lower_price": lv[0]["price"],
            "source_upper_pos": lv[0]["target_pos"], "source_lower_pos": lv[0]["target_pos"],
            "gamma_base": float(gamma_meta.get("gamma_base", gamma_base)),
            "gamma_final": float(gamma_meta.get("gamma_final", gamma_base)),
            "gamma_multiplier": float(gamma_meta.get("gamma_multiplier", 1.0)),
            "trend_raw": float(gamma_meta.get("trend_raw", 0.0)),
            "trend_side": float(gamma_meta.get("trend_side", 0.0)),
            "trend_ret20_component": float(gamma_meta.get("trend_ret20_component", 0.0)),
            "trend_score_component": float(gamma_meta.get("trend_score_component", 0.0)),
            "trend_streak_component": float(gamma_meta.get("trend_streak_component", 0.0)),
        }])
    latest = daily_profile.iloc[-1]
    close = safe_float(latest.get("close", np.nan))
    atr = safe_float(latest.get("ATR", np.nan))
    if grid_intensity is None:
        grid_intensity = compute_grid_intensity(daily_profile, side, micro=micro)
    base_points = int(grid_intensity["base_points_per_interval"])
    micro_score = safe_float((micro or {}).get("micro_score", np.nan))
    if gamma is None:
        gamma = BUY_GAMMA_DEFAULT if side == "buy" else SELL_GAMMA_DEFAULT
    gamma_meta = compute_trend_gamma_adjustment(daily_profile, side=side, base_gamma=gamma)
    gamma_effective = float(gamma_meta.get("gamma_final", gamma))
    rows, seen = [], set()
    for i in range(len(lv) - 1):
        left, right = lv[i], lv[i + 1]
        adv20 = safe_float(latest.get("ADV_20", np.nan))
        n_inner = choose_points_for_interval(
            left["price"],
            right["price"],
            atr=atr,
            close=close,
            base_points=base_points,
            micro_score=micro_score,
            pos_a=left.get("target_pos", np.nan),
            pos_b=right.get("target_pos", np.nan),
            adv20=adv20,
            side=side,
            latest_features=latest,
            grid_intensity=grid_intensity,
        )
        prices = generate_interval_prices(left["price"], right["price"], side=side, n_inner=n_inner, method=method)
        for px in prices:
            key = round(float(px), 10)
            if key in seen:
                continue
            seen.add(key)
            raw_target = interpolate_target_pos(px, lv, side=side, gamma=gamma_effective, outside="clamp")
            target, terminal_guard_applied, terminal_progress_x, terminal_completion_cap = _apply_terminal_target_guard(
                px, raw_target, lv, side=side, trend_side=float(gamma_meta.get("trend_side", 0.0))
            )
            rows.append({
                "side": side,
                "interval_id": i + 1,
                "is_manager_level": any(abs(px - x["price"]) <= 1e-8 for x in lv),
                "price": float(px),
                "target_pos": float(target),
                "raw_target_pos_before_terminal_guard": float(raw_target),
                "terminal_guard_applied": bool(terminal_guard_applied),
                "terminal_progress_x": float(terminal_progress_x) if not pd.isna(terminal_progress_x) else np.nan,
                "terminal_completion_cap": float(terminal_completion_cap) if not pd.isna(terminal_completion_cap) else np.nan,
                "source_upper_price": float(left["price"]),
                "source_lower_price": float(right["price"]),
                "source_upper_pos": float(left["target_pos"]),
                "source_lower_pos": float(right["target_pos"]),
                "gamma_base": float(gamma_meta.get("gamma_base", gamma)),
                "gamma_final": float(gamma_effective),
                "gamma_multiplier": float(gamma_meta.get("gamma_multiplier", 1.0)),
                "trend_raw": float(gamma_meta.get("trend_raw", 0.0)),
                "trend_side": float(gamma_meta.get("trend_side", 0.0)),
                "trend_ret20_component": float(gamma_meta.get("trend_ret20_component", 0.0)),
                "trend_score_component": float(gamma_meta.get("trend_score_component", 0.0)),
                "trend_streak_component": float(gamma_meta.get("trend_streak_component", 0.0)),
            })
    out = pd.DataFrame(rows)
    if len(out) == 0:
        return out
    out = out.sort_values("price", ascending=(side == "sell")).reset_index(drop=True)
    if side == "buy":
        out = out.sort_values("price", ascending=False).reset_index(drop=True)
        out["add_or_reduce_pos"] = out["target_pos"].diff().fillna(0.0)
    else:
        out = out.sort_values("price", ascending=True).reset_index(drop=True)
        out["add_or_reduce_pos"] = out["target_pos"].diff().fillna(0.0)
    cols = [
        "side", "interval_id", "is_manager_level", "price", "target_pos", "raw_target_pos_before_terminal_guard",
        "terminal_guard_applied", "terminal_progress_x", "terminal_completion_cap", "add_or_reduce_pos",
        "source_upper_price", "source_lower_price", "source_upper_pos", "source_lower_pos",
        "gamma_base", "gamma_final", "gamma_multiplier", "trend_raw", "trend_side",
        "trend_ret20_component", "trend_score_component", "trend_streak_component",
    ]
    for c in cols:
        if c not in out.columns:
            out[c] = np.nan
    return out[cols]



def build_leader_readable_summary(code: str, asof_date: str, diag: Dict[str, Any], gi_buy: Optional[Dict[str, Any]], gi_sell: Optional[Dict[str, Any]], split_meta: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    metrics = diag.get("metrics", {}) or {}
    assessment = diag.get("assessment", {}) or {}
    label = assessment.get("suitability_label", metrics.get("suitability_label", "未判断"))
    score = safe_float(assessment.get("suitability_score", metrics.get("suitability_score", np.nan)))
    style = assessment.get("execution_style", metrics.get("execution_style", ""))
    reason = assessment.get("suitability_reason", metrics.get("suitability_reason", ""))
    density_parts = []
    if gi_buy:
        density_parts.append(f"买入侧 GI={gi_buy['GridIntensity']:.2f}，默认做网格；实际档数由价格距离/ATR/仓位变化/流动性/趋势/方向性价格偏离决定")
    if gi_sell:
        density_parts.append(f"卖出侧 GI={gi_sell['GridIntensity']:.2f}，默认做网格；实际档数由价格距离/ATR/仓位变化/流动性/趋势/方向性价格偏离决定")
    density_text = "；".join(density_parts) if density_parts else "未生成买卖侧网格密度"
    auto_parts = []
    if split_meta and split_meta.get("buy_anchor_split"):
        m = split_meta["buy_anchor_split"]
        auto_parts.append(f"买入只给两端锚点时，程序已按 ATR 和博士权重补出中间档：价格约 {m.get('price_mid', np.nan):.4f}，目标仓位约 {m.get('pos_mid', np.nan):.2%}。")
    if split_meta and split_meta.get("sell_anchor_split"):
        m = split_meta["sell_anchor_split"]
        auto_parts.append(f"卖出只给两端锚点时，程序已按 ATR 和博士权重补出中间档：价格约 {m.get('price_mid', np.nan):.4f}，目标仓位约 {m.get('pos_mid', np.nan):.2%}。")
    auto_text = "；".join(auto_parts)
    summary_text = "\n".join([
        f"【{code} | {asof_date}】",
        f"网格适用性：{label}（评分 {score:.2f}）。{style}",
        f"主要依据：{reason}",
        f"网格拆分：{density_text}。",
        auto_text,
    ]).strip()
    return {
        "code": code,
        "asof_date": asof_date,
        "suitability_label": label,
        "suitability_score": score,
        "execution_style": style,
        "main_reason": reason,
        "density_recommendation": density_text,
        "auto_split_comment": auto_text,
        "summary_text": summary_text,
    }


def add_execution_reasons(table: pd.DataFrame, side: str, diag: Dict[str, Any], gi: Optional[Dict[str, Any]]) -> pd.DataFrame:
    if table is None or len(table) == 0:
        return table
    out = table.copy()
    metrics = diag.get("metrics", {}) or {}
    style = metrics.get("execution_style", "")
    density = gi.get("density", "") if gi else ""
    base = gi.get("base_points_per_interval", np.nan) if gi else np.nan
    atr_pct = safe_float(metrics.get("ATR_pct", np.nan))
    rank60 = safe_float(metrics.get("close_rank_60d", np.nan))
    reason = f"{style}；当前 ATR_pct={atr_pct:.2%}，近60日价格分位约 {rank60:.0%}。GI={gi.get('GridIntensity', np.nan):.2f} 只作为适用性参考；默认做网格，具体档数由价格距离/ATR/仓位变化/流动性/趋势/方向性价格偏离决定。"
    if side == "buy":
        out["action"] = np.where(out["add_or_reduce_pos"] > 0, "加仓", "观察/锚点")
    else:
        out["action"] = np.where(out["add_or_reduce_pos"] < 0, "减仓", "观察/锚点")
    out["action_reason"] = reason
    return out


def build_execution_table(buy_table, sell_table, current_pos, code, asof_date, diag, gi_buy, gi_sell):
    parts = []
    for side, table, gi in [("buy", buy_table, gi_buy), ("sell", sell_table, gi_sell)]:
        if table is None or len(table) == 0:
            continue
        tmp = add_execution_reasons(table, side, diag, gi)
        tmp.insert(0, "code", code)
        tmp.insert(1, "asof_date", asof_date)
        tmp["current_pos"] = current_pos
        tmp["GridIntensity"] = gi["GridIntensity"] if gi else np.nan
        tmp["density"] = gi["density"] if gi else ""
        for k, v in (diag.get("metrics", {}) or {}).items():
            tmp[k] = v
        parts.append(tmp)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def evaluate_manager_grid_splitter(code: str, asof_date: str, manager_buy_levels: Optional[List[Dict[str, float]]], manager_sell_levels: Optional[List[Dict[str, float]]], current_pos: float, daily_df: Optional[pd.DataFrame] = None, intraday_df: Optional[pd.DataFrame] = None, lookback_years: int = 3, price_adj: str = DEFAULT_PRICE_ADJ, split_method: str = "linear", buy_gamma: float = BUY_GAMMA_DEFAULT, sell_gamma: float = SELL_GAMMA_DEFAULT, auto_split_buy_anchors: bool = False, auto_split_sell_anchors: Optional[bool] = None, print_detail: bool = True, daily_profile: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    code = norm_code(code)
    current_pos = float(current_pos)
    if daily_profile is not None:
        daily = daily_profile.copy().sort_values("date").reset_index(drop=True)
    else:
        daily_raw = fetch_wind_daily_ohlcva(code, end_date=asof_date, lookback_years=lookback_years, price_adj=price_adj) if daily_df is None else daily_df
        daily = prepare_daily_features(daily_raw)
    if len(daily) < 80:
        raise ValueError("valid daily data is too short after feature calculation")
    if intraday_df is None:
        try:
            minute_df = fetch_wind_intraday_minute(code, date=asof_date, bar_size=1, price_adj=price_adj)
        except Exception:
            minute_df = pd.DataFrame(columns=["date", "close", "volume", "amount"])
    else:
        minute_df = normalize_intraday_df(intraday_df)
    micro = compute_intraday_micro_features(minute_df)
    diag = summarize_stock_diagnostics(daily, micro=micro)
    gi_buy = gi_sell = buy_table = sell_table = None
    skipped, skip_reason = False, {}
    split_meta = {}
    if auto_split_sell_anchors is None:
        auto_split_sell_anchors = bool(auto_split_buy_anchors)
    if manager_buy_levels:
        if auto_split_buy_anchors and len(manager_buy_levels) == 2:
            ordered = sorted(manager_buy_levels, key=lambda x: float(x["price"]), reverse=True)
            try:
                manager_buy_levels, split_meta["buy_anchor_split"] = auto_split_anchored_buy_levels(current_pos, ordered[0]["price"], ordered[0]["target_pos"], ordered[1]["price"], ordered[1]["target_pos"], daily)
            except Exception as e:
                skipped = True; skip_reason["buy_auto_split"] = str(e)
        ok, reason, _ = validate_manager_levels_soft(manager_buy_levels, "buy")
        if ok:
            gi_buy = compute_grid_intensity(daily, "buy", micro=micro)
            buy_table = split_manager_levels(manager_buy_levels, "buy", daily, micro=micro, grid_intensity=gi_buy, method=split_method, gamma=buy_gamma)
        else:
            skipped = True; skip_reason["buy"] = reason
    if manager_sell_levels:
        if auto_split_sell_anchors and len(manager_sell_levels) == 2:
            ordered = sorted(manager_sell_levels, key=lambda x: float(x["price"]))
            try:
                manager_sell_levels, split_meta["sell_anchor_split"] = auto_split_anchored_sell_levels(current_pos, ordered[0]["price"], ordered[0]["target_pos"], ordered[1]["price"], ordered[1]["target_pos"], daily)
            except Exception as e:
                skipped = True; skip_reason["sell_auto_split"] = str(e)
        ok, reason = validate_sell_consistency(manager_sell_levels, current_pos)
        if ok:
            gi_sell = compute_grid_intensity(daily, "sell", micro=micro)
            sell_table = split_manager_levels(manager_sell_levels, "sell", daily, micro=micro, grid_intensity=gi_sell, method=split_method, gamma=sell_gamma)
        else:
            skipped = True; skip_reason["sell"] = reason
    execution_table = build_execution_table(buy_table, sell_table, current_pos, code, asof_date, diag, gi_buy, gi_sell)
    leader_summary = build_leader_readable_summary(code, asof_date, diag, gi_buy, gi_sell, split_meta)
    if print_detail:
        print(f"\n=== {code} | {asof_date} | market={guess_market(code)} ===")
        print("\n=== Leader Summary ===")
        print(leader_summary["summary_text"])
        print("\n=== Diagnostics ===")
        print(diag["text"])
        if gi_buy:
            print("\n=== GI BUY ===")
            print(pd.Series({"GridIntensity": gi_buy["GridIntensity"], "density": gi_buy["density"], "base_points_per_interval": gi_buy["base_points_per_interval"]}))
        if gi_sell:
            print("\n=== GI SELL ===")
            print(pd.Series({"GridIntensity": gi_sell["GridIntensity"], "density": gi_sell["density"], "base_points_per_interval": gi_sell["base_points_per_interval"]}))
        if skipped:
            print("\n=== SKIPPED / INVALID ===")
            print(skip_reason)
        if buy_table is not None:
            print("\n--- BUY execution table ---")
            print(buy_table.to_string(index=False))
        if sell_table is not None:
            print("\n--- SELL execution table ---")
            print(sell_table.to_string(index=False))
    return {"code": code, "asof_date": asof_date, "diagnostics": diag, "leader_summary": leader_summary, "daily_profile": daily, "minute_df": minute_df, "micro": micro, "GI_buy": gi_buy, "GI_sell": gi_sell, "buy_table": buy_table, "sell_table": sell_table, "execution_table": execution_table, "skipped": skipped, "skip_reason": skip_reason, "split_meta": split_meta}
