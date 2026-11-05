# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Any, Dict, Optional, Mapping
import numpy as np
import pandas as pd

from config import (
    DEFAULT_ADV_MIN, DENSITY_DENSE, DENSITY_MEDIUM,
)
from utils import clip01, safe_float


def _get(d: Mapping[str, Any], key: str, default: float = np.nan) -> float:
    try:
        return safe_float(d.get(key, default), default)
    except Exception:
        return default


def score_liquidity_good(adv20: float, adv_low: float = DEFAULT_ADV_MIN, adv_high: float = 1e9) -> float:
    """0~1，越大表示成交额越好。"""
    adv20 = safe_float(adv20, np.nan)
    if np.isnan(adv20) or adv20 <= 0:
        return 0.40
    return clip01((np.log(max(adv20, 1.0)) - np.log(adv_low)) / (np.log(adv_high) - np.log(adv_low)))


def score_liquidity_need(adv20: float, adv_low: float = DEFAULT_ADV_MIN, adv_high: float = 1e9) -> float:
    """0~1，越大表示越需要用小单网格改善流动性执行。"""
    return clip01(1.0 - score_liquidity_good(adv20, adv_low=adv_low, adv_high=adv_high))


def score_volatility_space(atr_pct: float) -> float:
    """0~1，越大表示价格有网格触发空间。低波动触发机会少；中高波动触发空间大。"""
    atr_pct = safe_float(atr_pct, np.nan)
    if np.isnan(atr_pct):
        return 0.50
    return clip01((atr_pct - 0.012) / (0.080 - 0.012))


def score_volatility_high_for_spacing(atr_pct: float) -> float:
    """0~1，越大表示波动已经偏高，价格间距应拉宽、网格略稀。"""
    atr_pct = safe_float(atr_pct, np.nan)
    if np.isnan(atr_pct):
        return 0.30
    return clip01((atr_pct - 0.040) / (0.120 - 0.040))


def _deviation_z(features: Mapping[str, Any]) -> float:
    close = _get(features, "close")
    ma60 = _get(features, "MA_60")
    atr = _get(features, "ATR")
    dev60 = _get(features, "Deviation_60")
    if not np.isnan(dev60):
        return float(dev60)
    if not np.isnan(close) and not np.isnan(ma60) and not np.isnan(atr) and atr > 0:
        return float((close - ma60) / atr)
    q50 = _get(features, "q50_60")
    q05 = _get(features, "q05_60")
    q95 = _get(features, "q95_60")
    if not np.isnan(close) and not np.isnan(q50) and not np.isnan(q05) and not np.isnan(q95) and q95 > q05:
        return float((close - q50) / ((q95 - q05) / 3.29))
    return 0.0


def score_side_deviation(features: Mapping[str, Any], side: str) -> Dict[str, float]:
    """价格偏离的方向性评分。

    sell：价格高于正常水平是有利偏离；buy：价格低于正常水平是有利偏离。
    """
    z = _deviation_z(features)
    signed = z if side == "sell" else -z
    favorable = clip01((signed - 0.30) / (2.50 - 0.30))
    unfavorable = clip01((-signed - 0.30) / (2.50 - 0.30))
    neutral_score = clip01(0.50 + 0.50 * favorable - 0.40 * unfavorable)
    return {
        "deviation_z": float(z),
        "side_signed_deviation": float(signed),
        "side_favorable_deviation": float(favorable),
        "side_unfavorable_deviation": float(unfavorable),
        "side_deviation_score": float(neutral_score),
    }


def score_trend_components(features: Mapping[str, Any]) -> Dict[str, float]:
    """趋势风险拆成三部分：历史趋势经常性、最近趋势可能性、最近已有趋势。"""
    ts = _get(features, "TrendScore_60")
    up20 = _get(features, "UpRatio_20")
    streak_len = _get(features, "streak_len")
    dev_z = abs(_deviation_z(features))

    trend_frequency = clip01(abs(ts) / 0.12) if not np.isnan(ts) else 0.50
    trend_recent_probability = clip01(abs(up20 - 0.50) / 0.35) if not np.isnan(up20) else 0.50
    streak_score = clip01((streak_len - 2.0) / 10.0) if not np.isnan(streak_len) else 0.30
    deviation_score = clip01(dev_z / 3.0) if not np.isnan(dev_z) else 0.30
    trend_existing = max(streak_score, deviation_score)

    trend_risk = clip01(0.35 * trend_frequency + 0.35 * trend_recent_probability + 0.30 * trend_existing)
    return {
        "trend_frequency": float(trend_frequency),
        "trend_recent_probability": float(trend_recent_probability),
        "trend_existing": float(trend_existing),
        "trend_risk": float(trend_risk),
        "low_trend_score": float(1.0 - trend_risk),
    }


def compute_grid_components(features: Mapping[str, Any], side: str, adv_min: float = DEFAULT_ADV_MIN) -> Dict[str, float]:
    """统一计算网格适用性和密度调节所需组件。"""
    atr_pct = _get(features, "ATR_pct")
    adv20 = _get(features, "ADV_20")
    liq_good = score_liquidity_good(adv20, adv_low=adv_min, adv_high=1e9)
    liq_need = score_liquidity_need(adv20, adv_low=adv_min, adv_high=1e9)
    vol_space = score_volatility_space(atr_pct)
    vol_high = score_volatility_high_for_spacing(atr_pct)
    dev = score_side_deviation(features, side)
    trend = score_trend_components(features)
    applicability = clip01(
        0.25 * liq_need
        + 0.25 * vol_space
        + 0.25 * dev["side_deviation_score"]
        + 0.25 * trend["low_trend_score"]
    )
    out = {
        "liquidity_good_score": float(liq_good),
        "liquidity_need_score": float(liq_need),
        "volatility_space_score": float(vol_space),
        "volatility_high_spacing_score": float(vol_high),
        "grid_applicability_score": float(applicability),
    }
    out.update(dev)
    out.update(trend)
    return out


def compute_grid_intensity(df: pd.DataFrame, side: str, micro: Optional[Dict[str, float]] = None, adv_min: float = DEFAULT_ADV_MIN) -> Dict[str, Any]:
    """计算网格适用性。

    v2-default-grid 口径：默认所有可解析、仓位有变化空间的股票都做网格。
    这个分数不再决定是否做，也不直接决定每段拆几个；它只用于解释和密度/仓位分布调节。
    """
    if side not in ("buy", "sell"):
        raise ValueError("side must be buy or sell")
    latest = df.copy().sort_values("date").reset_index(drop=True).iloc[-1]
    components = compute_grid_components(latest, side, adv_min=adv_min)
    raw_gi = components["grid_applicability_score"]

    # 保留旧字段 GridIntensity，方便下游输出；这里不再做极端拉伸影响拆分。
    gi = clip01(raw_gi)

    if gi >= DENSITY_DENSE:
        density = "默认网格｜积极"
    elif gi >= DENSITY_MEDIUM:
        density = "默认网格｜正常"
    else:
        density = "默认网格｜保守"

    return {
        "GridIntensity": gi,
        "raw_grid_intensity": raw_gi,
        "density": density,
        # 兼容旧输出。真实档数在 grid_engine.choose_points_for_interval 中决定。
        "base_points_per_interval": 1,
        "grid_count_driver": "default_grid_price_atr_position_liquidity_trend_side_deviation",
        "default_grid_enabled": True,
        "components": {
            **components,
            "raw_gi": raw_gi,
            "stretched_gi": gi,
            "w_liquidity_need": 0.25,
            "w_volatility_space": 0.25,
            "w_side_deviation": 0.25,
            "w_low_trend": 0.25,
        },
    }
