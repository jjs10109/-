# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Sequence, Dict, Any, Optional
import numpy as np
import pandas as pd

from config import ATR_N, MA_WINDOW, TREND_WINDOW, UP_RATIO_WINDOW, ADV_SHORT_WINDOW, ADV_LONG_WINDOW, VOL_WINDOWS
from data_loader import normalize_daily_df
from utils import safe_float


def build_profile_daily(ts: pd.DataFrame, atr_n: int = ATR_N) -> pd.DataFrame:
    df = normalize_daily_df(ts)
    prev_close = df["close"].shift(1)
    tr = np.maximum.reduce([
        (df["high"] - df["low"]).values,
        (df["high"] - prev_close).abs().values,
        (df["low"] - prev_close).abs().values,
    ])
    df["ATR"] = pd.Series(tr).rolling(atr_n).mean()
    df["ATR_pct"] = df["ATR"] / df["close"]
    df["log_ret"] = np.log(df["close"] / df["close"].shift(1))
    df["ADV_20"] = df["amount"].rolling(ADV_SHORT_WINDOW).mean()
    df["ADV_60"] = df["amount"].rolling(ADV_LONG_WINDOW).mean()
    mean60 = df["amount"].rolling(ADV_LONG_WINDOW).mean()
    std60 = df["amount"].rolling(ADV_LONG_WINDOW).std()
    df["ADV_CV_60"] = std60 / (mean60 + 1e-12)
    return df


def compute_streak_series(signs: np.ndarray) -> np.ndarray:
    streak = np.zeros(len(signs), dtype=float)
    cur = 0
    cur_sign = 0
    for i, s in enumerate(signs):
        if np.isnan(s) or s == 0:
            cur = 0
            cur_sign = 0
            streak[i] = 0
        elif s == cur_sign:
            cur += 1
            streak[i] = cur
        else:
            cur_sign = s
            cur = 1
            streak[i] = cur
    return streak


def add_trend_metrics(df: pd.DataFrame, slope_window: int = TREND_WINDOW, ma_window: int = MA_WINDOW, up_ratio_window: int = UP_RATIO_WINDOW) -> pd.DataFrame:
    out = df.copy().sort_values("date").reset_index(drop=True)
    ret_diff = out["close"].diff()
    signs = np.sign(ret_diff.values)
    up = (ret_diff > 0).astype(int)
    out[f"MA_{ma_window}"] = out["close"].rolling(ma_window).mean()
    out[f"UpRatio_{up_ratio_window}"] = up.rolling(up_ratio_window).mean()
    out["streak_len"] = compute_streak_series(signs)
    out[f"Deviation_{ma_window}"] = (out["close"] - out[f"MA_{ma_window}"]) / out["ATR"]
    N = slope_window
    x = np.arange(N, dtype=float)
    xd = x - x.mean()
    denom = float((xd ** 2).sum())
    scores = [np.nan] * len(out)
    closes = out["close"].values.astype(float)
    atrs = out["ATR"].values.astype(float)
    for i in range(N - 1, len(out)):
        y = closes[i - N + 1 : i + 1]
        if np.any(np.isnan(y)):
            continue
        slope = float(np.sum(xd * (y - y.mean())) / denom)
        atr_now = atrs[i]
        if not np.isnan(atr_now) and atr_now != 0:
            scores[i] = slope / atr_now
    out[f"TrendScore_{slope_window}"] = scores
    return out


def add_volatility(df: pd.DataFrame, windows: Sequence[int] = VOL_WINDOWS) -> pd.DataFrame:
    out = df.copy().sort_values("date").reset_index(drop=True)
    out["log_ret"] = np.log(out["close"] / out["close"].shift(1))
    for w in windows:
        out[f"vol_{w}"] = out["log_ret"].rolling(w).std() * np.sqrt(252)
    if "vol_60" in out.columns and "vol_252" in out.columns:
        out["vol_regime"] = (out["vol_60"] / out["vol_252"]).replace([np.inf, -np.inf], np.nan)
    return out


def add_price_range_features(df: pd.DataFrame, window: int = 60, price_col: str = "close") -> pd.DataFrame:
    out = df.copy().sort_values("date").reset_index(drop=True)
    price = pd.to_numeric(out[price_col], errors="coerce")
    roll = price.rolling(window)

    high_n = roll.max()
    low_n = roll.min()
    out[f"High_{window}"] = high_n
    out[f"Low_{window}"] = low_n
    out[f"RangeRatio_{window}"] = ((high_n - low_n) / price).replace([np.inf, -np.inf], np.nan)

    # Final optimization: avoid the explicit Python for-loop for rolling
    # quantiles.  These use the same rolling window as the original code, but
    # let pandas do the heavy lifting in vectorized/compiled code.
    out[f"q05_{window}"] = roll.quantile(0.05)
    out[f"q50_{window}"] = roll.quantile(0.50)
    out[f"q95_{window}"] = roll.quantile(0.95)

    # Rolling rank of the current close inside the trailing window.  This is the
    # same definition as before: mean(recent <= current), including current day.
    out[f"rank_{window}"] = roll.apply(
        lambda x: float(np.nanmean(x <= x[-1])) if len(x) and np.isfinite(x[-1]) else np.nan,
        raw=True,
    )
    return out


def prepare_daily_features(daily_raw: pd.DataFrame) -> pd.DataFrame:
    daily = build_profile_daily(daily_raw, atr_n=ATR_N)
    daily = add_trend_metrics(daily, slope_window=TREND_WINDOW, ma_window=MA_WINDOW, up_ratio_window=UP_RATIO_WINDOW)
    daily = add_volatility(daily, windows=VOL_WINDOWS)
    daily = add_price_range_features(daily, window=60)
    return daily.dropna(subset=["ATR", "ATR_pct"]).reset_index(drop=True)



def _clip01(x: float) -> float:
    try:
        return float(np.clip(x, 0.0, 1.0))
    except Exception:
        return 0.5


def assess_grid_suitability(metrics: Dict[str, float]) -> Dict[str, Any]:
    """默认做网格的诊断口径。

    这里的 score 不再用于决定“做/不做”，而是用于解释网格应偏积极、正常还是保守。
    真正不做网格只由硬性条件决定：指令无效、仓位无变化、价格区间过窄、数据缺失等。
    """
    atr_pct = safe_float(metrics.get("ATR_pct", np.nan))
    trend = safe_float(metrics.get("TrendScore_60", np.nan))
    up20 = safe_float(metrics.get("UpRatio_20", np.nan))
    streak_len = safe_float(metrics.get("streak_len", np.nan))
    deviation_60 = safe_float(metrics.get("Deviation_60", np.nan))
    adv20 = safe_float(metrics.get("ADV_20", np.nan))
    adv_cv = safe_float(metrics.get("ADV_CV_60", np.nan))

    # 流动性不足：越低越需要小单拆分。
    if np.isnan(adv20) or adv20 <= 0:
        liq_good = 0.40
    else:
        liq_good = _clip01((np.log(max(adv20, 1.0)) - np.log(5e7)) / (np.log(1e9) - np.log(5e7)))
    liquidity_need = 1.0 - liq_good

    # 波动空间：波动偏高说明有网格触发空间，但后续密度会用 ATR 拉宽，不会机械加密。
    volatility_space = 0.50 if np.isnan(atr_pct) else _clip01((atr_pct - 0.012) / (0.080 - 0.012))

    # 趋势风险：历史趋势经常性 + 最近趋势可能性 + 最近已有趋势。
    trend_frequency = 0.50 if np.isnan(trend) else _clip01(abs(trend) / 0.12)
    trend_recent_probability = 0.50 if np.isnan(up20) else _clip01(abs(up20 - 0.50) / 0.35)
    streak_score = 0.30 if np.isnan(streak_len) else _clip01((streak_len - 2.0) / 10.0)
    deviation_score = 0.30 if np.isnan(deviation_60) else _clip01(abs(deviation_60) / 3.0)
    trend_existing = max(streak_score, deviation_score)
    trend_risk = _clip01(0.35 * trend_frequency + 0.35 * trend_recent_probability + 0.30 * trend_existing)
    low_trend_score = 1.0 - trend_risk

    # 这里无法区分买卖方向，所以用绝对偏离表示“存在可利用价格偏离”。
    price_deviation_score = 0.50 if np.isnan(deviation_60) else _clip01((abs(deviation_60) - 0.30) / (2.50 - 0.30))

    score = _clip01(
        0.25 * liquidity_need
        + 0.25 * volatility_space
        + 0.25 * price_deviation_score
        + 0.25 * low_trend_score
    )

    if score >= 0.65:
        label = "默认做网格｜偏积极"
        style = "可以相对积极地分批，提高换手率；仍保留经理原始锚点。"
    elif score >= 0.45:
        label = "默认做网格｜正常"
        style = "正常拆分；趋势和波动只调节疏密，不取消网格。"
    else:
        label = "默认做网格｜偏保守"
        style = "保守拆分、网格偏稀、仓位后置；除非硬性条件不满足，否则仍做网格。"

    reasons = []
    reasons.append(f"流动性不足分={liquidity_need:.2f}，用于控制单档仓位大小")
    reasons.append(f"波动空间分={volatility_space:.2f}，波动越高越有触发空间但价格步长会拉宽")
    reasons.append(f"价格偏离分={price_deviation_score:.2f}，买卖侧会再按方向判断")
    reasons.append(f"趋势风险={trend_risk:.2f}，趋势越强网格越稀但默认不取消")
    if not np.isnan(adv_cv) and adv_cv > 1.2:
        reasons.append("成交额波动较大，单档交易量要更小")

    return {
        "suitability_score": score,
        "suitability_label": label,
        "execution_style": style,
        "suitability_reason": "；".join(reasons),
        "volatility_score": float(volatility_space),
        "trend_score_for_grid": float(low_trend_score),
        "price_location_score": float(price_deviation_score),
        "liquidity_score": float(liq_good),
        "liquidity_need_score": float(liquidity_need),
        "trend_risk": float(trend_risk),
        "default_grid_enabled": True,
    }

def summarize_stock_diagnostics(daily: pd.DataFrame, micro: Optional[Dict[str, float]] = None, lookback_dist_days: int = 60) -> Dict[str, Any]:
    d = daily.dropna(subset=["close"]).copy().sort_values("date").reset_index(drop=True)
    if len(d) < 80:
        return {"text": "历史数据太少，无法稳定评估诊断指标。", "metrics": {}}
    latest = d.iloc[-1]
    metrics = {
        "close": safe_float(latest.get("close", np.nan)),
        "ATR_pct": safe_float(latest.get("ATR_pct", np.nan)),
        "vol_60": safe_float(latest.get("vol_60", np.nan)),
        "vol_252": safe_float(latest.get("vol_252", np.nan)),
        "vol_regime": safe_float(latest.get("vol_regime", np.nan)),
        "TrendScore_60": safe_float(latest.get("TrendScore_60", np.nan)),
        "UpRatio_20": safe_float(latest.get("UpRatio_20", np.nan)),
        "streak_len": safe_float(latest.get("streak_len", np.nan)),
        "Deviation_60": safe_float(latest.get("Deviation_60", np.nan)),
        "ADV_20": safe_float(latest.get("ADV_20", np.nan)),
        "ADV_60": safe_float(latest.get("ADV_60", np.nan)),
        "ADV_CV_60": safe_float(latest.get("ADV_CV_60", np.nan)),
        "intraday_vol": safe_float((micro or {}).get("intraday_vol", np.nan)),
        "illiq_proxy_1d": safe_float((micro or {}).get("illiq_proxy", np.nan)),
        "q05_60d": safe_float(latest.get("q05_60", np.nan)),
        "q50_60d": safe_float(latest.get("q50_60", np.nan)),
        "q95_60d": safe_float(latest.get("q95_60", np.nan)),
        "close_rank_60d": safe_float(latest.get("rank_60", np.nan)),
        "High_60": safe_float(latest.get("High_60", np.nan)),
        "Low_60": safe_float(latest.get("Low_60", np.nan)),
        "RangeRatio_60": safe_float(latest.get("RangeRatio_60", np.nan)),
    }
    assessment = assess_grid_suitability(metrics)
    metrics.update(assessment)
    text = [
        "【网格适用性结论】",
        f"- {assessment['suitability_label']}，评分={assessment['suitability_score']:.2f}；{assessment['execution_style']}",
        f"- 原因：{assessment['suitability_reason']}",
        "【波动】",
        f"- ATR_pct={metrics['ATR_pct']:.2%}, vol_60={metrics['vol_60']:.2f}, vol_252={metrics['vol_252']:.2f}, vol_regime={metrics['vol_regime']:.2f}",
        "【趋势/价格行为】",
        f"- TrendScore_60={metrics['TrendScore_60']:.4f}, UpRatio_20={metrics['UpRatio_20']:.2f}, streak_len={metrics['streak_len']:.0f}, Deviation_60={metrics['Deviation_60']:.2f}",
        "【流动性/冲击】",
        f"- ADV_20={metrics['ADV_20']:,.0f}, ADV_60={metrics['ADV_60']:,.0f}, ADV_CV_60={metrics['ADV_CV_60']:.2f}",
        f"- intraday_vol={metrics['intraday_vol']:.3e}, illiq_proxy={metrics['illiq_proxy_1d']:.3e}",
        f"【价格分布（近{lookback_dist_days}日）】",
        f"- q05={metrics['q05_60d']:.2f}, q50={metrics['q50_60d']:.2f}, q95={metrics['q95_60d']:.2f}, close={metrics['close']:.2f}, rank≈{metrics['close_rank_60d']:.0%}",
    ]
    return {"metrics": metrics, "assessment": assessment, "text": "\n".join(text)}
