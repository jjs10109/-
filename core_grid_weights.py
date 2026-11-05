# -*- coding: utf-8 -*-
from __future__ import annotations
from math import inf
from typing import Dict, Sequence, Tuple
import numpy as np
import pandas as pd

from config import B_LEVELS, CORE_STEP, MAX_SUB_WEIGHT, MAX_TOTAL_WEIGHT, COST_RATE, MA_WINDOW, ATR_N
from features_daily import build_profile_daily


def prepare_core_grid_df(df: pd.DataFrame, b_levels: Sequence[float] = B_LEVELS) -> pd.DataFrame:
    """准备博士权重学习需要的 MA、ATR、三层触发状态。"""
    prepared = build_profile_daily(df, atr_n=ATR_N).copy().sort_values("date").reset_index(drop=True)
    prepared[f"MA_{MA_WINDOW}"] = prepared["close"].rolling(MA_WINDOW).mean()
    prepared = build_states(prepared, b_levels=b_levels, ma_col=f"MA_{MA_WINDOW}", atr_col="ATR")
    return prepared.dropna(subset=["close", "ATR", f"MA_{MA_WINDOW}"]).reset_index(drop=True)


def build_states(df: pd.DataFrame, b_levels: Sequence[float] = B_LEVELS, ma_col: str = f"MA_{MA_WINDOW}", atr_col: str = "ATR") -> pd.DataFrame:
    out = df.copy().sort_values("date").reset_index(drop=True)
    for j, b in enumerate(b_levels, start=1):
        state = []
        cur = 0
        for _, row in out.iterrows():
            close = row.get("close", np.nan)
            ma = row.get(ma_col, np.nan)
            atr = row.get(atr_col, np.nan)
            if np.isnan(close) or np.isnan(ma) or np.isnan(atr):
                cur = 0
            elif close > ma:
                cur = 0
            elif close <= ma - b * atr:
                cur = 1
            state.append(cur)
        out[f"s{j}"] = state
    return out


def candidate_weights(step: float = CORE_STEP):
    vals = np.arange(0.0, MAX_SUB_WEIGHT + 1e-9, step)
    for p1 in vals:
        for p2 in vals:
            for p3 in vals:
                if p1 + p2 + p3 <= MAX_TOTAL_WEIGHT + 1e-9:
                    yield (float(p1), float(p2), float(p3))


def evaluate_with_weights(df: pd.DataFrame, weights: Tuple[float, float, float]) -> pd.DataFrame:
    p1, p2, p3 = weights
    out = df.copy()
    out["pos_target"] = p1 * out["s1"] + p2 * out["s2"] + p3 * out["s3"]
    out["ret_cc"] = out["close"].pct_change()
    out["cost"] = COST_RATE * (out["pos_target"] - out["pos_target"].shift(1)).abs().fillna(0)
    out["strategy_ret"] = out["pos_target"].shift(1).fillna(0) * out["ret_cc"].fillna(0) - out["cost"]
    return out


def calc_return_and_dd(strategy_ret: pd.Series) -> Tuple[float, float, float]:
    nav = (1 + strategy_ret.fillna(0)).cumprod()
    if len(nav) == 0:
        return 0.0, 0.0, -inf
    total_ret = float(nav.iloc[-1] - 1)
    peak = nav.cummax()
    drawdown = nav / peak - 1
    max_dd = float(drawdown.min())
    score = total_ret / abs(max_dd) if abs(max_dd) > 1e-12 else total_ret
    return total_ret, max_dd, float(score)


def optimize_weights(train_df: pd.DataFrame, step: float = CORE_STEP) -> Tuple[Tuple[float, float, float], float, float, float]:
    best_w = (0.0, 0.0, 0.0)
    best_score = -inf
    best_ret, best_dd = 0.0, 0.0
    for w in candidate_weights(step):
        tmp = evaluate_with_weights(train_df, w)
        total_ret, max_dd, score = calc_return_and_dd(tmp["strategy_ret"])
        if score > best_score:
            best_score = score
            best_w = w
            best_ret = total_ret
            best_dd = max_dd
    return best_w, best_ret, best_dd, best_score


def fit_core_grid_weights(df: pd.DataFrame, b_levels: Sequence[float] = B_LEVELS) -> Dict[str, object]:
    prepared = prepare_core_grid_df(df, b_levels=b_levels)
    if len(prepared) < 120:
        raise ValueError(f"Core-grid history too short after MA/ATR warmup: {len(prepared)} rows")
    weights, train_ret, train_dd, train_score = optimize_weights(prepared)
    weight_sum = float(sum(weights))
    fit_valid = bool(weight_sum > 1e-12)
    normalized = tuple(float(x / weight_sum) for x in weights) if fit_valid else (0.0, 0.0, 0.0)
    p2, p3 = weights[1], weights[2]
    if p2 + p3 > 1e-12:
        w2_deep = float(p2 / (p2 + p3))
        w3_deep = float(p3 / (p2 + p3))
    else:
        w2_deep, w3_deep = 0.5, 0.5
    latest = prepared.iloc[-1]
    trigger_span = float((max(b_levels) - min(b_levels)) * latest["ATR"]) if pd.notna(latest["ATR"]) else np.nan
    return {
        "weights": weights,
        "normalized_weights": normalized,
        "fit_valid": fit_valid,
        "train_return": train_ret,
        "train_max_drawdown": train_dd,
        "train_score": train_score,
        "w2_deep": w2_deep,
        "w3_deep": w3_deep,
        "latest_atr": float(latest["ATR"]),
        "trigger_span": trigger_span,
        "prepared": prepared,
    }
