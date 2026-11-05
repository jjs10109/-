# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Dict
import numpy as np
import pandas as pd

from data_loader import normalize_intraday_df
from utils import clip01


def compute_intraday_micro_features(min_df: pd.DataFrame) -> Dict[str, float]:
    empty = {"micro_score": np.nan, "intraday_vol": np.nan, "intraday_trend": np.nan, "jump_q95": np.nan, "illiq_proxy": np.nan}
    if min_df is None or len(min_df) < 60:
        return empty
    df = normalize_intraday_df(min_df)
    if len(df) < 60:
        return empty
    df["ret"] = np.log(df["close"] / df["close"].shift(1))
    r = df["ret"].dropna()
    if len(r) < 30:
        return {"micro_score": 0.5, "intraday_vol": np.nan, "intraday_trend": np.nan, "jump_q95": np.nan, "illiq_proxy": np.nan}
    intraday_vol = float(r.std())
    jump_q95 = float(r.abs().quantile(0.95))
    y = df["close"].values.astype(float)
    x = np.arange(len(df), dtype=float)
    mask = ~np.isnan(y)
    if mask.sum() > 30:
        x2 = x[mask]
        y2 = y[mask]
        x2 = x2 - x2.mean()
        denom = float((x2 ** 2).sum())
        intraday_trend = float((x2 * (y2 - y2.mean())).sum() / denom) if denom != 0 else np.nan
    else:
        intraday_trend = np.nan
    if df["amount"].notna().sum() > 30:
        a = df["amount"].abs()
        rabs = df["ret"].abs()
        threshold = a.quantile(0.10)
        mask2 = (a > max(threshold, 1e-6)) & rabs.notna()
        illiq = (rabs[mask2] / (a[mask2] + 1e-12)).replace([np.inf, -np.inf], np.nan).dropna()
        if len(illiq) > 0:
            illiq_proxy = float(illiq.clip(lower=illiq.quantile(0.05), upper=illiq.quantile(0.95)).tail(240).mean())
        else:
            illiq_proxy = np.nan
    else:
        illiq_proxy = np.nan

    def sfr(v, lo, hi):
        if np.isnan(v):
            return 0.5
        if v <= lo:
            return 1.0
        if v >= hi:
            return 0.0
        return clip01(1 - (v - lo) / (hi - lo))

    micro_score = clip01(
        0.35 * sfr(intraday_vol, 0.0003, 0.0020)
        + 0.25 * sfr(jump_q95, 0.0006, 0.0040)
        + 0.25 * sfr(abs(intraday_trend) if not np.isnan(intraday_trend) else np.nan, 0.0, 0.02)
        + 0.15 * sfr(illiq_proxy, 1e-13, 5e-12)
    )
    return {"micro_score": micro_score, "intraday_vol": intraday_vol, "intraday_trend": intraday_trend, "jump_q95": jump_q95, "illiq_proxy": illiq_proxy}
