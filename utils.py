# -*- coding: utf-8 -*-
from __future__ import annotations
import numpy as np
import pandas as pd


def clip01(x: float) -> float:
    if pd.isna(x):
        return 0.0
    return float(np.clip(x, 0.0, 1.0))


def clip(x: float, lo: float, hi: float) -> float:
    if pd.isna(x):
        return float(lo)
    return float(np.clip(x, lo, hi))


def safe_float(x, default=np.nan) -> float:
    try:
        if pd.isna(x):
            return float(default)
        return float(x)
    except Exception:
        return float(default)


def norm_code(code: str) -> str:
    return str(code).strip().upper()


def guess_market(code: str) -> str:
    code = norm_code(code)
    if code.endswith(".HK"):
        return "HK"
    if code.endswith(".SH") or code.endswith(".SZ") or code.endswith(".BJ"):
        return "A"
    return "A"
