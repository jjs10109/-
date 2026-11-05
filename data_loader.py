# -*- coding: utf-8 -*-
from __future__ import annotations
import json
from pathlib import Path
from typing import Any, List, Tuple, Optional, Dict

import numpy as np
import pandas as pd

from config import DEFAULT_PRICE_ADJ
from utils import norm_code, guess_market


def ensure_wind() -> Any:
    try:
        from WindPy import w
    except Exception as e:
        raise RuntimeError("WindPy is not available. Please run this in an environment with Wind installed.") from e
    if not w.isconnected():
        w.start()
    return w



def wind_code_candidates(code: str) -> List[str]:
    """Return Wind-friendly code candidates while keeping project-internal codes unchanged.

    The live project stores HK codes in a 5-digit canonical form (e.g. 00941.HK)
    for stable matching across holdings/Oplus/instructions.  Some Wind terminals,
    including the user's environment, return data for the 4-digit form instead
    (e.g. 0941.HK).  For HK we therefore try the one-leading-zero-removed form
    first, then the canonical form, then the fully stripped numeric form.
    """
    c = norm_code(code)
    out: List[str] = []

    def add(x: str) -> None:
        x = norm_code(x)
        if x and x not in out:
            out.append(x)

    if c.endswith('.HK'):
        stem = c[:-3]
        if stem.isdigit():
            if len(stem) == 5 and stem.startswith('0'):
                add(f"{stem[1:]}.HK")
            add(c)
            stripped = stem.lstrip('0') or '0'
            add(f"{stripped}.HK")
            return out
    add(c)
    return out

def intraday_window(market: str, date: str) -> List[Tuple[str, str]]:
    if market == "A":
        return [(f"{date} 09:30:00", f"{date} 11:30:00"), (f"{date} 13:00:00", f"{date} 15:00:00")]
    return [(f"{date} 09:30:00", f"{date} 12:00:00"), (f"{date} 13:00:00", f"{date} 16:00:00")]


def normalize_daily_df(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    rename = {}
    for c in data.columns:
        lc = str(c).strip().lower()
        if lc in {"time", "datetime", "trade_dt"}:
            rename[c] = "date"
        elif lc == "amt":
            rename[c] = "amount"
        elif lc == "vol":
            rename[c] = "volume"
        else:
            rename[c] = lc
    data = data.rename(columns=rename)
    required = ["high", "low", "close", "volume", "amount"]
    missing = [c for c in required if c not in data.columns]
    if missing:
        raise ValueError(f"daily data missing columns: {missing}")
    if "date" not in data.columns:
        data["date"] = pd.RangeIndex(len(data))
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    for c in required:
        data[c] = pd.to_numeric(data[c], errors="coerce")
    data = data.dropna(subset=["date", "high", "low", "close"]).sort_values("date").reset_index(drop=True)
    return data[["date", "high", "low", "close", "volume", "amount"]]


def normalize_intraday_df(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=["date", "high", "low", "close", "volume", "amount"])
    data = df.copy()
    rename = {}
    for c in data.columns:
        lc = str(c).strip().lower()
        if lc in {"time", "datetime", "trade_dt"}:
            rename[c] = "date"
        elif lc == "amt":
            rename[c] = "amount"
        elif lc == "vol":
            rename[c] = "volume"
        else:
            rename[c] = lc
    data = data.rename(columns=rename)
    required = ["close", "volume", "amount"]
    if any(c not in data.columns for c in required):
        return pd.DataFrame(columns=["date", "high", "low", "close", "volume", "amount"])
    if "date" not in data.columns:
        data["date"] = pd.RangeIndex(len(data))
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    for c in ["high", "low", "close", "volume", "amount"]:
        if c in data.columns:
            data[c] = pd.to_numeric(data[c], errors="coerce")
    if "high" not in data.columns:
        data["high"] = data["close"]
    if "low" not in data.columns:
        data["low"] = data["close"]
    data = data.dropna(subset=["date", "close"]).sort_values("date").reset_index(drop=True)
    return data[["date", "high", "low", "close", "volume", "amount"]]


def fetch_wind_daily_ohlcva(code: str, end_date: str, lookback_years: int = 3, price_adj: str = DEFAULT_PRICE_ADJ) -> pd.DataFrame:
    w = ensure_wind()
    code = norm_code(code)
    end = pd.to_datetime(end_date)
    start = end - pd.DateOffset(years=lookback_years)
    opt = f"PriceAdj={price_adj}" if str(price_adj or "").strip() else ""
    attempts = []
    for wind_code in wind_code_candidates(code):
        r = w.wsd(wind_code, "high,low,close,volume,amt", start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"), opt)
        err = getattr(r, "ErrorCode", -1)
        if err != 0:
            attempts.append(f"{wind_code}: ErrorCode={err}")
            continue
        try:
            df = pd.DataFrame(np.array(r.Data).T, columns=["high", "low", "close", "volume", "amount"])
            df["date"] = pd.to_datetime(r.Times)
            df = normalize_daily_df(df)
        except Exception as exc:
            attempts.append(f"{wind_code}: parse={type(exc).__name__}")
            continue
        if not df.empty:
            df.attrs["wind_fetch_code"] = wind_code
            return df
        attempts.append(f"{wind_code}: empty")
    raise RuntimeError(
        f"Wind wsd failed/no data for {code}; tried {wind_code_candidates(code)}; details: {'; '.join(attempts)}"
    )



def fetch_wind_latest_raw_close(code: str, end_date: str, lookback_days: int = 15, return_fetch_code: bool = False):
    """Return the latest unadjusted close on/before end_date.

    For HK securities Wind code candidates are tried automatically, e.g.
    internal 00941.HK -> Wind 0941.HK first.  Set return_fetch_code=True to
    receive ``(price, wind_code_used)`` for audit output.
    """
    w = ensure_wind()
    code = norm_code(code)
    end = pd.to_datetime(end_date)
    start = end - pd.Timedelta(days=max(int(lookback_days), 1))
    attempts = []
    for wind_code in wind_code_candidates(code):
        r = w.wsd(
            wind_code,
            "close",
            start.strftime("%Y-%m-%d"),
            end.strftime("%Y-%m-%d"),
            "",
        )
        err = getattr(r, "ErrorCode", -1)
        if err != 0:
            attempts.append(f"{wind_code}: ErrorCode={err}")
            continue
        vals = []
        if getattr(r, "Data", None):
            try:
                vals = list(r.Data[0])
            except Exception:
                vals = []
        ss = pd.to_numeric(pd.Series(vals), errors="coerce").dropna()
        if not ss.empty:
            px = float(ss.iloc[-1])
            if np.isfinite(px) and px > 0:
                return (px, wind_code) if return_fetch_code else px
        attempts.append(f"{wind_code}: empty")
    if return_fetch_code:
        return float("nan"), ""
    return float("nan")


def fetch_wind_intraday_minute(code: str, date: str, bar_size: int = 1, price_adj: str = DEFAULT_PRICE_ADJ) -> pd.DataFrame:
    """Fetch Wind minute bars, trying HK Wind code variants automatically."""
    w = ensure_wind()
    code = norm_code(code)
    market = guess_market(code)
    parts = []
    opt = f"BarSize={bar_size}"
    if str(price_adj or "").strip():
        opt += f";PriceAdj={price_adj}"
    candidates = wind_code_candidates(code)
    for st, ed in intraday_window(market, date):
        tmp = None
        used_code = ""
        for wind_code in candidates:
            r = w.wsi(wind_code, "high,low,close,volume,amt", st, ed, opt)
            if getattr(r, "ErrorCode", -1) != 0:
                r = w.wsi(wind_code, "close,volume,amt", st, ed, opt)
                if getattr(r, "ErrorCode", -1) != 0:
                    continue
                candidate = pd.DataFrame(np.array(r.Data).T, columns=["close", "volume", "amount"])
            else:
                candidate = pd.DataFrame(np.array(r.Data).T, columns=["high", "low", "close", "volume", "amount"])
            candidate["date"] = pd.to_datetime(r.Times)
            candidate = normalize_intraday_df(candidate)
            if not candidate.empty:
                tmp = candidate
                used_code = wind_code
                break
        if tmp is not None and not tmp.empty:
            tmp.attrs["wind_fetch_code"] = used_code
            parts.append(tmp)
    if not parts:
        return pd.DataFrame(columns=["date", "high", "low", "close", "volume", "amount"])
    return normalize_intraday_df(pd.concat(parts, ignore_index=True))


def load_table(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    raise ValueError("file must be .xlsx, .xls, or .csv")


def parse_levels(s: Optional[str]) -> List[Dict[str, float]]:
    if s is None or str(s).strip() == "":
        return []
    s = str(s).strip()
    path = Path(s)
    if path.exists():
        if path.suffix.lower() == ".json":
            data = json.loads(path.read_text(encoding="utf-8"))
        else:
            df = load_table(path)
            data = df.to_dict("records")
    else:
        data = json.loads(s)
    return [{"price": float(x["price"]), "target_pos": float(x["target_pos"])} for x in data]
