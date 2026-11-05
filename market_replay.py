# -*- coding: utf-8 -*-
"""
股数增量追踪 + A股反推净资产 + Wind盯市，重建每日（含1012）真实仓位。

- 股数：按指令序号增量回放（已验证完全准确）；
- 净资产：每个交易日用 A股指令  指令数量*指令价格/ratio 反推，多只取中位数；
- 盯市：Wind 每日收盘价（港股用港币汇率），仓位=市值/净资产；
- 用 1014 真实持仓对账验证。
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path
from typing import Dict
import numpy as np
import pandas as pd
from WindPy import w

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import live_grid_simulator_current as m
from incremental_replay import extract_orders


def parse_args():
    ap = argparse.ArgumentParser(description="股数增量追踪 + A股反推净资产 + Wind盯市，重建每日持仓")
    ap.add_argument("--base-date", default="2026-09-30",
                    help="基线交易日 YYYY-MM-DD（默认 2026-09-30）")
    ap.add_argument("--base-position-file", default="持仓数据/全球产品持仓查询_20261008.xls",
                    help="基线持仓文件（相对项目根或绝对路径，默认 持仓数据/全球产品持仓查询_20261008.xls）")
    ap.add_argument("--end-date", default=None,
                    help="增量截止交易日 YYYY-MM-DD，默认到最新快照")
    return ap.parse_args()


def discover_days_files(base_date: str, end_date: str | None):
    """自动扫描 oplus交易文件，返回 (DAYS, FILES)：从 base_date 起到 end_date（含），
    按真实业务日期排序，同一天只保留一个快照（文件名排序最后的，即最新导出）。"""
    items = []
    for p in sorted((ROOT / "oplus交易文件").glob("oplus*.xlsx")):
        raw = m.read_excel_any(p, dtype={"证券代码": str})
        if isinstance(raw, dict):
            raw = next(iter(raw.values()))
        dc = m._find_col(raw.columns, ["业务日期"])
        d = pd.to_datetime(raw[dc], errors="coerce").max() if dc else pd.NaT
        if pd.isna(d):
            continue
        items.append((d.strftime("%Y-%m-%d"), p.name))
    by_date = {}
    for d, f in items:
        if d >= base_date and (end_date is None or d <= end_date):
            by_date[d] = f  # 文件名已排序，后者覆盖
    ordered = sorted(by_date.items())
    return [d for d, _ in ordered], [f for _, f in ordered]


def discover_real_tags():
    """扫描 持仓数据/全球产品持仓查询_*.xls*，返回 {内部日期: 文件名标签}，供逐日对账自动对齐。"""
    out = {}
    for p in sorted((ROOT / "持仓数据").glob("全球产品持仓查询_*.xls*")):
        try:
            raw = m.read_excel_any(p)
        except Exception:
            continue
        pc = m._find_col(raw.columns, m.PRODUCT_COL_CANDIDATES)
        if pc is None:
            continue
        sub = raw[raw[pc].apply(m.is_sim_product)]
        if "日期" not in sub.columns:
            continue
        for d in sub["日期"].dropna().unique():
            ds = str(d)[:10]
            out[ds] = p.name
    return out


def infer_nav_from_a(raw: pd.DataFrame) -> float:
    cc = m._find_col(raw.columns, m.CODE_COL_CANDIDATES)
    pc = m._find_col(raw.columns, m.PRODUCT_COL_CANDIDATES)
    mkc = m._find_col(raw.columns, m.MARKET_COL_CANDIDATES)
    ratio_c = m._find_col(raw.columns, ["占净资产比例"])
    qty_c = m._find_col(raw.columns, ["指令数量"])
    px_c = m._find_col(raw.columns, ["指令价格"])
    vals = []
    for _, r in raw.iterrows():
        if not m.is_sim_product(r.get(pc)):
            continue
        code = m.canonical_code(r.get(cc), market=r.get(mkc, "") if mkc else "")
        if not m.market_group(code) == "A":
            continue
        qty = m._safe_float(r.get(qty_c), np.nan); px = m._safe_float(r.get(px_c), np.nan)
        ratio = m._safe_float(r.get(ratio_c), np.nan)
        if np.isfinite(ratio) and ratio > 1:
            ratio /= 100
        if np.isfinite(qty) and np.isfinite(px) and np.isfinite(ratio) and qty > 0 and px > 0 and ratio > 0:
            vals.append(qty * px / ratio)
    if len(vals) < 2:
        return np.nan
    arr = np.array(vals)
    if len(arr) >= 5:
        lo, hi = np.quantile(arr, [0.2, 0.8])
        arr = arr[(arr >= lo) & (arr <= hi)]
    return float(np.median(arr))


def track_shares_via_qty(days, files, base_position_file) -> Dict[str, Dict[str, float]]:
    """用 基线持仓数量 + 每文件 成交数量增量（按指令序号，取当日成交，跨天累加）。"""
    bp = Path(base_position_file)
    if not bp.is_absolute():
        bp = ROOT / bp
    base = m.read_excel_any(bp)
    cc = m._find_col(base.columns, m.CODE_COL_CANDIDATES)
    mkc = m._find_col(base.columns, m.MARKET_COL_CANDIDATES)
    pc = m._find_col(base.columns, m.PRODUCT_COL_CANDIDATES)
    qtyc = "持仓数量"
    shares: Dict[str, float] = {}
    names: Dict[str, str] = {}
    for _, r in base[base[pc].apply(m.is_sim_product)].iterrows():
        code = m.canonical_code(r.get(cc), market=r.get(mkc, "") if mkc else "")
        shares[code] = shares.get(code, 0) + m._safe_float(r.get(qtyc), 0)
        names[code] = r.get("证券名称", names.get(code, ""))
    daily = {days[0]: dict(shares)}
    applied_qty: Dict[str, float] = {}

    def ingest(fn: str, apply_delta: bool):
        raw = m.read_excel_any(ROOT / f"oplus交易文件/{fn}", dtype={"证券代码": str})
        if isinstance(raw, dict):
            raw = next(iter(raw.values()))
        cc2 = m._find_col(raw.columns, m.CODE_COL_CANDIDATES)
        pc2 = m._find_col(raw.columns, m.PRODUCT_COL_CANDIDATES)
        mkc2 = m._find_col(raw.columns, m.MARKET_COL_CANDIDATES)
        oidc = m._find_col(raw.columns, m.OPLUS_ORDER_ID_CANDIDATES)
        fqc = m._find_col(raw.columns, ["成交数量"])
        bc = m._find_col(raw.columns, ["指令批号"])
        dirc = m._find_col(raw.columns, ["委托方向", "交易方向", "方向"])
        for _, r in raw[raw[pc2].apply(m.is_sim_product)].iterrows():
            code = m.canonical_code(r.get(cc2), market=r.get(mkc2, "") if mkc2 else "")
            if m.market_group(code) == "OTHER":
                continue
            oid = str(r.get(oidc, "")).strip() or f"NOID:{r.get(bc)}|{code}"
            cum_fq = m._safe_float(r.get(fqc), 0)  # 累计成交数量（快照，正数）
            prev_q = applied_qty.get(oid, 0)
            # 成交数量无符号，必须按委托方向决定增减
            sgn = 1.0 if "买" in str(r.get(dirc, "")) else -1.0
            dq = sgn * (cum_fq - prev_q)
            if apply_delta and abs(dq) > 1e-9:
                shares[code] = shares.get(code, 0) + dq
                names[code] = r.get("证券名称", names.get(code, ""))
            applied_qty[oid] = cum_fq

    # 基线日 oplus 快照：只初始化各指令累计成交（不改持仓），
    # 否则基线日之前的历史成交会被当成增量重复计算。
    ingest(files[0], apply_delta=False)
    # 基线日之后：取相邻快照差并改持仓
    for day, fn in zip(days[1:], files[1:]):
        ingest(fn, apply_delta=True)
        daily[day] = {c: q for c, q in shares.items() if q > 1e-9}
    return daily, names


def _ensure_wind():
    if not w.isconnected():
        w.start()
    if not w.isconnected():
        raise RuntimeError(
            "Wind 终端未连接：请先启动 Wind 金融终端并登录账号，"
            "再重新运行本脚本。（WindPy w.start() 后 isconnected() 仍为 False）"
        )


def wind_close_matrix(codes, start, end):
    _ensure_wind()
    code_list = ",".join(m.manager_display_code(c) for c in codes)
    r = w.wsd(code_list, "close", start, end, "PriceAdj=F")
    if r.ErrorCode != 0:
        raise RuntimeError(f"wsd error {r.ErrorCode}（codes={code_list[:80]}...）")
    dates = [t.strftime("%Y-%m-%d") for t in r.Times]
    out = {}
    for i, code in enumerate(codes):
        vals = r.Data[i] if len(codes) > 1 else r.Data[0]
        out[code] = dict(zip(dates, vals))
    return out


def fx_hkd_cny(start, end):
    _ensure_wind()
    r = w.wsd("HKDCNY.EX", "close", start, end, "")
    dates = [t.strftime("%Y-%m-%d") for t in r.Times]
    return dict(zip(dates, r.Data[0]))


def main():
    args = parse_args()
    days, files = discover_days_files(args.base_date, args.end_date)
    if len(days) < 1:
        print("未发现符合日期范围的 oplus 快照，请检查 --base-date / --end-date。")
        return
    real_tags = discover_real_tags()
    print(f"基线 {days[0]}，增量至 {days[-1]}，共 {len(days)} 个交易日快照")
    is_default = (args.base_date == "2026-09-30" and args.end_date is None)
    if is_default:
        range_suffix = ""
    else:
        b = args.base_date.replace("-", "")[4:]
        e = days[-1].replace("-", "")[4:]
        range_suffix = f"_{b}至{e}"
    daily_shares, names = track_shares_via_qty(days, files, args.base_position_file)
    # 每日净资产
    navs = {}
    for day, fn in zip(days, files):
        raw = m.read_excel_any(ROOT / f"oplus交易文件/{fn}", dtype={"证券代码": str})
        if isinstance(raw, dict):
            raw = next(iter(raw.values()))
        navs[day] = infer_nav_from_a(raw)
    print("每日净资产（A股反推）：")
    for d in days:
        print(f"  {d}: {navs[d]:,.0f}")

    # 所有A/H代码
    all_codes = sorted({c for dd in daily_shares.values() for c in dd})
    ah_codes = [c for c in all_codes if m.market_group(c) in ("A", "HK")]
    close = wind_close_matrix(ah_codes, days[0], days[-1])
    fx = fx_hkd_cny(days[0], days[-1])

    # 每日盯市仓位
    market_daily = {}
    for day in days:
        rows = []
        nav = navs[day]
        for code, sh in daily_shares.get(day, {}).items():
            if m.market_group(code) not in ("A", "HK"):
                continue
            stem = code.split(".")[0]
            if stem.startswith(("204", "1318")):  # 国债逆回购，非网格标的
                continue
            px = close.get(code, {}).get(day, np.nan)
            if not np.isfinite(px):
                continue
            local_mv = sh * px
            mv = local_mv * fx.get(day, 1.0) if code.endswith(".HK") else local_mv
            rows.append({"code": code, "name": names.get(code, ""), "shares": sh,
                         "close": px, "mv_cny": mv, "pos": mv / nav if np.isfinite(nav) else np.nan})
        market_daily[day] = pd.DataFrame(rows)

    # 每日股数对账（增量重建 vs 真实持仓）
    print("\n==== 每日股数对账（增量重建 vs 真实持仓） ====")
    diff_rows = []
    print(f"{'日期':12} {'真实':>5} {'重建':>5} {'一致':>5} {'差异':>5}")
    for day in days:
        real_fn = real_tags.get(day)
        if real_fn is None:
            continue  # 该日无真实持仓文件可对比
        raw = m.read_excel_any(ROOT / "持仓数据" / real_fn)
        pc = m._find_col(raw.columns, m.PRODUCT_COL_CANDIDATES)
        cc = m._find_col(raw.columns, m.CODE_COL_CANDIDATES)
        mkc = m._find_col(raw.columns, m.MARKET_COL_CANDIDATES)
        real_sh = {}
        for _, r in raw[raw[pc].apply(m.is_sim_product)].iterrows():
            code = m.canonical_code(r.get(cc), market=r.get(mkc, "") if mkc else "")
            if m.market_group(code) not in ("A", "HK"):
                continue
            if code.split(".")[0].startswith(("204", "1318")):
                continue
            real_sh[code] = real_sh.get(code, 0) + m._safe_float(r.get("持仓数量"), 0)
        reb_sh = {c: q for c, q in daily_shares.get(day, {}).items()
                  if m.market_group(c) in ("A", "HK")}
        codes = sorted(set(real_sh) | set(reb_sh))
        exact = 0
        for c in codes:
            rq, bq = real_sh.get(c), reb_sh.get(c)
            if rq is None:
                diff_rows.append({"date": day, "code": c, "name": names.get(c, ""), "real": 0,
                                  "rebuilt": bq, "diff": bq, "kind": "仅重建"})
            elif bq is None:
                diff_rows.append({"date": day, "code": c, "name": names.get(c, ""), "real": rq,
                                  "rebuilt": 0, "diff": -rq, "kind": "仅真实(打新/非网格)"})
            elif abs(bq - rq) > 1e-6:
                diff_rows.append({"date": day, "code": c, "name": names.get(c, ""), "real": rq,
                                  "rebuilt": bq, "diff": bq - rq, "kind": "股数差异"})
            else:
                exact += 1
        print(f"{day:12} {len(real_sh):>5} {len(reb_sh):>5} {exact:>5} {len(codes)-exact:>5}")

    ddf = pd.DataFrame(diff_rows)
    if len(ddf):
        ddf["absdiff"] = ddf["diff"].abs()
        ddf = ddf.sort_values("absdiff", ascending=False)
        print("\n==== 差异/遗漏明细 ====")
        print(ddf[["date", "code", "name", "kind", "real", "rebuilt", "diff"]].to_string(index=False))
        p_audit = ROOT / f"持仓数据/增量重建_股数对账{range_suffix}.xlsx"
        ddf.to_excel(p_audit, index=False)
        print(f"明细已保存: {p_audit}")

    # 保存最新交易日持仓为标准格式（供次日网格使用）
    latest = days[-1]
    ltag = latest.replace("-", "")[4:]
    dl = market_daily[latest]
    out_l = pd.DataFrame({
        "日期": latest,
        "产品名称": "净值线3号",
        "单元名称": "单元",
        "交易市场": dl["code"].map(m.trade_market),
        "证券代码": dl["code"].map(lambda x: x.split(".")[0]),
        "证券名称": dl["name"],
        "当前仓位": dl["pos"],
    })
    p_l = ROOT / f"持仓数据/重建{ltag}持仓_盯市.xlsx"
    out_l.to_excel(p_l, index=False)
    print(f"\n最新日({latest})盯市持仓已保存: {p_l}（{len(dl)}行，总仓位={dl['pos'].sum():.4f}）")
    # 保存全部每日
    p_daily = ROOT / f"持仓数据/重建每日持仓_盯市{range_suffix}.xlsx"
    with pd.ExcelWriter(p_daily) as wri:
        for day in days:
            market_daily[day].to_excel(wri, sheet_name=day.replace("-", "")[4:], index=False)
    print(f"每日盯市持仓已保存: {p_daily}")


if __name__ == "__main__":
    main()
