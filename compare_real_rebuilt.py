# -*- coding: utf-8 -*-
"""对比重建持仓（重建每日持仓_盯市.xlsx）与真实持仓（全球产品持仓查询_*.xls）。
核心指标：每个交易日、每个 A/H 网格标的的【持仓股数】。"""
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd
import numpy as np

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import live_grid_simulator_current as m

REAL_TAGS = ["20261008", "20261009", "20261012", "20261013", "20261014", "20261015", "20261016"]
# 内部日期 -> 重建sheet名（MMDD）
REBUILT_SHEET = {
    "2026-09-30": "0930", "2026-10-08": "1008", "2026-10-09": "1009",
    "2026-10-12": "1012", "2026-10-13": "1013", "2026-10-14": "1014",
}


def read_real(tag):
    raw = m.read_excel_any(ROOT / f"持仓数据/全球产品持仓查询_{tag}.xls")
    pc = m._find_col(raw.columns, m.PRODUCT_COL_CANDIDATES)
    cc = m._find_col(raw.columns, m.CODE_COL_CANDIDATES)
    mkc = m._find_col(raw.columns, m.MARKET_COL_CANDIDATES)
    qtyc = m._find_col(raw.columns, ["持仓数量"])
    sub = raw[raw[pc].apply(m.is_sim_product)]
    idate = str(sub["日期"].dropna().iloc[0])[:10]
    d, names = {}, {}
    for _, r in sub.iterrows():
        code = m.canonical_code(r.get(cc), market=r.get(mkc, "") if mkc else "")
        if m.market_group(code) not in ("A", "HK"):
            continue
        if code.split(".")[0].startswith(("204", "1318")):
            continue
        d[code] = d.get(code, 0) + m._safe_float(r.get(qtyc), 0)
        names[code] = r.get("证券名称", names.get(code, ""))
    return idate, d, names


def read_rebuilt(sheet):
    df = pd.read_excel(ROOT / "持仓数据/重建每日持仓_盯市.xlsx", sheet_name=sheet)
    return dict(zip(df["code"], df["shares"])), dict(zip(df["code"], df["name"]))


def main():
    all_rows = []
    print(f"{'内部日期':12} {'真实档数':>6} {'重建档数':>6} {'股数一致':>6} {'有差异':>6} {'仅真实':>6} {'仅重建':>6}")
    for tag in REAL_TAGS:
        idate, real_d, real_names = read_real(tag)
        sheet = REBUILT_SHEET.get(idate)
        if sheet is None:
            print(f"{idate:12} {len(real_d):>6} {'--':>6}  重建未覆盖该日")
            continue
        reb_d, reb_names = read_rebuilt(sheet)
        codes = sorted(set(real_d) | set(reb_d))
        exact = mismatch = only_real = only_reb = 0
        for c in codes:
            rq, bq = real_d.get(c), reb_d.get(c)
            if rq is None:
                only_reb += 1
                all_rows.append({"date": idate, "code": c, "name": reb_names.get(c, ""),
                                 "real": 0, "rebuilt": bq, "diff": bq, "pct": np.nan, "type": "仅重建"})
            elif bq is None:
                only_real += 1
                all_rows.append({"date": idate, "code": c, "name": real_names.get(c, ""),
                                 "real": rq, "rebuilt": 0, "diff": -rq, "pct": np.nan, "type": "仅真实"})
            else:
                d = bq - rq
                if abs(d) < 1e-6:
                    exact += 1
                else:
                    mismatch += 1
                    all_rows.append({"date": idate, "code": c, "name": real_names.get(c, ""),
                                     "real": rq, "rebuilt": bq, "diff": d,
                                     "pct": d / rq if rq else np.nan, "type": "股数差异"})
        print(f"{idate:12} {len(real_d):>6} {len(reb_d):>6} {exact:>6} {mismatch:>6} {only_real:>6} {only_reb:>6}")

    diff_df = pd.DataFrame(all_rows)
    print("\n==== 全部差异明细（按 |diff| 排序） ====")
    if len(diff_df):
        diff_df["absdiff"] = diff_df["diff"].abs()
        diff_df = diff_df.sort_values("absdiff", ascending=False)
        show = diff_df[["date", "code", "name", "type", "real", "rebuilt", "diff", "pct"]]
        print(show.to_string(index=False))
    else:
        print("无任何差异。")

    out = ROOT / "持仓重建_真实对比_股数.xlsx"
    diff_df.to_excel(out, index=False)
    print(f"\n差异明细已保存: {out}")


if __name__ == "__main__":
    main()
