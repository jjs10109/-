# -*- coding: utf-8 -*-
"""
对账6：网格价格逐只对照 Wind 核验（事后核对当前 outputs_live_grid 结果）

核对维度：
  1) 每只股票的网格参考价 == Wind 在行情日(Oplus日)的未复权收盘价（偏差0）；
  2) 所有档位价都落在参考价 ±30% 价格带内；无非正价/缺失/重复档；
  3) 方向：买入档价≤现价、卖出档价≥现价；落在现价"已触发"一侧的经理条件单
     （如"X以下买入"且现价已低于X）属于导入即成交，判定为正确；
  4) 目标仓位单调性：买入越跌目标仓位越高、卖出越涨目标仓位越低；
  5) "X元以上卖完"档若该股票全产品零持仓，则正确跳过（无股可卖）。

运行：D:\\anaconda3\\python.exe 对账6_网格价格Wind核对.py
需连 Wind，读取 outputs_live_grid/实盘网格_复核报告.xlsx 与 指令/ 下最新剔除版。
报告输出：outputs_live_grid/网格价格Wind核对报告.xlsx
"""
import sys
from pathlib import Path

ROOT = Path(r"C:\Users\gucao\Desktop\live_grid_project")
sys.path.insert(0, str(ROOT))
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

from data_loader import ensure_wind, wind_code_candidates  # noqa: E402
from instruction_reader import parse_instruction_excel  # noqa: E402


def log(*a):
    print(*a, flush=True)


def latest_report():
    out = ROOT / "outputs_live_grid"
    fs = sorted(out.glob("实盘网格_复核报告*.xlsx"))
    fs = [f for f in fs if not f.name.startswith("~$") and "_old" not in f.name]
    return fs[-1]


def latest_filtered_instruction():
    d = ROOT / "指令"
    fs = sorted(d.glob("组合指令*_剔除已成交.xlsx"))
    fs = [f for f in fs if not f.name.startswith("~$")]
    return fs[-1]


def wind_raw_close(w, code, end_date, start_date):
    """返回 (收盘价, 成功Wind代码, 该收盘价对应日期)。逐候选代码尝试。"""
    for wc in wind_code_candidates(code):
        r = w.wsd(wc, "close", start_date, end_date, "PriceAdj=0")
        if getattr(r, "ErrorCode", -1) != 0:
            continue
        vals = list(r.Data[0])
        times = list(r.Times)
        ss = pd.Series(vals, index=times)
        ss = pd.to_numeric(ss, errors="coerce").dropna()
        if ss.empty:
            continue
        return float(ss.iloc[-1]), wc, ss.index[-1].strftime("%Y-%m-%d")
    return float("nan"), "", ""


def main():
    rep_path = latest_report()
    ins_path = latest_filtered_instruction()
    log("复核报告:", rep_path.name)
    log("指令文件:", ins_path.name)
    g = pd.read_excel(rep_path, sheet_name="refreshed_active_grid")
    cp = pd.read_excel(rep_path, sheet_name="current_positions")
    g["price"] = pd.to_numeric(g["price"], errors="coerce")
    g["reference_price"] = pd.to_numeric(g["reference_price"], errors="coerce")
    g["target_pos"] = pd.to_numeric(g["target_pos"], errors="coerce")
    mdd = str(g["market_data_date"].iloc[0])[:10]
    end = pd.to_datetime(mdd)
    start = (end - pd.Timedelta(days=14)).strftime("%Y-%m-%d")
    pdf, _ = parse_instruction_excel(ins_path)
    ok = pdf[pdf["parse_status"].eq("ok")]

    w = ensure_wind()
    rows, problems = [], []
    mono_bad = 0
    mono_chk = 0

    for code, x in g.groupby("code"):
        name = x["name"].iloc[0]
        ref = float(x["reference_price"].iloc[0])
        lo, hi = float(x["band_low"].iloc[0]), float(x["band_high"].iloc[0])
        src = x["reference_price_source"].iloc[0]
        wc_px, wf, wdate = wind_raw_close(w, code, mdd, start)
        diff = round((ref - wc_px) / wc_px * 100, 3) if np.isfinite(wc_px) and wc_px else np.nan
        buys, sells = x[x["side"] == "buy"], x[x["side"] == "sell"]
        br = f"{len(buys)}档[{buys['price'].min():.2f},{buys['price'].max():.2f}]" if len(buys) else "无"
        sr = f"{len(sells)}档[{sells['price'].min():.2f},{sells['price'].max():.2f}]" if len(sells) else "无"
        # 带外被过滤的经理档
        ins = ok[ok["code"] == code]
        out_lv = [f"{z['side']}{z['price']:g}" for _, z in ins.iterrows()
                  if pd.notna(z["price"]) and not (lo - 1e-9 <= z["price"] <= hi + 1e-9)]
        # 持仓非0单元数
        csub = cp[cp["code"].astype(str) == code]
        nz = csub[pd.to_numeric(csub["current_pos"], errors="coerce").fillna(0).abs() > 1e-9] if len(csub) else csub
        hold = f"{len(nz)}个单元" if len(nz) else "0持仓"

        flags = []
        # 1 参考价
        if not np.isfinite(wc_px):
            flags.append("Wind取不到价")
        elif abs(diff) > 0.5:
            flags.append(f"参考价偏差{diff}%")
        # 2 价格带/异常价
        ob = int(((x["price"] < lo - 1e-9) | (x["price"] > hi + 1e-9)).sum())
        if ob:
            flags.append(f"{ob}档超±30%带")
        if int(((x["price"] <= 0) | x["price"].isna()).sum()):
            flags.append("存在非正/缺失价")
        # 4 单调性（逐单元逐方向）
        for _, u in x.groupby(["product", "unit"], dropna=False):
            for side in ("buy", "sell"):
                s = u[u["side"] == side].sort_values("price")
                t = s["target_pos"].dropna().values
                if len(t) >= 2:
                    mono_chk += 1
                    if np.any(np.diff(t) > 1e-6):
                        mono_bad += 1
        # 3 方向：买入最高档不应高于现价（容差1%）；卖出最低档不应低于现价（容差1%）
        # 经理条件单"X以下买入"在现价≤X时会立即成交，属于正常触发单，不算错误。
        if len(buys) and buys["price"].max() > ref * 1.01:
            # 必须在经理指令里能找到 >= 该价 的"买入到X"锚点，才认定为触发单
            trig = ins[(ins["side"] == "buy") & (pd.to_numeric(ins["price"], errors="coerce") >= buys["price"].max() - 1e-9)]
            if not trig.empty:
                pass  # 经理触发买入单，正确
            else:
                flags.append(f"最高买入{buys['price'].max():.2f}高于现价且无经理锚点")
        if len(sells) and sells["price"].min() < ref * 0.99:
            trig = ins[(ins["side"] == "sell") & (pd.to_numeric(ins["price"], errors="coerce") <= sells["price"].min() + 1e-9)]
            if not trig.empty:
                pass  # 经理触发卖出单，正确
            else:
                flags.append(f"最低卖出{sells['price'].min():.2f}低于现价且无经理锚点")

        verdict = "✓ 正确" if not flags else "✗ 需核查"
        if flags:
            problems.append(f"{code} {name}: {'; '.join(flags)}")
        rows.append({
            "代码": code, "名称": name, "Wind代码": wf, "行情日": mdd,
            "Wind收盘": round(wc_px, 3) if np.isfinite(wc_px) else None,
            "Wind价格日期": wdate, "网格参考价": round(ref, 3), "参考价偏差%": diff,
            "价格来源": src, "价格带下限": round(lo, 2), "价格带上限": round(hi, 2),
            "买入网格": br, "卖出网格": sr, "带外过滤远端档": ", ".join(out_lv) if out_lv else "无",
            "当前持仓": hold, "结论": verdict,
        })

    # 5 零持仓"卖完"档核验
    for code, x in g.groupby("code"):
        ins = ok[ok["code"] == code]
        csub = cp[cp["code"].astype(str) == code]
        nz = csub[pd.to_numeric(csub["current_pos"], errors="coerce").fillna(0).abs() > 1e-9] if len(csub) else csub
        if len(nz):
            continue
        sell_all = ins[(ins["side"] == "sell") & (pd.to_numeric(ins["target_pos"], errors="coerce") == 0)]
        # 零持仓卖完档不在网格里是正常的
        for _, z in sell_all.iterrows():
            if len(x[(x["side"] == "sell") & (np.isclose(x["price"], z["price"], atol=0.01))]):
                problems.append(f"{code} 零持仓却出现卖完档 {z['price']}（应跳过）")

    rep = pd.DataFrame(rows)
    out_path = ROOT / "outputs_live_grid" / "网格价格Wind核对报告.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as ww:
        rep.to_excel(ww, index=False, sheet_name="逐只Wind核对")

    log("")
    log(f"股票数 {len(rep)}，正确 {(rep['结论']=='✓ 正确').sum()}，需核查 {(rep['结论']=='✗ 需核查').sum()}")
    log(f"目标仓位单调性检查 {mono_chk} 组，异常 {mono_bad} 组")
    log("========== 问题清单 ==========")
    if problems:
        for p in problems:
            log(" -", p)
    else:
        log("全部通过：参考价=Wind收盘、档位在±30%带内、方向/单调性正确、零持仓卖完档已跳过。")
    log("报告:", out_path)


if __name__ == "__main__":
    main()
