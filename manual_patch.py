# -*- coding: utf-8 -*-
"""
人工改动记忆库：让你在 *_剔除已成交.xlsx 里的手改"一次记下、以后自动套用"。

为什么需要它
------------
每次跑 filter_completed_instructions_hk_nosuffix_fixed.py，都会从原始组合指令
重新生成 *_剔除已成交.xlsx，把你上次的手改冲掉。本脚本把你的手改存成补丁，
下次自动套回去。

每天用法（接在"跑剔除脚本"之后）
------------------------------
1) 套用历史手改（刚跑完剔除、还没人工改时）：
       python manual_patch.py --apply
   它会：
     - 把刚生成的纯自动版 *_剔除已成交.xlsx 备份成 *_自动版基准.xlsx
     - 把"人工指令补丁.csv"里你以前的手改写回/清空到对应格子

2) 打开 *_剔除已成交.xlsx 做新的人工增删，保存。

3) 学习这次新的手改（改完之后跑）：
       python manual_patch.py --learn
   它会 diff"自动版基准"和你改完的版本，把新增的人工改动追加进补丁库，
   明天起自动生效。

补丁库文件：指令\\人工指令补丁.csv
   sheet,row,col 精确定位（同一份组合指令在迭代期间行列不变，所以能自动复原）。
   经理下发全新指令（表结构变了）后，旧补丁不再适用，人工重新核对一次即可。
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook

# 直接复用剔除脚本的识别逻辑，保证认股票、认指令的方式完全一致
import filter_completed_instructions_hk_nosuffix_fixed as F

ROOT = Path(__file__).resolve().parent
INST_DIR = ROOT / "指令"
PATCH_CSV = INST_DIR / "人工指令补丁.csv"

PATCH_COLS = ["sheet", "row", "col", "code", "side", "norm_text", "raw_text", "action", "learned_at"]


# ---------------------------------------------------------------- 找文件
def find_latest_filtered() -> Path:
    cands = [
        p for p in INST_DIR.glob("组合指令*_剔除已成交.xlsx")
        if "_自动版基准" not in p.name and not p.name.startswith("~$")
    ]
    if not cands:
        raise FileNotFoundError(f"在 {INST_DIR} 没找到 组合指令*_剔除已成交.xlsx，先跑剔除脚本。")
    return max(cands, key=lambda p: p.stat().st_mtime)


def baseline_path_of(filtered_path: Path) -> Path:
    return filtered_path.with_name(filtered_path.stem + "_自动版基准.xlsx")


# ---------------------------------------------------------------- 读补丁库
def load_patches() -> pd.DataFrame:
    if not PATCH_CSV.exists():
        return pd.DataFrame(columns=PATCH_COLS)
    df = pd.read_csv(PATCH_CSV, dtype=str).fillna("")
    for c in ["row", "col"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
    return df


def save_patches(df: pd.DataFrame) -> None:
    df = df[[c for c in PATCH_COLS if c in df.columns]]
    df.to_csv(PATCH_CSV, index=False, encoding="utf-8-sig")


# ---------------------------------------------------------------- 遍历指令格子
def read_instruction_cells(xlsx: Path) -> Dict[Tuple[str, int, int], Dict[str, Any]]:
    """返回 {(sheet,row,col): {code, raw, side, norm}} 只收指令文本格子。"""
    wb = load_workbook(xlsx)
    out: Dict[Tuple[str, int, int], Dict[str, Any]] = {}
    for ws in wb.worksheets:
        current_code = ""
        for r in range(1, ws.max_row + 1):
            rowvals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            new_code = F.detect_code_in_row(rowvals)
            if new_code:
                current_code = new_code
            if not current_code:
                continue
            for c, v in enumerate(rowvals, start=1):
                if not F.looks_like_instruction_text(v):
                    continue
                txt = F.clean_text(v)
                out[(ws.title, r, c)] = {
                    "code": current_code,
                    "raw": txt,
                    "side": F.instruction_side_from_text(txt),
                    "norm": F.normalize_instruction_text(txt),
                }
    return out


# ---------------------------------------------------------------- apply
def cmd_apply(filtered: Path) -> None:
    baseline = baseline_path_of(filtered)
    shutil.copy2(filtered, baseline)
    print(f"已备份纯自动版基准：{baseline.name}")

    patches = load_patches()
    if patches.empty:
        print("补丁库为空，没有历史手改需要套用。")
        return

    wb = load_workbook(filtered)
    applied_keep = applied_drop = 0
    for _, p in patches.iterrows():
        sheet, r, c = str(p["sheet"]), int(p["row"]), int(p["col"])
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        cur = ws.cell(r, c).value
        cur_txt = F.clean_text(cur) if cur is not None else ""
        if p["action"] == "keep":
            # 强制保留：若该格已被清空，则写回原文；本来就有则不动
            if not cur_txt:
                ws.cell(r, c).value = p["raw_text"]
                applied_keep += 1
                print(f"  [保留] {sheet}!R{r}C{c} 写回：{p['raw_text']}")
        elif p["action"] == "drop":
            # 强制剔除：若该格还有指令文本，则清空
            if cur_txt and F.looks_like_instruction_text(cur):
                ws.cell(r, c).value = None
                applied_drop += 1
                print(f"  [剔除] {sheet}!R{r}C{c} 清空：{cur_txt}")
    wb.save(filtered)
    print(f"\n套用完成：写回保留 {applied_keep} 条，清空剔除 {applied_drop} 条。")
    print("现在可以打开 xlsx 做新的人工核对；改完跑：python manual_patch.py --learn")


# ---------------------------------------------------------------- learn
def cmd_learn(filtered: Path) -> None:
    baseline = baseline_path_of(filtered)
    if not baseline.exists():
        print(f"找不到基准 {baseline.name}。请先跑：python manual_patch.py --apply")
        sys.exit(1)

    auto = read_instruction_cells(baseline)   # 纯自动版
    mine = read_instruction_cells(filtered)   # 你改完的版本

    today = date.today().isoformat()
    new_rows = []

    all_keys = set(auto) | set(mine)
    for key in all_keys:
        a = auto.get(key)
        m = mine.get(key)
        a_txt = a["raw"] if a else ""
        m_txt = m["raw"] if m else ""
        if a_txt == m_txt:
            continue  # 没动

        sheet, r, c = key
        base_info = m or a
        if not base_info:
            continue
        if not a_txt and m_txt:
            action = "keep"   # 自动版空、你补了 -> 以后要保留
        elif a_txt and not m_txt:
            action = "drop"   # 自动版有、你删了 -> 以后要剔除
        else:
            action = "keep"   # 文本改了，以你的为准
            base_info = m

        new_rows.append({
            "sheet": sheet, "row": r, "col": c,
            "code": base_info["code"], "side": base_info["side"],
            "norm_text": base_info["norm"], "raw_text": m_txt or a_txt,
            "action": action, "learned_at": today,
        })

    if not new_rows:
        print("和自动版基准对比，没有发现新的人工改动。")
        return

    new_df = pd.DataFrame(new_rows)[PATCH_COLS]
    old_df = load_patches()

    # 去重：同一 (sheet,row,col,norm_text) 以新学到的为准
    key_cols = ["sheet", "row", "col"]
    merged = pd.concat([old_df, new_df], ignore_index=True)
    merged = merged.drop_duplicates(subset=key_cols, keep="last")
    save_patches(merged)

    print(f"学到 {len(new_rows)} 条新的人工改动，已存入：{PATCH_CSV.name}")
    for _, n in new_df.iterrows():
        tag = "保留(以后别删)" if n["action"] == "keep" else "剔除(以后别留)"
        print(f"  [{tag}] {n['code']} {n['sheet']}!R{int(n['row'])}C{int(n['col'])} | {n['raw_text']}")
    print("\n明天起这些改动会自动套用，不用再手改。")


def main() -> None:
    ap = argparse.ArgumentParser(description="人工改动记忆库：apply 套用历史手改 / learn 学习新手改。")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--apply", action="store_true", help="跑剔除后执行：备份基准 + 套用历史补丁")
    g.add_argument("--learn", action="store_true", help="人工改完后执行：diff 学习新补丁")
    args = ap.parse_args()

    filtered = find_latest_filtered()
    print(f"当前剔除版：{filtered.name}\n")
    if args.apply:
        cmd_apply(filtered)
    else:
        cmd_learn(filtered)


if __name__ == "__main__":
    main()
