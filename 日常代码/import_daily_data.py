# -*- coding: utf-8 -*-
"""
每日数据一键导入（收件箱）。

用法：
    把当天新下载的所有文件（oplus 成交、持仓查询、新组合指令）
    一股脑丢进本项目根目录的「_待导入」文件夹，然后运行：

        python import_daily_data.py

    它会按文件名自动识别并归位：
        含 oplus      ->  oplus文件/        （剔除脚本和网格主程序共用的成交数据）
        含 组合指令   ->  指令/             （新下发的经理指令）
        含 持仓/全球产品/position -> 持仓数据/ （持仓基线）
        含 剔除已成交 / 以 ~$ 开头 -> 跳过   （这是产物或临时文件，不该再导入）

    只预览不动文件：
        python import_daily_data.py --dry-run

    同名文件已存在时不会覆盖，会跳过并提示，防止历史数据被误顶掉。
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

# 本脚本在 日常代码/ 子目录，项目根目录是父目录
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

INBOX = ROOT / "_待导入"

OPLUS_DIR = ROOT / "oplus文件"
INSTRUCTION_DIR = ROOT / "指令"
# 持仓文件统一放到 持仓数据/ 文件夹
POSITION_DIR = ROOT / "持仓数据"

POSITION_KEYWORDS = ("持仓", "全球产品", "position", "positions")
# 这些是流程产物 / 临时文件，绝不能再当输入导回去
IGNORE_KEYWORDS = ("剔除已成交",)


def classify(name: str) -> str:
    n = name.strip()
    low = n.lower()
    if n.startswith("~$"):
        return "ignore"  # Excel 临时锁文件
    if any(k in n for k in IGNORE_KEYWORDS):
        return "ignore"  # 剔除产物，不是新输入
    if low.startswith("oplus") or "oplus" in low:
        return "oplus"
    if "组合指令" in n:
        return "instruction"
    if any(k in low for k in POSITION_KEYWORDS):
        return "position"
    return "unknown"


TARGET = {
    "oplus": (OPLUS_DIR, "oplus文件/（成交数据）"),
    "instruction": (INSTRUCTION_DIR, "指令/（经理组合指令）"),
    "position": (POSITION_DIR, "持仓数据/（持仓基线）"),
}


def main() -> int:
    ap = argparse.ArgumentParser(description="把 _待导入/ 里的当天文件自动归位。")
    ap.add_argument("--dry-run", action="store_true", help="只预览，不移动文件。")
    args = ap.parse_args()

    if not INBOX.exists():
        INBOX.mkdir(parents=True, exist_ok=True)
        print(f"已创建收件箱：{INBOX}")
        print("请把今天下载的文件放进去，再重跑本脚本。")
        return 0

    files = [p for p in INBOX.iterdir() if p.is_file() and not p.name.startswith(".")]
    if not files:
        print("收件箱是空的：把今天的 oplus / 持仓 / 组合指令 丢进 _待导入\\ 再运行。")
        return 0

    print("=" * 64)
    print("每日数据导入（收件箱 -> 自动归位）")
    print("=" * 64)

    moved, skipped_overwrite, ignored, unknown = [], [], [], []

    for p in sorted(files, key=lambda x: x.name):
        kind = classify(p.name)
        if kind == "ignore":
            ignored.append(p)
            continue
        if kind == "unknown":
            unknown.append(p)
            continue

        dest_dir, label = TARGET[kind]
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / p.name
        if dest.exists():
            skipped_overwrite.append((p, dest))
            continue
        if args.dry_run:
            print(f"  [预览] {p.name:40s} -> {label}")
        else:
            shutil.move(str(p), str(dest))
            print(f"  [导入] {p.name:40s} -> {label}")
        moved.append((p, dest, label))

    print("-" * 64)
    print(f"成功导入 {len(moved)} 个文件。")
    if skipped_overwrite:
        print("\n[未导入] 以下文件目标位置已存在同名文件，未覆盖：")
        for p, dest in skipped_overwrite:
            print(f"   - {p.name}  （已存在 {dest}，如需替换请手动处理）")
    if ignored:
        print(f"\n[跳过] {len(ignored)} 个产物/临时文件未处理：")
        for p in ignored:
            print(f"   - {p.name}")
    if unknown:
        print("\n[待人工判断] 以下文件没认出是什么，请确认后手动放：")
        for p in unknown:
            print(f"   - {p.name}")
    print("=" * 64)
    if not args.dry_run:
        print("下一步（在项目根目录执行，脚本在 日常代码\\ 下）：")
        print("  1) 剔除 + 恢复历史手改：python 日常代码/step1_filter.py")
        print("  2) 打开 指令\\ 下最新的 *_剔除已成交.xlsx，人工核对增删并保存（只改这个xlsx）")
        print("  3) 跑网格（盘前，持仓已含前一天成交，务必加 --no-apply-oplus-trades）：")
        print("     python 日常代码/step2_grid.py --asof-date 今天 --oplus-date 前一日 "
              "--position-file \"持仓数据/最新全球持仓.xls\" "
              "--instruction-file \"指令/最新剔除版指令.xlsx\" --no-apply-oplus-trades")
    return 0


if __name__ == "__main__":
    sys.exit(main())
