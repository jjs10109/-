# -*- coding: utf-8 -*-
"""
第 2 步合并：学习你这次的手改 + 跑网格主程序。

用法（把要传给网格主程序的参数原样放在后面）：
    python step2_grid.py --asof-date 2026-09-19 --oplus-date 0918 --position-file "全球产品持仓查询_20260918.xls"

它会依次执行：
    1) manual_patch.py --learn      把你刚才在 xlsx 里的手改记进补丁库
    2) live_grid_simulator.py ...  按后面给的参数跑网格（参数透传）
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

# 本脚本在 日常代码/ 子目录，项目根目录是父目录
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def run(cmd: list[str]) -> None:
    print(">>> " + " ".join(cmd))
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f"[出错] 退出码 {r.returncode}，流程停止。")
        sys.exit(r.returncode)


def main() -> None:
    print("=" * 64)
    print("第 2 步：学习手改 + 生成网格")
    print("=" * 64)
    run([sys.executable, str(ROOT / "manual_patch.py"), "--learn"])
    print()
    # 把用户传的所有参数原样透传给网格主程序（sim版，含流动性模拟）
    grid_args = sys.argv[1:]
    # 默认价格带 ±30%（原版默认±10%太窄，会过滤掉大量档位）；用户显式传了则不覆盖
    if not any(a.startswith("--price-band") for a in grid_args):
        grid_args = ["--price-band", "0.30"] + grid_args
    run([sys.executable, str(ROOT / "live_grid_simulator_sim.py"), *grid_args])
    print("\n第 2 步完成。去 outputs_live_grid\\ 拿买卖清单。")


if __name__ == "__main__":
    main()
