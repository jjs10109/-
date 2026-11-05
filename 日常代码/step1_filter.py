# -*- coding: utf-8 -*-
"""
第 1 步合并：跑剔除脚本 + 自动恢复你历史的手改。

用法：
    python step1_filter.py

它会依次执行：
    1) filter_completed_instructions_hk_nosuffix_fixed.py   重新生成剔除版
    2) manual_patch.py --apply                              备份基准 + 套用历史手改

跑完后：打开 指令\\*_剔除已成交.xlsx 做新的人工核对并保存，
再跑 step2_grid.py。
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

# 本脚本在 日常代码/ 子目录，项目根目录是父目录
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def run(script: str, extra=None) -> None:
    cmd = [sys.executable, str(ROOT / script)] + (extra or [])
    print(">>> " + " ".join(cmd))
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f"[出错] {script} 退出码 {r.returncode}，后续步骤已停止。")
        sys.exit(r.returncode)


def main() -> None:
    print("=" * 64)
    print("第 1 步：剔除已成交 + 恢复历史手改")
    print("=" * 64)
    run("filter_completed_instructions_hk_nosuffix_fixed.py")
    print()
    run("manual_patch.py", ["--apply"])
    print("\n第 1 步完成。")
    print("下一步：打开 指令\\ 下最新的 *_剔除已成交.xlsx，人工核对并保存。")
    print("然后跑：python step2_grid.py --asof-date 今天 --oplus-date 前一日 --position-file 持仓文件名")


if __name__ == "__main__":
    main()
