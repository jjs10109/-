# -*- coding: utf-8 -*-
"""
一键清理本项目调试产生的临时文件（安全：只删下划线开头的临时文件和调试目录）。
保留：_待导入、对账工具、所有正式代码/数据/输出、state（含旧增量备份）。
双击运行，或在项目根目录：python 清理临时文件.py
"""
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TEMP_EXT = {".txt", ".py", ".json", ".pkl", ".xlsx"}
TEMP_DIRS = ["_state_old_incremental_backup", "_test_state", "_test_outputs",
             "state_fresh", "_base", "__pycache__"]
# 输出目录里的历史遗留
TEMP_OUTPUTS = ["实盘网格_基金经理格式_当前全部网格_old.xlsx"]

removed_files = 0
removed_dirs = 0

# 1) 根目录下划线开头的临时文件（保留 _待导入 收件箱；保留的对账脚本已在“对账工具/”里）
for p in ROOT.glob("_*"):
    if p.is_file() and p.suffix.lower() in TEMP_EXT:
        p.unlink()
        removed_files += 1

# 2) 临时/缓存目录
for d in TEMP_DIRS:
    p = ROOT / d
    if p.exists():
        shutil.rmtree(p, ignore_errors=True)
        removed_dirs += 1

# 3) 子目录里的 __pycache__
for p in ROOT.rglob("__pycache__"):
    shutil.rmtree(p, ignore_errors=True)
    removed_dirs += 1

# 4) 输出目录的历史遗留
out = ROOT / "outputs_live_grid"
if out.exists():
    for name in TEMP_OUTPUTS:
        p = out / name
        if p.exists():
            p.unlink()
            removed_files += 1

# 5) 根目录里误放的网格输出文件（正式输出只在 outputs_live_grid/）
for p in ROOT.glob("实盘网格_*.xlsx"):
    if p.is_file():
        p.unlink()
        removed_files += 1

print(f"清理完成：删除临时文件 {removed_files} 个，临时目录 {removed_dirs} 个。")
print("保留：_待导入、对账工具、正式代码/数据/输出、state。")
if sys.stdin.isatty():
    input("按回车退出。")
