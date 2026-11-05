# -*- coding: utf-8 -*-
"""
实盘网格化模拟 / 买卖清单生成器

核心目标：
1. 读取组合指令，按每只股票生成网格价位；
2. 读取 oplus 交易数据，按产品 + 单元 + 股票计算/更新当前仓位；
3. 严格按 A股/港股买卖清单导入模板输出；
4. 只输出当前价格上下 price_band（默认 10%）内的网格；
5. 生成 delta ledger，避免同一网格重复录入；
6. 指令更新时，只刷新发生变化的股票的网格，其余股票沿用 active grid master。

依赖：pandas, numpy, openpyxl；如需读取 .xls，Windows 推荐安装 pywin32 并本机有 Excel。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from copy import copy

from config import DEFAULT_PRICE_ADJ, BUY_GAMMA_DEFAULT, SELL_GAMMA_DEFAULT
from data_loader import fetch_wind_daily_ohlcva, fetch_wind_latest_raw_close, fetch_wind_intraday_minute
from features_daily import prepare_daily_features
from grid_engine import evaluate_manager_grid_splitter
from instruction_reader import parse_instruction_excel

TEMPLATE_HEADERS = ["证券代码", "证券名称", "产品名称", "单元名称", "交易市场", "委托方向", "数量/仓位", "极限价格"]


DEFAULT_INSTRUCTION_DIR_NAME = "指令"
DEFAULT_OPLUS_DIR_NAME = "oplus交易文件"
EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}


def project_root_from_args(args: argparse.Namespace) -> Path:
    return Path(args.project_root).expanduser().resolve() if getattr(args, "project_root", "") else Path(__file__).resolve().parent


def ensure_live_project_folders(project_root: Path) -> None:
    """创建用户约定的输入文件夹。"""
    for name in [DEFAULT_INSTRUCTION_DIR_NAME, DEFAULT_OPLUS_DIR_NAME]:
        folder = project_root / name
        folder.mkdir(parents=True, exist_ok=True)
        keep = folder / "README.txt"
        if not keep.exists():
            if name == DEFAULT_INSTRUCTION_DIR_NAME:
                keep.write_text("把当天的组合指令 Excel 放在这个文件夹里。程序会自动读取最新文件。\n", encoding="utf-8")
            else:
                keep.write_text("把每天的 oplus 交易文件放在这个文件夹里。程序会询问日期，并读取该日期最新文件。\n", encoding="utf-8")


def normalize_date_token(x: Any, default_year: Optional[int] = None) -> str:
    """把 0623 / 20260623 / 2026-06-23 等输入统一为 YYYY-MM-DD。"""
    s = str(x or "").strip()
    if not s:
        return ""
    s = s.replace("年", "-").replace("月", "-").replace("日", "")
    s = s.replace("/", "-").replace(".", "-").replace("_", "-")
    digits = re.sub(r"\D", "", s)
    if len(digits) == 8:
        try:
            return datetime.strptime(digits, "%Y%m%d").strftime("%Y-%m-%d")
        except Exception:
            return s
    if len(digits) == 4:
        year = int(default_year or datetime.today().year)
        try:
            return datetime.strptime(f"{year}{digits}", "%Y%m%d").strftime("%Y-%m-%d")
        except Exception:
            return s
    try:
        return pd.to_datetime(s).strftime("%Y-%m-%d")
    except Exception:
        return s


def date_match_tokens(date_text: Any) -> List[str]:
    d = normalize_date_token(date_text)
    if not d:
        return []
    try:
        dt = datetime.strptime(d, "%Y-%m-%d")
    except Exception:
        raw = str(date_text or "").strip()
        return [raw] if raw else []
    return [
        dt.strftime("%Y%m%d"),
        dt.strftime("%y%m%d"),
        dt.strftime("%m%d"),
        dt.strftime("%Y-%m-%d"),
        dt.strftime("%Y_%m_%d"),
        dt.strftime("%m-%d"),
    ]


def _excel_files(folder: Path) -> List[Path]:
    if not folder.exists():
        return []
    files = []
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$"):
            files.append(p)
    return files


def pick_latest_file(files: Sequence[Path]) -> Optional[Path]:
    if not files:
        return None
    return max(files, key=lambda x: (x.stat().st_mtime, x.name))


def find_latest_instruction_file(project_root: Path, asof_date: str = "") -> Path:
    """自动选择指定日期的组合指令文件。

    只接受文件名同时满足：
    1. 包含“组合指令”；
    2. 包含运行日期（支持 0624、20260624、2026-06-24 等格式）。

    文件名可以继续带 v2/v3、最终版、(1) 等任意后缀。同一天匹配到
    多个文件时，使用最后修改时间最新的一份；不再退回读取其他 Excel，
    防止误把持仓、模板或上一日指令当成当天组合指令。
    """
    folder = project_root / DEFAULT_INSTRUCTION_DIR_NAME
    files = _excel_files(folder)
    if not files:
        raise FileNotFoundError(f"指令文件夹为空：{folder}。请把当天组合指令 Excel 放进去。")

    tokens = date_match_tokens(asof_date)
    if not tokens:
        raise ValueError("没有有效的运行日期，无法自动识别组合指令文件。请传入 --asof-date，例如 2026-06-24。")

    keyword_files = [p for p in files if "组合指令" in p.stem.replace(" ", "")]
    matched = [p for p in keyword_files if any(token and token in p.stem for token in tokens)]
    if not matched:
        available = ", ".join(sorted(p.name for p in files)) or "（无）"
        raise FileNotFoundError(
            f"在 {folder} 没有找到文件名同时包含‘组合指令’和日期 {asof_date} 的 Excel。"
            f"支持日期格式：{tokens}。当前 Excel：{available}"
        )

    chosen = pick_latest_file(matched)
    assert chosen is not None
    return chosen


def find_latest_oplus_file(project_root: Path, oplus_date: str) -> Path:
    folder = project_root / DEFAULT_OPLUS_DIR_NAME
    files = _excel_files(folder)
    if not files:
        raise FileNotFoundError(f"oplus交易文件夹为空：{folder}。请把 oplus Excel 放进去。")
    tokens = date_match_tokens(oplus_date)
    if not tokens:
        raise ValueError("没有输入有效 oplus 日期，例如 0623 或 20260623。")
    matched = [p for p in files if any(t and t in p.name for t in tokens)]
    if not matched:
        raise FileNotFoundError(
            f"在 {folder} 没有找到日期 {oplus_date} 对应的 oplus 文件。"
            f"支持文件名包含这些日期格式：{tokens}"
        )
    chosen = pick_latest_file(matched)
    assert chosen is not None
    return chosen


def find_template_file(project_root: Path, market: str) -> Path:
    patterns = ["a股买卖清单导入模板*.xlsx", "A股买卖清单导入模板*.xlsx"] if market == "A" else ["港股买卖清单导入模板*.xlsx"]
    candidates: List[Path] = []
    for pat in patterns:
        candidates.extend([p for p in project_root.glob(pat) if p.is_file() and not p.name.startswith("~$")])
    chosen = pick_latest_file(candidates)

    # 用户当前模拟目录只提供一份 8 列“买卖清单导入模板”。
    # A/H 的字段完全一致，因此没有单独港股模板时，港股也复用 A 股模板。
    if not chosen and market == "HK":
        fallback_patterns = ["a股买卖清单导入模板*.xlsx", "A股买卖清单导入模板*.xlsx"]
        fallback: List[Path] = []
        for pat in fallback_patterns:
            fallback.extend([p for p in project_root.glob(pat) if p.is_file() and not p.name.startswith("~$")])
        chosen = pick_latest_file(fallback)

    if not chosen:
        label = "A股" if market == "A" else "港股"
        raise FileNotFoundError(f"没有在 {project_root} 找到 {label}买卖清单导入模板，请放在 live_grid_project 根目录，或用参数手动指定。")
    return chosen


def find_initial_position_file(project_root: Path) -> str:
    patterns = ["初始持仓*.xls*", "*初始*持仓*.xls*", "*持仓*.xls*", "position*.xls*", "positions*.xls*"]
    candidates: List[Path] = []
    excluded_keywords = ["active_grid", "delta_ledger", "买卖清单", "模板", "复核报告"]
    for pat in patterns:
        for p in project_root.glob(pat):
            if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$"):
                if not any(k.lower() in p.name.lower() for k in excluded_keywords):
                    candidates.append(p)
    chosen = pick_latest_file(list(dict.fromkeys(candidates)))
    return str(chosen) if chosen else ""


def resolve_runtime_inputs(args: argparse.Namespace) -> argparse.Namespace:
    root = project_root_from_args(args)
    ensure_live_project_folders(root)
    args.project_root = str(root)

    if not getattr(args, "asof_date", ""):
        args.asof_date = datetime.today().strftime("%Y-%m-%d")
    else:
        args.asof_date = normalize_date_token(args.asof_date)

    if not getattr(args, "instruction_file", ""):
        args.instruction_file = str(find_latest_instruction_file(root, args.asof_date))

    if not getattr(args, "oplus_file", ""):
        oplus_date = getattr(args, "oplus_date", "") or input("请输入要读取的 oplus 日期（如 0623 或 20260623）：").strip()
        args.oplus_date = normalize_date_token(oplus_date, default_year=datetime.strptime(args.asof_date, "%Y-%m-%d").year)
        args.oplus_file = str(find_latest_oplus_file(root, args.oplus_date))

    if not getattr(args, "a_template", ""):
        args.a_template = str(find_template_file(root, "A"))
    if not getattr(args, "hk_template", ""):
        args.hk_template = str(find_template_file(root, "HK"))
    # 默认输出和状态目录放在 live_grid_project 下面，方便每天复用 state。
    if not getattr(args, "output_dir", ""):
        args.output_dir = str(root / "outputs_live_grid")
    if not getattr(args, "state_dir", ""):
        args.state_dir = str(root / "state")

    # 默认采用“固定持仓基线重算”：每次都从 position_file 重新建立仓位，
    # 再完整应用本次指定的 Oplus 文件。只有显式启用
    # --incremental-state-mode 时，才允许沿用 state/current_positions.csv。
    state_position_path = Path(args.state_dir).expanduser().resolve() / "current_positions.csv"
    if not getattr(args, "position_file", ""):
        if getattr(args, "incremental_state_mode", False) and state_position_path.exists():
            args.position_file = ""
        else:
            args.position_file = find_initial_position_file(root)

    print("\n========== 本次读取文件 ==========")
    print(f"项目目录：{root}")
    print(f"运行日期：{args.asof_date}")
    print(f"指令文件：{args.instruction_file}")
    print(f"oplus文件：{args.oplus_file}")
    print(f"A股模板：{args.a_template}")
    print(f"港股模板：{args.hk_template}")
    mode_label = "增量状态模式" if getattr(args, "incremental_state_mode", False) else "固定持仓基线重算模式"
    position_label = args.position_file if args.position_file else "未提供；将使用 state/current_positions.csv"
    print(f"计算模式：{mode_label}")
    print(f"持仓基线：{position_label}")
    print(f"输出目录：{args.output_dir}")
    print(f"状态目录：{args.state_dir}")
    return args

SIM_UNIT_NAME = "单元"
SIM_DEFAULT_PARTICIPATION_RATE = 0.01
# 现实成交回放：默认只占单只股票每分钟市场成交量的 1%，所有网格共享该容量。

# 分产品配置：一次运行分别输出「净值线3号」「聚砾1号」各自的网格，不做合并。
# source_name = 系统/oplus 里维护的真实产品简称，导出模板“产品名称”列必须用它。
PRODUCT_CONFIGS = [
    {
        "key": "净值线3",
        "source_names": {"净值线3号", "巨杉资产净值线3号"},
        "source_name": "净值线3号",
        "product_name": "巨杉网格化",
        "tag": "净值线3",
    },
    {
        "key": "聚砾1",
        "source_names": {"聚砾1号", "巨杉资产聚砾1号"},
        "source_name": "聚砾1号",
        "product_name": "巨杉网格化聚砾",
        "tag": "聚砾",
    },
]
# 以下全局量在运行时由 _activate_product 按产品切换；单产品函数统一引用它们。
SIM_SOURCE_PRODUCT_NAMES = PRODUCT_CONFIGS[0]["source_names"]
SIM_PRODUCT_NAME = PRODUCT_CONFIGS[0]["product_name"]
SIM_SOURCE_NAME = PRODUCT_CONFIGS[0]["source_name"]


def _activate_product(pcfg: dict) -> None:
    """把全局产品名/源产品集合切换到当前正在计算的产品。"""
    global SIM_SOURCE_PRODUCT_NAMES, SIM_PRODUCT_NAME, SIM_SOURCE_NAME
    SIM_SOURCE_PRODUCT_NAMES = pcfg["source_names"]
    SIM_PRODUCT_NAME = pcfg["product_name"]
    SIM_SOURCE_NAME = pcfg["source_name"]
PRODUCT_ALIAS = {
    "巨杉资产净值线1号": "净值线1号",
    "巨杉资产净值线3号": "净值线3号",
    "净值线3号": "净值线3号",
    "巨杉资产聚砾1号": "聚砾1号",
    "聚砾1号": "聚砾1号",
    "巨杉资产鸿6号": "鸿6号",
    "巨杉资产鸿7号": "鸿7号",
    "巨杉资产民生4号": "巨杉资产民生4号",
    "巨杉资产申港证券1号": "申港1号",
    "巨杉资产申港证券3号": "申港3号",
    "海外视野1A号": "海外视野1A号",
    "好买新方程": "新方程",
}

NAV_COL_CANDIDATES = [
    "产品净资产", "净资产", "基金净资产", "资产净值", "产品资产净值",
    "净资产(元)", "净资产（元）", "资产净值(元)", "资产净值（元）",
]
MARKET_VALUE_COL_CANDIDATES = [
    "证券市值", "持仓市值", "股票市值", "市值(元)", "市值（元）", "市值",
]

POSITION_COL_CANDIDATES = [
    "当前仓位", "持仓仓位", "股票仓位", "目标仓位", "仓位", "持仓占净资产",
    "市值/净资产(%)", "市值/净资产（%）", "市值/净资产", "市值/产品净资产(%)", "占净资产比例",
]
CODE_COL_CANDIDATES = ["证券代码", "代码", "stock_code", "code"]
NAME_COL_CANDIDATES = ["证券名称", "名称", "stock_name", "name"]
PRODUCT_COL_CANDIDATES = ["产品名称", "产品", "基金名称", "fund", "product"]
UNIT_COL_CANDIDATES = ["单元名称", "单元", "账户", "券商", "unit"]
MARKET_COL_CANDIDATES = ["交易市场", "市场", "market"]
POSITION_DATE_COL_CANDIDATES = ["日期", "交易日期", "业务日期", "date"]
OPLUS_ORDER_ID_CANDIDATES = ["指令序号", "委托序号", "订单编号", "指令编号", "指令批号", "流水号"]
OPLUS_BUSINESS_DATE_CANDIDATES = ["业务日期", "交易日期", "日期"]
OPLUS_ORDER_TIME_CANDIDATES = ["下达时间", "委托时间", "时间"]


def _is_blank(x: Any) -> bool:
    return x is None or (isinstance(x, float) and np.isnan(x)) or str(x).strip() == ""


def _safe_float(x: Any, default: float = np.nan) -> float:
    if _is_blank(x):
        return default
    try:
        s = str(x).replace("%", "").replace("％", "").replace(",", "").strip()
        return float(s)
    except Exception:
        return default


def _find_col(columns: Iterable[Any], candidates: Sequence[str]) -> Optional[str]:
    cols = [str(c).strip() for c in columns]
    lower = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand in columns:
            return cand
        if cand.lower() in lower:
            return lower[cand.lower()]
    for c in columns:
        cs = str(c)
        if any(k in cs for k in candidates):
            return c
    return None


def _market_suffix(market: Any) -> str:
    """根据持仓/Oplus 的交易市场字段确定证券代码后缀。"""
    s = str(market or "").strip().upper().replace(" ", "")
    if not s or s in {"NAN", "NONE"}:
        return ""
    if "北交所" in s or s in {"BJ", "BSE", "北京"}:
        return ".BJ"
    if "港股" in s or "香港" in s or "港通" in s or s in {"HK", "HKEQ"}:
        return ".HK"
    if "上海" in s or s in {"SH", "SSE"}:
        return ".SH"
    if "深圳" in s or s in {"SZ", "SZSE"}:
        return ".SZ"
    if "美国" in s or s in {"US", "NASDAQ", "NYSE"}:
        return ".US"
    return ""


def canonical_code(code: Any, market: Any = "") -> str:
    if _is_blank(code):
        return ""
    if isinstance(code, (int, np.integer)):
        raw = str(int(code))
    elif isinstance(code, (float, np.floating)) and np.isfinite(code) and float(code).is_integer():
        raw = str(int(code))
    else:
        raw = str(code).strip()
    raw = raw.upper().replace(" ", "").replace("．", ".")
    raw = re.sub(r"\.0+$", "", raw)
    if raw in {"NAN", "NONE"}:
        return ""
    raw = raw.replace("HKEQUITY", ".HK").replace("HKEQ", ".HK")
    raw = raw.replace("SHEQUITY", ".SH").replace("SZEQUITY", ".SZ")
    raw = raw.replace("USEQUITY", ".US")
    raw = re.sub(r"_US$", ".US", raw)
    digits = re.sub(r"\D", "", raw)

    # 持仓查询文件里的证券代码通常没有后缀，因此优先相信“交易市场”列。
    suffix = _market_suffix(market)
    if suffix == ".US":
        ticker = re.sub(r"(?:\.US|_US)$", "", raw)
        return f"{ticker}.US" if ticker else raw
    if suffix and digits:
        width = 5 if suffix == ".HK" else 6
        return f"{digits.zfill(width)}{suffix}"

    if raw.endswith(".HK") or ".HK" in raw:
        return f"{digits.zfill(5)}.HK" if digits else raw
    if raw.endswith(".SH") or ".SH" in raw:
        return f"{digits.zfill(6)}.SH" if digits else raw
    if raw.endswith(".SZ") or ".SZ" in raw:
        return f"{digits.zfill(6)}.SZ" if digits else raw
    if raw.endswith(".BJ") or ".BJ" in raw:
        return f"{digits.zfill(6)}.BJ" if digits else raw
    if raw.endswith((".O", ".N", ".US", ".NASDAQ", ".NYSE")):
        return raw
    if digits:
        if len(digits) <= 5:
            return f"{digits.zfill(5)}.HK"
        if len(digits) == 6:
            if digits.startswith(("43", "83", "87", "88", "92")):
                return f"{digits}.BJ"
            return f"{digits}.SH" if digits.startswith(("5", "6", "9")) else f"{digits}.SZ"
    return raw


def market_group(code: str) -> str:
    code = canonical_code(code)
    if code.endswith(".HK"):
        return "HK"
    if code.endswith((".SH", ".SZ", ".BJ")):
        return "A"
    return "OTHER"


# 港股代码 -> 港股通通道（港股通(沪)/港股通(深)），从持仓文件“交易市场”列构建。
HK_CHANNEL_MAP: Dict[str, str] = {}


def trade_market(code: str) -> str:
    code = canonical_code(code)
    if code.endswith(".SH"):
        return "上海"
    if code.endswith(".SZ"):
        return "深圳"
    if code.endswith(".BJ"):
        return "北交所"
    if code.endswith(".HK"):
        # 模板要求港股必须是“港股通(沪)”或“港股通(深)”，不能只写“港股”。
        return HK_CHANNEL_MAP.get(code, "港股通(沪)")
    return ""


def trade_direction(side: str) -> str:
    return "买入" if str(side).lower() == "buy" else "卖出"


def normalize_product_name(x: Any) -> str:
    s = str(x or "").strip()
    return PRODUCT_ALIAS.get(s, s)


def is_sim_product(x: Any) -> bool:
    """本模拟底层持仓 = 净值线3号 + 聚砾1号，在系统内统一命名为“巨杉网格化聚砾”。"""
    return normalize_product_name(x) in SIM_SOURCE_PRODUCT_NAMES


def normalize_unit_name(x: Any) -> str:
    """单产品模拟不再区分真实交易单元，所有净值线3号仓位/成交统一归并到“单元”。"""
    return SIM_UNIT_NAME


def convert_xls_to_xlsx(path: Path) -> Path:
    """把老 .xls 转成临时 .xlsx，优先用 Windows Excel COM；其次尝试 libreoffice。"""
    path = Path(path).resolve()
    tmpdir = Path(tempfile.mkdtemp(prefix="xls_convert_"))
    out = tmpdir / (path.stem + ".xlsx")

    # Windows + Excel：用户实盘环境最稳。
    try:
        import win32com.client  # type: ignore
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(path))
        wb.SaveAs(str(out), FileFormat=51)  # xlsx
        wb.Close(False)
        excel.Quit()
        if out.exists():
            return out
    except Exception:
        pass

    # 服务器/本机装了 libreoffice 的 fallback。
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(tmpdir), str(path)], check=True)
        if out.exists():
            return out

    raise RuntimeError(
        f"无法读取 .xls 文件：{path}。请在 Windows + Excel 环境运行，或先手工另存为 .xlsx。"
    )


def read_excel_any(path: str | Path, sheet_name: Any = 0, header: Any = 0, dtype: Any = None) -> Any:
    path = Path(path)
    if not path.exists():
        # 容错：命令行里路径分隔符两侧误加空格（如 "持仓数据 / xxx.xls"）时自动修正。
        # 目录/文件名正常不会以空格结尾，仅在去空格版本确实存在时才采用。
        stripped = Path(*[p.strip() for p in path.parts])
        if stripped.exists():
            path = stripped
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header, dtype=dtype)
    except Exception as e:
        if path.suffix.lower() == ".xls":
            converted = convert_xls_to_xlsx(path)
            return pd.read_excel(converted, sheet_name=sheet_name, header=header, dtype=dtype, engine="openpyxl")
        raise e


def read_template_accounts(template_path: str | Path) -> pd.DataFrame:
    wb = load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    headers = [ws.cell(1, c).value for c in range(1, 9)]
    if headers != TEMPLATE_HEADERS:
        raise ValueError(f"模板表头不匹配：{template_path}，当前={headers}，期望={TEMPLATE_HEADERS}")
    rows = []
    for r in range(2, ws.max_row + 1):
        product = ws.cell(r, 3).value
        unit = ws.cell(r, 4).value
        if not _is_blank(product) and not _is_blank(unit):
            rows.append({"product": normalize_product_name(product), "unit": normalize_unit_name(unit)})
    if not rows:
        raise ValueError(f"模板 {template_path} 没有读到产品/单元。")
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


def load_template_accounts(a_template: str | Path, hk_template: str | Path) -> pd.DataFrame:
    """
    模拟环境固定为一个产品、一个单元：
        净值线3号 -> 巨杉网格化 / 单元

    交易模板只负责提供格式，不再从模板中的产品/单元行推导账户，
    因而不会因为模板仍保留历史产品名称而把网格拆成多个单元。
    """
    # 仍读取一次模板，尽早检查模板路径/表头是否有效。
    read_template_accounts(a_template)
    read_template_accounts(hk_template)
    return pd.DataFrame([{
        "product": SIM_PRODUCT_NAME,
        "unit": SIM_UNIT_NAME,
    }])


def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int = 8) -> None:
    for col in range(1, max_col + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
        if src.hyperlink:
            dst._hyperlink = copy(src.hyperlink)
        if src.comment:
            dst.comment = copy(src.comment)
    if ws.row_dimensions[src_row].height is not None:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def export_import_template(orders_df: pd.DataFrame, template_path: str | Path, output_path: str | Path, freeze: bool = True) -> Path:
    """导出买卖清单；证券代码去除市场后缀，“数量/仓位”按每档 target_pos（目标仓位）填写。"""
    template_path = Path(template_path)
    output_path = Path(output_path)
    wb = load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    headers = [ws.cell(1, c).value for c in range(1, 9)]
    if headers != TEMPLATE_HEADERS:
        raise ValueError(f"模板表头不匹配：{headers}")
    style_row = 2 if ws.max_row >= 2 else 1

    # 清空旧数据区域，后面重写。产品/单元也一起重写，避免多档位订单行数超过模板账号行数。
    for r in range(2, ws.max_row + 1):
        for c in range(1, 9):
            ws.cell(r, c).value = None

    rows = []
    if orders_df is not None and not orders_df.empty:
        for _, o in orders_df.iterrows():
            # 导入模拟环境：产品/单元/方向/极限价格保持原样，
            # 仅“交易市场”港股改为港股通(沪/深)、“数量/仓位”改为百分数（三位小数）。
            rows.append([
                str(canonical_code(o.get("code"))).split(".", 1)[0],
                o.get("name", ""),
                SIM_PRODUCT_NAME,
                SIM_UNIT_NAME,
                trade_market(o.get("code", "")),
                trade_direction(o.get("side", "")),
                # 模板按目标仓位：填百分数数值、保留三位小数（例 23.751 表示 23.751%）。
                round(float(o.get("target_pos", 0.0)) * 100.0, 3),
                float(o.get("price", 0.0)),
            ])

    needed = max(len(rows), 1)
    for r in range(ws.max_row + 1, needed + 2):
        _copy_row_style(ws, style_row, r, 8)

    for idx, values in enumerate(rows, start=2):
        _copy_row_style(ws, style_row, idx, 8)
        for c, v in enumerate(values, start=1):
            ws.cell(idx, c).value = v

    # 多余行清空
    for r in range(len(rows) + 2, ws.max_row + 1):
        for c in range(1, 9):
            ws.cell(r, c).value = None

    ws.freeze_panes = "A2" if freeze else None
    # 模板保存时视图可能停在靠后的行（港股模板 topLeftCell=A11），冻结首行后会导致
    # 打开时前十几行被滚到冻结区上方、看起来“前十几行消失/第一行不显示”。
    # 清除 sheetView 自身的视图位置，交给 pane 的 topLeftCell 决定顶部。
    ws.sheet_view.topLeftCell = None
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path



def manager_display_code(code: Any) -> str:
    """Display HK codes in the manager/Wind-friendly form, e.g. 00941.HK -> 0941.HK.

    Internal matching continues to use canonical 5-digit HK codes.  This function is
    for human-readable manager-style output only.
    """
    c = canonical_code(code)
    if c.endswith(".HK"):
        stem = c[:-3]
        if len(stem) == 5 and stem.startswith("0"):
            return f"{stem[1:]}.HK"
    return c


def _fmt_compact_number(x: Any, decimals: int = 4) -> str:
    try:
        v = float(x)
    except Exception:
        return ""
    if not np.isfinite(v):
        return ""
    text = f"{v:.{decimals}f}".rstrip("0").rstrip(".")
    return text or "0"


def _manager_instruction_text(side: str, price: float, target_pos: float) -> str:
    price_text = _fmt_compact_number(price, 4)
    pct_text = _fmt_compact_number(float(target_pos) * 100.0, 4)
    if str(side).lower() == "buy":
        return f"{price_text}元以下买入到{pct_text}%"
    if float(target_pos) <= 1e-12:
        return f"{price_text}元以上卖完"
    return f"{price_text}元以上卖出到{pct_text}%"


def export_manager_instruction_style(active_grid: pd.DataFrame, output_path: str | Path) -> Path:
    """Export active grid in a fund-manager-style instruction workbook.

    Each product + unit + security uses up to two rows, matching the original
    manager instruction reading pattern: buy levels on the first row and sell
    levels on the following row.  Generated target positions are preserved per
    product/unit, so different accounts are never collapsed together.
    """
    output_path = Path(output_path)
    wb = Workbook()
    ws = wb.active
    ws.title = "股票指令"

    base_headers = ["产品名称", "单元名称", "code", "证券名称", "当前仓位", "参考价"]
    groups = []
    max_levels = 1
    if active_grid is not None and not active_grid.empty:
        df = active_grid.copy()
        for c in ["price", "target_pos", "current_pos_before_level", "reference_price"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        sort_cols = [c for c in ["product", "unit", "code", "side", "price"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(sort_cols).reset_index(drop=True)
        group_cols = ["product", "unit", "code"]
        for keys, g in df.groupby(group_cols, sort=True, dropna=False):
            product, unit, code = keys
            name_series = g.get("name", pd.Series(dtype=object)).dropna().astype(str)
            name = name_series.iloc[0] if not name_series.empty else ""
            ref_series = pd.to_numeric(g.get("reference_price", pd.Series(dtype=float)), errors="coerce").dropna()
            ref_price = float(ref_series.iloc[0]) if not ref_series.empty else np.nan

            buy = g[g["side"].astype(str).str.lower() == "buy"].sort_values("price", ascending=False)
            sell = g[g["side"].astype(str).str.lower() == "sell"].sort_values("price", ascending=True)
            # Each side starts from the same account-level current position.  For buys the
            # first executable level is the highest price; for sells it is the lowest price.
            # Reading that row avoids accidentally reporting an intermediate ladder position.
            current_pos = np.nan
            start_table = buy if not buy.empty else sell
            if not start_table.empty and "current_pos_before_level" in start_table.columns:
                cur = pd.to_numeric(start_table["current_pos_before_level"], errors="coerce").dropna()
                if not cur.empty:
                    current_pos = float(cur.iloc[0])
            buy_texts = [_manager_instruction_text("buy", r["price"], r["target_pos"]) for _, r in buy.iterrows()]
            sell_texts = [_manager_instruction_text("sell", r["price"], r["target_pos"]) for _, r in sell.iterrows()]
            max_levels = max(max_levels, len(buy_texts), len(sell_texts))
            groups.append({
                "product": str(product or ""), "unit": str(unit or ""), "code": manager_display_code(code),
                "name": name, "current_pos": current_pos, "reference_price": ref_price,
                "buy": buy_texts, "sell": sell_texts,
            })

    headers = base_headers + [f"指令{i}档" for i in range(1, max_levels + 1)]
    ws.append(headers)

    # Header style, intentionally close to the simple manager instruction workbook.
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="D9D9D9")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)

    for item in groups:
        first_written = False
        for side_key in ["buy", "sell"]:
            texts = item[side_key]
            if not texts:
                continue
            identity = not first_written
            row = [
                item["product"] if identity else "",
                item["unit"] if identity else "",
                item["code"] if identity else "",
                item["name"] if identity else "",
                item["current_pos"] if identity and np.isfinite(item["current_pos"]) else "",
                item["reference_price"] if identity and np.isfinite(item["reference_price"]) else "",
            ] + texts + [""] * (max_levels - len(texts))
            ws.append(row)
            first_written = True

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(1, len(headers)).column_letter}{ws.max_row}"
    widths = {"A": 16, "B": 14, "C": 13, "D": 18, "E": 12, "F": 12}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for col_idx in range(7, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = 28
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    if ws.max_row >= 2:
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(2, ws.max_row + 1):
            ws.row_dimensions[r].height = 22
        # Fractional positions are displayed as percentages, reference prices as numbers.
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 5).number_format = "0.00%"
            ws.cell(r, 6).number_format = "0.0000"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path

def split_orders_by_market(orders: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if orders is None or orders.empty:
        return pd.DataFrame(), pd.DataFrame()
    df = orders.copy()
    df["market_group"] = df["code"].apply(market_group)
    a = df[df["market_group"] == "A"].drop(columns=["market_group"], errors="ignore")
    hk = df[df["market_group"] == "HK"].drop(columns=["market_group"], errors="ignore")
    return a, hk


def _stable_trade_key(parts: Sequence[Any]) -> str:
    text = "|".join(str(x or "").strip() for x in parts)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:24]


def load_oplus_trades(oplus_file: str | Path) -> Tuple[pd.DataFrame, Dict[str, float]]:
    """读取 Oplus 指令成交进度。

    返回的 trade_delta_weight 是每条 Oplus 指令截至当前文件的“累计已成交仓位”，
    main() 会再和 state/oplus_trade_state.csv 比较，只把新增成交部分计入持仓。
    """
    if not oplus_file:
        return pd.DataFrame(), {}
    df = read_excel_any(oplus_file, dtype={"证券代码": str})
    if isinstance(df, dict):
        df = next(iter(df.values()))
    if df is None or df.empty:
        return pd.DataFrame(), {}

    code_col = _find_col(df.columns, CODE_COL_CANDIDATES)
    name_col = _find_col(df.columns, NAME_COL_CANDIDATES)
    product_col = _find_col(df.columns, PRODUCT_COL_CANDIDATES)
    unit_col = _find_col(df.columns, UNIT_COL_CANDIDATES)
    direction_col = _find_col(df.columns, ["委托方向", "交易方向", "方向"])
    ratio_col = _find_col(df.columns, ["占净资产比例", "保证金占用/产品净资产", "市值/产品净资产(%)"])
    actual_col = _find_col(df.columns, ["成交进度", "实际进度"])
    filled_qty_col = _find_col(df.columns, ["成交数量", "已成交数量"])
    order_qty_col = _find_col(df.columns, ["指令数量", "委托数量", "目标数量"])
    price_col = _find_col(df.columns, ["最新价", "现价", "价格"])
    status_col = _find_col(df.columns, ["指令状态", "状态"])
    remark_col = _find_col(df.columns, ["备注", "说明"])
    order_id_col = _find_col(df.columns, OPLUS_ORDER_ID_CANDIDATES)
    business_date_col = _find_col(df.columns, OPLUS_BUSINESS_DATE_CANDIDATES)
    order_time_col = _find_col(df.columns, OPLUS_ORDER_TIME_CANDIDATES)
    market_col = _find_col(df.columns, MARKET_COL_CANDIDATES)

    required = [code_col, product_col, direction_col, ratio_col]
    if any(c is None for c in required):
        raise ValueError(f"oplus 缺少必要列。当前列={list(df.columns)}")

    rows = []
    latest_price: Dict[str, float] = {}
    for row_no, (_, r) in enumerate(df.iterrows(), start=2):
        code = canonical_code(r.get(code_col), market=r.get(market_col, "") if market_col else "")
        if not code or code == "混合" or market_group(code) == "OTHER":
            continue

        # 最新价与产品行无关，混合汇总行以外的任意有效证券行都可提供参考价。
        px = _safe_float(r.get(price_col), np.nan) if price_col else np.nan
        if np.isfinite(px) and px > 0:
            latest_price[code] = float(px)

        product = normalize_product_name(r.get(product_col))
        unit = normalize_unit_name(r.get(unit_col)) if unit_col else ""
        # 只把净值线3号（系统内映射为“巨杉网格化”）的成交计入模拟持仓。
        if product != SIM_PRODUCT_NAME:
            continue
        direction_raw = str(r.get(direction_col, "")).strip()
        if "买" in direction_raw:
            direction = "买入"
        elif "卖" in direction_raw or "清仓" in direction_raw:
            direction = "卖出"
        else:
            direction = ""
        if product in {"", "混合"} or not direction:
            continue

        ratio = _safe_float(r.get(ratio_col), np.nan)
        if not np.isfinite(ratio):
            continue
        if ratio > 1.0 and ratio <= 100.0:
            ratio = ratio / 100.0
        ratio = abs(float(ratio))

        actual_raw = r.get(actual_col) if actual_col else np.nan
        actual = _safe_float(actual_raw, np.nan)
        status = str(r.get(status_col, "") if status_col else "").strip()
        if np.isfinite(actual):
            progress = float(actual) / 100.0
        else:
            filled = _safe_float(r.get(filled_qty_col), np.nan) if filled_qty_col else np.nan
            ordered = _safe_float(r.get(order_qty_col), np.nan) if order_qty_col else np.nan
            if np.isfinite(filled) and np.isfinite(ordered) and ordered > 0:
                progress = float(filled / ordered)
            elif any(k in status for k in ["已完成", "完成", "全部成交"]):
                progress = 1.0
            else:
                # 缺少成交证据时不再默认 100%，避免把未成交委托计入持仓。
                progress = 0.0
        progress = float(np.clip(progress, 0.0, 1.0))

        signed_delta = ratio * progress * (1.0 if direction == "买入" else -1.0)
        order_id = str(r.get(order_id_col, "") if order_id_col else "").strip()
        business_date = str(r.get(business_date_col, "") if business_date_col else "").strip()
        order_time = str(r.get(order_time_col, "") if order_time_col else "").strip()
        # 优先用 Oplus 指令序号；缺失时加入日期、时间和原始行号，至少保证同文件重跑稳定。
        identity = order_id or f"row={row_no}"
        trade_key = _stable_trade_key([
            business_date, order_time, identity, product, unit, code, direction,
            round(ratio, 10), str(r.get(remark_col, "") if remark_col else "").strip(),
        ])
        rows.append({
            "trade_key": trade_key,
            "business_date": business_date,
            "order_id": order_id,
            "product": product,
            "unit": unit,
            "code": code,
            "name": r.get(name_col, "") if name_col else "",
            "side": "buy" if direction == "买入" else "sell",
            "trade_delta_weight": float(signed_delta),
            "raw_ratio": ratio,
            "progress": progress,
            "status": status,
            "remark": r.get(remark_col, "") if remark_col else "",
        })
    trades = pd.DataFrame(rows)
    if not trades.empty:
        # 同一 trade_key 如在文件中重复出现，以最后一条累计进度为准，不能求和。
        trades = trades.drop_duplicates(subset=["trade_key"], keep="last").reset_index(drop=True)
    return trades, latest_price


def load_oplus_trade_state(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=["trade_key", "applied_delta_weight"])
    state = pd.read_csv(path, dtype={"trade_key": str, "code": str})
    if "trade_key" not in state.columns:
        return pd.DataFrame(columns=["trade_key", "applied_delta_weight"])
    state["applied_delta_weight"] = pd.to_numeric(state.get("applied_delta_weight"), errors="coerce").fillna(0.0)
    return state.drop_duplicates(subset=["trade_key"], keep="last")


def incremental_oplus_trades(
    current_trades: pd.DataFrame,
    previous_state: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """把 Oplus 累计成交进度转换成相对上次运行的新增成交。"""
    if current_trades is None or current_trades.empty:
        return pd.DataFrame(columns=list(current_trades.columns) if current_trades is not None else []), previous_state.copy()
    cur = current_trades.copy()
    prev = previous_state[["trade_key", "applied_delta_weight"]].copy() if previous_state is not None and not previous_state.empty else pd.DataFrame(columns=["trade_key", "applied_delta_weight"])
    cur = cur.merge(prev, on="trade_key", how="left")
    cur["applied_delta_weight"] = pd.to_numeric(cur["applied_delta_weight"], errors="coerce").fillna(0.0)
    cur["cumulative_delta_weight"] = pd.to_numeric(cur["trade_delta_weight"], errors="coerce").fillna(0.0)
    cur["trade_delta_weight"] = cur["cumulative_delta_weight"] - cur["applied_delta_weight"]
    incremental = cur[cur["trade_delta_weight"].abs() > 1e-12].copy()

    state_cols = [c for c in current_trades.columns if c != "trade_delta_weight"]
    current_state = current_trades.copy().rename(columns={"trade_delta_weight": "applied_delta_weight"})
    current_state = current_state[["trade_key", "applied_delta_weight", *[c for c in state_cols if c != "trade_key"]]]
    if previous_state is not None and not previous_state.empty:
        keep_old = previous_state[~previous_state["trade_key"].isin(set(current_state["trade_key"]))].copy()
        next_state = pd.concat([keep_old, current_state], ignore_index=True, sort=False)
    else:
        next_state = current_state
    next_state = next_state.drop_duplicates(subset=["trade_key"], keep="last")
    return incremental, next_state


def save_csv_atomic(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    df.to_csv(tmp, index=False, encoding="utf-8-sig")
    tmp.replace(path)


def _position_value_to_fraction(value: Any, column_name: Any) -> float:
    """把持仓文件中的仓位统一成 0~1 小数。

    “全球产品持仓查询”中的 `市值/净资产(%)` 是百分数点：
    2.55 表示 2.55%，0.28 表示 0.28%，所以必须一律除以 100。
    其他不带百分号的通用仓位列继续兼容 0.05 / 5 / 5% 三种写法。
    """
    if _is_blank(value):
        return np.nan
    raw = str(value).strip()
    v = _safe_float(value, np.nan)
    if not np.isfinite(v):
        return np.nan
    col = str(column_name or "").replace(" ", "").replace("（", "(").replace("）", ")")
    if "%" in raw or "％" in raw:
        return float(v / 100.0)
    if "%" in col or "％" in col:
        return float(v / 100.0)
    if "市值/净资产" in col or "市值/产品净资产" in col:
        return float(v / 100.0)
    if abs(v) > 1.0:
        return float(v / 100.0)
    return float(v)


def _normalize_position_sheet(
    df: pd.DataFrame,
    accounts: pd.DataFrame,
    sheet_name: str = "",
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """把一张持仓表规范成 product/unit/code/current_pos。"""
    empty = pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])
    stats: Dict[str, Any] = {
        "sheet": sheet_name,
        "input_rows": int(len(df)) if df is not None else 0,
        "loaded_rows": 0,
        "skipped_missing_position": 0,
        "skipped_invalid_key": 0,
        "skipped_other_product": 0,
        "latest_date": "",
    }
    if df is None or df.empty:
        return empty, stats

    code_col = _find_col(df.columns, CODE_COL_CANDIDATES)
    name_col = _find_col(df.columns, NAME_COL_CANDIDATES)
    product_col = _find_col(df.columns, PRODUCT_COL_CANDIDATES)
    unit_col = _find_col(df.columns, UNIT_COL_CANDIDATES)
    pos_col = _find_col(df.columns, POSITION_COL_CANDIDATES)
    market_col = _find_col(df.columns, MARKET_COL_CANDIDATES)
    date_col = _find_col(df.columns, POSITION_DATE_COL_CANDIDATES)

    stats.update({
        "code_col": str(code_col or ""),
        "product_col": str(product_col or ""),
        "unit_col": str(unit_col or ""),
        "position_col": str(pos_col or ""),
        "market_col": str(market_col or ""),
    })
    if code_col is None or product_col is None or pos_col is None:
        stats["skip_reason"] = "missing required columns"
        return empty, stats

    work = df.copy()
    # 如果同一工作表意外包含多个日期，只使用最新持仓快照，避免跨日重复相加。
    if date_col is not None:
        parsed_dates = pd.to_datetime(work[date_col], errors="coerce")
        if parsed_dates.notna().any():
            latest = parsed_dates.max()
            work = work.loc[parsed_dates == latest].copy()
            stats["latest_date"] = latest.strftime("%Y-%m-%d")

    rows: List[Dict[str, Any]] = []
    for _, r in work.iterrows():
        market_value = r.get(market_col, "") if market_col else ""
        code = canonical_code(r.get(code_col), market=market_value)
        # 记录港股通通道（沪/深），供导出模板的“交易市场”列使用。
        if code.endswith(".HK") and market_col:
            mv_text = str(market_value).strip()
            if mv_text in ("港股通(沪)", "港股通(深)"):
                HK_CHANNEL_MAP[code] = mv_text
        # 用源产品集合判断（净值线3号 / 聚砾1号），系统内统一命名为 SIM_PRODUCT_NAME。
        if not is_sim_product(r.get(product_col)):
            stats["skipped_other_product"] += 1
            continue
        product = SIM_PRODUCT_NAME
        unit = normalize_unit_name(r.get(unit_col)) if unit_col else ""
        pos = _position_value_to_fraction(r.get(pos_col), pos_col)
        if not np.isfinite(pos):
            stats["skipped_missing_position"] += 1
            continue
        if not code or not product:
            stats["skipped_invalid_key"] += 1
            continue
        rows.append({
            "product": product,
            "unit": unit,
            "code": code,
            "name": str(r.get(name_col, "") or "") if name_col else "",
            "current_pos": float(np.clip(pos, 0.0, 1.0)),
        })

    if not rows:
        return empty, stats
    out = pd.DataFrame(rows)

    # 单产品模拟统一归并到固定“单元”；保留这段兼容旧状态文件。
    no_unit = out["unit"].astype(str).str.strip() == ""
    if no_unit.any():
        expanded = []
        for _, row in out[no_unit].iterrows():
            units = accounts[accounts["product"] == row["product"]]["unit"].drop_duplicates().tolist()
            for u in units or [""]:
                x = row.to_dict()
                x["unit"] = u
                expanded.append(x)
        out = pd.concat([out[~no_unit], pd.DataFrame(expanded)], ignore_index=True)

    # 如果同一产品/单元/证券被拆成多行，仓位应相加，而不是任意取最后一行。
    out = out.groupby(["product", "unit", "code"], as_index=False).agg({
        "name": "last",
        "current_pos": "sum",
    })
    out["current_pos"] = pd.to_numeric(out["current_pos"], errors="coerce").fillna(0.0).clip(0.0, 1.0)
    stats["loaded_rows"] = int(len(out))
    return out, stats

def load_position_file(path: Optional[str | Path], accounts: pd.DataFrame) -> pd.DataFrame:
    if not path:
        return pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"找不到 position file: {p}")
    sheets = read_excel_any(p, sheet_name=None)
    frames = []
    if not isinstance(sheets, dict):
        sheets = {"Sheet1": sheets}
    load_stats: List[Dict[str, Any]] = []
    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue
        normalized, stats = _normalize_position_sheet(df, accounts, str(sheet_name))
        load_stats.append(stats)
        if not normalized.empty:
            frames.append(normalized)
    if not frames:
        details = "; ".join(
            f"{x.get('sheet')}: {x.get('skip_reason', '0 rows')}"
            for x in load_stats
        )
        print(f"WARNING: 没有从持仓文件读取到有效仓位。{details}")
        return pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])
    out = pd.concat(frames, ignore_index=True)
    out = out.groupby(["product", "unit", "code"], as_index=False).agg({"name": "last", "current_pos": "sum"})
    out["current_pos"] = pd.to_numeric(out["current_pos"], errors="coerce").fillna(0.0).clip(0.0, 1.0)
    loaded = sum(int(x.get("loaded_rows", 0)) for x in load_stats)
    missing_pos = sum(int(x.get("skipped_missing_position", 0)) for x in load_stats)
    used_sheets = [x for x in load_stats if int(x.get("loaded_rows", 0)) > 0]
    print(
        f"Loaded position file: {p.name}; rows={len(out)}; "
        f"source_rows={loaded}; missing_position_skipped={missing_pos}; "
        f"sheets={[x.get('sheet') for x in used_sheets]}"
    )
    return out


def _find_cny_mv_col(columns: Iterable[Any]) -> Optional[str]:
    """找“持仓市值（人民币）”列（统一折算后的市值，港股/美股可直接相加）。"""
    for c in columns:
        cs = str(c).replace("（", "(").replace("）", ")")
        if "人民币" in cs and "市值" in cs and "/" not in cs:
            return c
    return None


def _infer_product_nav_rows(work: pd.DataFrame, mv_cny_col: str, pct_col: str) -> float:
    """用 人民币市值 / (市值/净资产%) 反推单个产品净资产，取稳健中位数。"""
    mv = pd.to_numeric(work[mv_cny_col], errors="coerce")
    pct = pd.to_numeric(work[pct_col], errors="coerce")
    est = mv / (pct / 100.0)
    est = est[(pct > 0.1) & np.isfinite(est) & (est > 0)]
    if est.empty:
        return np.nan
    if len(est) >= 5:  # 去掉极端值，避免个别脏数据拉偏
        lo, hi = np.nanquantile(est, [0.1, 0.9])
        est = est[(est >= lo) & (est <= hi)]
    return float(np.nanmedian(est))


def load_combined_position_file(
    path: Optional[str | Path],
    accounts: pd.DataFrame,
) -> Tuple[pd.DataFrame, float, Dict[str, float]]:
    """读取「净值线3号 + 聚砾1号」合并持仓（市值加权口径）。

    关键：两个产品的“市值/净资产(%)”分母不同，绝不能直接相加。
    做法：
      1) 分别反推每个源产品的真实净资产 nav_i；
      2) 取每个证券在各产品中的“人民币市值”直接相加；
      3) 合并仓位 = 人民币市值合计 / 合并净资产。
    返回 (positions, combined_nav, {源产品: 净资产})。
    """
    empty_cols = ["product", "unit", "code", "name", "current_pos"]
    if not path:
        return pd.DataFrame(columns=empty_cols), np.nan, {}
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"找不到 position file: {p}")
    sheets = read_excel_any(p, sheet_name=None)
    if not isinstance(sheets, dict):
        sheets = {"Sheet1": sheets}

    # 源产品 -> 该产品各证券 {code: (name, mv_cny)}
    product_mv: Dict[str, Dict[str, Tuple[str, float]]] = {}
    product_nav: Dict[str, float] = {}

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue
        product_col = _find_col(df.columns, PRODUCT_COL_CANDIDATES)
        code_col = _find_col(df.columns, CODE_COL_CANDIDATES)
        name_col = _find_col(df.columns, NAME_COL_CANDIDATES)
        market_col = _find_col(df.columns, MARKET_COL_CANDIDATES)
        date_col = _find_col(df.columns, POSITION_DATE_COL_CANDIDATES)
        pct_col = _find_col(df.columns, POSITION_COL_CANDIDATES)
        mv_cny_col = _find_cny_mv_col(df.columns)
        if code_col is None or product_col is None or pct_col is None or mv_cny_col is None:
            continue

        work = df.copy()
        if date_col is not None:
            dt = pd.to_datetime(work[date_col], errors="coerce")
            if dt.notna().any():
                work = work.loc[dt == dt.max()].copy()
        work["_src"] = work[product_col].apply(normalize_product_name)
        work = work[work["_src"].isin(SIM_SOURCE_PRODUCT_NAMES)]

        for src, grp in work.groupby("_src"):
            nav = _infer_product_nav_rows(grp, mv_cny_col, pct_col)
            if np.isfinite(nav):
                # 同一产品可能在多个 sheet 出现；净资产取较大（更全）口径
                product_nav[src] = max(nav, product_nav.get(src, 0.0))
            mv_map = product_mv.setdefault(src, {})
            for _, r in grp.iterrows():
                code = canonical_code(r.get(code_col), market=r.get(market_col))
                if not code:
                    continue
                mv = _safe_float(r.get(mv_cny_col), 0.0)
                nm = str(r.get(name_col, "") or "")
                if code in mv_map:
                    old_nm, old_mv = mv_map[code]
                    mv_map[code] = (nm or old_nm, old_mv + mv)
                else:
                    mv_map[code] = (nm, mv)

    if not product_nav:
        print("WARNING: 未能从持仓文件反推出净值线3号/聚砾1号净资产，合并持仓为空。")
        return pd.DataFrame(columns=empty_cols), np.nan, {}

    combined_nav = float(sum(product_nav.values()))
    # 合并各产品证券市值
    merged: Dict[str, Tuple[str, float]] = {}
    for src, mv_map in product_mv.items():
        for code, (nm, mv) in mv_map.items():
            if code in merged:
                old_nm, old_mv = merged[code]
                merged[code] = (nm or old_nm, old_mv + mv)
            else:
                merged[code] = (nm, mv)

    rows = []
    for code, (nm, mv) in merged.items():
        pos = float(np.clip(mv / combined_nav, 0.0, 1.0))
        rows.append({"product": SIM_PRODUCT_NAME, "unit": SIM_UNIT_NAME,
                     "code": code, "name": nm, "current_pos": pos})
    out = pd.DataFrame(rows).sort_values("code").reset_index(drop=True)
    print(
        f"Loaded COMBINED position: {p.name}; rows={len(out)}; "
        f"合并净资产={combined_nav:,.0f}；分产品净资产="
        + "，".join(f"{k}={v:,.0f}" for k, v in product_nav.items())
    )
    return out, combined_nav, product_nav


def load_previous_positions(path: Path) -> pd.DataFrame:
    if path.exists():
        df = pd.read_csv(path, dtype={"code": str})
        if not df.empty:
            df["code"] = df["code"].apply(canonical_code)
            df["product"] = df["product"].apply(normalize_product_name)
            df = df[df["product"] == SIM_PRODUCT_NAME].copy()
            df["unit"] = df["unit"].apply(normalize_unit_name)
            df["current_pos"] = pd.to_numeric(df["current_pos"], errors="coerce").fillna(0.0)
            return df
    return pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])


def build_current_positions(
    accounts: pd.DataFrame,
    position_file: Optional[str | Path],
    previous_position_path: Path,
    oplus_trades: pd.DataFrame,
    apply_oplus_trades: bool = True,
    reset_position_state: bool = False,
    use_previous_state: bool = False,
) -> pd.DataFrame:
    """构建本次计算使用的最新仓位。

    默认固定基线重算：始终读取 position_file，再完整应用本次传入的 Oplus。
    只有 use_previous_state=True 且没有要求 reset 时，才沿用
    state/current_positions.csv 做跨日增量更新。
    """
    previous = load_previous_positions(previous_position_path)
    if use_previous_state and not reset_position_state and not previous.empty:
        base = previous
        if position_file:
            print(
                f"Using existing position state: {previous_position_path}; "
                f"position file {Path(position_file).name} is treated as initialization-only."
            )
    else:
        base = load_position_file(position_file, accounts) if position_file else previous
    if base.empty:
        base = pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])

    if apply_oplus_trades and oplus_trades is not None and not oplus_trades.empty:
        t = oplus_trades.copy()
        # 单产品模拟所有 Oplus 成交统一归并到固定“单元”；这里仅兼容旧状态文件的空单元。
        no_unit = t["unit"].astype(str).str.strip() == ""
        if no_unit.any():
            expanded = []
            for _, row in t[no_unit].iterrows():
                units = accounts[accounts["product"] == row["product"]]["unit"].drop_duplicates().tolist()
                for u in units or [""]:
                    x = row.to_dict(); x["unit"] = u; expanded.append(x)
            t = pd.concat([t[~no_unit], pd.DataFrame(expanded)], ignore_index=True)
        delta = t.groupby(["product", "unit", "code"], as_index=False).agg({
            "name": "last", "trade_delta_weight": "sum"
        })
        if base.empty:
            base = delta.rename(columns={"trade_delta_weight": "current_pos"})
            base["current_pos"] = base["current_pos"].clip(lower=0.0, upper=1.0)
        else:
            base = base.merge(delta, on=["product", "unit", "code"], how="outer", suffixes=("", "_trade"))
            base["name"] = base["name"].combine_first(base.get("name_trade"))
            base["current_pos"] = pd.to_numeric(base["current_pos"], errors="coerce").fillna(0.0)
            base["trade_delta_weight"] = pd.to_numeric(base["trade_delta_weight"], errors="coerce").fillna(0.0)
            base["current_pos"] = (base["current_pos"] + base["trade_delta_weight"]).clip(lower=0.0, upper=1.0)
            base = base.drop(columns=[c for c in ["name_trade", "trade_delta_weight"] if c in base.columns])

    if base.empty:
        return pd.DataFrame(columns=["product", "unit", "code", "name", "current_pos"])
    base["product"] = base["product"].apply(normalize_product_name)
    base["unit"] = base["unit"].apply(normalize_unit_name)
    base["code"] = base["code"].apply(canonical_code)
    base["current_pos"] = pd.to_numeric(base["current_pos"], errors="coerce").fillna(0.0).clip(lower=0.0, upper=1.0)
    return base.groupby(["product", "unit", "code"], as_index=False).agg({"name": "last", "current_pos": "last"})


def levels_json(levels: Sequence[Dict[str, Any]]) -> str:
    clean = []
    for x in levels or []:
        clean.append({
            "price": round(float(x.get("price", np.nan)), 8),
            "target_pos": None if pd.isna(x.get("target_pos")) else round(float(x.get("target_pos")), 8),
            "order_basis": str(x.get("order_basis", "target")),
            "delta_pos": None if pd.isna(x.get("delta_pos", np.nan)) else round(float(x.get("delta_pos")), 8),
            "raw_instruction": str(x.get("raw_instruction", "")),
        })
    return json.dumps(clean, ensure_ascii=False, sort_keys=True)


def instruction_hash(row: pd.Series) -> str:
    raw = {
        "code": canonical_code(row.get("code")),
        "buy_levels": json.loads(levels_json(row.get("buy_levels") or [])),
        "sell_levels": json.loads(levels_json(row.get("sell_levels") or [])),
        "unparsed": row.get("unparsed_instructions") or [],
    }
    s = json.dumps(raw, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:16]


def build_instruction_snapshot(instruction_file: str | Path, include_sheets: Optional[List[str]]) -> pd.DataFrame:
    instruction_path = Path(instruction_file)
    try:
        parsed, groups = parse_instruction_excel(str(instruction_path), include_sheets=include_sheets)
    except Exception as exc:
        # 老 .xls 在很多 pandas 环境下缺 xlrd；这里用本文件里的 Excel COM / libreoffice fallback 再读。
        if instruction_path.suffix.lower() == ".xls":
            converted = convert_xls_to_xlsx(instruction_path)
            parsed, groups = parse_instruction_excel(str(converted), include_sheets=include_sheets)
        else:
            raise exc
    if groups is None or groups.empty:
        raise RuntimeError("没有从组合指令中解析出有效股票指令。")
    groups = groups.copy()
    groups["code"] = groups["code"].apply(canonical_code)
    groups["instruction_hash"] = groups.apply(instruction_hash, axis=1)
    groups["buy_levels_json"] = groups["buy_levels"].apply(levels_json)
    groups["sell_levels_json"] = groups["sell_levels"].apply(levels_json)
    return groups


def load_prev_snapshot(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_csv(path, dtype={"code": str})
    return pd.DataFrame(columns=["code", "instruction_hash"])


def decide_changed_codes(snapshot: pd.DataFrame, prev_snapshot: pd.DataFrame, changed_codes_arg: str = "", incremental: bool = True) -> List[str]:
    if changed_codes_arg:
        return list(dict.fromkeys(canonical_code(x) for x in re.split(r"[,，;；\s]+", changed_codes_arg) if x.strip()))
    current_codes = set(snapshot["code"].dropna().apply(canonical_code))
    if not incremental or prev_snapshot.empty:
        return sorted(current_codes)
    prev_map = dict(zip(prev_snapshot["code"].apply(canonical_code), prev_snapshot["instruction_hash"].astype(str)))
    changed = []
    for _, row in snapshot.iterrows():
        code = canonical_code(row["code"])
        if prev_map.get(code) != str(row["instruction_hash"]):
            changed.append(code)
    # 指令文件中被删除的股票也必须刷新（即从 active master 中移除）。
    removed = set(prev_map) - current_codes
    return sorted(set(changed) | removed)


def should_force_full_refresh(
    prev_snapshot: pd.DataFrame,
    prev_active: pd.DataFrame,
    active_master_exists: bool,
    changed_codes_arg: str = "",
    full_refresh: bool = False,
) -> bool:
    """判断状态是否不完整，需要自动重建全部当前指令网格。"""
    return bool(
        not changed_codes_arg
        and not full_refresh
        and prev_snapshot is not None
        and not prev_snapshot.empty
        and (not active_master_exists or prev_active is None or prev_active.empty)
    )


def build_effective_levels(levels: Sequence[Dict[str, Any]], current_pos: float, side: str) -> List[Dict[str, Any]]:
    """把 delta 指令动态转换为用于生成网格和导入模板输出的有效目标仓位。"""
    if not levels:
        return []
    if side == "buy":
        ordered = sorted(levels, key=lambda x: float(x["price"]), reverse=True)
    else:
        ordered = sorted(levels, key=lambda x: float(x["price"]))
    running = float(current_pos)
    out = []
    for x in ordered:
        y = dict(x)
        basis = str(y.get("order_basis", "target") or "target").lower()
        if basis == "delta":
            delta = _safe_float(y.get("delta_pos"), 0.0)
            if side == "buy":
                running = min(1.0, running + abs(delta))
            else:
                running = max(0.0, running - abs(delta))
            y["effective_target_pos"] = running
            y["target_pos"] = running
        else:
            target = _safe_float(y.get("target_pos"), running)
            running = float(np.clip(target, 0.0, 1.0))
            y["effective_target_pos"] = running
            y["target_pos"] = running
        out.append(y)
    return out


def source_basis_for_row(row: pd.Series, effective_levels: List[Dict[str, Any]]) -> str:
    # 找与 price/target 接近的经理锚点；如果任一来源锚点是 delta，则记 delta，用于审计。
    bases = []
    for lvl in effective_levels:
        lp = _safe_float(lvl.get("price"), np.nan)
        lt = _safe_float(lvl.get("target_pos"), np.nan)
        for p_col, t_col in [("price", "target_pos"), ("source_upper_price", "source_upper_pos"), ("source_lower_price", "source_lower_pos")]:
            rp = _safe_float(row.get(p_col), np.nan)
            rt = _safe_float(row.get(t_col), np.nan)
            if np.isfinite(lp) and np.isfinite(rp) and abs(lp - rp) <= 1e-6:
                if not np.isfinite(lt) or not np.isfinite(rt) or abs(lt - rt) <= 1e-6:
                    bases.append(str(lvl.get("order_basis", "target")))
    return "delta" if "delta" in [b.lower() for b in bases] else "target"


def get_position(positions: pd.DataFrame, product: str, unit: str, code: str) -> Tuple[float, str]:
    if positions is None or positions.empty:
        return 0.0, ""
    code = canonical_code(code)
    m = positions[(positions["product"] == product) & (positions["unit"] == unit) & (positions["code"] == code)]
    if m.empty:
        return 0.0, ""
    r = m.iloc[-1]
    return float(r.get("current_pos", 0.0)), str(r.get("name", "") or "")


def fetch_daily_profile_cached(code: str, asof_date: str, lookback_years: int, price_adj: str, cache: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    code = canonical_code(code)
    if code in cache:
        return cache[code]
    raw = fetch_wind_daily_ohlcva(code, end_date=asof_date, lookback_years=lookback_years, price_adj=price_adj)
    prof = prepare_daily_features(raw)
    cache[code] = prof
    return prof


def latest_close_from_profile(profile: pd.DataFrame) -> float:
    if profile is None or profile.empty:
        return np.nan
    for col in ["close", "最新价"]:
        if col in profile.columns:
            s = pd.to_numeric(profile[col], errors="coerce").dropna()
            if not s.empty:
                return float(s.iloc[-1])
    return np.nan


def resolve_market_data_date(asof_date: str, oplus_date: str = "", oplus_file: str = "") -> str:
    """确定生成当日网格时允许使用的市场数据截止日。

    业务约定：D 日网格使用 D-1 日 Oplus 更新仓位，因此历史行情也截止到
    Oplus 所属日期，避免历史重跑时误用 D 日收盘数据。
    """
    asof = pd.to_datetime(asof_date)
    if str(oplus_date or "").strip():
        try:
            d = normalize_date_token(oplus_date, default_year=asof.year)
            if d:
                return d
        except Exception:
            pass

    # 显式传 --oplus-file 但没传 --oplus-date 时，从文件名提取 YYYYMMDD/MMDD。
    name = Path(str(oplus_file or "")).stem
    m8 = re.search(r"(?<!\d)(20\d{6})(?!\d)", name)
    if m8:
        try:
            return datetime.strptime(m8.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except Exception:
            pass
    m4s = re.findall(r"(?<!\d)(\d{4})(?!\d)", name)
    for token in reversed(m4s):
        try:
            return normalize_date_token(token, default_year=asof.year)
        except Exception:
            continue

    # 最保守 fallback：不用 asof 日，至少退回前一自然日；Wind raw-close 会继续向前找最近交易日。
    return (asof - pd.Timedelta(days=1)).strftime("%Y-%m-%d")


def order_dedupe_key(product: str, unit: str, code: str, side: str, price: float, order_weight: float) -> str:
    s = f"{product}|{unit}|{canonical_code(code)}|{side}|{round(float(price), 4)}|{round(float(order_weight), 6)}"
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:20]


def load_ledger(path: Path) -> pd.DataFrame:
    if path.exists():
        try:
            return pd.read_excel(path, sheet_name="delta_ledger", dtype={"code": str})
        except Exception:
            return pd.read_excel(path, dtype={"code": str})
    return pd.DataFrame()


def load_active_master(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_excel(path, dtype={"code": str})
    return pd.DataFrame()


def existing_keys_from_ledger(ledger: pd.DataFrame) -> set:
    if ledger is None or ledger.empty or "dedupe_key" not in ledger.columns:
        return set()
    status = ledger.get("status", pd.Series(["active"] * len(ledger))).astype(str).str.lower()
    active = ledger[~status.isin(["cancelled", "canceled", "作废", "已取消"])]
    return set(active["dedupe_key"].dropna().astype(str))


def build_orders_for_code(
    row: pd.Series,
    accounts: pd.DataFrame,
    positions: pd.DataFrame,
    asof_date: str,
    price_band: float,
    latest_price_map: Dict[str, float],
    daily_cache: Dict[str, pd.DataFrame],
    args: argparse.Namespace,
    existing_keys: set,
    market_data_date: str = "",
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """生成某股票的当前完整网格、本次新增网格及复核明细。

    ledger 只决定一档是否进入“本次新增”模板，不能把它从 active master 删除。
    否则经理只修改一个远端锚点时，未变化的近端网格会被历史 ledger 拦截并消失。
    """
    code = canonical_code(row["code"])
    name = str(row.get("name", "") or "")
    market_data_date = str(market_data_date or asof_date)
    buy_levels_raw = row.get("buy_levels") or []
    sell_levels_raw = row.get("sell_levels") or []
    instruction_id = str(row.get("instruction_hash", ""))

    # D 日网格只能使用 Oplus 日（通常 D-1）及以前的历史行情，避免历史重跑前视。
    profile = fetch_daily_profile_cached(code, market_data_date, args.lookback_years, args.price_adj, daily_cache)

    # 价格带必须基于真实交易价格：优先 Oplus 文件里的最新价；
    # Oplus 没有该股票时，单独向 Wind 拉 market_data_date 或此前最近交易日的未复权收盘价。
    ref_price = latest_price_map.get(code, np.nan)
    ref_price_source = "oplus_latest"
    wind_fetch_code = ""
    if not np.isfinite(ref_price) or ref_price <= 0:
        ref_price, wind_fetch_code = fetch_wind_latest_raw_close(code, market_data_date, return_fetch_code=True)
        ref_price_source = "wind_raw_close"
    if not np.isfinite(ref_price) or ref_price <= 0:
        raise ValueError(
            f"{code} 没有可用参考价格：Oplus无有效最新价，Wind在 {market_data_date} 或此前也无未复权收盘价，"
            f"无法筛选上下 {price_band:.0%} 网格。"
        )
    lo, hi = ref_price * (1 - price_band), ref_price * (1 + price_band)

    active_rows: List[Dict[str, Any]] = []
    new_rows: List[Dict[str, Any]] = []
    check_rows: List[Dict[str, Any]] = []

    for _, acct in accounts.iterrows():
        product = str(acct["product"])
        unit = str(acct["unit"])
        current_pos, pos_name = get_position(positions, product, unit, code)
        sec_name = name or pos_name

        buy_levels = build_effective_levels(buy_levels_raw, current_pos, "buy")
        sell_levels = build_effective_levels(sell_levels_raw, current_pos, "sell")
        if not buy_levels and not sell_levels:
            continue

        try:
            res = evaluate_manager_grid_splitter(
                code=code,
                asof_date=asof_date,
                manager_buy_levels=buy_levels,
                manager_sell_levels=sell_levels,
                current_pos=current_pos,
                daily_profile=profile,
                intraday_df=pd.DataFrame(columns=["date", "close", "volume", "amount"]),
                lookback_years=args.lookback_years,
                price_adj=args.price_adj,
                split_method=args.split_method,
                buy_gamma=args.buy_gamma,
                sell_gamma=args.sell_gamma,
                auto_split_buy_anchors=args.auto_split_buy_anchors,
                print_detail=False,
            )
        except Exception as exc:
            check_rows.append({
                "asof_date": asof_date, "product": product, "unit": unit, "code": code, "name": sec_name,
                "side": "", "price": np.nan, "current_pos": current_pos, "target_pos": np.nan,
                "order_weight": np.nan, "pass_position_check": False, "is_new_order": False,
                "skip_reason": f"网格生成失败：{type(exc).__name__}: {exc}",
            })
            continue

        side_tables = []
        if res.get("buy_table") is not None and not res["buy_table"].empty:
            t = res["buy_table"].copy()
            t["side"] = "buy"
            t["source_basis"] = t.apply(lambda z: source_basis_for_row(z, buy_levels), axis=1)
            side_tables.append(t)
        if res.get("sell_table") is not None and not res["sell_table"].empty:
            t = res["sell_table"].copy()
            t["side"] = "sell"
            t["source_basis"] = t.apply(lambda z: source_basis_for_row(z, sell_levels), axis=1)
            side_tables.append(t)
        if not side_tables:
            continue

        table = pd.concat(side_tables, ignore_index=True)
        table["price"] = pd.to_numeric(table["price"], errors="coerce")
        table["target_pos"] = pd.to_numeric(table["target_pos"], errors="coerce")
        table = table.dropna(subset=["price", "target_pos"])
        table = table[(table["price"] >= lo - 1e-10) & (table["price"] <= hi + 1e-10)]
        if table.empty:
            continue

        for side, sub in table.groupby("side", sort=False):
            sub = sub.sort_values("price", ascending=(side == "sell")).copy()
            if side == "buy":
                sub = sub.sort_values("price", ascending=False)
            running_pos = float(current_pos)
            for _, lvl in sub.iterrows():
                target = float(np.clip(lvl["target_pos"], 0.0, 1.0))
                delta = target - running_pos
                desired_side = "buy" if delta > 0 else "sell"
                order_weight = abs(delta)
                valid = True
                reason = "OK"
                if abs(delta) <= 1e-12:
                    valid = False; reason = "目标仓位与当前阶梯仓位相同，无需下单"
                elif desired_side != side:
                    valid = False; reason = f"方向不一致：grid side={side}, delta side={desired_side}"
                elif order_weight < float(args.min_order_weight):
                    valid = False; reason = f"单档差额 {order_weight:.6f} 小于 min_order_weight"
                elif side == "sell" and running_pos + 1e-12 < order_weight:
                    valid = False; reason = "卖出数量超过当前可卖仓位"
                elif side == "buy" and running_pos + order_weight > 1.0 + 1e-12:
                    valid = False; reason = "买入后仓位超过100%"

                key = order_dedupe_key(product, unit, code, side, float(lvl["price"]), order_weight)
                is_new = bool(valid and (args.force or key not in existing_keys))
                new_reason = "新网格，可进入本次新增模板" if is_new else (
                    "历史 ledger 已存在；仍保留在当前全部网格，但不重复进入本次新增模板" if valid else "无效网格"
                )
                after_pos = running_pos + (order_weight if side == "buy" else -order_weight)
                check_row = {
                    "asof_date": asof_date,
                    "product": product,
                    "unit": unit,
                    "code": code,
                    "name": sec_name,
                    "side": side,
                    "price": float(lvl["price"]),
                    "reference_price": ref_price,
                    "reference_price_source": ref_price_source,
                    "wind_fetch_code": wind_fetch_code,
                    "market_data_date": market_data_date,
                    "band_low": lo,
                    "band_high": hi,
                    "current_pos_before_level": running_pos,
                    "target_pos": target,
                    "order_weight": order_weight,
                    "position_after_level": after_pos,
                    "source_basis": lvl.get("source_basis", ""),
                    "is_manager_level": bool(lvl.get("is_manager_level", False)),
                    "instruction_hash": instruction_id,
                    "dedupe_key": key,
                    "pass_position_check": valid,
                    "is_new_order": is_new,
                    "skip_reason": reason,
                    "new_order_reason": new_reason,
                }
                check_rows.append(check_row)
                if valid:
                    active_row = {
                        **check_row,
                        "direction": trade_direction(side),
                        "trade_market": trade_market(code),
                        "status": "active",
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    active_rows.append(active_row)
                    running_pos = after_pos
                    if is_new:
                        new_rows.append(active_row)
                        existing_keys.add(key)

    return pd.DataFrame(active_rows), pd.DataFrame(new_rows), pd.DataFrame(check_rows)


def _find_nav_col(columns: Iterable[Any]) -> Optional[str]:
    """找真正的净资产绝对值列，排除“市值/净资产(%)”等比例字段。"""
    cols = list(columns)
    for cand in NAV_COL_CANDIDATES:
        for c in cols:
            cs = str(c).strip().replace("（", "(").replace("）", ")")
            cc = cand.replace("（", "(").replace("）", ")")
            if cs == cc and "/" not in cs and "%" not in cs and "％" not in cs and "比例" not in cs:
                return c
    for c in cols:
        cs = str(c).strip()
        if ("净资产" in cs or "资产净值" in cs) and "/" not in cs and "%" not in cs and "％" not in cs and "比例" not in cs:
            return c
    return None


def _find_market_value_col(columns: Iterable[Any]) -> Optional[str]:
    """找证券市值列，但明确排除“市值/净资产”这类比例列。"""
    cols = list(columns)
    normalized = {str(c).strip().replace("（", "(").replace("）", ")"): c for c in cols}
    for cand in MARKET_VALUE_COL_CANDIDATES:
        key = cand.replace("（", "(").replace("）", ")")
        if key in normalized:
            c = normalized[key]
            cs = str(c)
            if "/" not in cs and "比例" not in cs and "%" not in cs and "％" not in cs:
                return c
    for c in cols:
        cs = str(c).strip()
        if "市值" in cs and "/" not in cs and "比例" not in cs and "%" not in cs and "％" not in cs:
            return c
    return None


def infer_sim_product_nav(position_file: Optional[str | Path]) -> Tuple[float, str]:
    """推断当前产品（由 _activate_product 切换）的产品净资产。

    优先直接读取净资产列；若没有，则用 证券市值 / 持仓占净资产比例 反推。
    返回 (nav, source)。若无法推断则返回 (nan, reason)。
    """
    if not position_file:
        return np.nan, "未提供 position_file"
    p = Path(position_file)
    if not p.exists():
        return np.nan, f"position_file 不存在: {p}"
    try:
        sheets = read_excel_any(p, sheet_name=None)
    except Exception as exc:
        return np.nan, f"读取持仓文件失败: {type(exc).__name__}: {exc}"
    if not isinstance(sheets, dict):
        sheets = {"Sheet1": sheets}

    direct_values: List[float] = []
    derived_values: List[float] = []
    direct_sources: List[str] = []
    derived_sources: List[str] = []

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue
        product_col = _find_col(df.columns, PRODUCT_COL_CANDIDATES)
        if product_col is None:
            continue
        date_col = _find_col(df.columns, POSITION_DATE_COL_CANDIDATES)
        pos_col = _find_col(df.columns, POSITION_COL_CANDIDATES)
        nav_col = _find_nav_col(df.columns)
        mv_col = _find_market_value_col(df.columns)

        work = df.copy()
        if date_col is not None:
            dt = pd.to_datetime(work[date_col], errors="coerce")
            if dt.notna().any():
                latest = dt.max()
                work = work.loc[dt == latest].copy()

        mask = work[product_col].apply(is_sim_product)
        work = work.loc[mask].copy()
        if work.empty:
            continue

        if nav_col is not None:
            vals = pd.to_numeric(
                work[nav_col].astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            )
            vals = vals[(vals > 0) & np.isfinite(vals)]
            if not vals.empty:
                direct_values.extend(vals.tolist())
                direct_sources.append(f"{sheet_name}:{nav_col}")

        if mv_col is not None and pos_col is not None:
            for _, r in work.iterrows():
                mv = _safe_float(r.get(mv_col), np.nan)
                pos = _position_value_to_fraction(r.get(pos_col), pos_col)
                if np.isfinite(mv) and mv > 0 and np.isfinite(pos) and pos > 1e-8:
                    nav_est = float(mv / pos)
                    if np.isfinite(nav_est) and nav_est > 0:
                        derived_values.append(nav_est)
                        derived_sources.append(f"{sheet_name}:{mv_col}/{pos_col}")

    if direct_values:
        nav = float(np.nanmedian(np.array(direct_values, dtype=float)))
        return nav, "直接净资产列: " + ", ".join(sorted(set(direct_sources)))
    if derived_values:
        arr = np.array(derived_values, dtype=float)
        if len(arr) >= 5:
            lo, hi = np.nanquantile(arr, [0.1, 0.9])
            arr = arr[(arr >= lo) & (arr <= hi)]
        nav = float(np.nanmedian(arr)) if len(arr) else np.nan
        return nav, "证券市值/持仓比例反推: " + ", ".join(sorted(set(derived_sources)))
    return np.nan, "持仓文件中未找到可用于推断净资产的列"


def _infer_volume_share_multiplier(minute_df: pd.DataFrame) -> float:
    """自动判断 Wind 分钟 volume 是“股”还是“手”。

    用 amount / (close * volume) 判断：接近 1 表示 volume 已是股数；
    接近 100 表示 A 股常见的“手”，此时乘 100 转成股数。
    """
    if minute_df is None or minute_df.empty:
        return 1.0
    close = pd.to_numeric(minute_df.get("close"), errors="coerce")
    volume = pd.to_numeric(minute_df.get("volume"), errors="coerce")
    amount = pd.to_numeric(minute_df.get("amount"), errors="coerce")
    denom = close * volume
    ratio = (amount / denom.replace(0, np.nan)).replace([np.inf, -np.inf], np.nan).dropna()
    ratio = ratio[(ratio > 0.05) & (ratio < 5000)]
    if ratio.empty:
        return 1.0
    med = float(ratio.median())
    candidates = np.array([1.0, 10.0, 100.0, 1000.0])
    # 在 log 空间找最接近的常见单位倍数。
    idx = int(np.argmin(np.abs(np.log(candidates) - np.log(max(med, 1e-12)))))
    chosen = float(candidates[idx])
    # 过于离谱时保守按 1 处理，避免错误放大流动性。
    if med / chosen > 3.5 or chosen / med > 3.5:
        return 1.0
    return chosen


def _bar_vwap_or_close(row: pd.Series, volume_multiplier: float) -> float:
    close = _safe_float(row.get("close"), np.nan)
    volume = _safe_float(row.get("volume"), np.nan)
    amount = _safe_float(row.get("amount"), np.nan)
    px = np.nan
    if np.isfinite(amount) and amount > 0 and np.isfinite(volume) and volume > 0 and volume_multiplier > 0:
        px = amount / (volume * volume_multiplier)
    low = _safe_float(row.get("low"), close)
    high = _safe_float(row.get("high"), close)
    # Wind amount/volume 单位若异常，VWAP 可能离开 K 线范围；这种情况回退 close。
    if not np.isfinite(px) or px <= 0:
        px = close
    if np.isfinite(low) and np.isfinite(high) and low > 0 and high >= low:
        if not (low * 0.98 <= px <= high * 1.02):
            px = close
    return float(px) if np.isfinite(px) else np.nan


def _eligible_volume_fraction(row: pd.Series, side: str, limit_price: float) -> float:
    """估计一分钟成交量中有多少比例处在限价可成交区域。

    buy: 价格 <= 限价；sell: 价格 >= 限价。
    当整根分钟 K 都位于可成交区域时记 100%；只穿价时按 high/low 区间线性估计。
    """
    low = _safe_float(row.get("low"), np.nan)
    high = _safe_float(row.get("high"), np.nan)
    close = _safe_float(row.get("close"), np.nan)
    if not np.isfinite(low):
        low = close
    if not np.isfinite(high):
        high = close
    if not np.isfinite(low) or not np.isfinite(high) or not np.isfinite(limit_price):
        return 0.0
    if high < low:
        low, high = high, low
    if abs(high - low) <= 1e-12:
        if side == "buy":
            return 1.0 if low <= limit_price + 1e-12 else 0.0
        return 1.0 if high >= limit_price - 1e-12 else 0.0
    if side == "buy":
        if low > limit_price:
            return 0.0
        if high <= limit_price:
            return 1.0
        return float(np.clip((limit_price - low) / (high - low), 0.0, 1.0))
    if high < limit_price:
        return 0.0
    if low >= limit_price:
        return 1.0
    return float(np.clip((high - limit_price) / (high - low), 0.0, 1.0))


def _execution_price_from_bar(row: pd.Series, side: str, limit_price: float, volume_multiplier: float) -> float:
    """给模拟环境使用的分钟成交价估计；保证不劣于原限价。"""
    base = _bar_vwap_or_close(row, volume_multiplier)
    low = _safe_float(row.get("low"), base)
    high = _safe_float(row.get("high"), base)
    if side == "buy":
        px = min(limit_price, base) if np.isfinite(base) else limit_price
    else:
        px = max(limit_price, base) if np.isfinite(base) else limit_price
    if np.isfinite(low) and np.isfinite(high) and high >= low:
        px = float(np.clip(px, low, high))
        if side == "buy":
            px = min(px, limit_price)
        else:
            px = max(px, limit_price)
    return float(px)


def fetch_previous_trading_minute_bars(
    code: str,
    asof_date: str,
    max_lookback_days: int = 10,
) -> Tuple[pd.DataFrame, str]:
    """向前找最近一个有 Wind 1 分钟数据的交易日。

    使用未复权分钟价，因为限价单判断必须与真实交易价格一致。
    """
    end = pd.to_datetime(asof_date)
    for lag in range(1, max(int(max_lookback_days), 1) + 1):
        d = end - pd.Timedelta(days=lag)
        if d.weekday() >= 5:
            continue
        date_str = d.strftime("%Y-%m-%d")
        try:
            minute = fetch_wind_intraday_minute(code, date_str, bar_size=1, price_adj="")
        except Exception:
            minute = pd.DataFrame()
        if minute is not None and not minute.empty:
            minute = minute.copy().sort_values("date").reset_index(drop=True)
            return minute, date_str
    return pd.DataFrame(columns=["date", "high", "low", "close", "volume", "amount"]), ""


def simulate_previous_day_execution(
    orders_df: pd.DataFrame,
    positions: pd.DataFrame,
    asof_date: str,
    product_nav: float,
    participation_rate: float = SIM_DEFAULT_PARTICIPATION_RATE,
    max_lookback_days: int = 10,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """用上一交易日分钟线核实网格是否具备现实可成交性。

    规则：
    - 买单只有分钟 low <= 网格买价时才有成交机会；卖单只有 high >= 网格卖价时才有成交机会；
    - 穿价分钟按 high/low 线性估计真正处于可成交区域的成交量比例；
    - 本策略在任意一分钟最多占该股票分钟成交量 participation_rate（默认 1%）；
    - 同一股票的所有网格/单元共享这 1% 容量，防止重复使用同一份市场流动性；
    - 生成“模拟成交价”并把部分成交对应的目标仓位重新计算，用于仿真环境导入；
    - 原始网格表完全不修改，另行输出可执行模拟文件。
    """
    audit_cols = [
        "product", "unit", "code", "name", "side", "limit_price", "order_weight",
        "desired_shares", "filled_shares", "fill_ratio", "filled_weight",
        "sim_target_pos", "sim_fill_price", "previous_trade_date",
        "participation_rate", "volume_share_multiplier", "first_fill_time", "last_fill_time",
        "simulation_status", "simulation_reason",
    ]
    if orders_df is None or orders_df.empty:
        return pd.DataFrame(), pd.DataFrame(columns=audit_cols), pd.DataFrame()

    orders = orders_df.copy().reset_index(drop=True)
    if "pass_position_check" in orders.columns:
        orders = orders[orders["pass_position_check"].fillna(False)].copy().reset_index(drop=True)
    if orders.empty:
        return pd.DataFrame(), pd.DataFrame(columns=audit_cols), pd.DataFrame()

    if not np.isfinite(product_nav) or product_nav <= 0:
        audit = orders.copy()
        audit["limit_price"] = pd.to_numeric(audit.get("price"), errors="coerce")
        audit["desired_shares"] = np.nan
        audit["filled_shares"] = 0.0
        audit["fill_ratio"] = 0.0
        audit["filled_weight"] = 0.0
        audit["sim_target_pos"] = np.nan
        audit["sim_fill_price"] = np.nan
        audit["previous_trade_date"] = ""
        audit["participation_rate"] = participation_rate
        audit["volume_share_multiplier"] = np.nan
        audit["first_fill_time"] = ""
        audit["last_fill_time"] = ""
        audit["simulation_status"] = "SKIP"
        audit["simulation_reason"] = "无法获得巨杉网格化（净值线3号）的产品净资产"
        return pd.DataFrame(), audit[[c for c in audit_cols if c in audit.columns]], pd.DataFrame()

    participation_rate = float(np.clip(participation_rate, 0.0, 1.0))
    orders["_sim_oid"] = np.arange(len(orders), dtype=int)
    orders["price"] = pd.to_numeric(orders["price"], errors="coerce")
    orders["order_weight"] = pd.to_numeric(orders["order_weight"], errors="coerce")
    orders = orders.dropna(subset=["price", "order_weight"]).copy()

    # 当前真实仓位（净值线3号）用于将模拟部分成交重新折算为目标仓位。
    pos_map: Dict[Tuple[str, str, str], float] = {}
    if positions is not None and not positions.empty:
        for _, r in positions.iterrows():
            key = (normalize_product_name(r.get("product")), normalize_unit_name(r.get("unit")), canonical_code(r.get("code")))
            pos_map[key] = float(np.clip(_safe_float(r.get("current_pos"), 0.0), 0.0, 1.0))

    all_states: Dict[int, Dict[str, Any]] = {}
    minute_summary_rows: List[Dict[str, Any]] = []
    global_fill_seq = 0

    for code, code_orders in orders.groupby("code", sort=False):
        code = canonical_code(code)
        minute, prev_date = fetch_previous_trading_minute_bars(code, asof_date, max_lookback_days=max_lookback_days)
        if minute.empty:
            for _, o in code_orders.iterrows():
                oid = int(o["_sim_oid"])
                all_states[oid] = {
                    "filled_shares": 0.0, "value_sum": 0.0, "desired_shares": np.nan,
                    "fill_ratio": 0.0, "filled_weight": 0.0, "prev_date": "",
                    "volume_multiplier": np.nan, "first_fill_time": "", "last_fill_time": "",
                    "first_fill_seq": np.inf, "status": "NO_MINUTE_DATA",
                    "reason": "向前未找到可用的 Wind 1分钟行情",
                }
            minute_summary_rows.append({
                "code": code, "previous_trade_date": "", "minute_rows": 0,
                "raw_volume": np.nan, "volume_share_multiplier": np.nan,
                "estimated_volume_shares": np.nan, "participation_rate": participation_rate,
                "max_strategy_volume_shares": np.nan,
            })
            continue

        volume_multiplier = _infer_volume_share_multiplier(minute)
        raw_volume = pd.to_numeric(minute["volume"], errors="coerce").fillna(0.0).clip(lower=0.0)
        est_volume_shares = raw_volume * volume_multiplier
        minute_summary_rows.append({
            "code": code,
            "previous_trade_date": prev_date,
            "minute_rows": int(len(minute)),
            "raw_volume": float(raw_volume.sum()),
            "volume_share_multiplier": volume_multiplier,
            "estimated_volume_shares": float(est_volume_shares.sum()),
            "participation_rate": participation_rate,
            "max_strategy_volume_shares": float(est_volume_shares.sum() * participation_rate),
        })

        states: Dict[int, Dict[str, Any]] = {}
        for _, o in code_orders.iterrows():
            oid = int(o["_sim_oid"])
            limit_price = float(o["price"])
            order_weight = abs(float(o["order_weight"]))
            desired_shares = order_weight * product_nav / max(limit_price, 1e-12)
            states[oid] = {
                "side": str(o.get("side", "")).lower(),
                "limit_price": limit_price,
                "order_weight": order_weight,
                "desired_shares": max(float(desired_shares), 0.0),
                "remaining_shares": max(float(desired_shares), 0.0),
                "filled_shares": 0.0,
                "value_sum": 0.0,
                "first_fill_time": "",
                "last_fill_time": "",
                "first_fill_seq": np.inf,
                "prev_date": prev_date,
                "volume_multiplier": volume_multiplier,
                "status": "NOT_FILLED",
                "reason": f"昨日价格未触发，或触发后的{participation_rate:.0%}分钟流动性容量不足",
            }

        for _, bar in minute.iterrows():
            vol_raw = max(_safe_float(bar.get("volume"), 0.0), 0.0)
            minute_volume_shares = vol_raw * volume_multiplier
            total_capacity = participation_rate * minute_volume_shares
            if total_capacity <= 1e-12:
                continue

            eligible = []
            for oid, st in states.items():
                if st["remaining_shares"] <= 1e-12:
                    continue
                frac = _eligible_volume_fraction(bar, st["side"], st["limit_price"])
                if frac <= 1e-12:
                    continue
                close = _safe_float(bar.get("close"), st["limit_price"])
                # 同一分钟若多个档位同时可成交，优先完全处于可成交区域的档位，
                # 再优先离分钟价格更近的档位，避免任意顺序造成容量偏差。
                if st["side"] == "buy":
                    side_priority = -st["limit_price"]
                else:
                    side_priority = st["limit_price"]
                eligible.append((oid, frac, -frac, abs(st["limit_price"] - close), side_priority))

            eligible.sort(key=lambda x: (x[2], x[3], x[4], x[0]))
            remaining_total_capacity = total_capacity
            for oid, frac, *_ in eligible:
                if remaining_total_capacity <= 1e-12:
                    break
                st = states[oid]
                order_side_cap = total_capacity * float(frac)
                alloc = min(st["remaining_shares"], remaining_total_capacity, order_side_cap)
                if alloc <= 1e-12:
                    continue
                exec_px = _execution_price_from_bar(bar, st["side"], st["limit_price"], volume_multiplier)
                if not np.isfinite(exec_px) or exec_px <= 0:
                    exec_px = st["limit_price"]
                st["remaining_shares"] -= alloc
                st["filled_shares"] += alloc
                st["value_sum"] += alloc * exec_px
                ts = pd.to_datetime(bar.get("date"), errors="coerce")
                ts_text = ts.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(ts) else ""
                if not st["first_fill_time"]:
                    st["first_fill_time"] = ts_text
                    st["first_fill_seq"] = global_fill_seq
                    global_fill_seq += 1
                st["last_fill_time"] = ts_text
                remaining_total_capacity -= alloc

        for oid, st in states.items():
            desired = st["desired_shares"]
            filled = st["filled_shares"]
            ratio = float(np.clip(filled / desired, 0.0, 1.0)) if desired > 1e-12 else 0.0
            st["fill_ratio"] = ratio
            st["filled_weight"] = st["order_weight"] * ratio
            if filled > 1e-12:
                st["status"] = "FILLED" if ratio >= 1.0 - 1e-9 else "PARTIAL"
                st["reason"] = f"昨日分钟线触发；按最多{participation_rate:.0%}分钟流动性模拟成交"
            all_states[oid] = st

    audit_rows: List[Dict[str, Any]] = []
    executable = orders.copy()
    executable["sim_target_pos"] = np.nan
    executable["sim_fill_price"] = np.nan
    executable["sim_fill_ratio"] = 0.0
    executable["sim_filled_weight"] = 0.0
    executable["sim_first_fill_seq"] = np.inf

    # 先写每档成交结果。
    for idx, o in executable.iterrows():
        oid = int(o["_sim_oid"])
        st = all_states.get(oid, {})
        desired = _safe_float(st.get("desired_shares"), np.nan)
        filled = _safe_float(st.get("filled_shares"), 0.0)
        fill_ratio = _safe_float(st.get("fill_ratio"), 0.0)
        filled_weight = _safe_float(st.get("filled_weight"), 0.0)
        sim_fill_price = (
            _safe_float(st.get("value_sum"), 0.0) / filled
            if filled > 1e-12 else np.nan
        )
        executable.at[idx, "sim_fill_price"] = sim_fill_price
        executable.at[idx, "sim_fill_ratio"] = fill_ratio
        executable.at[idx, "sim_filled_weight"] = filled_weight
        executable.at[idx, "sim_first_fill_seq"] = st.get("first_fill_seq", np.inf)

        audit_rows.append({
            "product": SIM_PRODUCT_NAME,
            "unit": SIM_UNIT_NAME,
            "code": canonical_code(o.get("code")),
            "name": o.get("name", ""),
            "side": o.get("side", ""),
            "limit_price": _safe_float(o.get("price"), np.nan),
            "order_weight": _safe_float(o.get("order_weight"), np.nan),
            "desired_shares": desired,
            "filled_shares": filled,
            "fill_ratio": fill_ratio,
            "filled_weight": filled_weight,
            "sim_target_pos": np.nan,
            "sim_fill_price": sim_fill_price,
            "previous_trade_date": st.get("prev_date", ""),
            "participation_rate": participation_rate,
            "volume_share_multiplier": st.get("volume_multiplier", np.nan),
            "first_fill_time": st.get("first_fill_time", ""),
            "last_fill_time": st.get("last_fill_time", ""),
            "simulation_status": st.get("status", "SKIP"),
            "simulation_reason": st.get("reason", "未模拟"),
            "_sim_oid": oid,
        })

    # 模拟环境会按我们填入的价格直接成交，因此导入文件按“昨日真实成交发生顺序”
    # 重新生成目标仓位；部分成交也只推进对应的仓位增量。
    executable = executable[executable["sim_fill_ratio"] > 1e-12].copy()
    if not executable.empty:
        for keys, g in executable.groupby(["product", "unit", "code"], sort=False):
            product, unit, code = keys
            base_key = (normalize_product_name(product), normalize_unit_name(unit), canonical_code(code))
            running_pos = float(np.clip(pos_map.get(base_key, 0.0), 0.0, 1.0))
            order_idx = g.sort_values(["sim_first_fill_seq", "price"]).index.tolist()
            for idx in order_idx:
                side = str(executable.at[idx, "side"]).lower()
                fw = float(executable.at[idx, "sim_filled_weight"])
                running_pos = running_pos + fw if side == "buy" else running_pos - fw
                running_pos = float(np.clip(running_pos, 0.0, 1.0))
                executable.at[idx, "sim_target_pos"] = running_pos
                oid = int(executable.at[idx, "_sim_oid"])
                for rec in audit_rows:
                    if rec.get("_sim_oid") == oid:
                        rec["sim_target_pos"] = running_pos
                        break

        executable["original_limit_price"] = executable["price"]
        executable["original_target_pos"] = executable["target_pos"]
        executable["target_pos"] = executable["sim_target_pos"]
        executable["price"] = executable["sim_fill_price"].round(4)
        executable["product"] = SIM_PRODUCT_NAME
        executable["unit"] = SIM_UNIT_NAME
        executable = executable.sort_values(["sim_first_fill_seq", "code", "unit"]).reset_index(drop=True)

    audit = pd.DataFrame(audit_rows)
    if not audit.empty:
        audit = audit.drop(columns=["_sim_oid"], errors="ignore")
    minute_summary = pd.DataFrame(minute_summary_rows)
    return executable.drop(columns=["_sim_oid"], errors="ignore"), audit, minute_summary


def save_excel_report(path: Path, sheets: Dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            sheet = name[:31]
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                letter = col[0].column_letter
                header = str(col[0].value or "")
                width = min(max(len(header) + 4, 12), 36)
                ws.column_dimensions[letter].width = width
                if any(k in header for k in ["pos", "weight", "仓位", "比例", "delta", "order_weight"]):
                    for cell in col[1:]:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0.00%"
                if "price" in header.lower() or "价格" in header:
                    for cell in col[1:]:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = "0.0000"


def run_one_product(args, pcfg: dict) -> int:
    """对单个产品（净值线3 或 聚砾1）完整跑一遍网格生成，输出到独立子目录。"""
    _activate_product(pcfg)
    base_out = Path(args.output_dir).resolve()
    outdir = base_out / pcfg["tag"]
    state_dir = (Path(args.state_dir).resolve() / pcfg["tag"]) if args.state_dir else outdir / "state"
    outdir.mkdir(parents=True, exist_ok=True)
    state_dir.mkdir(parents=True, exist_ok=True)
    # 文件名前缀带产品 tag，确保两个产品的文件互不覆盖。
    args.prefix = f"{args.prefix}_{pcfg['tag']}"

    include_sheets = [x.strip() for x in args.sheets.split(",") if x.strip()] if args.sheets else None
    accounts = load_template_accounts(args.a_template, args.hk_template)
    oplus_cumulative, latest_price_map = load_oplus_trades(args.oplus_file) if args.oplus_file else (pd.DataFrame(), {})
    market_data_date = resolve_market_data_date(args.asof_date, getattr(args, "oplus_date", ""), args.oplus_file)
    print(f"市场数据截止日：{market_data_date}（参考价/历史特征不使用运行日之后数据）")

    previous_position_path = state_dir / "current_positions.csv"
    oplus_state_path = state_dir / "oplus_trade_state.csv"
    baseline_recompute = not args.incremental_state_mode

    if baseline_recompute:
        if not args.position_file:
            raise FileNotFoundError(
                "固定持仓基线重算模式必须提供 --position-file，或把名称含‘持仓’的 Excel 放在项目根目录。"
            )
        # 每次从同一持仓基线开始，因此应应用 Oplus 文件中的全部累计成交，
        # 不能再减去上次运行记录，否则第二次运行会少算成交。
        if args.no_apply_oplus_trades:
            oplus_applied = pd.DataFrame(columns=oplus_cumulative.columns)
        else:
            oplus_applied = oplus_cumulative.copy()
        positions = build_current_positions(
            accounts=accounts,
            position_file=args.position_file,
            previous_position_path=previous_position_path,
            oplus_trades=oplus_applied,
            apply_oplus_trades=not args.no_apply_oplus_trades,
            reset_position_state=True,
            use_previous_state=False,
        )
        # 保存的状态只用于审计/查看，不作为下一次固定基线重算的输入。
        save_csv_atomic(positions, previous_position_path)
        if not args.no_apply_oplus_trades:
            audit_state = oplus_cumulative.copy().rename(columns={"trade_delta_weight": "applied_delta_weight"})
            if not audit_state.empty:
                save_csv_atomic(audit_state, oplus_state_path)
        oplus_incremental = oplus_applied
    else:
        previous_oplus_state = (
            pd.DataFrame(columns=["trade_key", "applied_delta_weight"])
            if args.reset_oplus_state
            else load_oplus_trade_state(oplus_state_path)
        )
        if args.no_apply_oplus_trades:
            oplus_incremental = pd.DataFrame(columns=oplus_cumulative.columns)
            next_oplus_state = previous_oplus_state
        else:
            oplus_incremental, next_oplus_state = incremental_oplus_trades(oplus_cumulative, previous_oplus_state)
        positions = build_current_positions(
            accounts=accounts,
            position_file=args.position_file,
            previous_position_path=previous_position_path,
            oplus_trades=oplus_incremental,
            apply_oplus_trades=not args.no_apply_oplus_trades,
            reset_position_state=args.reset_position_state,
            use_previous_state=True,
        )
        save_csv_atomic(positions, previous_position_path)
        if not args.no_apply_oplus_trades:
            save_csv_atomic(next_oplus_state, oplus_state_path)

    # 单产品模拟：净值线3号 -> 巨杉网格化。流动性模拟需要产品净资产来把仓位差换成股数。
    if float(args.sim_nav) > 0:
        sim_product_nav = float(args.sim_nav)
        sim_nav_source = "命令行 --sim-nav"
    else:
        sim_product_nav, sim_nav_source = infer_sim_product_nav(args.position_file)
    if np.isfinite(sim_product_nav) and sim_product_nav > 0:
        print(f"模拟产品：{SIM_PRODUCT_NAME} / {SIM_UNIT_NAME}（底层持仓={pcfg['key']}）")
        print(f"模拟产品净资产：{sim_product_nav:,.2f}；来源={sim_nav_source}")
    else:
        print(f"WARNING: 无法自动得到 {SIM_PRODUCT_NAME} 净资产：{sim_nav_source}")
        print("原始网格仍会正常输出；昨日分钟流动性模拟文件将为空。可用 --sim-nav 手动指定净资产。")

    snapshot_all = build_instruction_snapshot(args.instruction_file, include_sheets=include_sheets)
    supported_mask = snapshot_all["code"].apply(market_group).isin(["A", "HK"])
    unsupported_snapshot = snapshot_all[~supported_mask].copy()
    snapshot = snapshot_all[supported_mask].copy().reset_index(drop=True)
    if snapshot.empty:
        raise RuntimeError("组合指令中没有可用于 A股/港股模板的有效股票指令。")

    snapshot_path = state_dir / "instruction_snapshot.csv"
    prev_snapshot = load_prev_snapshot(snapshot_path)

    active_master_path = state_dir / "active_grid_master.xlsx"
    ledger_path = state_dir / "delta_ledger.xlsx"
    prev_active = load_active_master(active_master_path)

    if baseline_recompute:
        # 固定基线重算时，旧 snapshot/active master 不能参与决定本次计算范围。
        # 未指定 --changed-codes 就重算全部当前 A/H 指令；指定时只做测试代码。
        changed_codes = decide_changed_codes(
            snapshot,
            pd.DataFrame(columns=["code", "instruction_hash"]),
            args.changed_codes,
            incremental=False,
        )
        state_needs_rebuild = False
    else:
        # 增量模式下保留状态恢复保护。
        state_needs_rebuild = should_force_full_refresh(
            prev_snapshot=prev_snapshot,
            prev_active=prev_active,
            active_master_exists=active_master_path.exists(),
            changed_codes_arg=args.changed_codes,
            full_refresh=args.full_refresh,
        )
        if state_needs_rebuild:
            print(
                "WARNING: instruction snapshot exists but active grid master is missing/empty; "
                "forcing a full refresh of all current instruction codes."
            )
        changed_codes = decide_changed_codes(
            snapshot,
            prev_snapshot,
            args.changed_codes,
            incremental=not (args.full_refresh or state_needs_rebuild),
        )
    changed_codes = [c for c in changed_codes if c]

    ledger = load_ledger(ledger_path)
    existing_keys = existing_keys_from_ledger(ledger)

    current_codes = set(snapshot["code"].apply(canonical_code))
    previous_codes = set(prev_snapshot["code"].apply(canonical_code)) if prev_snapshot is not None and not prev_snapshot.empty else set()
    removed_codes = previous_codes - current_codes
    if baseline_recompute:
        # 本次 active grid 必须完全来自“0622基线 + 指定Oplus + 当前指令”的新计算。
        # 不继承任何旧 active master，避免旧市场后缀、旧仓位或旧指令污染结果。
        base_active = pd.DataFrame()
    else:
        # 被指令删除的股票立即移除；已修改股票等新网格成功后再替换。
        base_active = prev_active.copy() if prev_active is not None else pd.DataFrame()
        if not base_active.empty and removed_codes:
            base_active = base_active[~base_active["code"].apply(canonical_code).isin(removed_codes)].copy()

    daily_cache: Dict[str, pd.DataFrame] = {}
    refreshed_active_all = []
    new_orders_all = []
    checks_all = []
    errors = []
    successful_refresh_codes: set[str] = set()
    changed_rows = snapshot[snapshot["code"].isin(changed_codes)].copy()

    for _, row in changed_rows.iterrows():
        code = canonical_code(row["code"])
        try:
            active_orders, new_orders_for_code, checks = build_orders_for_code(
                row=row,
                accounts=accounts,
                positions=positions,
                asof_date=args.asof_date,
                market_data_date=market_data_date,
                price_band=args.price_band,
                latest_price_map=latest_price_map,
                daily_cache=daily_cache,
                args=args,
                existing_keys=existing_keys,
            )
            successful_refresh_codes.add(code)
            if not active_orders.empty:
                refreshed_active_all.append(active_orders)
            if not new_orders_for_code.empty:
                new_orders_all.append(new_orders_for_code)
            if not checks.empty:
                checks_all.append(checks)
        except Exception as exc:
            errors.append({
                "code": code,
                "name": row.get("name", ""),
                "error": f"{type(exc).__name__}: {exc}",
                "old_active_grid_retained": bool(not baseline_recompute),
            })

    refreshed_active = pd.concat(refreshed_active_all, ignore_index=True) if refreshed_active_all else pd.DataFrame()
    new_orders = pd.concat(new_orders_all, ignore_index=True) if new_orders_all else pd.DataFrame()
    checks = pd.concat(checks_all, ignore_index=True) if checks_all else pd.DataFrame()
    errors_df = pd.DataFrame(errors)

    if not base_active.empty and successful_refresh_codes:
        base_active = base_active[~base_active["code"].apply(canonical_code).isin(successful_refresh_codes)].copy()
    active_master = pd.concat([base_active, refreshed_active], ignore_index=True)
    if not active_master.empty:
        active_master["code"] = active_master["code"].apply(canonical_code)
        active_master["product"] = active_master["product"].apply(normalize_product_name)
        active_master = active_master[active_master["product"] == SIM_PRODUCT_NAME].copy()
        active_master = active_master.drop_duplicates(subset=["dedupe_key"], keep="last")

    ledger_out = pd.concat([ledger if ledger is not None else pd.DataFrame(), new_orders], ignore_index=True)
    if not ledger_out.empty and "dedupe_key" in ledger_out.columns:
        ledger_out = ledger_out.drop_duplicates(subset=["dedupe_key"], keep="first")

    save_csv_atomic(snapshot, snapshot_path)
    save_excel_report(active_master_path, {"active_grid_master": active_master})
    save_excel_report(ledger_path, {"delta_ledger": ledger_out})

    # 昨日真实分钟行情成交核实。原始 active_master 不修改；另外生成“可模拟成交”版本。
    if args.disable_liquidity_simulation:
        simulated_orders = pd.DataFrame()
        simulation_audit = pd.DataFrame()
        simulation_minute_summary = pd.DataFrame()
    else:
        simulated_orders, simulation_audit, simulation_minute_summary = simulate_previous_day_execution(
            orders_df=active_master,
            positions=positions,
            asof_date=args.asof_date,
            product_nav=sim_product_nav,
            participation_rate=args.liquidity_participation,
            max_lookback_days=args.minute_lookback_days,
        )

    simulation_report_path = outdir / f"{args.prefix}_昨日分钟流动性模拟.xlsx"
    save_excel_report(simulation_report_path, {
        "可模拟成交网格": simulated_orders,
        "逐档成交核实": simulation_audit,
        "分钟流动性概览": simulation_minute_summary,
    })

    report_path = outdir / f"{args.prefix}_复核报告.xlsx"
    run_meta = pd.DataFrame([{
        "calculation_mode": "baseline_recompute" if baseline_recompute else "incremental_state",
        "position_file": str(args.position_file or ""),
        "oplus_file": str(args.oplus_file or ""),
        "instruction_file": str(args.instruction_file or ""),
        "asof_date": str(args.asof_date),
        "simulation_product": SIM_PRODUCT_NAME,
        "simulation_source_product": pcfg["key"],
        "simulation_nav": sim_product_nav,
        "simulation_nav_source": sim_nav_source,
        "liquidity_participation": float(args.liquidity_participation),
        "oplus_rows_applied": int(len(oplus_incremental)),
        "position_rows_after_update": int(len(positions)),
    }])
    save_excel_report(report_path, {
        "run_meta": run_meta,
        "changed_codes": pd.DataFrame({"changed_code": changed_codes}),
        "refreshed_active_grid": refreshed_active,
        "new_orders": new_orders,
        "position_check": checks,
        "current_positions": positions,
        "oplus_incremental": oplus_incremental,
        "oplus_cumulative": oplus_cumulative,
        "unsupported_codes": unsupported_snapshot,
        "liquidity_sim_audit": simulation_audit,
        "liquidity_minute_summary": simulation_minute_summary,
        "errors": errors_df,
    })

    active_a, active_hk = split_orders_by_market(active_master)
    new_a, new_hk = split_orders_by_market(new_orders)
    simulated_a, simulated_hk = split_orders_by_market(simulated_orders)
    paths = {
        "active_a_template": outdir / f"{args.prefix}_A股_当前全部网格_买卖清单导入.xlsx",
        "active_hk_template": outdir / f"{args.prefix}_港股_当前全部网格_买卖清单导入.xlsx",
        "new_a_template": outdir / f"{args.prefix}_A股_本次新增网格_买卖清单导入.xlsx",
        "new_hk_template": outdir / f"{args.prefix}_港股_本次新增网格_买卖清单导入.xlsx",
        "simulated_a_template": outdir / f"{args.prefix}_A股_昨日分钟模拟成交_可导入.xlsx",
        "simulated_hk_template": outdir / f"{args.prefix}_港股_昨日分钟模拟成交_可导入.xlsx",
        "manager_instruction": outdir / f"{args.prefix}_基金经理格式_当前全部网格.xlsx",
    }
    export_import_template(active_a, args.a_template, paths["active_a_template"])
    export_import_template(active_hk, args.hk_template, paths["active_hk_template"], freeze=False)
    export_import_template(new_a, args.a_template, paths["new_a_template"])
    export_import_template(new_hk, args.hk_template, paths["new_hk_template"], freeze=False)
    export_import_template(simulated_a, args.a_template, paths["simulated_a_template"])
    export_import_template(simulated_hk, args.hk_template, paths["simulated_hk_template"], freeze=False)
    export_manager_instruction_style(active_master, paths["manager_instruction"])

    print("\n========== 实盘网格生成完成 ==========")
    print(f"指令股票数（A/H）：{len(snapshot)}")
    print(f"不支持并跳过的代码数：{len(unsupported_snapshot)}")
    print(f"本次刷新股票数：{len(changed_codes)} -> {changed_codes[:20]}{'...' if len(changed_codes)>20 else ''}")
    oplus_label = "本次应用 Oplus 成交记录" if baseline_recompute else "本次 Oplus 新增成交记录"
    print(f"{oplus_label}：{len(oplus_incremental)}")
    print(f"本次新增订单：{len(new_orders)}")
    print(f"当前 active grid：{len(active_master)}")
    print(f"昨日分钟核实后可模拟成交网格：{len(simulated_orders)}")
    print(f"昨日分钟流动性模拟报告：{simulation_report_path}")
    print(f"复核报告：{report_path}")
    print(f"delta ledger：{ledger_path}")
    print(f"active master：{active_master_path}")
    for k, path in paths.items():
        print(f"{k}: {path}")
    if not errors_df.empty:
        print("\n有错误，请看复核报告 errors sheet。")
    return 0

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="实盘网格化模拟：按产品/单元生成 A股/港股买卖清单导入模板")
    p.add_argument("--instruction-file", default="", help="组合指令 Excel；不填则自动读取 live_grid_project/指令/ 中最新文件")
    p.add_argument("--oplus-file", default="", help="oplus 交易数据；不填则询问日期，并自动读取 live_grid_project/oplus交易文件/ 中该日期最新文件")
    p.add_argument("--oplus-date", default="", help="oplus 文件日期，例如 0623 或 20260623；不填则运行时 input 询问")
    p.add_argument("--a-template", default="", help="A股买卖清单导入模板.xlsx；不填则从 live_grid_project 根目录自动寻找")
    p.add_argument("--hk-template", default="", help="港股买卖清单导入模板.xlsx；不填则从 live_grid_project 根目录自动寻找")
    p.add_argument("--asof-date", default="", help="运行日期，例如 2026-06-23；不填默认今天")
    p.add_argument("--output-dir", default="", help="输出目录；默认 live_grid_project/outputs_live_grid")
    p.add_argument("--state-dir", default="", help="状态目录；默认 live_grid_project/state，建议不要随便删除")
    p.add_argument("--project-root", default="", help="项目根目录；默认当前 live_grid_simulator.py 所在文件夹")
    p.add_argument("--prefix", default="实盘网格", help="输出文件名前缀")
    p.add_argument("--sheets", default="股票指令", help="读取指令 sheet，逗号分隔；默认 股票指令")
    p.add_argument("--position-file", default="", help="固定持仓基线文件；默认每次都从该文件重新计算，再完整应用指定 Oplus")
    p.add_argument("--incremental-state-mode", action="store_true", help="改用旧的跨日增量状态模式：沿用 state/current_positions.csv 并只应用 Oplus 新增进度")
    p.add_argument("--no-apply-oplus-trades", action="store_true", help="不把 oplus 成交数据累加到持仓基线")
    p.add_argument("--reset-oplus-state", action="store_true", help="清空 Oplus 已应用进度记录；仅在重新建立持仓基线时使用")
    p.add_argument("--reset-position-state", action="store_true", help="忽略 state/current_positions.csv，重新用 --position-file 建立仓位基线")
    p.add_argument("--changed-codes", default="", help="手动指定本次只刷新的代码，逗号分隔；不填则自动比较指令 hash")
    p.add_argument("--full-refresh", action="store_true", help="全量刷新所有指令股票")
    p.add_argument("--force", action="store_true", help="忽略 delta ledger 去重，强制输出")
    p.add_argument("--price-band", type=float, default=0.10, help="只输出参考价上下该比例内的网格，默认 0.10")
    p.add_argument("--min-order-weight", type=float, default=0.0001, help="最小单档仓位差额，默认 0.0001=0.01%%")
    p.add_argument("--sim-nav", type=float, default=0.0, help="巨杉网格化聚砾模拟产品净资产；默认从持仓文件自动推断（净值线3号+聚砾1号）")
    p.add_argument("--liquidity-participation", type=float, default=SIM_DEFAULT_PARTICIPATION_RATE, help="昨日分钟成交量最大参与率，默认 0.01=1%%")
    p.add_argument("--minute-lookback-days", type=int, default=10, help="向前寻找最近交易日分钟线的最大日历天数，默认10")
    p.add_argument("--disable-liquidity-simulation", action="store_true", help="只生成原始网格，不做昨日分钟流动性成交模拟")
    p.add_argument("--lookback-years", type=int, default=3)
    p.add_argument("--price-adj", default=DEFAULT_PRICE_ADJ)
    p.add_argument("--split-method", default="linear", choices=["linear", "geometric"])
    p.add_argument("--buy-gamma", type=float, default=BUY_GAMMA_DEFAULT)
    p.add_argument("--sell-gamma", type=float, default=SELL_GAMMA_DEFAULT)
    p.add_argument("--auto-split-buy-anchors", action="store_true", help="买卖两端锚点自动补中间档")
    return p.parse_args()


def main() -> int:
    """一次运行：分别输出「净值线3」「聚砾1」各自的网格，两套结果分开存放。"""
    args = resolve_runtime_inputs(parse_args())
    if not (0.0 < float(args.price_band) < 1.0):
        raise ValueError("price_band 必须在 0 和 1 之间。")
    if float(args.min_order_weight) < 0:
        raise ValueError("min_order_weight 不能为负数。")
    if int(args.lookback_years) < 1:
        raise ValueError("lookback_years 至少为 1。")
    if not (0.0 < float(args.liquidity_participation) <= 1.0):
        raise ValueError("liquidity_participation 必须在 (0, 1]，默认 0.01。")
    if int(args.minute_lookback_days) < 1:
        raise ValueError("minute_lookback_days 至少为 1。")

    rc = 0
    original_prefix = args.prefix
    for pcfg in PRODUCT_CONFIGS:
        args.prefix = original_prefix  # 每个产品从原始前缀开始，避免跨产品累积
        print(f"\n########## 开始计算产品：{pcfg['key']} ##########")
        try:
            one_rc = run_one_product(args, pcfg)
        except Exception as exc:
            print(f"ERROR: 产品 {pcfg['key']} 计算失败：{type(exc).__name__}: {exc}")
            one_rc = 1
        rc = rc or one_rc
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
