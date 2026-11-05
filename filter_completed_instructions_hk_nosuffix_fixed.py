# -*- coding: utf-8 -*-
"""
根据历史 Oplus 成交文件，剔除组合指令中已经累计完成 100% 的指令。

目录约定
--------
把本文件放在项目根目录，例如：

项目根目录/
├─ filter_completed_instructions.py
├─ 组合指令0624v1.xls            # 也可以放在“指令”文件夹
├─ oplus文件/
│  ├─ oplus0623.xlsx
│  ├─ oplus0624.xlsx
│  └─ oplus0625.xlsx
└─ 指令/                         # 可选

运行：
    python filter_completed_instructions.py

也可以明确指定组合指令：
    python filter_completed_instructions.py --instruction-file "组合指令0624v1.xls"

输出：
    组合指令0624v1_剔除已成交.xlsx

核心规则
--------
1. 读取“oplus文件”文件夹下所有 oplus<日期>.xls/xlsx/xlsm。
2. 对每个 Oplus 文件：
   - 优先使用“产品名称=混合 且 单元名称=混合”的汇总行；
   - 若没有汇总行，则按指令批号用 成交数量 / 指令数量 计算该批次进度；
   - 同一条指令同一天若存在多个批次，则当天进度相加，最多按 100% 计；
   - 不同日期不相加，直接取最新日期的累计成交进度。
3. 同一条指令最新累计成交进度 >= 100%，视为已成交完成。
4. 通过“证券代码 + 买卖方向 + 备注/指令文本”与组合指令匹配。
5. 已完成的指令只清空对应“指令单元格”，不删除股票整行，也不改其他未完成指令。
6. 原始组合指令为 .xls 时，Windows 下调用本机 Excel 转为临时 .xlsx；
   输出统一为 .xlsx，并尽量保留原格式。
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import load_workbook


# ============================================================
# 配置
# ============================================================

OPLUS_FOLDER_CANDIDATES = ["oplus交易文件", "oplus文件"]
INSTRUCTION_FOLDER_NAME = "指令"

OPLUS_PREFIX = "oplus"
COMPLETION_THRESHOLD = 100.0
OUTPUT_SUFFIX = "_剔除已成交"

EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}

# 组合指令中用于判断“这是交易指令”的关键词。
INSTRUCTION_KEYWORDS = (
    "买入", "买回", "加仓",
    "卖出", "减仓", "卖完", "清仓",
)

# 一些无需用来匹配组合指令的 Oplus 业务。
IGNORE_DIRECTION_KEYWORDS = (
    "回购",
)


# ============================================================
# 基础清洗
# ============================================================

def _is_blank(x: Any) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    return str(x).strip() == ""


def clean_text(x: Any) -> str:
    if _is_blank(x):
        return ""
    s = str(x).strip()
    s = (
        s.replace("％", "%")
        .replace("，", ",")
        .replace("。", ".")
        .replace("．", ".")
        .replace("：", ":")
        .replace("；", ";")
        .replace("（", "(")
        .replace("）", ")")
        .replace("　", "")
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .replace("\t", "")
    )
    return s


def normalize_instruction_text(x: Any) -> str:
    """
    用于匹配 Oplus“备注”和组合指令单元格。

    原则：只做轻度标准化，不擅自改变数字和指令含义。
    """
    s = clean_text(x).lower()
    if not s:
        return ""

    # 只清掉对含义没有影响的单位/标点差异。
    s = s.replace("人民币", "")
    s = s.replace("美元", "")
    s = s.replace("元", "")
    s = s.replace(",", "")
    s = s.replace(";", "")
    s = s.replace(":", "")
    return s


def normalize_side(x: Any) -> str:
    s = clean_text(x)
    if not s:
        return ""

    if any(k in s for k in IGNORE_DIRECTION_KEYWORDS):
        return ""

    if "买" in s:
        return "buy"
    if "卖" in s or "清仓" in s:
        return "sell"
    return ""


def instruction_side_from_text(x: Any) -> str:
    s = clean_text(x)
    if not s:
        return ""
    if "买" in s and "卖" not in s:
        return "buy"
    if "卖" in s or "清仓" in s:
        return "sell"
    return ""


def looks_like_instruction_text(x: Any) -> bool:
    s = clean_text(x)
    if not s:
        return False
    return any(k in s for k in INSTRUCTION_KEYWORDS)



def instruction_semantic_signature(x: Any) -> Optional[Tuple[str, str, float, float]]:
    """
    常见组合指令语义签名：
      (side, price_condition, trigger_price, target_position_pct)

    仅在能明确解析“价格阈值 + 目标仓位”时启用。
    """
    s = normalize_instruction_text(x)
    side = instruction_side_from_text(x)
    if not s or not side:
        return None

    m = re.search(r"(-?\d+(?:\.\d+)?)\s*(?:及|或)?(以上|以下)", s)
    if not m:
        return None

    try:
        price = round(float(m.group(1)), 8)
    except Exception:
        return None

    cond = ">=" if m.group(2) == "以上" else "<="

    if "清仓" in s or "卖完" in s:
        target = 0.0
    else:
        mt = re.search(r"(?:到|至)(-?\d+(?:\.\d+)?)%", s)
        if not mt:
            return None
        try:
            target = round(float(mt.group(1)), 8)
        except Exception:
            return None

    return (side, cond, price, target)


def _digits_only(x: Any) -> str:
    s = clean_text(x).upper()
    return "".join(re.findall(r"\d", s))


def normalize_code(code: Any, market: Any = "") -> str:
    """
    统一证券代码。

    港股兼容示例：
      0197.hk / 0197.HK / 00197.HK / 0197 HK / 0197 HK Equity
      197 + 香港市场
      -> 00197

    A股示例：
      2422 + 深圳 -> 002422
      002422.SZ -> 002422
    """
    if _is_blank(code):
        return ""

    raw = clean_text(code).upper().strip("'\"")
    market_text = clean_text(market).upper()

    if re.fullmatch(r"\d+\.0", raw):
        raw = raw[:-2]

    raw = (
        raw.replace("HKEQUITY", ".HK")
        .replace("HK EQUITY", ".HK")
        .replace("SHEQUITY", ".SH")
        .replace("SH EQUITY", ".SH")
        .replace("SZEQUITY", ".SZ")
        .replace("SZ EQUITY", ".SZ")
        .replace("BJEQUITY", ".BJ")
        .replace("BJ EQUITY", ".BJ")
        .replace("USEQUITY", ".US")
        .replace("US EQUITY", ".US")
    )

    raw = re.sub(r"^(\d{1,5})[-_/]?HK$", r"\1.HK", raw)
    raw = re.sub(r"^(\d{6})[-_/]?(SH|SZ|BJ)$", r"\1.\2", raw)

    m = re.fullmatch(r"(\d{1,5})\.HK", raw)
    if m:
        return m.group(1).zfill(5)

    m = re.fullmatch(r"(\d{6})\.(SH|SZ|BJ)", raw)
    if m:
        return m.group(1)

    m = re.fullmatch(r"([A-Z]{1,12})\.(O|N|US|NASDAQ|NYSE)", raw)
    if m:
        return m.group(1)

    if re.fullmatch(r"\d+", raw):
        hk_market = (
            "香港" in market_text
            or "港股" in market_text
            or "港通" in market_text
            or "联交所" in market_text
            or market_text in {"HK", "HKG", "SEHK"}
        )
        a_market = (
            "上海" in market_text
            or "深圳" in market_text
            or "北交" in market_text
            or "上交所" in market_text
            or "深交所" in market_text
            or market_text in {"SH", "SZ", "BJ", "SSE", "SZSE", "BSE"}
        )

        if hk_market:
            return raw.zfill(5)
        if a_market:
            return raw.zfill(6)

        if len(raw) == 6:
            return raw
        if len(raw) <= 5:
            return raw.zfill(5)
        return raw

    if re.fullmatch(r"[A-Z]{1,12}", raw):
        return raw

    return re.sub(r"\.(SH|SZ|BJ|HK|O|N|US|NASDAQ|NYSE)$", "", raw)


# ============================================================
# Excel 读取 / .xls 转换
# ============================================================

def excel_xls_to_xlsx_with_excel(xls_path: Path, output_path: Path) -> Path:
    """
    Windows + 本机 Excel：
    把旧 .xls 转为 .xlsx，以便 openpyxl 在保留格式的基础上修改。

    需要 pywin32：
        pip install pywin32
    """
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            "输入是 .xls，但当前环境无法直接保留格式编辑。\n"
            "请在 Windows + 已安装 Excel 的电脑运行，并安装 pywin32：\n"
            "    pip install pywin32\n"
            "或者先手工把组合指令另存为 .xlsx。"
        ) from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(xls_path.resolve()))
        # 51 = xlOpenXMLWorkbook (.xlsx)
        workbook.SaveAs(str(output_path.resolve()), FileFormat=51)
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if not output_path.exists():
        raise RuntimeError(f".xls 转 .xlsx 失败：{xls_path}")
    return output_path


def make_editable_xlsx_copy(instruction_path: Path) -> Tuple[Path, Optional[Path]]:
    """
    返回可由 openpyxl 编辑的 xlsx 路径。

    第二个返回值是需要最后清理的临时目录；若原文件本来就是 xlsx/xlsm，则为 None。
    """
    suffix = instruction_path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        return instruction_path, None

    if suffix == ".xls":
        tmpdir = Path(tempfile.mkdtemp(prefix="instruction_filter_"))
        converted = tmpdir / f"{instruction_path.stem}.xlsx"
        excel_xls_to_xlsx_with_excel(instruction_path, converted)
        return converted, tmpdir

    raise ValueError(f"不支持的组合指令格式：{instruction_path.suffix}")


def read_excel_any(path: Path, sheet_name: Any = 0, header: Any = 0) -> Any:
    """
    读取 Oplus。xlsx/xlsm 直接 pandas；旧 .xls 尝试 pandas，
    若环境缺 xlrd，则在 Windows 下临时用 Excel 转成 xlsx。
    """
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header)
    except Exception as first_exc:
        if path.suffix.lower() != ".xls":
            raise

        tmpdir = Path(tempfile.mkdtemp(prefix="oplus_xls_"))
        try:
            converted = tmpdir / f"{path.stem}.xlsx"
            excel_xls_to_xlsx_with_excel(path, converted)
            return pd.read_excel(converted, sheet_name=sheet_name, header=header, engine="openpyxl")
        except Exception:
            raise RuntimeError(f"无法读取 Oplus 文件：{path}\n原始错误：{first_exc}") from first_exc
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# Oplus 解析
# ============================================================

REQUIRED_OPLUS_COLUMNS = {
    "证券代码",
    "成交进度",
    "委托方向",
}


def find_header_row(path: Path, sheet_name: str) -> int:
    """
    自动寻找 Oplus 表头行。
    """
    raw = read_excel_any(path, sheet_name=sheet_name, header=None)
    if raw is None or len(raw) == 0:
        raise ValueError(f"Oplus 文件为空：{path} / {sheet_name}")

    scan_n = min(len(raw), 30)
    for i in range(scan_n):
        vals = {clean_text(v) for v in raw.iloc[i].tolist() if not _is_blank(v)}
        if "证券代码" in vals and "成交进度" in vals:
            return i

    raise ValueError(f"找不到 Oplus 表头（需要至少包含 证券代码、成交进度）：{path} / {sheet_name}")


def read_oplus_file(path: Path) -> pd.DataFrame:
    """
    读取一个 Oplus 文件。默认取第一个包含正确表头的 sheet。
    """
    # 先获取 sheet 名。
    try:
        sheets = read_excel_any(path, sheet_name=None, header=None)
    except Exception as exc:
        raise RuntimeError(f"读取 Oplus 失败：{path}") from exc

    if not isinstance(sheets, dict):
        sheets = {"Sheet1": sheets}

    last_error = None
    for sheet_name in sheets.keys():
        try:
            header_row = find_header_row(path, sheet_name)
            df = read_excel_any(path, sheet_name=sheet_name, header=header_row)
            df.columns = [clean_text(c) for c in df.columns]

            if REQUIRED_OPLUS_COLUMNS.issubset(set(df.columns)):
                df["_source_file"] = path.name
                df["_source_sheet"] = sheet_name
                return df
        except Exception as exc:
            last_error = exc
            continue

    raise ValueError(f"没有找到有效 Oplus sheet：{path}；最后错误：{last_error}")


def _progress_to_float(x: Any) -> float:
    if _is_blank(x):
        return 0.0
    s = clean_text(x).replace("%", "")
    try:
        v = float(s)
    except Exception:
        return 0.0

    # Oplus 示例里成交进度是 76 / 100 这种 0~100 数字。
    # 如果遇到 0.76 这种比例，则转成 76。
    if 0.0 <= v <= 1.0 and "." in s:
        v *= 100.0

    return float(np.clip(v, 0.0, 100.0))


def _safe_num(x: Any) -> float:
    try:
        if _is_blank(x):
            return np.nan
        return float(str(x).replace(",", ""))
    except Exception:
        return np.nan


def _business_date_series(df: pd.DataFrame, path: Path) -> pd.Series:
    if "业务日期" in df.columns:
        d = pd.to_datetime(df["业务日期"], errors="coerce")
        if d.notna().any():
            return d.dt.strftime("%Y-%m-%d")

    # 文件名兜底：oplusYYYYMMDD / oplusMMDD。
    stem = path.stem.lower()
    m8 = re.search(r"oplus[^0-9]*(\d{8})", stem)
    if m8:
        dt = pd.to_datetime(m8.group(1), format="%Y%m%d", errors="coerce")
        return pd.Series([dt.strftime("%Y-%m-%d")] * len(df), index=df.index)

    m4 = re.search(r"oplus[^0-9]*(\d{4})", stem)
    if m4:
        mmdd = m4.group(1)
        year = pd.Timestamp.today().year
        dt = pd.to_datetime(f"{year}{mmdd}", format="%Y%m%d", errors="coerce")
        return pd.Series([dt.strftime("%Y-%m-%d")] * len(df), index=df.index)

    return pd.Series([""] * len(df), index=df.index)


def code_candidates(code: Any) -> List[str]:
    """
    给组合指令匹配用的代码候选。

    Excel 可能把：
      00179 -> 179
      002422 -> 2422

    在组合指令本身缺少“交易市场”时无法只靠 2422 判断 A/H，
    因此同时尝试 5 位港股和 6 位 A 股补零形式。
    """
    base = normalize_code(code)
    out: List[str] = []

    def add(x: str) -> None:
        x = clean_text(x).upper()
        if x and x not in out:
            out.append(x)

    add(base)

    if re.fullmatch(r"\d+", base):
        stripped = base.lstrip("0") or "0"
        if len(stripped) <= 5:
            add(stripped.zfill(5))
            add(stripped.zfill(6))

    return out


def _make_instruction_key(
    code: str,
    side: str,
    instruction_text: str,
) -> Tuple[str, str, str]:
    candidates = code_candidates(code)
    canonical = candidates[0] if candidates else normalize_code(code)
    return (
        canonical,
        side,
        normalize_instruction_text(instruction_text),
    )


def extract_batch_progress(df: pd.DataFrame, path: Path) -> pd.DataFrame:
    """
    将一个 Oplus 文件整理成“每个批次的一条进度”。

    优先级：
      1) 同一批次存在 产品名称=混合 且 单元名称=混合 -> 直接用这条汇总行；
      2) 否则使用该批次 sum(成交数量)/sum(指令数量)；
      3) 再不行用该批次“成交进度”的最大值。
    """
    work = df.copy()

    for col in ["证券代码", "交易市场", "委托方向", "备注", "指令批号", "产品名称", "单元名称"]:
        if col not in work.columns:
            work[col] = ""

    work["_date"] = _business_date_series(work, path)
    work["_code"] = [
        normalize_code(c, m)
        for c, m in zip(work["证券代码"], work["交易市场"])
    ]
    work["_side"] = work["委托方向"].map(normalize_side)
    work["_remark"] = work["备注"].map(clean_text)
    work["_remark_norm"] = work["备注"].map(normalize_instruction_text)
    work["_progress"] = work["成交进度"].map(_progress_to_float)

    # 没有可匹配方向、代码、备注的不参与组合指令剔除。
    work = work[
        (work["_code"] != "")
        & (work["_side"] != "")
        & (work["_remark_norm"] != "")
    ].copy()

    if work.empty:
        return pd.DataFrame(
            columns=[
                "date", "batch_id", "code", "side",
                "remark", "remark_norm", "progress", "source_file"
            ]
        )

    # Oplus 里正常会有指令批号；没有时用 code+side+remark 当批次。
    work["_batch_id"] = work["指令批号"].map(clean_text)
    missing_batch = work["_batch_id"] == ""
    work.loc[missing_batch, "_batch_id"] = (
        work.loc[missing_batch, "_code"]
        + "|"
        + work.loc[missing_batch, "_side"]
        + "|"
        + work.loc[missing_batch, "_remark_norm"]
    )

    rows: List[Dict[str, Any]] = []

    group_cols = ["_date", "_batch_id", "_code", "_side", "_remark_norm"]
    for (date, batch_id, code, side, remark_norm), g in work.groupby(group_cols, dropna=False, sort=False):
        product = g["产品名称"].map(clean_text)
        unit = g["单元名称"].map(clean_text)

        summary = g[(product == "混合") & (unit == "混合")]

        if not summary.empty:
            # 同一批次理论上只有一条混合汇总行；若重复，取最大成交进度防止重复累计。
            progress = float(summary["_progress"].max())
            remark = clean_text(summary.iloc[0]["备注"])
            method = "混合汇总行"
        else:
            instr_qty = (
                g["指令数量"].map(_safe_num)
                if "指令数量" in g.columns
                else pd.Series(np.nan, index=g.index)
            )
            filled_qty = (
                g["成交数量"].map(_safe_num)
                if "成交数量" in g.columns
                else pd.Series(np.nan, index=g.index)
            )

            valid_qty = instr_qty.notna() & filled_qty.notna() & (instr_qty > 0)
            if valid_qty.any() and float(instr_qty[valid_qty].sum()) > 0:
                progress = 100.0 * float(filled_qty[valid_qty].sum()) / float(instr_qty[valid_qty].sum())
                progress = float(np.clip(progress, 0.0, 100.0))
                method = "成交数量/指令数量"
            else:
                progress = float(g["_progress"].max())
                method = "成交进度最大值"

            remark = clean_text(g.iloc[0]["备注"])

        rows.append(
            {
                "date": date,
                "batch_id": batch_id,
                "code": code,
                "side": side,
                "remark": remark,
                "remark_norm": remark_norm,
                "progress": progress,
                "progress_method": method,
                "source_file": path.name,
            }
        )

    return pd.DataFrame(rows)


def build_cumulative_progress(oplus_files: Iterable[Path]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    返回：
      daily_df      每条指令在各交易日看到的“累计成交进度”快照
      cumulative_df 每条指令最新交易日的累计成交进度

    重要口径：
    Oplus 的“成交进度”本身就是累计值，不是当天新增成交比例。
    因此：
      1. 同一天同一条指令若出现多条记录/批次，不把百分比相加；
         取当天可见的最高累计进度作为该日快照。
      2. 不同日期之间也不相加；按日期排序后，取最新一天的累计进度。
      3. 正式进度始终限制在 0~100%。
    """
    batch_parts: List[pd.DataFrame] = []

    for path in sorted(oplus_files, key=lambda p: p.name.lower()):
        batch = extract_batch_progress(read_oplus_file(path), path)
        if not batch.empty:
            batch_parts.append(batch)

    if not batch_parts:
        empty_daily = pd.DataFrame(
            columns=[
                "date", "code", "side", "remark", "remark_norm",
                "daily_progress", "batch_count", "source_files",
            ]
        )
        empty_cum = pd.DataFrame(
            columns=[
                "code", "side", "remark", "remark_norm",
                "cumulative_progress", "completed", "latest_date",
                "dates", "day_count",
            ]
        )
        return empty_daily, empty_cum

    batch_df = pd.concat(batch_parts, ignore_index=True)
    batch_df["progress"] = pd.to_numeric(batch_df["progress"], errors="coerce").fillna(0.0)
    batch_df["progress"] = batch_df["progress"].clip(lower=0.0, upper=100.0)

    # Oplus 的成交进度本身是累计值。
    # 同一天同一条指令即使出现多个批次/重复快照，也不能把百分比相加。
    # 累计进度理论上单调不减，因此当天取最高值作为当天最终快照。
    daily = (
        batch_df.groupby(
            ["date", "code", "side", "remark_norm"],
            dropna=False,
            as_index=False,
        )
        .agg(
            daily_progress=("progress", "max"),
            remark=("remark", "first"),
            batch_count=("batch_id", "nunique"),
            source_files=("source_file", lambda x: ",".join(sorted(set(map(str, x))))),
        )
    )
    daily["daily_progress"] = daily["daily_progress"].clip(lower=0.0, upper=100.0)

    # 为了正确判断“最新一天”，使用可解析日期排序；无法解析的日期放在最前，
    # 这样只要存在正常日期，就一定会由正常日期作为最新快照。
    daily["_date_sort"] = pd.to_datetime(daily["date"], errors="coerce")
    daily["_date_sort"] = daily["_date_sort"].fillna(pd.Timestamp.min)
    daily = daily.sort_values(
        ["code", "side", "remark_norm", "_date_sort"],
        kind="stable",
    ).reset_index(drop=True)

    # 记录所有出现日期，便于核对；正式成交进度取“最新日期”的累计值。
    history = (
        daily.groupby(
            ["code", "side", "remark_norm"],
            dropna=False,
            as_index=False,
        )
        .agg(
            dates=("date", lambda x: ",".join(dict.fromkeys(str(v) for v in x if str(v)))),
            day_count=("date", "nunique"),
            max_seen_progress=("daily_progress", "max"),
        )
    )

    latest = (
        daily.groupby(
            ["code", "side", "remark_norm"],
            dropna=False,
            as_index=False,
            sort=False,
        )
        .tail(1)
        [["code", "side", "remark_norm", "remark", "date", "daily_progress"]]
        .rename(
            columns={
                "date": "latest_date",
                "daily_progress": "cumulative_progress",
            }
        )
    )

    cumulative = latest.merge(
        history,
        on=["code", "side", "remark_norm"],
        how="left",
    )
    cumulative["cumulative_progress"] = (
        pd.to_numeric(cumulative["cumulative_progress"], errors="coerce")
        .fillna(0.0)
        .clip(lower=0.0, upper=100.0)
    )
    cumulative["completed"] = cumulative["cumulative_progress"] >= COMPLETION_THRESHOLD

    # 如果历史上曾经出现更高进度、但最新快照反而下降，保留一个核对标记。
    # 正式判断仍以最新日期为准，不擅自用旧值覆盖最新 Oplus 数据。
    cumulative["progress_decreased_vs_history"] = (
        cumulative["max_seen_progress"] > cumulative["cumulative_progress"] + 1e-9
    )

    cumulative = cumulative.sort_values(
        ["completed", "cumulative_progress", "code"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    daily = daily.drop(columns=["_date_sort"])
    return daily, cumulative


# ============================================================
# 组合指令识别
# ============================================================

def detect_code_in_cell(x: Any) -> str:
    """
    识别组合指令里的证券代码。
    """
    if _is_blank(x):
        return ""

    raw = clean_text(x).upper().strip("'\"")
    if not raw:
        return ""

    prepared = (
        raw.replace("HKEQUITY", ".HK")
        .replace("SHEQUITY", ".SH")
        .replace("SZEQUITY", ".SZ")
        .replace("BJEQUITY", ".BJ")
        .replace("USEQUITY", ".US")
    )
    prepared = re.sub(r"^(\d{1,5})[-_/]?HK$", r"\1.HK", prepared)
    prepared = re.sub(r"^(\d{6})[-_/]?(SH|SZ|BJ)$", r"\1.\2", prepared)

    if re.fullmatch(r"\d{1,5}\.HK", prepared):
        return normalize_code(prepared)
    if re.fullmatch(r"\d{6}\.(SH|SZ|BJ)", prepared):
        return normalize_code(prepared)
    if re.fullmatch(r"[A-Z]{1,12}\.(O|N|US|NASDAQ|NYSE)", prepared):
        return normalize_code(prepared)

    if re.fullmatch(r"\d+\.0", prepared):
        prepared = prepared[:-2]
    if re.fullmatch(r"\d{6}", prepared):
        return prepared
    if re.fullmatch(r"\d{1,5}", prepared):
        return prepared.zfill(5)

    return ""


def detect_code_in_row(values: List[Any]) -> str:
    """
    优先扫描整行带市场后缀的代码，再只检查前两列的裸数字代码。
    """
    for v in values:
        if _is_blank(v):
            continue

        raw = clean_text(v).upper().strip("'\"")
        prepared = (
            raw.replace("HKEQUITY", ".HK")
            .replace("SHEQUITY", ".SH")
            .replace("SZEQUITY", ".SZ")
            .replace("BJEQUITY", ".BJ")
            .replace("USEQUITY", ".US")
        )
        prepared = re.sub(r"^(\d{1,5})[-_/]?HK$", r"\1.HK", prepared)
        prepared = re.sub(r"^(\d{6})[-_/]?(SH|SZ|BJ)$", r"\1.\2", prepared)

        if (
            re.fullmatch(r"\d{1,5}\.HK", prepared)
            or re.fullmatch(r"\d{6}\.(SH|SZ|BJ)", prepared)
            or re.fullmatch(r"[A-Z]{1,12}\.(O|N|US|NASDAQ|NYSE)", prepared)
        ):
            return normalize_code(prepared)

    for v in values[:2]:
        code = detect_code_in_cell(v)
        if code:
            return code

    return ""


def completed_lookup(cumulative_df: pd.DataFrame) -> Dict[Tuple[str, str, str], Dict[str, Any]]:
    out: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

    if cumulative_df is None or cumulative_df.empty:
        return out

    done = cumulative_df[cumulative_df["completed"]].copy()

    for _, row in done.iterrows():
        key = _make_instruction_key(
            row["code"],
            row["side"],
            row["remark_norm"],
        )
        out[key] = row.to_dict()

    return out


def match_completed_instruction(
    code: str,
    instruction_text: str,
    done_lookup: Dict[Tuple[str, str, str], Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    side = instruction_side_from_text(instruction_text)
    text_norm = normalize_instruction_text(instruction_text)

    if not code or not side or not text_norm:
        return None

    candidates = code_candidates(code)

    # 1) 完全匹配
    for candidate in candidates:
        key = (candidate, side, text_norm)
        if key in done_lookup:
            return done_lookup[key]

    # 2) 安全同义词
    aliases = {text_norm}
    if "卖完" in text_norm:
        aliases.add(text_norm.replace("卖完", "清仓"))
    if "清仓" in text_norm:
        aliases.add(text_norm.replace("清仓", "卖完"))

    for candidate in candidates:
        for alias in aliases:
            key2 = (candidate, side, alias)
            if key2 in done_lookup:
                return done_lookup[key2]

    # 3) 语义匹配：代码 + 买卖方向 + 价格阈值 + 目标仓位
    sig = instruction_semantic_signature(instruction_text)
    if sig is not None:
        candidate_set = set(candidates)
        seen_ids = set()

        for row in done_lookup.values():
            obj_id = id(row)
            if obj_id in seen_ids:
                continue
            seen_ids.add(obj_id)

            if str(row.get("side", "")) != side:
                continue

            row_code_candidates = set(code_candidates(row.get("code", "")))
            if not (candidate_set & row_code_candidates):
                continue

            row_sig = instruction_semantic_signature(
                row.get("remark", "") or row.get("remark_norm", "")
            )
            if row_sig is not None and row_sig == sig:
                return row

    return None


# ============================================================
# 输出筛好的组合指令
# ============================================================

def filter_instruction_workbook(
    instruction_path: Path,
    cumulative_df: pd.DataFrame,
    output_path: Path,
) -> pd.DataFrame:
    """
    保留原工作簿布局，只把确认累计完成 >=100% 的指令单元格清空。
    返回删除记录 DataFrame。
    """
    editable_path, tmpdir = make_editable_xlsx_copy(instruction_path)

    try:
        keep_vba = editable_path.suffix.lower() == ".xlsm"
        wb = load_workbook(editable_path, keep_vba=keep_vba)

        done = completed_lookup(cumulative_df)
        removed_records: List[Dict[str, Any]] = []

        for ws in wb.worksheets:
            current_code = ""

            for row_idx in range(1, ws.max_row + 1):
                row_values = [
                    ws.cell(row_idx, col_idx).value
                    for col_idx in range(1, ws.max_column + 1)
                ]

                new_code = detect_code_in_row(row_values)
                if new_code:
                    current_code = new_code

                if not current_code:
                    continue

                for col_idx, value in enumerate(row_values, start=1):
                    if not looks_like_instruction_text(value):
                        continue

                    matched = match_completed_instruction(
                        current_code,
                        value,
                        done,
                    )
                    if matched is None:
                        continue

                    removed_records.append(
                        {
                            "sheet": ws.title,
                            "excel_row": row_idx,
                            "excel_col": col_idx,
                            "code": current_code,
                            "instruction": clean_text(value),
                            "oplus_remark": matched.get("remark", ""),
                            "side": matched.get("side", ""),
                            "cumulative_progress": matched.get("cumulative_progress", np.nan),
                            "dates": matched.get("dates", ""),
                        }
                    )

                    # 只清空值，不删除整行、不动格式。
                    ws.cell(row_idx, col_idx).value = None

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)

        return pd.DataFrame(removed_records)

    finally:
        if tmpdir is not None:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# 自动找文件
# ============================================================

def _mmdd_from_filename(p: Path) -> str:
    """从 oplusYYYYMMDD / oplusMMDD 文件名里取 MMDD。"""
    digits = "".join(re.findall(r"\d", p.stem))
    if len(digits) == 8:
        digits = digits[4:]
    return digits[-4:]


def find_oplus_files(project_root: Path, through_date: str = "") -> List[Path]:
    folder = None
    for name in OPLUS_FOLDER_CANDIDATES:
        cand = project_root / name
        if cand.exists():
            folder = cand
            break
    if folder is None:
        raise FileNotFoundError(
            f"找不到 Oplus 文件夹：在项目根下尝试了 {OPLUS_FOLDER_CANDIDATES}\n"
            f"请把 oplus<日期>.xls/xlsx 放入“{OPLUS_FOLDER_CANDIDATES[0]}”。"
        )

    files = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXCEL_SUFFIXES
        and p.stem.lower().startswith(OPLUS_PREFIX.lower())
        and not p.name.startswith("~$")
        and (not through_date or _mmdd_from_filename(p) <= through_date)
    ]

    if not files:
        raise FileNotFoundError(
            f"{folder} 里没有找到符合条件的 oplus<日期>.xls/xlsx 文件"
            f"（through_date={through_date or '不限'}）。"
        )

    return sorted(files, key=lambda p: (p.stat().st_mtime, p.name))


def find_latest_instruction_file(project_root: Path) -> Path:
    candidates: List[Path] = []

    for folder in [project_root, project_root / INSTRUCTION_FOLDER_NAME]:
        if not folder.exists():
            continue
        for p in folder.iterdir():
            if (
                p.is_file()
                and p.suffix.lower() in EXCEL_SUFFIXES
                and "组合指令" in p.stem.replace(" ", "")
                and OUTPUT_SUFFIX not in p.stem
                and not p.name.startswith("~$")
            ):
                candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            "没有找到组合指令 Excel。\n"
            "请把“组合指令*.xls/xlsx”放在项目根目录或“指令”文件夹，"
            "或者使用 --instruction-file 明确指定。"
        )

    return max(candidates, key=lambda p: (p.stat().st_mtime, p.name))


def default_output_path(instruction_path: Path) -> Path:
    return instruction_path.with_name(
        f"{instruction_path.stem}{OUTPUT_SUFFIX}.xlsx"
    )


# ============================================================
# 主程序
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="根据历史 Oplus 成交进度，剔除组合指令中已经累计完成 100% 的指令。"
    )
    parser.add_argument(
        "--project-root",
        default="",
        help="项目根目录；默认使用本 Python 文件所在目录。",
    )
    parser.add_argument(
        "--instruction-file",
        default="",
        help="组合指令文件路径；不填则自动找最新的“组合指令*.xls/xlsx”。",
    )
    parser.add_argument(
        "--output",
        default="",
        help="输出 xlsx 路径；不填则自动生成 *_剔除已成交.xlsx。",
    )
    parser.add_argument(
        "--through-date",
        default="",
        help="只使用业务日期<=该日的 oplus，格式 MMDD，例如 1014；不填则用到全部最新文件。",
    )

    args = parser.parse_args()
    through_date = args.through_date.strip()

    project_root = (
        Path(args.project_root).expanduser().resolve()
        if args.project_root
        else Path(__file__).resolve().parent
    )

    if args.instruction_file:
        instruction_path = Path(args.instruction_file).expanduser()
        if not instruction_path.is_absolute():
            instruction_path = project_root / instruction_path
        instruction_path = instruction_path.resolve()
    else:
        instruction_path = find_latest_instruction_file(project_root)

    if not instruction_path.exists():
        raise FileNotFoundError(f"找不到组合指令：{instruction_path}")

    oplus_files = find_oplus_files(project_root, through_date=through_date)

    output_path = (
        Path(args.output).expanduser()
        if args.output
        else default_output_path(instruction_path)
    )
    if not output_path.is_absolute():
        output_path = project_root / output_path
    output_path = output_path.resolve()

    print("=" * 72)
    print("组合指令成交筛选")
    print("=" * 72)
    print(f"项目目录：{project_root}")
    print(f"组合指令：{instruction_path}")
    print(f"Oplus 文件夹：{oplus_files[0].parent}（截止 {through_date or '最新'}）")
    print(f"Oplus 文件数：{len(oplus_files)}")
    for p in oplus_files:
        print(f"  - {p.name}")

    daily_df, cumulative_df = build_cumulative_progress(oplus_files)

    completed_count = int(cumulative_df["completed"].sum()) if not cumulative_df.empty else 0
    print()
    print(f"识别到的历史指令数：{len(cumulative_df)}")
    print(f"累计成交进度 >= {COMPLETION_THRESHOLD:g}%：{completed_count}")

    removed_df = filter_instruction_workbook(
        instruction_path=instruction_path,
        cumulative_df=cumulative_df,
        output_path=output_path,
    )

    print()
    print("=" * 72)
    print("完成")
    print("=" * 72)
    print(f"输出文件：{output_path}")
    print(f"从组合指令中剔除的已完成指令：{len(removed_df)} 条")

    if not removed_df.empty:
        print()
        print("剔除明细：")
        for _, r in removed_df.iterrows():
            print(
                f"  [{r['sheet']}!R{int(r['excel_row'])}C{int(r['excel_col'])}] "
                f"{r['code']} | {r['instruction']} | "
                f"累计进度={float(r['cumulative_progress']):.2f}%"
            )

    # 同目录额外生成两个核对 CSV，方便检查，不影响正式 Excel。
    audit_dir = output_path.parent
    daily_csv = audit_dir / f"{output_path.stem}_每日成交进度.csv"
    cum_csv = audit_dir / f"{output_path.stem}_累计成交进度.csv"

    daily_df.to_csv(daily_csv, index=False, encoding="utf-8-sig")
    cumulative_df.to_csv(cum_csv, index=False, encoding="utf-8-sig")

    print(f"每日进度核对：{daily_csv}")
    print(f"累计进度核对：{cum_csv}")


if __name__ == "__main__":
    main()
