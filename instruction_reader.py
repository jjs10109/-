# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


def _is_blank(x: Any) -> bool:
    return x is None or (isinstance(x, float) and np.isnan(x)) or str(x).strip() == ""


def _clean_text(x: Any) -> str:
    if _is_blank(x):
        return ""
    return str(x).replace("％", "%").replace("，", ",").replace(" ", "").strip()


def _norm_basic_text(x: Any) -> str:
    if _is_blank(x):
        return ""
    s = str(x).strip().replace("％", "%").replace("，", ",").replace(" ", "")
    s = s.replace("．", ".").replace("。", ".")
    return s


def _looks_like_code(x: Any) -> bool:
    return bool(_detect_code_in_value(x))


def _norm_code_from_excel(x: Any) -> str:
    code = _detect_code_in_value(x)
    return code or _clean_text(x).upper()


def _detect_code_in_value(x: Any) -> str:
    """识别单个单元格里的证券代码。

    关键修复：不仅支持 01810.HK / 600519.SH，也支持 EWTX.O 这种美股代码。
    这样遇到下一只股票时会立刻切换，不会把下一只的指令串到上一只。
    """
    s = _norm_basic_text(x).upper()
    if not s or s == "NAN":
        return ""
    if any(k in s for k in ["资产", "备注", "合计", "代码", "名称", "方向", "产品"]):
        return ""

    s = s.replace("HKEQUITY", ".HK").replace("HKEQ", ".HK")
    s = s.replace("SHEQUITY", ".SH").replace("SZEQUITY", ".SZ")
    s = s.replace("USEQUITY", ".US")

    m = re.fullmatch(r"(\d{1,5})\.HK", s)
    if m:
        return f"{m.group(1).zfill(5)}.HK"
    m = re.fullmatch(r"(\d{6})\.(SH|SZ|BJ)", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}"
    m = re.fullmatch(r"[A-Z]{1,8}\.(O|N|US|NASDAQ|NYSE)", s)
    if m:
        return s
    # 裸 6 位 A 股代码：只当它出现在明显代码列/前两列时由 row-level 逻辑处理
    # 这里不直接识别，避免把价格 600000 之类误判。
    return ""


def _detect_code_in_row(row: List[Any]) -> Tuple[str, Optional[int]]:
    """扫描整行，返回新证券代码和所在列。

    规则：
    1. 优先识别带后缀的代码，如 01810.HK / EWTX.O / 600519.SH；
    2. 再识别前两列里的裸代码，如 01810 / 600519；
    3. 不扫描后面的指令列裸数字，避免把 49、53、24 这种价格当成代码。
    """
    for idx, v in enumerate(row):
        code = _detect_code_in_value(v)
        if code:
            return code, idx

    # 前两列裸数字作为代码候选。持仓/指令表常见：第一列或第二列是证券代码。
    for idx in range(min(2, len(row))):
        s = _norm_basic_text(row[idx]).upper()
        if not s or s == "NAN":
            continue
        if re.fullmatch(r"\d{1,5}", s):
            return f"{s.zfill(5)}.HK", idx
        if re.fullmatch(r"\d{6}", s):
            if s.startswith(("43", "83", "87", "88", "92")):
                return f"{s}.BJ", idx
            if s.startswith(("5", "6", "9")):
                return f"{s}.SH", idx
            return f"{s}.SZ", idx
    return "", None


def _is_instruction_text(s: str) -> bool:
    return any(k in s for k in ["买", "卖", "以下", "以上", "市价", "清仓", "加仓", "减仓", "前期", "上一行"])


def _is_ah_universe_scope_limited_instruction(text: str) -> bool:
    """本回测按可交易 A/H/美股组合处理，跳过只适用于 A股-only 组合的条件指令。"""
    s = _clean_text(text).upper()
    if not s:
        return False
    patterns = [
        "只能买A股", "只买A股", "仅买A股", "仅A股",
        "A股的几个组合", "只能买A股的几个组合",
        "只适用A股", "仅适用A股", "A股ONLY", "A股-ONLY",
    ]
    return any(p.upper() in s for p in patterns)


def _detect_name_in_row(row: List[Any], code_idx: Optional[int]) -> str:
    """根据代码位置找证券名称。

    支持两种常见格式：
    - code 在第0列，name 在第1列；
    - name 在第0列，code 在第1列（例如 Edgewise Therapeutics / EWTX.O）。
    """
    candidates: List[str] = []
    if code_idx is not None:
        for j in (code_idx + 1, code_idx - 1):
            if 0 <= j < len(row):
                t = _clean_text(row[j])
                if t and not _detect_code_in_value(t) and not _is_instruction_text(t):
                    candidates.append(t)
    for v in row[:4]:
        t = _clean_text(v)
        if not t or _detect_code_in_value(t) or _is_instruction_text(t):
            continue
        if re.search(r"[\u4e00-\u9fa5A-Za-z]", t):
            candidates.append(t)
    return candidates[0] if candidates else ""

def _convert_xls_to_xlsx(path: Path) -> Path:
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("读取 .xls 失败，且本机没有 libreoffice/soffice 可用于转换。建议把指令文件另存为 .xlsx。")
    tmpdir = Path(tempfile.mkdtemp(prefix="grid_instruction_"))
    cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(tmpdir), str(path)]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f".xls 转 .xlsx 失败：{proc.stderr or proc.stdout}")
    files = list(tmpdir.glob("*.xlsx"))
    if not files:
        raise RuntimeError(".xls 转换后没有找到 xlsx 文件。")
    return files[0]


def read_instruction_workbook(path: str | Path) -> Dict[str, pd.DataFrame]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"找不到指令文件：{path}")
    try:
        return pd.read_excel(path, sheet_name=None, header=None)
    except Exception as e1:
        if path.suffix.lower() == ".xls":
            converted = _convert_xls_to_xlsx(path)
            try:
                return pd.read_excel(converted, sheet_name=None, header=None, engine="openpyxl")
            except Exception as e2:
                raise RuntimeError(f"读取指令 Excel 失败：{e1}; 转换后读取也失败：{e2}") from e2
        raise RuntimeError(f"读取指令 Excel 失败：{e1}") from e1


def parse_instruction_cell(text: Any) -> Dict[str, Any]:
    raw = _clean_text(text)
    out = {
        "raw_instruction": raw,
        "side": "unknown",
        "price": np.nan,
        "target_pos": np.nan,
        "delta_pos": np.nan,
        "reduce_pos": np.nan,
        "order_basis": "unknown",  # target=目标仓位模板；delta=差额模板
        "is_market": False,
        "is_buyback_note": False,
        "parse_status": "unparsed",
        "parse_note": "",
    }
    if not raw:
        out["parse_status"] = "empty"
        return out

    if "买回" in raw or "上一行" in raw or "前期卖出" in raw or "重新买回" in raw:
        out.update({"side": "buy", "is_buyback_note": True})
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*(?:元|美元)?\s*(?:以下)?", raw)
        if m:
            out["price"] = float(m.group(1))
        m_back_to = re.search(r"(?:买回|买入|买回到|买入到|重新买回到)\s*到?\s*([0-9]+(?:\.[0-9]+)?)%", raw)
        if not m_back_to:
            m_back_to = re.search(r"到\s*([0-9]+(?:\.[0-9]+)?)%", raw)
        if m_back_to and not np.isnan(out["price"]):
            out["target_pos"] = float(m_back_to.group(1)) / 100.0
            out["order_basis"] = "target"
            out["parse_status"] = "ok"
            out["parse_note"] = "买回/重新买回指令已按买入目标仓位处理"
        else:
            out["parse_status"] = "note_only"
            out["parse_note"] = "买回上一行或前期卖出权重，但没有明确目标仓位，暂时作为备注"
        return out

    if "市价" in raw:
        out["is_market"] = True
    if "买" in raw and "卖" not in raw:
        out["side"] = "buy"
    elif "卖" in raw:
        out["side"] = "sell"

    m_price = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*(?:元|美元)?\s*(以上|以下)", raw)
    if m_price:
        out["price"] = float(m_price.group(1))
    else:
        m_price2 = re.search(r"^([0-9]+(?:\.[0-9]+)?)\s*(?:元|美元)?(?=.*(?:买入|卖出|买回|卖完|加仓|减仓))", raw)
        if m_price2:
            out["price"] = float(m_price2.group(1))
        elif out["is_market"]:
            out["price"] = np.nan

    if "卖完" in raw:
        out["side"] = "sell"
        out["target_pos"] = 0.0
        out["order_basis"] = "target"
        out["parse_status"] = "ok" if out["is_market"] or not np.isnan(out["price"]) else "market_only"
        return out

    m_buy_to = re.search(r"买入?到\s*([0-9]+(?:\.[0-9]+)?)%", raw)
    if m_buy_to:
        out["side"] = "buy"
        out["target_pos"] = float(m_buy_to.group(1)) / 100.0
        out["order_basis"] = "target"
        out["parse_status"] = "ok" if not np.isnan(out["price"]) else "missing_price"
        return out

    m_sell_to = re.search(r"卖出?到\s*([0-9]+(?:\.[0-9]+)?)%", raw)
    if m_sell_to:
        out["side"] = "sell"
        out["target_pos"] = float(m_sell_to.group(1)) / 100.0
        out["order_basis"] = "target"
        out["parse_status"] = "ok" if not np.isnan(out["price"]) else "missing_price"
        return out

    m_within = re.search(r"([0-9]+(?:\.[0-9]+)?)%以内", raw)
    if m_within and out["side"] == "sell":
        out["target_pos"] = float(m_within.group(1)) / 100.0
        out["order_basis"] = "target"
        out["parse_status"] = "ok" if not np.isnan(out["price"]) else "missing_price"
        return out

    m_reduce = re.search(r"(?:卖出|减仓)\s*([0-9]+(?:\.[0-9]+)?)%", raw)
    if m_reduce:
        out["side"] = "sell"
        out["reduce_pos"] = float(m_reduce.group(1)) / 100.0
        out["delta_pos"] = out["reduce_pos"]
        out["order_basis"] = "delta"
        out["parse_status"] = "reduce_only"
        out["parse_note"] = "这是卖出比例/差额，不是卖出后目标仓位；会输出到差额模板"
        return out

    m_add = re.search(r"(?:买入|加仓)\s*([0-9]+(?:\.[0-9]+)?)%", raw)
    if m_add:
        out["side"] = "buy"
        out["delta_pos"] = float(m_add.group(1)) / 100.0
        out["order_basis"] = "delta"
        out["parse_status"] = "add_only"
        out["parse_note"] = "这是买入比例/差额，不是买入后目标仓位；会输出到差额模板"
        return out

    out["parse_note"] = "没有识别出价格/方向/目标仓位"
    return out


def parse_instruction_excel(path: str | Path, include_sheets: Optional[List[str]] = None, sell_current_pos_default: float = 1.0, buy_current_pos_default: float = 0.0) -> Tuple[pd.DataFrame, pd.DataFrame]:
    sheets = read_instruction_workbook(path)
    parsed_rows: List[Dict[str, Any]] = []
    groups: Dict[str, Dict[str, Any]] = {}

    for sheet_name, df in sheets.items():
        if include_sheets and sheet_name not in include_sheets:
            continue
        if df.empty:
            continue

        current_code, current_name = "", ""
        start_row = 1
        for r in range(start_row, len(df)):
            row = df.iloc[r].tolist()

            # 关键修复：整行扫描新证券代码，而不是只看第一列。
            # 这样遇到 EWTX.O 这种下一只美股时，会重置 current_code，
            # 不会把 EWTX 的 24/28/32/36 误归到上一只小米集团。
            new_code, code_idx = _detect_code_in_row(row)
            if new_code:
                current_code = new_code
                detected_name = _detect_name_in_row(row, code_idx)
                if detected_name:
                    current_name = detected_name
                else:
                    current_name = ""

            if not current_code:
                continue

            # 指令单元格：跳过代码列和名称列，但扫描整行其余单元格。
            # 原版本固定 row[2:]，如果代码在第1列、名称在第0列，容易把行结构读错。
            for i, cell in enumerate(row):
                if code_idx is not None and i == code_idx:
                    continue
                text = _clean_text(cell)
                if not text:
                    continue
                if _is_ah_universe_scope_limited_instruction(text):
                    info = {
                        "raw_instruction": text,
                        "side": "unknown",
                        "price": np.nan,
                        "target_pos": np.nan,
                        "delta_pos": np.nan,
                        "reduce_pos": np.nan,
                        "order_basis": "unknown",
                        "is_market": False,
                        "is_buyback_note": False,
                        "parse_status": "scope_limited_ah_universe_skipped",
                        "parse_note": "只适用于只能买A股组合；本回测按可交易A/H/美股组合处理，跳过",
                        "sheet": sheet_name,
                        "excel_row": r + 1,
                        "instruction_col": f"col_{i + 1}",
                        "code": current_code,
                        "name": current_name,
                    }
                    parsed_rows.append(info)
                    continue
                if _detect_code_in_value(text):
                    continue
                if text == current_name:
                    continue
                # 没有交易关键词的不解析，避免把证券名称/产品名当作 unparsed instruction。
                if not _is_instruction_text(text):
                    continue

                info = parse_instruction_cell(text)
                info.update({
                    "sheet": sheet_name,
                    "excel_row": r + 1,
                    "instruction_col": f"col_{i + 1}",
                    "code": current_code,
                    "name": current_name,
                })
                parsed_rows.append(info)

    parsed_df = pd.DataFrame(parsed_rows)
    if parsed_df.empty:
        return parsed_df, pd.DataFrame()

    for code, sub in parsed_df.groupby("code", sort=False):
        name = sub["name"].dropna().iloc[0] if len(sub["name"].dropna()) else ""
        buy_levels = []
        sell_levels = []
        unparsed = []
        cumulative_reduce = 0.0
        cumulative_add = 0.0
        for _, row in sub.iterrows():
            side = row.get("side")
            price = row.get("price")
            target = row.get("target_pos")
            status = row.get("parse_status")
            raw = row.get("raw_instruction")
            if status == "ok" and side == "buy" and pd.notna(price) and pd.notna(target):
                buy_levels.append({"price": float(price), "target_pos": float(target), "raw_instruction": raw, "order_basis": "target", "delta_pos": np.nan})
            elif status == "ok" and side == "sell" and pd.notna(price) and pd.notna(target):
                sell_levels.append({"price": float(price), "target_pos": float(target), "raw_instruction": raw, "order_basis": "target", "delta_pos": np.nan})
            elif status == "reduce_only" and side == "sell" and pd.notna(price):
                delta = float(row.get("reduce_pos") or row.get("delta_pos") or 0.0)
                cumulative_reduce += delta
                sell_levels.append({
                    "price": float(price),
                    "target_pos": max(float(sell_current_pos_default) - cumulative_reduce, 0.0),
                    "raw_instruction": raw,
                    "order_basis": "delta",
                    "delta_pos": delta,
                })
            elif status == "add_only" and side == "buy" and pd.notna(price):
                delta = float(row.get("delta_pos") or 0.0)
                cumulative_add += delta
                buy_levels.append({
                    "price": float(price),
                    "target_pos": min(float(buy_current_pos_default) + cumulative_add, 1.0),
                    "raw_instruction": raw,
                    "order_basis": "delta",
                    "delta_pos": delta,
                })
            else:
                unparsed.append(raw)
        buy_levels = sorted(buy_levels, key=lambda x: x["price"], reverse=True)
        sell_levels = sorted(sell_levels, key=lambda x: x["price"])
        groups[code] = {
            "code": code,
            "name": name,
            "buy_levels": buy_levels,
            "sell_levels": sell_levels,
            "unparsed_instructions": [x for x in unparsed if x],
            "buy_count": len(buy_levels),
            "sell_count": len(sell_levels),
        }

    group_df = pd.DataFrame(list(groups.values()))

    # 数据质量检查：同一代码出现多个不同名称，说明原始指令代码列可能填错
    name_counts = parsed_df.groupby("code")["name"].nunique()
    bad_codes = name_counts[name_counts > 1].index.tolist()
    if bad_codes:
        print("[警告] 以下代码对应多个不同名称，原始指令代码列可能填错（会导致指令混档）：")
        for code in bad_codes:
            names = parsed_df[parsed_df["code"] == code]["name"].unique().tolist()
            print(f"  {code}: {names}")

    return parsed_df, group_df
