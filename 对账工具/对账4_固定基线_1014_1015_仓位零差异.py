# -*- coding: utf-8 -*-
import sys
from pathlib import Path
ROOT = Path(r'C:\Users\gucao\Desktop\live_grid_project')
sys.path.insert(0, str(ROOT))
import pandas as pd, numpy as np, warnings, glob
warnings.filterwarnings('ignore')
m = __import__('live_grid_simulator_sim')

def raw_long_positions(path):
    sheets = m.read_excel_any(path, sheet_name=None)
    if not isinstance(sheets, dict):
        sheets = {'Sheet1': sheets}
    frames = []
    for sn, df in sheets.items():
        if df is None or df.empty:
            continue
        norm, _ = m._normalize_position_sheet(df, pd.DataFrame(columns=['product', 'unit']), str(sn))
        frames.append(norm)
    return pd.concat(frames, ignore_index=True)

files = {
    '1014': ROOT / '持仓数据' / '全球产品持仓查询_20261014.xls',
    '1015': ROOT / '持仓数据' / '全球产品持仓查询_20261015.xls',
}
for tag, p in files.items():
    if not p.exists():
        print(tag, 'missing', p.name)
        continue
    accounts = m.extract_accounts_from_position_file(p)
    built = m.build_current_positions(accounts, p, ROOT / '_test_state' / '_none.csv',
                                      oplus_trades=pd.DataFrame(), apply_oplus_trades=False,
                                      use_previous_state=False)
    raw = raw_long_positions(p)
    merged = built.merge(raw, on=['product', 'unit', 'code'], suffixes=('_built', '_raw'))
    diff = (merged['current_pos_built'] - merged['current_pos_raw']).abs()
    nz = merged[diff > 1e-6]
    print('=== %s (%s) ===' % (tag, p.name))
    print('  built=%d raw=%d matched=%d' % (len(built), len(raw), len(merged)))
    print('  mismatch rows=%d maxdiff=%.6f' % (len(nz), diff.max()))
    print('  short rows=%d' % int((built['short_qty'] > 0).sum()))
    x = built[(built['product'] == '净值线3号') & (built['unit'] == '东方') & (built['code'].str.startswith('603067'))]
    if len(x):
        print('  603067 zhenhua current_pos=%.4f (expect ~0.1117)' % x.iloc[0]['current_pos'])
