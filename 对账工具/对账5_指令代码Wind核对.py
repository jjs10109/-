# -*- coding: utf-8 -*-
import sys
from pathlib import Path
ROOT = Path(r'C:\Users\gucao\Desktop\live_grid_project')
sys.path.insert(0, str(ROOT))
import pandas as pd, warnings
warnings.filterwarnings('ignore')
m = __import__('live_grid_simulator_sim')
from data_loader import ensure_wind

p = ROOT / '指令' / '组合指令1013.xlsx'
sheets = m.read_excel_any(p, sheet_name=None)
df = sheets['股票指令']
seen = {}
for _, r in df.iterrows():
    raw = str(r.get('code', '') or '').strip()
    nm = str(r.get('证券名称', '') or '').strip()
    if not raw or raw.lower() == 'nan':
        continue
    code = m.canonical_code(raw)
    if not code or m.market_group(code) == 'OTHER':
        continue
    seen.setdefault(code, nm)

def wind_code(code):
    if code.endswith('.HK'):
        return code.split('.')[0].zfill(5) + '.HK'
    return code

w = ensure_wind()
def get_name(code):
    wc = wind_code(code)
    for attempt in (wc,):
        r = w.wss(attempt, 'sec_name')
        if r.Data and r.Data[0] and r.Data[0][0] not in (None, ''):
            return r.Data[0][0]
    # 港股试4位
    if code.endswith('.HK'):
        wc2 = code.split('.')[0].lstrip('0') or '0'
        wc2 = wc2.zfill(4) + '.HK'
        r = w.wss(wc2, 'sec_name')
        if r.Data and r.Data[0] and r.Data[0][0] not in (None, ''):
            return r.Data[0][0]
    return ''

def norm(s):
    return str(s or '').replace(' ', '').replace('Ａ','A').replace('Ｂ','B').replace('股份','').replace('集团','').replace('控股','').replace('中国','').strip()

bad = []
for code in seen:
    ins = seen[code]
    wn = get_name(code)
    ok = wn and (norm(ins) in norm(wn) or norm(wn) in norm(ins))
    tag = 'OK' if ok else 'XX'
    if not ok:
        bad.append((code, ins, wn))
    print(f'{tag} {code:12} 指令:{ins:12} Wind:{wn}')
print('\n不一致', len(bad), '只:')
for c, i, wn in bad:
    print(f'  {c} 指令填【{i}】 Wind【{wn}】')
