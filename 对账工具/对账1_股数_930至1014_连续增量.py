import sys
from pathlib import Path
ROOT = Path(r'C:\Users\gucao\Desktop\live_grid_project')
sys.path.insert(0, str(ROOT))
import pandas as pd, numpy as np
import warnings
warnings.filterwarnings('ignore')
mod=__import__('live_grid_simulator_sim')
def load(fn):
    r=mod.read_excel_any(ROOT/'持仓数据'/fn,dtype={'证券代码':str})
    if isinstance(r,dict):r=next(iter(r.values()))
    return r
def real(r):
    ucol=mod._find_col(r.columns,mod.UNIT_COL_CANDIDATES);L={};S={}
    for _,x in r.iterrows():
        code=mod.canonical_code(x['证券代码'],market=x.get('交易市场',''))
        if not code:continue
        k=(mod.normalize_product_name(x['产品']),mod.normalize_unit_name(x[ucol]),code)
        q=pd.to_numeric(x['持仓数量'],errors='coerce');q=q if np.isfinite(q) else 0
        if x['多空标志']=='多仓':L[k]=L.get(k,0)+q
        else:S[k]=S.get(k,0)+q
    return L,S
r0=load('全球产品持仓查询_20261008.xls');rt=load('全球产品持仓查询_20261014.xls')
L0,S0=real(r0);Lt,St=real(rt)
L=dict(L0);S=dict(S0)
def dl_ds(d):
    d=str(d)
    if '担保品买入' in d:return (+1,0)
    if '担保品卖出' in d:return (-1,0)
    if '融券卖出' in d:return (0,+1)
    if '买券还券' in d:return (0,-1)
    if '买' in d:return (+1,0)
    if '卖' in d:return (-1,0)
    return (0,0)
seen={}
for f in ['oplus1008.xlsx','oplus1009.xlsx','oplus1012.xlsx','oplus1013.xlsx','oplus1014.xlsx']:
    df=pd.read_excel(ROOT/'oplus文件'/f,dtype={'证券代码':str})
    for _,x in df.iterrows():
        code=mod.canonical_code(x['证券代码'],market=x.get('交易市场',''))
        if not code or code=='混合':continue
        p=mod.normalize_product_name(x['产品名称']);u=mod.normalize_unit_name(x.get('单元名称',''))
        if p=='混合':continue
        oid=mod._clean_cell_str(x.get('指令序号'))
        tk=(oid,p,u,code,str(x['委托方向']))
        q=pd.to_numeric(x.get('成交数量'),errors='coerce');q=q if np.isfinite(q) else 0
        dl,ds=dl_ds(x['委托方向']);k=(p,u,code)
        if tk not in seen:seen[tk]=q;L[k]=L.get(k,0)+dl*q;S[k]=S.get(k,0)+ds*q
        else:dq=q-seen[tk];seen[tk]=q;L[k]=L.get(k,0)+dl*dq;S[k]=S.get(k,0)+ds*dq
badL=[(k,L.get(k,0),Lt.get(k,0),L.get(k,0)-Lt.get(k,0)) for k in set(L)|set(Lt) if abs(L.get(k,0)-Lt.get(k,0))>1]
badS=[(k,S.get(k,0),St.get(k,0),S.get(k,0)-St.get(k,0)) for k in set(S)|set(St) if abs(S.get(k,0)-St.get(k,0))>1]
bp={}
for b in badL:bp[b[0][0]]=bp.get(b[0][0],0)+1
for b in badS:bp[b[0][0]]=bp.get(b[0][0],0)+1
print(f'多仓差异{len(badL)} 融券差异{len(badS)}')
for p,c in sorted(bp.items(),key=lambda x:-x[1]):print(f'  {p}: {c}')
print('--- 多仓差异（全部）---')
for b in sorted(badL,key=lambda x:-abs(x[3])):print(b[0][0],b[0][1],b[0][2],f'模型{b[1]:.0f} 真实{b[2]:.0f} {b[3]:+.0f}')
