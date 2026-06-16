import pandas as pd, warnings
warnings.filterwarnings('ignore')

RAW='5月上险数据底表.xlsx'
FULL = pd.read_excel(RAW, sheet_name='Sheet1')
DF14 = FULL[FULL['月份'].between(1,4)]
CUR  = DF14[DF14['年份']==2026].copy()   # 2026 1-4月
PRE  = DF14[DF14['年份']==2025].copy()   # 2025 1-4月
Y25  = FULL[FULL['年份']==2025].copy()   # 2025全年

MSEG = '米段（轴距划分）'
VOL  = '容积段（新）'

def ssum(frame, f):
    if not f: return float(frame['数量'].sum())
    m = pd.Series(True, index=frame.index)
    for k,v in f.items():
        if isinstance(v,(list,tuple,set)): m &= frame[k].isin(v)
        else: m &= (frame[k]==v)
    return float(frame.loc[m,'数量'].sum())

def sub(frame, f):
    if not f: return frame
    m = pd.Series(True, index=frame.index)
    for k,v in f.items():
        if isinstance(v,(list,tuple,set)): m &= frame[k].isin(v)
        else: m &= (frame[k]==v)
    return frame.loc[m]

def fmt_pct(x):
    if x is None: return None
    if x=='净增长': return '净增长'
    return round(x*100,1)

def yoy(c,p):
    if p==0: return '净增长' if c>0 else None
    return c/p-1

def share(v,tot):
    return (v/tot) if (tot and tot>0) else None

def brands_by_sales(frame, f=None, base=None):
    fr = frame
    if base: fr = sub(fr, base)
    if f: fr = sub(fr, f)
    g = fr.groupby('企业品牌')['数量'].sum().sort_values(ascending=False)
    return [b for b,v in g.items() if v>0]
