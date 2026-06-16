import sys, warnings
warnings.filterwarnings('ignore')
sys.path.insert(0,'/tmp')
from core import *
from wr import *
from openpyxl import Workbook
from openpyxl.utils import get_column_letter, column_index_from_string as ci

def C(s): return ci(s)

def merged(*dicts):
    out={}
    for d in dicts:
        if d: out.update(d)
    return out

BASE = {}  # sheet-wide scope (e.g. 大VAN), prepended to every filter

def cell(rgf, brand, cf, want, denom_rgf=None, curF=None, preF=None):
    """rgf: row-group filter; brand: 企业品牌 or None; cf: column condition; want: metric.
    curF/preF override the current/prior frames (e.g. Y25 for 2025全年 columns)."""
    if denom_rgf is None: denom_rgf = rgf
    if curF is None: curF = CUR
    if preF is None: preF = PRE
    bf = {'企业品牌':brand} if brand is not None else {}
    cur_v = ssum(curF, merged(BASE,rgf,cf,bf))
    pre_v = ssum(preF, merged(BASE,rgf,cf,bf))
    if want=='销量':
        return cur_v if cur_v>0 else None
    if want=='同比':
        return fmt_pct(yoy(cur_v,pre_v))
    if want=='占有率':
        tot=ssum(curF, merged(BASE,denom_rgf,cf)); s=share(cur_v,tot)
        return fmt_pct(s) if (s is not None and cur_v>0) else None
    if want=='占同比':
        sc=share(cur_v, ssum(curF, merged(BASE,denom_rgf,cf)))
        sp=share(pre_v, ssum(preF, merged(BASE,denom_rgf,cf)))
        if cur_v==0 and pre_v==0: return None
        return round(((sc or 0)-(sp or 0))*100,1)
    if want=='比重':
        # denom_rgf carries the PARENT row-group filter; brand filter applies to both
        den=ssum(curF, merged(BASE,denom_rgf,cf,bf)); s=share(cur_v,den)
        return fmt_pct(s) if (s is not None and cur_v>0) else None
    raise ValueError(want)

# COLSPEC: list of (col_index, want, cf). Build a row dict from a (rgf,brand) using a colspec.
def build_values(colspec, rgf, brand, denom_rgf=None):
    vals={}
    for spec in colspec:
        if len(spec)==4:
            col, want, cf, opt = spec
        else:
            col, want, cf = spec; opt={}
        dr = opt.get('denom', denom_rgf)
        b  = opt.get('brand', brand)
        vals[col]=cell(rgf, b, cf, want, dr,
                       curF=opt.get('curF'), preF=opt.get('preF'))
    return vals

def pct_set(colspec):
    return {spec[0] for spec in colspec if spec[1] in ('同比','占有率','占同比','比重')}

def bizhong(rgf, cf, parent_rgf, brand=None):
    """比重 = 本行销量(cf) / 上一级分类汇总销量(cf, 全品牌)."""
    bf = {'企业品牌':brand} if brand is not None else {}
    num = ssum(CUR, merged(BASE, rgf, cf, bf))
    den = ssum(CUR, merged(BASE, parent_rgf, cf))
    s = share(num, den)
    return fmt_pct(s) if (s is not None and num>0) else None

def ztb_set(colspec):
    return {spec[0] for spec in colspec if spec[1]=='占同比'}
