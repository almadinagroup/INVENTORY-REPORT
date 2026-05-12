"""
AL MADINA GROUP — Inventory Intelligence Dashboard
Single-file Streamlit app. Upload Excel → generate HTML dashboard.

pip install streamlit pandas python-calamine numpy
streamlit run app.py
"""

import streamlit as st
import json, gc, re, zipfile
from io import BytesIO
from collections import defaultdict
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DATA PROCESSING
# ═══════════════════════════════════════════════════════════════════════════════

MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

COLUMN_ALIASES = {
    'item bar code':'Item Bar Code','barcode':'Item Bar Code','item barcode':'Item Bar Code',
    'item name':'Item Name','name':'Item Name',
    'cost':'Cost','selling':'Selling','selling price':'Selling','selling1':'Selling',
    'stock':'Stock','qty':'Stock','quantity':'Stock',
    'brand':'Brand','category':'Category','class':'Class','group':'Group',
    'supplier':'Supplier','margin%':'Margin%','margin':'Margin%',
    'stock value':'Stock Value','profit':'Profit',
    'lp date':'LP Date','last purchase date':'LP Date',
    'lp qty':'LP Qty','last purchase qty':'LP Qty',
    'lp supplier':'LP Supplier','last purchase supplier':'LP Supplier',
    'total sales':'Total Sales','sales':'Total Sales',
}

def normalise_col(h):
    return COLUMN_ALIASES.get(str(h).strip().lower(), str(h).strip()) if h else ''

def detect_months(cols):
    pat = re.compile(r'(%s)[,.\s]+(\d{4})' % '|'.join(MONTH_NAMES), re.IGNORECASE)
    return [c for c in cols if pat.search(str(c))]

def month_sort_key(m):
    try: return datetime.strptime(re.sub(r'\s+',', ',str(m).strip().replace(',,',',')), '%b, %Y')
    except: return datetime.min


def read_excel_df(file_bytes):
    """Read Excel → DataFrame. Uses calamine (Python 3.14 compatible), falls back to stdlib."""
    import pandas as pd

    bio = BytesIO(file_bytes)
    engine = None
    for eng in ('calamine', 'openpyxl', 'xlrd'):
        try:
            pd.read_excel(bio, header=0, nrows=0, engine=eng)
            engine = eng
            bio.seek(0)
            break
        except Exception:
            bio.seek(0)

    if engine is None:
        return _read_excel_stdlib(file_bytes)

    df_h = pd.read_excel(bio, header=0, nrows=0, engine=engine)
    bio.seek(0)
    raw_cols  = list(df_h.columns)
    norm_cols = [normalise_col(c) for c in raw_cols]
    month_cols = detect_months(norm_cols)

    NEEDED = {'Item Bar Code','Item Name','Category','Class','Group','Brand','Supplier',
              'Cost','Selling','Stock','Stock Value','Margin%',
              'LP Date','LP Qty','LP Supplier','Total Sales'} | set(month_cols)

    orig_to_norm = dict(zip(raw_cols, norm_cols))
    keep_orig    = [c for c in raw_cols if orig_to_norm[c] in NEEDED]

    df = pd.read_excel(bio, header=0, usecols=keep_orig, engine=engine)
    df.rename(columns=orig_to_norm, inplace=True)
    df = df.loc[:, ~df.columns.duplicated()]

    num_cols = ['Cost','Selling','Stock','Stock Value','Margin%','LP Qty','Total Sales'] + month_cols
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('float32')

    for col in ['Category','Class','Group','Brand']:
        if col in df.columns:
            df[col] = df[col].fillna('').astype('category')

    for col in ['Item Name','Item Bar Code','Supplier','LP Supplier']:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str)

    df.dropna(how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df, month_cols


def _read_excel_stdlib(file_bytes):
    """Pure Python stdlib xlsx reader — zero dependencies."""
    import pandas as pd
    import xml.etree.ElementTree as ET

    NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    TR="{%s}row"%NS; TC="{%s}c"%NS; TV="{%s}v"%NS
    TT="{%s}t"%NS; TSI="{%s}si"%NS; TIS="{%s}is"%NS

    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names = set(i.filename for i in z.infolist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(TSI):
                shared.append("".join(t.text or "" for t in si.iter(TT)))
        sheet = next((n for n in sorted(names)
                      if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")), None)
        sb = z.read(sheet)

    def ci(ref):
        i=0
        for ch in ref:
            if not ch.isalpha(): break
            i=i*26+ord(ch.upper())-64
        return i-1

    rows, cur = [], {}
    for ev, el in ET.iterparse(BytesIO(sb), events=("start","end")):
        if ev=="start" and el.tag==TR: cur={}
        elif ev=="end" and el.tag==TC:
            ref=el.get("r","A1"); c=ci(ref); t=el.get("t","n")
            v=el.find(TV); is_=el.find(TIS)
            if is_ is not None:
                te=is_.find(TT); val=te.text if te is not None else ""
            elif v is not None:
                rv=v.text or ""
                if t=="s": val=shared[int(rv)] if shared else rv
                elif t=="b": val=bool(int(rv))
                else:
                    try: val=float(rv) if "." in rv else int(rv)
                    except: val=rv
            else: val=None
            cur[c]=val; el.clear()
        elif ev=="end" and el.tag==TR:
            if cur: rows.append([cur.get(i) for i in range(max(cur)+1)])
            el.clear()

    if not rows: raise RuntimeError("Excel file is empty.")
    raw_h = [str(v) if v is not None else "" for v in rows[0]]
    norm_h = [normalise_col(h) for h in raw_h]
    month_cols = detect_months(norm_h)
    data = []
    for row in rows[1:]:
        while len(row)<len(norm_h): row.append(None)
        if all(v is None or str(v).strip()=="" for v in row): continue
        data.append({norm_h[j]:row[j] for j in range(len(norm_h)) if norm_h[j]})
    import pandas as pd
    df = pd.DataFrame(data)
    num_cols = ['Cost','Selling','Stock','Stock Value','Margin%','LP Qty','Total Sales']+month_cols
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col],errors='coerce').fillna(0).astype('float32')
    for col in ['Category','Class','Group','Brand']:
        if col in df.columns: df[col]=df[col].fillna('').astype('category')
    for col in ['Item Name','Item Bar Code','Supplier','LP Supplier']:
        if col in df.columns: df[col]=df[col].fillna('').astype(str)
    df.dropna(how='all',inplace=True); df.reset_index(drop=True,inplace=True)
    return df, month_cols


def build_kpis(df, month_cols):
    last3 = sorted(month_cols, key=month_sort_key)[-3:] if len(month_cols)>=3 else month_cols
    l3 = df[[c for c in last3 if c in df.columns]].sum(axis=1).astype('float32') if last3 else 0

    stock = df['Stock']       if 'Stock'       in df.columns else 0
    ts    = df['Total Sales'] if 'Total Sales' in df.columns else 0
    sv    = df['Stock Value'] if 'Stock Value' in df.columns else 0
    lpq   = df['LP Qty']      if 'LP Qty'      in df.columns else 0
    lpd   = df['LP Date']     if 'LP Date'     in df.columns else None

    mask_zero = (stock>0)&(ts==0)
    mask_neg  = stock<0
    mask_oos  = (l3>0)&(stock<=0)
    mask_risk = (sv>200)&(ts<10)&(stock>0)
    mask_uw   = mask_zero&(lpd.notna() if lpd is not None else False)&(lpq<stock)

    return {
        'total_items': len(df),
        'in_stock':    int((stock>0).sum()),
        'total_sv':    round(float(sv.sum()),0),
        'total_ts':    round(float(ts.sum()),0),
        'neg_count':   int(mask_neg.sum()),
        'neg_sv':      round(float(sv[mask_neg].abs().sum()),0),
        'zero_count':  int(mask_zero.sum()),
        'zero_sv':     round(float(sv[mask_zero].sum()),0),
        'oos_count':   int(mask_oos.sum()),
        'risk_count':  int(mask_risk.sum()),
        'risk_sv':     round(float(sv[mask_risk].sum()),0),
        'uw_count':    int(mask_uw.sum()),
        'uw_sv':       round(float(sv[mask_uw].sum()),0),
    }, last3, mask_zero, mask_neg, mask_oos, mask_risk, mask_uw


def fmtN(n, d=1):
    try:
        n=float(n)
        if abs(n)>=1_000_000: return f"{n/1_000_000:.{d}f}M"
        if abs(n)>=1_000:     return f"{n/1_000:.{d}f}K"
        return f"{n:,.{d}f}"
    except: return "—"
def fmtA(n,d=0): return f"AED {fmtN(n,d)}"

def esc(s): return str(s).replace('&','&amp;').replace('"','&quot;').replace('<','&lt;').replace('>','&gt;')

def lp_label(lpd_val, lpq_val):
    """Return LP date string. If blank/NaT → 'Before 2025'."""
    import pandas as pd
    if lpd_val is None or (hasattr(lpd_val,'__class__') and 'NaT' in str(type(lpd_val))) or str(lpd_val).strip() in ('','nan','NaT','None','0'):
        return 'Before 2025'
    s = str(lpd_val).strip()
    if s in ('','nan','NaT','None'): return 'Before 2025'
    return s[:20]


def df_to_compact_json(df, fields, last3=None):
    """
    Convert DataFrame to compact list of dicts for JS rendering.
    fields: list of (key, col_name, type) — type: 'str','num','bc','lpd'
    Returns JSON string.
    """
    import pandas as pd
    rows = []
    for _, r in df.iterrows():
        row = {}
        for key, col, typ in fields:
            if col not in df.columns:
                row[key] = '' if typ in ('str','bc','lpd') else 0
                continue
            val = r[col]
            if typ == 'num':
                row[key] = round(float(val) if val is not None and str(val) not in ('','nan') else 0, 2)
            elif typ == 'lpd':
                row[key] = lp_label(val, 0)
            else:  # str, bc
                row[key] = str(val)[:50] if val is not None else ''
        # Add last3 months if needed
        if last3:
            row['m0'] = round(float(r.get(last3[2],0) or 0), 1) if len(last3)>2 else 0
            row['m1'] = round(float(r.get(last3[1],0) or 0), 1) if len(last3)>1 else 0
            row['m2'] = round(float(r.get(last3[0],0) or 0), 1) if last3 else 0
        rows.append(row)
    return json.dumps(rows, separators=(',',':'))


def cls_summary(df, mask_zero, mask_neg, mask_oos):
    """Build class summary table HTML."""
    if 'Class' not in df.columns: return ''
    ts_col = 'Total Sales' if 'Total Sales' in df.columns else None
    sv_col = 'Stock Value' if 'Stock Value' in df.columns else None

    agg = df.groupby('Class', observed=True).agg(
        sales=('Total Sales','sum') if ts_col else ('Class','count'),
        sv=('Stock Value','sum') if sv_col else ('Class','count'),
        items=('Item Name','count'),
    ).reset_index().sort_values('sales',ascending=False).head(12)

    zs  = df[mask_zero].groupby('Class',observed=True).size() if 'Class' in df.columns else {}
    neg = df[mask_neg].groupby('Class',observed=True).size()  if 'Class' in df.columns else {}
    oos = df[mask_oos].groupby('Class',observed=True).size()  if 'Class' in df.columns else {}

    rows = []
    for _, r in agg.iterrows():
        c=str(r['Class']); it=int(r['items'])
        z=int(zs.get(c,0)); ng=int(neg.get(c,0)); os_=int(oos.get(c,0) if hasattr(oos,'get') else 0)
        st = f'<span class="b-red">High Dead Stock</span>' if z>it*0.1 else (f'<span class="b-amber">OOS Risk</span>' if os_>it*0.1 else f'<span class="b-green">Healthy</span>')
        rows.append(f'<tr><td><span class="b-brand">{esc(c)}</span></td>'
                    f'<td class="tr">{fmtN(r["sales"],0)}</td><td class="tr">{fmtN(r["sv"],0)}</td>'
                    f'<td class="tr">{fmtN(it,0)}</td>'
                    f'<td class="tr {"c-amber" if z>50 else ""}">{fmtN(z,0)}</td>'
                    f'<td class="tr {"c-red" if ng>20 else ""}">{fmtN(ng,0)}</td>'
                    f'<td class="tr {"c-red" if os_>50 else ""}">{fmtN(os_,0)}</td>'
                    f'<td>{st}</td></tr>')
    return ''.join(rows)


# ─── HTML GENERATOR ────────────────────────────────────────────────────────────

def generate_html(file_bytes, source_filename=''):
    import pandas as pd

    df, month_cols = read_excel_df(file_bytes)
    months_sorted  = sorted(month_cols, key=month_sort_key)
    last3          = months_sorted[-3:] if len(months_sorted)>=3 else months_sorted
    last3_labels   = [months_sorted[-1] if months_sorted else '',
                      months_sorted[-2] if len(months_sorted)>1 else '',
                      months_sorted[-3] if len(months_sorted)>2 else '']

    df['L3'] = df[[c for c in last3 if c in df.columns]].sum(axis=1).astype('float32') if last3 else 0.0
    df['_bc'] = df['Item Bar Code'].astype(str) if 'Item Bar Code' in df.columns else ''

    K, last3, mask_zero, mask_neg, mask_oos, mask_risk, mask_uw = build_kpis(df, month_cols)

    sv_col = 'Stock Value' if 'Stock Value' in df.columns else 'Cost'
    ts_col = 'Total Sales' if 'Total Sales' in df.columns else 'L3'

    df_top  = df[df[ts_col]>0].sort_values(ts_col,ascending=False)
    df_zero = df[mask_zero].sort_values(sv_col,ascending=False)
    df_neg  = df[mask_neg].sort_values('Stock')
    df_oos  = df[mask_oos].sort_values('L3',ascending=False)
    df_risk = df[mask_risk].sort_values(sv_col,ascending=False)
    df_uw   = df[mask_uw].sort_values(sv_col,ascending=False)

    # ── Chart data ─────────────────────────────────────────────────────────
    ml  = [m.replace(', 20',"'") for m in months_sorted]
    mv  = [round(float(df[m].sum()),1) if m in df.columns else 0 for m in months_sorted]
    mc  = ['#1B2B4B']*(len(months_sorted)-1)+['rgba(37,99,235,0.45)']
    if 'Class' in df.columns and ts_col in df.columns:
        tc  = df.groupby('Class',observed=True)[ts_col].sum().sort_values(ascending=False).head(10)
        cl  = [c[:16]+'…' if len(c)>16 else c for c in tc.index.astype(str).tolist()]
        cv  = [round(float(v),1) for v in tc.values]
    else: cl,cv = [],[]
    cat_c = ['#1B2B4B','#1e3a5f','#1e40af','#1d4ed8','#2563eb','#3b82f6','#60a5fa','#0f766e','#0d9488','#14b8a6']
    hv = [max(0,K['in_stock']-K['zero_count']-K['risk_count']),K['zero_count'],K['neg_count'],K['oos_count'],K['risk_count']]

    # ── Compact JSON data for each section ────────────────────────────────
    TOP_FIELDS  = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('grp','Group','str'),('brand','Brand','str'),('sup','Supplier','str'),
                   ('ts','Total Sales','num'),('l3','L3','num'),('stock','Stock','num'),
                   ('sv','Stock Value','num'),('cost','Cost','num'),('sell','Selling','num'),('mg','Margin%','num')]
    ZERO_FIELDS = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('grp','Group','str'),('sup','Supplier','str'),
                   ('stock','Stock','num'),('sv','Stock Value','num'),
                   ('cost','Cost','num'),('sell','Selling','num'),
                   ('lpd','LP Date','lpd'),('lpq','LP Qty','num')]
    NEG_FIELDS  = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('cost','Cost','num'),('sell','Selling','num'),('stock','Stock','num'),
                   ('sv','Stock Value','num'),('ts','Total Sales','num'),('l3','L3','num'),
                   ('lpd','LP Date','lpd'),('lpq','LP Qty','num')]
    OOS_FIELDS  = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('sup','Supplier','str'),('sell','Selling','num'),('stock','Stock','num'),
                   ('ts','Total Sales','num'),('l3','L3','num')]
    RISK_FIELDS = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('cost','Cost','num'),('sell','Selling','num'),('stock','Stock','num'),
                   ('sv','Stock Value','num'),('ts','Total Sales','num'),('mg','Margin%','num'),
                   ('l3','L3','num'),('lpd','LP Date','lpd'),('lpq','LP Qty','num')]
    UW_FIELDS   = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('cost','Cost','num'),('sell','Selling','num'),('stock','Stock','num'),
                   ('sv','Stock Value','num'),('lpq','LP Qty','num'),('lpd','LP Date','lpd'),
                   ('sup','Supplier','str')]

    top_json  = df_to_compact_json(df_top,  TOP_FIELDS)
    zero_json = df_to_compact_json(df_zero, ZERO_FIELDS)
    neg_json  = df_to_compact_json(df_neg,  NEG_FIELDS)
    oos_json  = df_to_compact_json(df_oos,  OOS_FIELDS, last3=last3)
    risk_json = df_to_compact_json(df_risk, RISK_FIELDS)
    uw_json   = df_to_compact_json(df_uw,   UW_FIELDS)

    # Category list for filters
    cats = sorted(df['Category'].dropna().unique().astype(str).tolist()) if 'Category' in df.columns else []
    catcls = {}
    if 'Category' in df.columns and 'Class' in df.columns:
        for cat, cdf in df.groupby('Category',observed=True):
            catcls[str(cat)] = sorted(cdf['Class'].dropna().unique().astype(str).tolist())

    cls_table = cls_summary(df, mask_zero, mask_neg, mask_oos)

    # Free memory
    del df, df_top, df_zero, df_neg, df_oos, df_risk, df_uw
    gc.collect()

    # ── KPI cards ──────────────────────────────────────────────────────────
    kpi_defs = [
        ('Total SKUs',           fmtN(K['total_items'],0), 'Products in portfolio',                  'brand','kv-brand'),
        ('Items In Stock',       fmtN(K['in_stock'],0),    fmtA(K['total_sv'],0)+' stock value',    'green','kv-green'),
        ('Total Units Sold',     fmtN(K['total_ts'],0),    'Cumulative period total',                'blue', 'kv-blue'),
        ('Negative Stock Items', fmtN(K['neg_count'],0),   'Audit required urgently',               'red',  'kv-red'),
        ('Zero Sale — In Stock', fmtN(K['zero_count'],0),  fmtA(K['zero_sv'],0)+' capital at risk', 'red',  'kv-red'),
        ('OOS Active Demand',    fmtN(K['oos_count'],0),   'Selling but unavailable',               'red',  'kv-red'),
        ('High Risk Slow Movers',fmtN(K['risk_count'],0),  fmtA(K['risk_sv'],0)+' exposure',        'amber','kv-amber'),
        ('Unwanted Repurchases', fmtN(K['uw_count'],0),    fmtA(K['uw_sv'],0)+' re-bought, unsold', 'amber','kv-amber'),
    ]
    kpi_html = ''.join(
        f'<div class="kcard kc-{b}"><div class="klbl">{l}</div><div class="kval {v}">{va}</div><div class="ksub">{s}</div></div>'
        for l,va,s,b,v in kpi_defs)

    action_rows = [
        ('b-red',  'P1 — Critical','OOS Active Demand',     f"{K['oos_count']:,} items", 'c-red',  'Daily revenue loss',      'Emergency POs — top items by last-3M velocity',        'Buying Team',    'Today'),
        ('b-red',  'P1 — Critical','Unwanted Repurchases',  f"{K['uw_count']:,} items",  'c-red',  fmtA(K['uw_sv'],0),        'Supplier return + credit note negotiations',           'Category Mgmt',  'This Week'),
        ('b-amber','P2 — High',    'Negative Stock',        f"{K['neg_count']:,} items", 'c-amber','Reporting inaccuracy',    'Physical count + GRN audit + corrections',             'Operations',     '5 Business Days'),
        ('b-amber','P2 — High',    'Slow Mover Capital',    f"{K['risk_count']:,} items",'c-amber', fmtA(K['risk_sv'],0),     'Tiered markdown + bundle offers',                      'Merchandising',  '14 Days'),
        ('b-blue', 'P3 — Medium',  'Zero Sale Dead Stock',  f"{K['zero_count']:,} items",'c-blue',  fmtA(K['zero_sv'],0),    'Range review — discontinue or clearance',              'Category Mgmt',  '30 Days'),
        ('b-slate','P4 — Standard','New Listing Policy',    'Process change',            '',        'Future prevention',      'Mandatory 90-day sell-through KPI on new listings',    'Commercial Dir.','Next Quarter'),
    ]
    action_html = ''.join(
        f'<tr><td><span class="{b}">{p}</span></td><td>{i}</td><td>{sc}</td><td class="{ec}">{ex}</td><td>{ac}</td><td>{ow}</td><td>{dl}</td></tr>'
        for b,p,i,sc,ec,ex,ac,ow,dl in action_rows)

    rdate = datetime.now().strftime('%d %b %Y')
    dlabel = f'Source: {esc(source_filename)} &nbsp;|&nbsp; Generated: {rdate}' if source_filename else f'Generated: {rdate}'
    oos_h = [esc(last3_labels[0]),esc(last3_labels[1]),esc(last3_labels[2])]

    cat_opts = ''.join(f'<option value="{esc(c)}">{esc(c)}</option>' for c in ['All Categories']+cats)

    CSS = """:root{--bg:#F4F5F7;--surf:#FFF;--surf2:#F9FAFB;--bdr:#E5E7EB;--bdr2:#D1D5DB;--txt:#111827;--txt2:#374151;--mut:#6B7280;--brand:#1B2B4B;--brand-t:#EEF2FF;--red:#DC2626;--red-t:#FEF2F2;--red-b:#FECACA;--amb:#D97706;--amb-t:#FFFBEB;--amb-b:#FDE68A;--grn:#059669;--grn-t:#ECFDF5;--grn-b:#A7F3D0;--blu:#2563EB;--blu-t:#EFF6FF;--blu-b:#BFDBFE;--slt:#475569;--r:8px;--sh:0 1px 3px rgba(0,0,0,.08)}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--txt);font-family:"Inter",sans-serif;font-size:13.5px;line-height:1.55;-webkit-font-smoothing:antialiased}
.hdr{background:var(--brand);color:#fff;padding:0 32px;display:flex;align-items:center;justify-content:space-between;height:60px;position:sticky;top:0;z-index:200;border-bottom:1px solid rgba(255,255,255,.08)}
.hdr-l{display:flex;align-items:center;gap:16px}.logo{display:flex;align-items:center;gap:10px}
.lm{width:34px;height:34px;background:linear-gradient(135deg,#3B82F6,#60A5FA);border-radius:7px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;color:#fff}
.lt{font-weight:700;font-size:15px;letter-spacing:-.2px}.ls{font-size:10.5px;color:rgba(255,255,255,.5)}
.hsep{width:1px;height:28px;background:rgba(255,255,255,.12)}.hlbl{font-size:12px;color:rgba(255,255,255,.55)}
.hr{font-size:11px;color:rgba(255,255,255,.4);font-family:"IBM Plex Mono",monospace}
.nav{background:var(--brand);border-bottom:1px solid rgba(255,255,255,.08);display:flex;padding:0 32px;overflow-x:auto;gap:2px}
.nav::-webkit-scrollbar{height:0}
.nbtn{background:none;border:none;color:rgba(255,255,255,.45);font-family:"Inter",sans-serif;font-size:12px;font-weight:500;padding:10px 14px;cursor:pointer;white-space:nowrap;border-bottom:2px solid transparent;transition:color .15s;display:flex;align-items:center;gap:6px}
.nbtn:hover{color:rgba(255,255,255,.75)}.nbtn.active{color:#fff;border-bottom-color:#60A5FA}
.npill{display:inline-flex;align-items:center;justify-content:center;background:rgba(255,255,255,.12);border-radius:10px;font-size:10px;font-weight:600;min-width:18px;height:16px;padding:0 5px}
.npill.red{background:rgba(239,68,68,.35)}.npill.amb{background:rgba(245,158,11,.35)}
.page{display:none}.page.active{display:block}
.pi{max-width:1440px;margin:0 auto;padding:24px 32px 40px}
.ptitle{font-size:18px;font-weight:700;letter-spacing:-.3px;margin-bottom:3px}.pdesc{font-size:12.5px;color:var(--mut);margin-bottom:20px}
.fbar{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:13px 18px;display:flex;align-items:center;gap:12px;margin-bottom:20px;flex-wrap:wrap;box-shadow:var(--sh)}
.flbl{font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em}
.fsep{width:1px;height:20px;background:var(--bdr)}.fg{display:flex;align-items:center;gap:8px}.fg label{font-size:12px;color:var(--mut);font-weight:500}
select.fsel{background:var(--bg);border:1px solid var(--bdr2);border-radius:5px;color:var(--txt2);font-family:"Inter",sans-serif;font-size:12.5px;font-weight:500;padding:6px 10px;cursor:pointer;outline:none;min-width:160px}
select.fsel:focus{border-color:var(--blu)}.freset{background:none;border:1px solid var(--bdr2);border-radius:5px;color:var(--mut);font-family:"Inter",sans-serif;font-size:12px;font-weight:500;padding:6px 12px;cursor:pointer}
.freset:hover{background:var(--bg)}.fcnt{font-size:12px;color:var(--mut);font-family:"IBM Plex Mono",monospace}
.csv-btn{background:var(--brand);color:#fff;border:none;border-radius:5px;font-family:"Inter",sans-serif;font-size:12px;font-weight:600;padding:6px 14px;cursor:pointer;margin-left:auto;display:flex;align-items:center;gap:6px}
.csv-btn:hover{background:var(--brand-l)}
.kgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(195px,1fr));gap:12px;margin-bottom:24px}
.kcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:16px 18px;box-shadow:var(--sh);position:relative;overflow:hidden}
.kcard::after{content:"";position:absolute;top:0;left:0;right:0;height:3px;border-radius:8px 8px 0 0}
.kc-brand::after{background:var(--brand)}.kc-green::after{background:var(--grn)}.kc-red::after{background:var(--red)}.kc-amber::after{background:var(--amb)}.kc-blue::after{background:var(--blu)}
.klbl{font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px}
.kval{font-size:24px;font-weight:700;letter-spacing:-.5px;line-height:1;margin-bottom:4px}
.kv-brand{color:var(--brand)}.kv-green{color:var(--grn)}.kv-red{color:var(--red)}.kv-amber{color:var(--amb)}.kv-blue{color:var(--blu)}
.ksub{font-size:11.5px;color:var(--mut)}
.crow{display:grid;gap:16px;margin-bottom:20px}.c1{grid-template-columns:1fr}.c21{grid-template-columns:2fr 1fr}
.cbox{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:20px;box-shadow:var(--sh)}
.chead{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.ctitle{font-size:13px;font-weight:600;color:var(--txt)}.cmeta{font-size:11px;color:var(--mut)}
.ch{position:relative}.ch260{height:260px}
.tcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);box-shadow:var(--sh);overflow:hidden;margin-bottom:20px}
.thd{padding:13px 18px;border-bottom:1px solid var(--bdr);display:flex;align-items:center;justify-content:space-between;background:var(--surf2)}
.ttitle{font-size:13px;font-weight:600;color:var(--txt)}.tinfo{display:flex;align-items:center;gap:10px}
.tsc{overflow-x:auto}.tsc::-webkit-scrollbar{height:4px}.tsc::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:2px}
.tmh{max-height:520px;overflow-y:auto}.tmh::-webkit-scrollbar{width:5px}.tmh::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:3px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead tr{background:#F9FAFB;position:sticky;top:0;z-index:1}
th{text-align:left;padding:9px 12px;font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;border-bottom:1px solid var(--bdr)}
tbody tr{border-bottom:1px solid var(--bdr);transition:background .1s}tbody tr:last-child{border-bottom:none}tbody tr:hover{background:#F8FAFF}
td{padding:9px 12px;color:var(--txt2);vertical-align:middle}
.tn{font-weight:500;color:var(--txt);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}
.bc{font-family:"IBM Plex Mono",monospace;font-size:11px;color:var(--mut)}
.mono{font-family:"IBM Plex Mono",monospace;font-size:12px}.tr{text-align:right;font-family:"IBM Plex Mono",monospace;font-size:12px}
.c-red{color:var(--red)}.c-amber{color:var(--amb)}.c-green{color:var(--grn)}.c-blue{color:var(--blu)}.muted{color:var(--mut)}.fw6{font-weight:600}.bold-green{color:var(--grn);font-weight:700}
.sup{max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}
.b-red{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--red-t);color:var(--red);border:1px solid var(--red-b);white-space:nowrap}
.b-amber{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--amb-t);color:var(--amb);border:1px solid var(--amb-b);white-space:nowrap}
.b-green{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--grn-t);color:var(--grn);border:1px solid var(--grn-b);white-space:nowrap}
.b-blue{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--blu-t);color:var(--blu);border:1px solid var(--blu-b);white-space:nowrap}
.b-slate{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:#F1F5F9;color:var(--slt);border:1px solid #CBD5E1;white-space:nowrap}
.b-brand{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--brand-t);color:var(--brand);border:1px solid #C7D2FE;white-space:nowrap}
.arow{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-bottom:20px}
.abox{border-radius:var(--r);padding:14px 16px;border:1px solid}
.abox.red{background:var(--red-t);border-color:var(--red-b)}.abox.amb{background:var(--amb-t);border-color:var(--amb-b)}
.abox.grn{background:var(--grn-t);border-color:var(--grn-b)}.abox.blu{background:var(--blu-t);border-color:var(--blu-b)}
.ah{display:flex;align-items:flex-start;gap:8px;margin-bottom:5px}.ai{font-size:14px;flex-shrink:0;font-style:normal}
.at{font-size:12.5px;font-weight:600;color:var(--txt)}.ab{font-size:12px;color:var(--txt2);line-height:1.5;padding-left:22px}
.strip{background:var(--brand);color:#fff;border-radius:var(--r);padding:12px 20px;display:flex;align-items:center;gap:24px;margin-bottom:20px;flex-wrap:wrap}
.si{display:flex;flex-direction:column;gap:1px}.sl{font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em}
.sv{font-size:16px;font-weight:700;font-family:"IBM Plex Mono",monospace;letter-spacing:-.5px}
.sv.red{color:#FCA5A5}.sv.amb{color:#FCD34D}.sv.grn{color:#6EE7B7}.ss{width:1px;height:32px;background:rgba(255,255,255,.12)}
.rank-list{display:flex;flex-direction:column;gap:7px}
.rank-row{display:grid;grid-template-columns:28px 1fr 120px 70px;align-items:center;gap:10px}
.rn{font-size:11px;color:var(--mut);font-weight:600;text-align:right}
.rnm{font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rb{background:var(--bdr);border-radius:3px;height:8px}.rf{height:100%;border-radius:3px;background:var(--brand)}
.rv{font-size:11.5px;font-family:"IBM Plex Mono",monospace;text-align:right;color:var(--mut)}
.footer{text-align:center;padding:20px 32px;border-top:1px solid var(--bdr);font-size:11.5px;color:var(--mut)}
@media(max-width:900px){.pi{padding:16px}.hdr,.nav{padding-left:16px;padding-right:16px}.crow.c21{grid-template-columns:1fr}.kgrid{grid-template-columns:repeat(2,1fr)}}"""

    JS = f"""
var CATS={json.dumps(cats,separators=(',',':'))};
var CATCLS={json.dumps(catcls,separators=(',',':'))};
var TOP_DATA={top_json};
var ZERO_DATA={zero_json};
var NEG_DATA={neg_json};
var OOS_DATA={oos_json};
var RISK_DATA={risk_json};
var UW_DATA={uw_json};
var ML={json.dumps(ml)};var MV={json.dumps(mv)};var MC={json.dumps(mc)};
var CL={json.dumps(cl)};var CV={json.dumps(cv)};var CC={json.dumps(cat_c[:len(cl)])};
var HV={json.dumps(hv)};
var OOS_LABELS={json.dumps(oos_h)};
var K={json.dumps(K)};

function SP(id){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nbtn').forEach(b=>b.classList.remove('active'));
  var pg=document.getElementById('pg-'+id),nb=document.getElementById('nb-'+id);
  if(pg)pg.classList.add('active');if(nb)nb.classList.add('active');
  window.scrollTo(0,0);
}}
var CF={{family:"'Inter',sans-serif",size:11}};
Chart.defaults.font=CF;Chart.defaults.color='#6B7280';
function fmtT(v){{return v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':Number(v).toLocaleString('en');}}
function fmtN(v,d=1){{v=parseFloat(v)||0;if(Math.abs(v)>=1e6)return(v/1e6).toFixed(d)+'M';if(Math.abs(v)>=1e3)return(v/1e3).toFixed(d)+'K';return v.toLocaleString('en',{{maximumFractionDigits:d}});}}
function esc(s){{return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}}

new Chart(document.getElementById('chM'),{{type:'bar',data:{{labels:ML,datasets:[{{data:MV,backgroundColor:MC,borderRadius:4,borderSkipped:false}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>' '+Number(c.parsed.y).toLocaleString()+' units'}}}}}},scales:{{x:{{grid:{{display:false}},ticks:{{font:CF}}}},y:{{grid:{{color:'#F3F4F6'}},ticks:{{font:CF,callback:v=>fmtT(v)}},border:{{display:false}}}}}}}}}});
new Chart(document.getElementById('chC'),{{type:'bar',data:{{labels:CL,datasets:[{{data:CV,backgroundColor:CC,borderRadius:4}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>' '+Number(c.parsed.x).toLocaleString()+' units'}}}}}},scales:{{x:{{grid:{{color:'#F3F4F6'}},ticks:{{font:CF,callback:v=>fmtT(v)}},border:{{display:false}}}},y:{{grid:{{display:false}},ticks:{{font:CF}}}}}}}}}});
new Chart(document.getElementById('chH'),{{type:'doughnut',data:{{labels:['Healthy','Zero Sale','Negative','OOS','High Risk'],datasets:[{{data:HV,backgroundColor:['#059669','#D97706','#DC2626','#7C3AED','#1B2B4B'],borderWidth:2,borderColor:'#fff'}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'62%',plugins:{{legend:{{position:'bottom',labels:{{font:{{family:"'Inter'",size:10}},padding:8,boxWidth:10}}}}}}}}}});

// ── FILTER LOGIC ──
function filterData(data,cat,cls){{
  if((!cat||cat==='All Categories')&&(!cls||cls==='All Classes'))return data;
  return data.filter(r=>{{
    var cm=!cat||cat==='All Categories'||r.cat===cat;
    var clm=!cls||cls==='All Classes'||r.cls===cls;
    return cm&&clm;
  }});
}}
function updClsOpts(selId,cat){{
  var el=document.getElementById(selId);if(!el)return;
  var list=(cat&&cat!=='All Categories'&&CATCLS[cat])||[];
  el.innerHTML='<option value="All Classes">All Classes</option>'+list.map(c=>'<option value="'+esc(c)+'">'+esc(c)+'</option>').join('');
}}
function fmtS(n){{n=parseFloat(n)||0;if(Math.abs(n)>=1e6)return(n/1e6).toFixed(1)+'M';if(Math.abs(n)>=1e3)return(n/1e3).toFixed(1)+'K';return n.toLocaleString('en',{{maximumFractionDigits:0}});}}
function mkStrip(id,items){{document.getElementById(id).innerHTML='<div class="strip">'+items.map((it,i)=>(i?'<div class="ss"></div>':'')+'<div class="si"><div class="sl">'+it[0]+'</div><div class="sv '+it[2]+'">'+it[1]+'</div></div>').join('')+'</div>';}}

// ── CSV DOWNLOAD ──
function downloadCSV(data,headers,filename){{
  var rows=[headers.join(',')].concat(data.map(r=>headers.map(h=>{{var v=r[h]===undefined?'':r[h];return '"'+String(v).replace(/"/g,'""')+'"';}}).join(',')));
  var blob=new Blob([rows.join('\\n')],{{type:'text/csv;charset=utf-8;'}});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);
}}

// ── ROW RENDERERS ──
function rowTop(r,i){{var mc=r.mg<10?'c-red':(r.mg>30?'c-green fw6':'');return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="muted">'+esc(r.brand||'—')+'</td><td class="tr bold-green">'+fmtN(r.ts,0)+'</td><td class="tr c-green">'+fmtN(r.l3,0)+'</td><td class="tr">'+fmtN(r.stock,1)+'</td><td class="tr">'+fmtN(r.sv,0)+'</td><td class="tr">'+r.cost.toFixed(2)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td class="tr '+mc+'">'+r.mg.toFixed(1)+'%</td></tr>';}}
function rowZero(r,i){{return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="muted">'+esc(r.grp||'—')+'</td><td class="tr c-amber">'+fmtN(r.stock,1)+'</td><td class="tr c-red fw6">'+fmtN(r.sv,0)+'</td><td class="tr">'+r.cost.toFixed(2)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td>'+esc(r.lpd)+'</td><td class="tr">'+(r.lpq||'—')+'</td><td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';}}
function rowNeg(r,i){{return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="tr c-red fw6">'+fmtN(r.stock,1)+'</td><td class="tr c-red">'+fmtN(r.sv,0)+'</td><td class="tr">'+fmtN(r.ts,0)+'</td><td class="tr c-green">'+fmtN(r.l3,0)+'</td><td class="tr">'+r.cost.toFixed(2)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td>'+esc(r.lpd)+'</td><td class="tr">'+(r.lpq||'—')+'</td></tr>';}}
function rowOOS(r,i){{return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="tr c-green fw6">'+fmtN(r.l3,0)+'</td><td class="tr">'+fmtN(r.m0,1)+'</td><td class="tr">'+fmtN(r.m1,1)+'</td><td class="tr">'+fmtN(r.m2,1)+'</td><td class="tr c-red fw6">'+fmtN(r.stock,1)+'</td><td class="tr">'+fmtN(r.ts,0)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';}}
function rowRisk(r,i){{return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="tr c-amber fw6">'+fmtN(r.sv,0)+'</td><td class="tr">'+fmtN(r.stock,1)+'</td><td class="tr c-red">'+fmtN(r.ts,0)+'</td><td class="tr">'+(r.l3||0).toFixed(0)+'</td><td class="tr">'+r.cost.toFixed(2)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td class="tr">'+r.mg.toFixed(1)+'%</td><td>'+esc(r.lpd)+'</td><td class="tr">'+(r.lpq||'—')+'</td></tr>';}}
function rowUW(r,i){{var sup=r.sup||'—';return '<tr><td class="mono">'+(i+1)+'</td><td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td><td><span class="b-slate">'+esc(r.cat)+'</span></td><td><span class="b-brand">'+esc(r.cls)+'</span></td><td class="tr c-amber">'+fmtN(r.stock,1)+'</td><td class="tr c-red fw6">'+fmtN(r.sv,0)+'</td><td class="tr">'+r.cost.toFixed(2)+'</td><td class="tr">'+r.sell.toFixed(2)+'</td><td class="tr">'+(r.lpq||'—')+'</td><td>'+esc(r.lpd)+'</td><td><span class="sup">'+esc(sup)+'</span></td></tr>';}}

function renderTable(tbodyId,data,rowFn){{document.getElementById(tbodyId).innerHTML=data.map((r,i)=>rowFn(r,i)).join('');}}
function renderRank(id,data){{var mx=data.length?data[0].ts:1;document.getElementById(id).innerHTML=data.slice(0,10).map((r,i)=>'<div class="rank-row"><div class="rn">'+(i+1)+'</div><div class="rnm" title="'+esc(r.n)+'">'+esc(r.n)+'</div><div class="rb"><div class="rf" style="width:'+(r.ts/mx*100).toFixed(1)+'%"></div></div><div class="rv">'+fmtN(r.ts,0)+'</div></div>').join('');}}

// ── SECTION CONTROLLERS ──
var _state={{}};
function initSection(id,allData,rowFn,stripCfg,csvHeaders){{
  var catEl=document.getElementById(id+'-cat'),clsEl=document.getElementById(id+'-cls');
  function run(){{
    var cat=catEl?catEl.value:'All Categories',cls=clsEl?clsEl.value:'All Classes';
    var d=filterData(allData,cat,cls);
    renderTable(id+'-tb',d,rowFn);
    if(id==='top'){{renderRank('top-rank',d);document.getElementById('top-cnt').textContent='Showing '+d.length+' items';return;}}
    document.getElementById(id+'-cnt').textContent='Showing '+d.length+' items';
    if(stripCfg){{
      var sv=d.reduce((s,r)=>s+(r.sv||0),0);
      var l3=d.reduce((s,r)=>s+(r.l3||0),0);
      var items=stripCfg(d,sv,l3);
      mkStrip(id+'-strip',items);
    }}
    _state[id]=d;
  }}
  if(catEl)catEl.addEventListener('change',()=>{{updClsOpts(id+'-cls',catEl.value);run();}});
  if(clsEl)clsEl.addEventListener('change',run);
  window[id+'Reset']=()=>{{if(catEl)catEl.value='All Categories';updClsOpts(id+'-cls','All Categories');run();}};
  window[id+'CSV']=()=>{{var d=_state[id]||allData;downloadCSV(d,csvHeaders,id+'_export.csv');}};
  if(clsEl)updClsOpts(id+'-cls','All Categories');
  _state[id]=allData;
  run();
}}

initSection('top',TOP_DATA,rowTop,null,['bc','n','cat','cls','brand','ts','l3','stock','sv','cost','sell','mg']);
initSection('zero',ZERO_DATA,rowZero,
  (d,sv)=>[['Items Shown',fmtS(d.length),''],['Stock Value at Risk','AED '+fmtS(sv),'amb'],['Avg / Item',d.length?'AED '+fmtS(sv/d.length):'—',''],['Total Zero-Sale',fmtS(K.zero_count),'red']],
  ['bc','n','cat','cls','grp','sup','stock','sv','cost','sell','lpd','lpq']);
initSection('neg',NEG_DATA,rowNeg,
  (d,sv)=>[['Items Shown',fmtS(d.length),''],['Neg. Value Exposure','AED '+fmtS(Math.abs(sv)),'red'],['Total Negative',fmtS(K.neg_count),'red'],['Priority','CRITICAL','amb']],
  ['bc','n','cat','cls','cost','sell','stock','sv','ts','l3','lpd','lpq']);
initSection('oos',OOS_DATA,rowOOS,
  (d,sv,l3)=>[['OOS Items',fmtS(d.length),'red'],['Last 3M Sales',fmtS(l3)+' units','grn'],['Total OOS Portfolio',fmtS(K.oos_count),'red'],['Revenue Risk','HIGH','amb']],
  ['bc','n','cat','cls','sup','sell','stock','ts','l3','m0','m1','m2']);
initSection('risk',RISK_DATA,rowRisk,
  (d,sv)=>[['Items Shown',fmtS(d.length),''],['Value at Risk','AED '+fmtS(sv),'amb'],['Portfolio Exposure','AED '+fmtS(K.risk_sv),'red'],['Total High Risk',fmtS(K.risk_count),'amb']],
  ['bc','n','cat','cls','cost','sell','stock','sv','ts','mg','l3','lpd','lpq']);
initSection('uw',UW_DATA,rowUW,
  (d,sv)=>[['Items Shown',fmtS(d.length),''],['Stock Value','AED '+fmtS(sv),'red'],['Portfolio Total','AED '+fmtS(K.uw_sv),'red'],['Total Unwanted',fmtS(K.uw_count),'red']],
  ['bc','n','cat','cls','cost','sell','stock','sv','lpq','lpd','sup']);
"""

    def fbar(prefix, has_rank=False):
        rank_id = f'id="{prefix}-rank"' if has_rank else ''
        return f'''<div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="{prefix}-cat">{cat_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="{prefix}-cls"></select></div>
    <button class="freset" onclick="{prefix}Reset()">Reset</button>
    <span class="fcnt" id="{prefix}-cnt"></span>
    <button class="csv-btn" onclick="{prefix}CSV()">&#8595; CSV</button>
  </div>'''

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AL MADINA GROUP — Inventory Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="hdr">
  <div class="hdr-l">
    <div class="logo"><div class="lm">AM</div><div><div class="lt">AL MADINA GROUP</div><div class="ls">Inventory Management System</div></div></div>
    <div class="hsep"></div><div class="hlbl">Stock Intelligence Report</div>
  </div>
  <div class="hr">{fmtN(K['total_items'],0)} SKUs &nbsp;|&nbsp; {dlabel}</div>
</header>
<nav class="nav">
  <button class="nbtn active" id="nb-ov"   onclick="SP('ov')">Overview</button>
  <button class="nbtn"        id="nb-top"  onclick="SP('top')">Top Movers</button>
  <button class="nbtn"        id="nb-zero" onclick="SP('zero')">Zero Sale Stock <span class="npill amb">{fmtN(K['zero_count'],0)}</span></button>
  <button class="nbtn"        id="nb-neg"  onclick="SP('neg')">Negative Stock <span class="npill red">{fmtN(K['neg_count'],0)}</span></button>
  <button class="nbtn"        id="nb-oos"  onclick="SP('oos')">OOS / Active Demand <span class="npill red">{fmtN(K['oos_count'],0)}</span></button>
  <button class="nbtn"        id="nb-risk" onclick="SP('risk')">High Risk Slow Movers <span class="npill amb">{fmtN(K['risk_count'],0)}</span></button>
  <button class="nbtn"        id="nb-uw"   onclick="SP('uw')">Unwanted Repurchases <span class="npill red">{fmtN(K['uw_count'],0)}</span></button>
  <button class="nbtn"        id="nb-ins"  onclick="SP('ins')">Insights &amp; Actions</button>
</nav>

<div id="pg-ov" class="page active"><div class="pi">
  <div class="ptitle">Inventory Overview</div>
  <div class="pdesc">{fmtN(K['total_items'],0)} SKUs — all categories and classes.</div>
  <div class="kgrid">{kpi_html}</div>
  <div class="crow c1"><div class="cbox">
    <div class="chead"><span class="ctitle">Monthly Sales Volume (Units)</span><span class="cmeta">Most recent month may be partial</span></div>
    <div class="ch ch260"><canvas id="chM"></canvas></div>
  </div></div>
  <div class="crow c21">
    <div class="cbox"><div class="chead"><span class="ctitle">Top 10 Classes by Sales Volume</span></div><div class="ch ch260"><canvas id="chC"></canvas></div></div>
    <div class="cbox"><div class="chead"><span class="ctitle">Portfolio Health</span></div><div class="ch ch260"><canvas id="chH"></canvas></div></div>
  </div>
  <div class="tcard"><div class="thd"><span class="ttitle">Class Performance Summary</span><span class="b-slate">Top 12 Classes</span></div>
  <div class="tsc"><table><thead><tr><th>Class</th><th class="tr">Total Sales</th><th class="tr">Stock Value (AED)</th><th class="tr">SKUs</th><th class="tr">Zero Sale</th><th class="tr">Neg Stock</th><th class="tr">OOS</th><th>Status</th></tr></thead>
  <tbody>{cls_table}</tbody></table></div></div>
</div></div>

<div id="pg-top" class="page"><div class="pi">
  <div class="ptitle">Top Moving Items</div>
  <div class="pdesc">All items with recorded sales, ranked by total units sold. Barcode shown below each item name.</div>
  {fbar('top')}
  <div class="cbox" style="margin-bottom:20px">
    <div class="chead"><span class="ctitle">Top 10 — Sales Volume</span></div>
    <div class="rank-list" id="top-rank"></div>
  </div>
  <div class="tcard"><div class="thd"><span class="ttitle">All Items by Sales Volume</span><div class="tinfo"><span class="fcnt" id="top-cnt"></span><span class="b-brand">Sorted by Total Sales</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th>Brand</th><th class="tr">Total Sales</th><th class="tr">Last 3M</th><th class="tr">Stock</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost</th><th class="tr">Selling</th><th class="tr">Margin%</th></tr></thead>
  <tbody id="top-tb"></tbody></table></div></div>
</div></div>

<div id="pg-zero" class="page"><div class="pi">
  <div class="ptitle">Zero Sale Items &mdash; In Stock</div>
  <div class="pdesc">Items with no sales in the full data period. Blank last purchase date = purchased before 2025.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9888;</i><span class="at">Dead Stock &mdash; {fmtN(K['zero_count'],0)} Items</span></div><div class="ab">Zero sales with no demonstrated demand. Capital tied up with no return. Blank LP Date = purchased before 2025.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#8377;</i><span class="at">Capital at Risk &mdash; {fmtA(K['zero_sv'],0)}</span></div><div class="ab">Immediate action required: supplier return, clearance promotion, or write-off to recover capital.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Action Required</span></div><div class="ab">Classify: (1) Supplier return eligible, (2) Promotional clearance, (3) Write-off. All actioned within 30 days.</div></div>
  </div>
  {fbar('zero')}
  <div id="zero-strip"></div>
  <div class="tcard"><div class="thd"><span class="ttitle">Zero Sale Items with Stock</span><div class="tinfo"><span class="b-red">Sorted by Stock Value</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th>Group</th><th class="tr">Stock</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost</th><th class="tr">Selling</th><th>Last Purchase</th><th class="tr">LP Qty</th><th>Supplier</th></tr></thead>
  <tbody id="zero-tb"></tbody></table></div></div>
</div></div>

<div id="pg-neg" class="page"><div class="pi">
  <div class="ptitle">Negative Stock Items</div>
  <div class="pdesc">Items with negative on-hand quantities. Blank LP Date = purchased before 2025.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9888;</i><span class="at">{fmtN(K['neg_count'],0)} Items &mdash; Negative Quantity</span></div><div class="ab">Negative stock distorts financial reports and reorder calculations. Root cause identification required immediately.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Probable Causes</span></div><div class="ab">Unposted GRNs, manual sales errors, missing supplier invoices. Run pending PO report against each item.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Resolution Steps</span></div><div class="ab">1. Physical count top items. 2. Match GRNs. 3. Post verified receipts. 4. Raise adjustment journals.</div></div>
  </div>
  {fbar('neg')}
  <div id="neg-strip"></div>
  <div class="tcard"><div class="thd"><span class="ttitle">Negative Stock Items</span><div class="tinfo"><span class="b-red">Sorted by Negative Qty (Worst First)</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th class="tr">Stock Qty</th><th class="tr">Stock Value (AED)</th><th class="tr">Total Sales</th><th class="tr">Last 3M</th><th class="tr">Cost</th><th class="tr">Selling</th><th>Last Purchase</th><th class="tr">LP Qty</th></tr></thead>
  <tbody id="neg-tb"></tbody></table></div></div>
</div></div>

<div id="pg-oos" class="page"><div class="pi">
  <div class="ptitle">Out-of-Stock &mdash; Active Demand</div>
  <div class="pdesc">Items that sold in the last 3 months but currently have zero or negative stock. Active lost sales.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9873;</i><span class="at">{fmtN(K['oos_count'],0)} Items Actively Out of Stock</span></div><div class="ab">Confirmed recent demand but unavailable. Every day without stock is direct revenue loss.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#8679;</i><span class="at">Prioritise by Last 3M Velocity</span></div><div class="ab">Highest recent sales = greatest daily revenue risk. Generate emergency POs for top items immediately.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">&#10003;</i><span class="at">Immediate Reorder Candidates</span></div><div class="ab">Proven demand. Predictable replenishment ROI. This is the buying team's action list for today.</div></div>
  </div>
  {fbar('oos')}
  <div id="oos-strip"></div>
  <div class="tcard"><div class="thd"><span class="ttitle">OOS Items with Recent Sales</span><div class="tinfo"><span class="b-red">Sorted by Last 3M Sales</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th class="tr">Last 3M</th><th class="tr">{oos_h[0]}</th><th class="tr">{oos_h[1]}</th><th class="tr">{oos_h[2]}</th><th class="tr">Current Stock</th><th class="tr">Total Sales</th><th class="tr">Selling</th><th>Supplier</th></tr></thead>
  <tbody id="oos-tb"></tbody></table></div></div>
</div></div>

<div id="pg-risk" class="page"><div class="pi">
  <div class="ptitle">High-Value Slow Movers</div>
  <div class="pdesc">Stock value &gt; AED 200 and total sales &lt; 10 units. High capital, minimal turnover. Blank LP Date = purchased before 2025.</div>
  <div class="arow">
    <div class="abox amb"><div class="ah"><i class="ai">&#8595;</i><span class="at">{fmtA(K['risk_sv'],0)} &mdash; Low Turnover Capital</span></div><div class="ab">{fmtN(K['risk_count'],0)} items with fewer than 10 units sold — severely underperforming capital allocation.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">&#8856;</i><span class="at">Holding Cost Alert</span></div><div class="ab">At 15% annual holding cost, this capital costs approximately {fmtA(K['risk_sv']*0.15,0)}/year. Redeploy to fast-moving lines.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Exit Strategy</span></div><div class="ab">Items &gt;6 months, &lt;3 units: immediate markdown. Return-eligible: negotiate credit or exchange for faster lines.</div></div>
  </div>
  {fbar('risk')}
  <div id="risk-strip"></div>
  <div class="tcard"><div class="thd"><span class="ttitle">High-Value Slow Moving Items</span><div class="tinfo"><span class="b-amber">Sorted by Stock Value</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th class="tr">Stock Value (AED)</th><th class="tr">Stock</th><th class="tr">Total Sales</th><th class="tr">Last 3M</th><th class="tr">Cost</th><th class="tr">Selling</th><th class="tr">Margin%</th><th>Last Purchase</th><th class="tr">LP Qty</th></tr></thead>
  <tbody id="risk-tb"></tbody></table></div></div>
</div></div>

<div id="pg-uw" class="page"><div class="pi">
  <div class="ptitle">Unwanted Repurchases</div>
  <div class="pdesc">LP Qty &lt; Current Stock + zero sales = item had existing stock when re-purchased and has never sold. Blank LP Date = purchased before 2025.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#8856;</i><span class="at">{fmtA(K['uw_sv'],0)} &mdash; Confirmed Re-purchase Waste</span></div><div class="ab">{fmtN(K['uw_count'],0)} items had prior stock, were re-purchased, and still haven't sold. Buyers ordered without checking stock.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Supplier Leverage</span></div><div class="ab">Prior stock + repurchase + zero sales = strong evidence for supplier buy-back, credit notes, or extended terms.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Process Fix</span></div><div class="ab">Require buyers to check stock before any PO. System alert when ordering zero-sale items with existing stock.</div></div>
  </div>
  {fbar('uw')}
  <div id="uw-strip"></div>
  <div class="tcard"><div class="thd"><span class="ttitle">Re-purchased Items with Zero Sales</span><div class="tinfo"><span class="b-red">Sorted by Stock Value</span></div></div>
  <div class="tsc tmh"><table><thead><tr><th>#</th><th>Item Name / Barcode</th><th>Category</th><th>Class</th><th class="tr">Current Stock</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost</th><th class="tr">Selling</th><th class="tr">Last PO Qty</th><th>Last Purchase Date</th><th>Last Purchase Supplier</th></tr></thead>
  <tbody id="uw-tb"></tbody></table></div></div>
</div></div>

<div id="pg-ins" class="page"><div class="pi">
  <div class="ptitle">Strategic Analysis &amp; Action Plan</div>
  <div class="pdesc">Prioritised management recommendations from the full inventory dataset.</div>
  <div class="arow" style="grid-template-columns:1fr 1fr">
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">OOS Active Demand &mdash; {fmtN(K['oos_count'],0)} Items</span></div><div class="ab">Actively selling but unavailable. Generate emergency purchase orders for top items by last-3-month velocity immediately.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">Repurchase Waste &mdash; {fmtA(K['uw_sv'],0)}</span></div><div class="ab">{fmtN(K['uw_count'],0)} items re-purchased with existing stock and zero sales. Engage suppliers for buy-back or credit notes.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">Negative Stock &mdash; {fmtN(K['neg_count'],0)} Items</span></div><div class="ab">Compromises financial reporting. Operations team: 5-business-day resolution. Physical count + GRN reconciliation.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">{fmtA(K['risk_sv'],0)} Slow Mover Capital</span></div><div class="ab">Tiered markdown: 30% at 90 days, 50% at 180 days, supplier credit at 365 days.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">P3</i><span class="at">Zero Sale Dead Stock &mdash; {fmtA(K['zero_sv'],0)}</span></div><div class="ab">{fmtN(K['zero_count'],0)} items — classify as supplier return, clearance, or write-off within 30 days.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">&#10003;</i><span class="at">Food &amp; Essentials Strength</span></div><div class="ab">Grocery, Fresh Produce, and Beverages drive footfall and cross-category conversion. Zero OOS tolerance here.</div></div>
  </div>
  <div class="tcard"><div class="thd"><span class="ttitle">Prioritised Action Plan</span><span class="b-brand">Management Decision Matrix</span></div>
  <div class="tsc"><table><thead><tr><th>Priority</th><th>Issue</th><th>Scale</th><th>Financial Exposure</th><th>Action</th><th>Owner</th><th>Deadline</th></tr></thead>
  <tbody>{action_html}</tbody></table></div></div>
  <div class="arow" style="grid-template-columns:1fr 1fr 1fr;margin-top:8px">
    <div class="abox grn"><div class="ah"><i class="ai">&#8721;</i><span class="at">Supplier Leverage</span></div><div class="ab">Purchased-but-unsold evidence. Engage for buy-back, credit, or exchange. Estimated recovery: 40–60% of exposed capital.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8853;</i><span class="at">Clearance Strategy</span></div><div class="ab">Bundle dead stock with fast movers. End-cap placements. Dedicated clearance zone. Target: 25% moved in 60 days.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Governance</span></div><div class="ab">Monthly slow-mover reviews. 3 consecutive zero-sale months = mandatory review. 6 months + &lt;5 units = discontinuation flag.</div></div>
  </div>
</div></div>

<div class="footer">AL MADINA GROUP &nbsp;|&nbsp; Inventory Intelligence Report &nbsp;|&nbsp; {dlabel}</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script>{JS}</script>
</body></html>"""

    return html


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STREAMLIT PORTAL
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="AL MADINA GROUP — Inventory Dashboard",page_icon="📦",layout="wide",initial_sidebar_state="collapsed")
st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#F4F5F7}[data-testid="stHeader"]{background:#1B2B4B}
.block-container{padding-top:2rem;max-width:900px}h1{color:#1B2B4B!important}
</style>""",unsafe_allow_html=True)

c1,c2=st.columns([1,7])
with c1: st.markdown('<div style="background:linear-gradient(135deg,#3B82F6,#60A5FA);width:52px;height:52px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:18px;color:white;margin-top:4px">AM</div>',unsafe_allow_html=True)
with c2: st.markdown("# AL MADINA GROUP"); st.caption("Inventory Intelligence Dashboard Generator")
st.divider()

with st.expander("ℹ️  How to use",expanded=False):
    st.markdown("""
**Step 1** — Export your inventory report as `.xlsx`.  
**Required columns:** `Item Bar Code`, `Item Name`, `Stock`, `Total Sales`, `Cost`, `Selling`, `Stock Value`, `Category`, `Class`, `Group`, `Supplier`, `Margin%`, `LP Date`, `LP Qty`, `LP Supplier`, + monthly columns like `May, 2025` … `May, 2026`.  
**Blank LP Date** = item was last purchased before 2025 (shown as "Before 2025" in the dashboard).  
**Step 2** — Upload. **Step 3** — Generate. **Step 4** — Download HTML, open in any browser.  
**CSV Download** — Each page has a ↓ CSV button that exports the current filtered view.
    """)

st.subheader("Upload Inventory Excel File")
uploaded_file=st.file_uploader("Drop your Excel file",type=["xlsx","xls"],label_visibility="collapsed")

if uploaded_file:
    sz=len(uploaded_file.getvalue())/1024/1024
    c1,c2,c3=st.columns(3)
    c1.metric("File",uploaded_file.name[:30]); c2.metric("Size",f"{sz:.1f} MB"); c3.metric("Format",uploaded_file.name.split('.')[-1].upper())
    st.divider()

    if st.button("⚡  Generate Dashboard",type="primary",use_container_width=True):
        progress=st.progress(0,text="Starting…")
        try:
            progress.progress(10,text="Reading Excel file…")
            file_bytes=uploaded_file.getvalue()

            progress.progress(30,text="Analysing inventory data…")
            df_p,mc_p=read_excel_df(file_bytes)
            K_p,*_=build_kpis(df_p,mc_p)
            del df_p; gc.collect()

            st.markdown(f"""<div style="background:white;border:1px solid #E5E7EB;border-radius:8px;padding:14px 18px;font-size:13px">
            ✅ <b>{K_p['total_items']:,}</b> SKUs &nbsp;|&nbsp; 🔴 <b>{K_p['neg_count']:,}</b> negative &nbsp;|&nbsp;
            ⚠️ <b>{K_p['zero_count']:,}</b> zero sale &nbsp;|&nbsp; 🚨 <b>{K_p['oos_count']:,}</b> OOS &nbsp;|&nbsp;
            🛑 <b>{K_p['uw_count']:,}</b> unwanted repurchases</div>""",unsafe_allow_html=True)

            progress.progress(50,text="Building dashboard — this may take 30-60 seconds for large files…")
            html_content=generate_html(file_bytes,source_filename=uploaded_file.name)

            progress.progress(100,text="Done!")
            st.success("✅  Dashboard generated successfully!")

            out_name=uploaded_file.name.rsplit('.',1)[0]+'_dashboard.html'
            st.download_button("⬇️  Download Dashboard HTML",data=html_content.encode('utf-8'),
                               file_name=out_name,mime="text/html",use_container_width=True,type="primary")
            st.caption(f"Open `{out_name}` in any browser. No internet required after download.")

            st.divider()
            st.subheader("Dashboard Preview")
            st.caption("Full interactive dashboard with all pages, filters, charts and CSV downloads.")
            st.components.v1.html(html_content,height=820,scrolling=True)

            st.divider()
            st.subheader("Summary Metrics")
            r1=st.columns(4); r2=st.columns(4)
            r1[0].metric("Total SKUs",f"{K_p['total_items']:,}")
            r1[1].metric("Stock Value",f"AED {K_p['total_sv']:,.0f}")
            r1[2].metric("Total Units Sold",f"{K_p['total_ts']:,.0f}")
            r1[3].metric("Items In Stock",f"{K_p['in_stock']:,}")
            r2[0].metric("Negative Stock",f"{K_p['neg_count']:,}",delta=f"-AED {K_p['neg_sv']:,.0f}",delta_color="inverse")
            r2[1].metric("Zero Sale Stock",f"{K_p['zero_count']:,}",delta=f"-AED {K_p['zero_sv']:,.0f}",delta_color="inverse")
            r2[2].metric("High Risk Slow",f"{K_p['risk_count']:,}",delta=f"-AED {K_p['risk_sv']:,.0f}",delta_color="inverse")
            r2[3].metric("Unwanted Repurch.",f"{K_p['uw_count']:,}",delta=f"-AED {K_p['uw_sv']:,.0f}",delta_color="inverse")

        except Exception as e:
            progress.empty()
            st.error(f"❌  Error: {e}")
            with st.expander("Technical details"):
                import traceback; st.code(traceback.format_exc())
else:
    st.markdown("""<div style="background:white;border:2px dashed #D1D5DB;border-radius:12px;padding:48px 32px;text-align:center;margin:20px 0">
      <div style="font-size:48px;margin-bottom:12px">📊</div>
      <div style="font-size:16px;font-weight:600;color:#1B2B4B;margin-bottom:6px">Upload your inventory Excel file</div>
      <div style="font-size:13px;color:#6B7280">Supports .xlsx from any ERP or POS system</div>
    </div>""",unsafe_allow_html=True)

st.divider()
st.caption("AL MADINA GROUP Inventory Intelligence Portal  |  All data processed locally")
