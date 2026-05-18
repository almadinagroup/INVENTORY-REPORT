"""
AL MADINA GROUP — Inventory Intelligence Dashboard
Single-file Streamlit app. Upload Excel → generate HTML dashboard.

pip install streamlit pandas python-calamine
streamlit run app.py
"""

import streamlit as st
import json, gc, re, zipfile
from io import BytesIO
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 0 — PAGE CONFIG  (must be the very first Streamlit call)
# ═══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="AL MADINA GROUP — Inventory Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── PERMANENT DARK MODE + HIDE ALL STREAMLIT CHROME ─────────────────────────
st.markdown("""
<style>
/* ── Hide ALL Streamlit chrome ── */
[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"],
[data-testid="stStatusWidget"],[data-testid="stDeployButton"],
[data-testid="collapsedControl"],
#MainMenu, header, footer, .stDeployButton { display:none!important; }

/* ── Global dark background ── */
html,body,[data-testid="stAppViewContainer"],[data-testid="stMain"],
[data-testid="stMainBlockContainer"],.main,.block-container {
    background:#0D1117!important; color:#E6EDF3!important;
}
[data-testid="stSidebar"] { background:#161B22!important; }

/* ── Block container ── */
.block-container { padding-top:1.8rem; max-width:960px; }

/* ── Typography ── */
h1,h2,h3,h4 { color:#E6EDF3!important; }
p,li,label,.stCaption { color:#8B949E!important; }
hr { border-color:#30363D!important; }

/* ── Metrics ── */
[data-testid="stMetric"] {
    background:#161B22; border:1px solid #30363D;
    border-radius:8px; padding:12px 16px;
}
[data-testid="stMetricLabel"]  { color:#8B949E!important; font-size:12px!important; }
[data-testid="stMetricValue"]  { color:#E6EDF3!important; font-size:20px!important; font-weight:700!important; }
[data-testid="stMetricDelta"] svg { display:none; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background:#161B22!important; border:2px dashed #30363D!important;
    border-radius:12px!important; outline:none!important; box-shadow:none!important;
}
[data-testid="stFileUploader"]:hover  { border-color:#388BFD!important; }
[data-testid="stFileUploaderDropzone"] { background:transparent!important; border:none!important; }
[data-testid="stFileDropzoneInstructions"] { color:#8B949E!important; }

/* ── Primary button ── */
.stButton>button[kind="primary"] {
    background:linear-gradient(135deg,#1B2B4B,#2563EB)!important;
    color:white!important; border:none!important; border-radius:8px!important;
    font-weight:600!important; font-size:14px!important; padding:10px 20px!important;
}
.stButton>button[kind="primary"]:hover { opacity:.9!important; }
.stButton>button[kind="secondary"] {
    background:#21262D!important; color:#E6EDF3!important;
    border:1px solid #30363D!important; border-radius:8px!important;
}

/* ── Download button ── */
.stDownloadButton>button {
    background:#238636!important; color:white!important; border:none!important;
    border-radius:8px!important; font-weight:600!important; font-size:14px!important;
    padding:10px 20px!important; width:100%!important;
}
.stDownloadButton>button:hover { background:#2EA043!important; }

/* ── Expander — kill red focus ring ── */
[data-testid="stExpander"],[data-testid="stExpander"]>details,
[data-testid="stExpander"]>details>summary,[data-testid="stExpanderDetails"],
details,details>summary {
    background:#161B22!important; border:1px solid #30363D!important;
    border-radius:8px!important; outline:none!important; box-shadow:none!important;
}
details>summary { color:#8B949E!important; padding:10px 14px; cursor:pointer; }
details>summary:focus,details>summary:focus-visible,
details:focus,details:focus-within,
[data-testid="stExpander"]:focus,[data-testid="stExpander"]:focus-within {
    outline:none!important; box-shadow:none!important; border-color:#30363D!important;
}
*:focus { outline:none!important; }
*:focus-visible { outline:1px solid #388BFD!important; box-shadow:none!important; }

/* ── Progress bar ── */
[data-testid="stProgress"]>div>div { background:#2563EB!important; }
[data-testid="stProgress"] { background:#21262D!important; border-radius:4px; }

/* ── Alert boxes ── */
[data-testid="stAlert"] { border-radius:8px!important; }
.stSuccess { background:#0D2818!important; border:1px solid #238636!important; color:#3FB950!important; border-radius:8px!important; }
.stError   { background:#2D1117!important; border:1px solid #DA3633!important; color:#F85149!important; border-radius:8px!important; }
.stWarning { background:#2D2000!important; border:1px solid #9E6A03!important; color:#D29922!important; border-radius:8px!important; }
.stInfo    { background:#0D1F3C!important; border:1px solid #388BFD!important; color:#79C0FF!important; border-radius:8px!important; }

/* ── Code blocks ── */
.stCodeBlock { background:#161B22!important; border:1px solid #30363D!important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:#161B22; }
::-webkit-scrollbar-thumb { background:#30363D; border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:#484F58; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DATA PROCESSING  (identical logic to reference app__3_.py)
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

REQUIRED_COLUMNS = [
    'Item Bar Code','Item Name','Stock','Total Sales',
    'Cost','Selling','Stock Value','Category','Class',
    'Supplier','Margin%','LP Date','LP Qty',
]
OPTIONAL_COLUMNS = ['Group','Brand','LP Supplier']


def normalise_col(h):
    return COLUMN_ALIASES.get(str(h).strip().lower(), str(h).strip()) if h else ''


def detect_months(cols):
    pat = re.compile(r'(%s)[,.\s]+(\d{4})' % '|'.join(MONTH_NAMES), re.IGNORECASE)
    return [c for c in cols if pat.search(str(c))]


def month_sort_key(m):
    try: return datetime.strptime(re.sub(r'\s+', ', ', str(m).strip().replace(',,', ',')), '%b, %Y')
    except: return datetime.min


def check_missing_columns(norm_cols):
    present = set(norm_cols)
    return ([c for c in REQUIRED_COLUMNS if c not in present],
            [c for c in OPTIONAL_COLUMNS  if c not in present])


def read_excel_df(file_bytes):
    """Read Excel -> DataFrame.
    Tries calamine first (fast, pip install python-calamine),
    then openpyxl/xlrd if available, then falls back to pure stdlib zip reader.
    """
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
    raw_cols   = list(df_h.columns)
    norm_cols  = [normalise_col(c) for c in raw_cols]
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

    # ── LP Date: parse to datetime so .isna() reliably catches blanks ──
    if 'LP Date' in df.columns:
        import pandas as pd
        # First convert any Excel serial numbers (numeric floats) to NaT/dates
        def _parse_lpd(v):
            if v is None: return pd.NaT
            s = str(v).strip()
            if s in ('', 'nan', 'NaT', 'None', '0', 'NaTType') or 'NaT' in s:
                return pd.NaT
            # Try direct pandas parse
            try:
                return pd.to_datetime(s, dayfirst=True, errors='coerce')
            except Exception:
                return pd.NaT
        # If column is numeric (Excel serial), convert via pd origin
        if pd.api.types.is_numeric_dtype(df['LP Date']):
            df['LP Date'] = pd.to_datetime(
                df['LP Date'].replace(0, pd.NaT), unit='D',
                origin='1899-12-30', errors='coerce'
            )
        else:
            df['LP Date'] = df['LP Date'].apply(_parse_lpd)

    df.dropna(how='all', inplace=True)
    if 'Item Name' in df.columns:
        df = df[df['Item Name'].notna() &
                (df['Item Name'].astype(str).str.strip() != '') &
                (df['Item Name'].astype(str).str.lower() != 'nan')]
    df.reset_index(drop=True, inplace=True)
    return df, month_cols


def _read_excel_stdlib(file_bytes):
    """Pure-Python stdlib xlsx reader — zero extra dependencies."""
    import pandas as pd
    import xml.etree.ElementTree as ET

    NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    TR  = "{%s}row" % NS; TC = "{%s}c" % NS; TV = "{%s}v" % NS
    TT  = "{%s}t"   % NS; TSI = "{%s}si" % NS; TIS = "{%s}is" % NS

    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names  = {i.filename for i in z.infolist()}
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(TSI):
                shared.append("".join(t.text or "" for t in si.iter(TT)))
        sheet = next((n for n in sorted(names)
                      if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")), None)
        sb = z.read(sheet)

    def ci(ref):
        i = 0
        for ch in ref:
            if not ch.isalpha(): break
            i = i * 26 + ord(ch.upper()) - 64
        return i - 1

    rows, cur = [], {}
    for ev, el in ET.iterparse(BytesIO(sb), events=("start", "end")):
        if ev == "start" and el.tag == TR: cur = {}
        elif ev == "end" and el.tag == TC:
            ref = el.get("r","A1"); c = ci(ref); t = el.get("t","n")
            v = el.find(TV); is_ = el.find(TIS)
            if is_ is not None:
                te = is_.find(TT); val = te.text if te is not None else ""
            elif v is not None:
                rv = v.text or ""
                if   t == "s": val = shared[int(rv)] if shared else rv
                elif t == "b": val = bool(int(rv))
                else:
                    try: val = float(rv) if "." in rv else int(rv)
                    except: val = rv
            else: val = None
            cur[c] = val; el.clear()
        elif ev == "end" and el.tag == TR:
            if cur: rows.append([cur.get(i) for i in range(max(cur)+1)])
            el.clear()

    if not rows: raise RuntimeError("Excel file is empty.")
    raw_h  = [str(v) if v is not None else "" for v in rows[0]]
    norm_h = [normalise_col(h) for h in raw_h]
    month_cols = detect_months(norm_h)
    data = []
    for row in rows[1:]:
        while len(row) < len(norm_h): row.append(None)
        if all(v is None or str(v).strip() == "" for v in row): continue
        data.append({norm_h[j]: row[j] for j in range(len(norm_h)) if norm_h[j]})

    df = pd.DataFrame(data)
    num_cols = ['Cost','Selling','Stock','Stock Value','Margin%','LP Qty','Total Sales'] + month_cols
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('float32')
    for col in ['Category','Class','Group','Brand']:
        if col in df.columns: df[col] = df[col].fillna('').astype('category')
    for col in ['Item Name','Item Bar Code','Supplier','LP Supplier']:
        if col in df.columns: df[col] = df[col].fillna('').astype(str)
    # ── LP Date: parse to datetime so isna() catches all blank variants ──
    if 'LP Date' in df.columns:
        def _parse_lpd2(v):
            if v is None: return pd.NaT
            s = str(v).strip()
            if s in ('', 'nan', 'NaT', 'None', '0', 'NaTType') or 'NaT' in s:
                return pd.NaT
            try:
                f = float(s)
                if f == 0: return pd.NaT
                return pd.Timestamp('1899-12-30') + pd.Timedelta(days=f)
            except ValueError:
                pass
            return pd.to_datetime(s, dayfirst=True, errors='coerce')
        df['LP Date'] = df['LP Date'].apply(_parse_lpd2)
    df.dropna(how='all', inplace=True)
    if 'Item Name' in df.columns:
        df = df[df['Item Name'].notna() &
                (df['Item Name'].astype(str).str.strip() != '') &
                (df['Item Name'].astype(str).str.lower() != 'nan')]
    df.reset_index(drop=True, inplace=True)
    return df, month_cols


def build_kpis(df, month_cols):
    last3 = sorted(month_cols, key=month_sort_key)[-3:] if len(month_cols) >= 3 else month_cols
    l3    = df[[c for c in last3 if c in df.columns]].sum(axis=1).astype('float32') if last3 else 0

    stock = df['Stock']       if 'Stock'       in df.columns else 0
    ts    = df['Total Sales'] if 'Total Sales' in df.columns else 0
    sv    = df['Stock Value'] if 'Stock Value' in df.columns else 0
    lpq   = df['LP Qty']      if 'LP Qty'      in df.columns else 0
    lpd   = df['LP Date']     if 'LP Date'     in df.columns else None

    mask_zero = (stock > 0) & (ts == 0)
    mask_neg  = stock < 0
    mask_oos  = (l3 > 0) & (stock <= 0)
    mask_risk = (sv > 200) & (ts < 10) & (stock > 0)
    # mask_uw: zero-sale items that were re-purchased (LP Date present, LP Qty < current stock)
    mask_uw   = mask_zero & (lpd.notna() if lpd is not None else False) & (lpq < stock)
    # mask_pre: items with stock but blank/missing LP Date (purchased pre-2025 / no PO record)
    # After LP Date normalisation in read_excel_df, blanks are NaT — isna() catches them.
    # Extra safety: also catch string blanks in case dtype is object.
    if lpd is not None:
        import pandas as _pd
        if _pd.api.types.is_datetime64_any_dtype(lpd):
            _lpd_blank = lpd.isna()
        else:
            _lpd_blank = lpd.isna() | lpd.astype(str).str.strip().isin(['','nan','NaT','None','0','NaTType'])
        mask_pre = (stock > 0) & _lpd_blank
    else:
        mask_pre = (stock > 0) & False

    return {
        'total_items':    len(df),
        'in_stock':       int((stock > 0).sum()),
        'total_sv':       round(float(sv.sum()), 0),
        'total_ts':       round(float(ts.sum()), 0),
        'neg_count':      int(mask_neg.sum()),
        'neg_sv':         round(float(sv[mask_neg].abs().sum()), 0),
        'zero_count':     int(mask_zero.sum()),
        'zero_sv':        round(float(sv[mask_zero].sum()), 0),
        'oos_count':      int(mask_oos.sum()),
        'risk_count':     int(mask_risk.sum()),
        'risk_sv':        round(float(sv[mask_risk].sum()), 0),
        'uw_count':       int(mask_uw.sum()),
        'uw_sv':          round(float(sv[mask_uw].sum()), 0),
        'pre_count':      int(mask_pre.sum()),
        'pre_sv':         round(float(sv[mask_pre].sum()), 0),
        'pre_zero_count': int((mask_pre & (ts <= 0)).sum()),
        'pre_zero_sv':    round(float(sv[mask_pre & (ts <= 0)].sum()), 0),
        'pre_sales_count':int((mask_pre & (ts > 0)).sum()),
        'pre_sales_sv':   round(float(sv[mask_pre & (ts > 0)].sum()), 0),
    }, last3, mask_zero, mask_neg, mask_oos, mask_risk, mask_uw, mask_pre


def fmtN(n, d=1):
    try:
        n = float(n)
        if abs(n) >= 1_000_000: return f"{n/1_000_000:.{d}f}M"
        if abs(n) >= 1_000:     return f"{n/1_000:.{d}f}K"
        return f"{n:,.{d}f}"
    except: return "—"

def fmtA(n, d=0): return f"AED {fmtN(n,d)}"
def esc(s): return str(s).replace('&','&amp;').replace('"','&quot;').replace('<','&lt;').replace('>','&gt;')


def lp_label(lpd_val, lpq_val=0):
    """Return LP date string. Blank/NaT -> empty string (dash in table)."""
    if lpd_val is None: return ''
    s = str(lpd_val).strip()
    if s in ('', 'nan', 'NaT', 'None', '0', 'NaTType'): return ''
    if 'NaT' in s: return ''
    return s[:20]


# Report reference date — use today dynamically
def age_days(lpd_val):
    """Return days since last purchase. None if blank/unparseable."""
    if lpd_val is None: return None
    s = str(lpd_val).strip()
    if s in ('','nan','NaT','None','0','NaTType') or 'NaT' in s: return None
    today = datetime.now()
    for fmt in ('%d-%b-%Y','%Y-%m-%d','%d/%m/%Y','%m/%d/%Y','%d-%m-%Y','%b %d, %Y','%d %b %Y'):
        try: return (today - datetime.strptime(s, fmt)).days
        except: pass
    return None


def age_bucket(days):
    """Return ageing band label from days integer."""
    if days is None: return 'Pre-2025'
    if days <= 30:  return '0–30 days'
    if days <= 90:  return '31–90 days'
    if days <= 180: return '91–180 days'
    if days <= 360: return '181–360 days'
    return '360+ days'


def age_bucket_order(label):
    order = {'0–30 days':0,'31–90 days':1,'91–180 days':2,'181–360 days':3,'360+ days':4,'Pre-2025':5}
    return order.get(label, 9)


def df_to_compact_json(df, fields, last3=None):
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
                lpd_str = lp_label(val, 0)
                row[key] = lpd_str
                # also compute ageing
                d = age_days(val)
                row['age']  = d if d is not None else -1   # -1 = pre-2025/unknown
                row['ageb'] = age_bucket(d)                # human label
            else:
                row[key] = str(val)[:50] if val is not None else ''
        if last3:
            row['m0'] = round(float(r.get(last3[2], 0) or 0), 1) if len(last3) > 2 else 0
            row['m1'] = round(float(r.get(last3[1], 0) or 0), 1) if len(last3) > 1 else 0
            row['m2'] = round(float(r.get(last3[0], 0) or 0), 1) if last3 else 0
        # ensure ageb exists even if no lpd field was in fields
        if 'ageb' not in row: row['ageb'] = ''
        rows.append(row)
    return json.dumps(rows, separators=(',',':'))


def cls_summary(df, mask_zero, mask_neg, mask_oos):
    if 'Class' not in df.columns: return ''
    ts_col = 'Total Sales' if 'Total Sales' in df.columns else None
    sv_col = 'Stock Value'  if 'Stock Value'  in df.columns else None

    agg = df.groupby('Class', observed=True).agg(
        sales=('Total Sales','sum') if ts_col else ('Class','count'),
        sv   =('Stock Value', 'sum') if sv_col else ('Class','count'),
        items=('Item Name',   'count'),
    ).reset_index().sort_values('sales', ascending=False).head(12)

    zs  = df[mask_zero].groupby('Class', observed=True).size()
    neg = df[mask_neg ].groupby('Class', observed=True).size()
    oos = df[mask_oos ].groupby('Class', observed=True).size()

    rows = []
    for _, r in agg.iterrows():
        c = str(r['Class']); it = int(r['items'])
        z = int(zs.get(c,0)); ng = int(neg.get(c,0)); os_ = int(oos.get(c,0))
        tag = ('<span class="b-red">High Dead Stock</span>'  if z  > it*0.1 else
               '<span class="b-amber">OOS Risk</span>'       if os_ > it*0.1 else
               '<span class="b-green">Healthy</span>')
        rows.append(
            f'<tr><td><span class="b-brand">{esc(c)}</span></td>'
            f'<td class="tr">{fmtN(r["sales"],0)}</td>'
            f'<td class="tr">{fmtN(r["sv"],0)}</td>'
            f'<td class="tr">{fmtN(it,0)}</td>'
            f'<td class="tr {"c-amber" if z>50 else ""}">{fmtN(z,0)}</td>'
            f'<td class="tr {"c-red" if ng>20 else ""}">{fmtN(ng,0)}</td>'
            f'<td class="tr {"c-red" if os_>50 else ""}">{fmtN(os_,0)}</td>'
            f'<td>{tag}</td></tr>'
        )
    return ''.join(rows)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — HTML GENERATOR (rewritten — robust, searchable, sortable)
# ═══════════════════════════════════════════════════════════════════════════════

def generate_html(file_bytes, source_filename=''):
    import pandas as pd

    df, month_cols = read_excel_df(file_bytes)
    months_sorted  = sorted(month_cols, key=month_sort_key)
    last3          = months_sorted[-3:] if len(months_sorted) >= 3 else months_sorted
    last3_labels   = [months_sorted[-1] if months_sorted else '',
                      months_sorted[-2] if len(months_sorted) > 1 else '',
                      months_sorted[-3] if len(months_sorted) > 2 else '']

    df['L3']  = df[[c for c in last3 if c in df.columns]].sum(axis=1).astype('float32') if last3 else 0.0
    df['_bc'] = df['Item Bar Code'].astype(str) if 'Item Bar Code' in df.columns else ''

    K, last3, mask_zero, mask_neg, mask_oos, mask_risk, mask_uw, mask_pre = build_kpis(df, month_cols)

    sv_col = 'Stock Value' if 'Stock Value' in df.columns else 'Cost'
    ts_col = 'Total Sales' if 'Total Sales' in df.columns else 'L3'

    df_top  = df[df[ts_col] > 0].sort_values(ts_col, ascending=False)
    df_zero = df[mask_zero].sort_values(sv_col, ascending=False)
    df_neg  = df[mask_neg ].sort_values('Stock')
    df_oos  = df[mask_oos ].sort_values('L3',    ascending=False)
    df_risk = df[mask_risk].sort_values(sv_col,  ascending=False)
    df_uw   = df[mask_uw  ].sort_values(sv_col,  ascending=False)
    df_pre  = df[mask_pre ].sort_values(sv_col,  ascending=False)

    ml  = [m.replace(', 20', "'") for m in months_sorted]
    mv  = [round(float(df[m].sum()), 1) if m in df.columns else 0 for m in months_sorted]
    if 'Class' in df.columns and ts_col in df.columns:
        tc = df.groupby('Class', observed=True)[ts_col].sum().sort_values(ascending=False).head(10)
        cl = [c[:16]+'…' if len(c) > 16 else c for c in tc.index.astype(str).tolist()]
        cv = [round(float(v), 1) for v in tc.values]
    else: cl, cv = [], []
    hv = [max(0, K['in_stock']-K['zero_count']-K['risk_count']),
          K['zero_count'], K['neg_count'], K['oos_count'], K['risk_count']]

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
    PRE_FIELDS  = [('bc','_bc','bc'),('n','Item Name','str'),('cat','Category','str'),('cls','Class','str'),
                   ('grp','Group','str'),('brand','Brand','str'),('sup','Supplier','str'),
                   ('stock','Stock','num'),('sv','Stock Value','num'),('ts','Total Sales','num'),
                   ('cost','Cost','num'),('sell','Selling','num'),('mg','Margin%','num'),('l3','L3','num')]

    top_json  = df_to_compact_json(df_top,  TOP_FIELDS)
    zero_json = df_to_compact_json(df_zero, ZERO_FIELDS)
    neg_json  = df_to_compact_json(df_neg,  NEG_FIELDS)
    oos_json  = df_to_compact_json(df_oos,  OOS_FIELDS, last3=last3)
    risk_json = df_to_compact_json(df_risk, RISK_FIELDS)
    uw_json   = df_to_compact_json(df_uw,   UW_FIELDS)
    pre_json  = df_to_compact_json(df_pre,  PRE_FIELDS)

    cats   = sorted(df['Category'].dropna().unique().astype(str).tolist()) if 'Category' in df.columns else []
    catcls = {}
    if 'Category' in df.columns and 'Class' in df.columns:
        for cat, cdf in df.groupby('Category', observed=True):
            catcls[str(cat)] = sorted(cdf['Class'].dropna().unique().astype(str).tolist())

    cls_table = cls_summary(df, mask_zero, mask_neg, mask_oos)

    # Top suppliers & top categories (NEW — for Insights page)
    top_sup = []
    if 'Supplier' in df.columns and ts_col in df.columns:
        sg = df.groupby('Supplier', observed=True).agg(
            sales=(ts_col, 'sum'), sv=(sv_col, 'sum'), items=('Item Name', 'count')
        ).reset_index()
        sg = sg[sg['Supplier'].astype(str).str.strip() != ''].sort_values('sales', ascending=False).head(10)
        for _, r in sg.iterrows():
            top_sup.append({
                'name': str(r['Supplier'])[:40], 'sales': round(float(r['sales']), 0),
                'sv': round(float(r['sv']), 0), 'items': int(r['items'])
            })

    del df, df_top, df_zero, df_neg, df_oos, df_risk, df_uw, df_pre
    gc.collect()

    kpi_defs = [
        ('Total SKUs',            fmtN(K['total_items'],0),  'Products in portfolio',                   'brand','kv-brand'),
        ('Items In Stock',        fmtN(K['in_stock'],0),     fmtA(K['total_sv'],0)+' stock value',      'green','kv-green'),
        ('Total Units Sold',      fmtN(K['total_ts'],0),     'Cumulative period total',                 'blue', 'kv-blue'),
        ('Negative Stock Items',  fmtN(K['neg_count'],0),    'Audit required urgently',                 'red',  'kv-red'),
        ('Zero Sale — In Stock',  fmtN(K['zero_count'],0),   fmtA(K['zero_sv'],0)+' capital at risk',   'red',  'kv-red'),
        ('OOS Active Demand',     fmtN(K['oos_count'],0),    'Selling but unavailable',                 'red',  'kv-red'),
        ('High Risk Slow Movers', fmtN(K['risk_count'],0),   fmtA(K['risk_sv'],0)+' exposure',          'amber','kv-amber'),
        ('Unwanted Repurchases',  fmtN(K['uw_count'],0),     fmtA(K['uw_sv'],0)+' re-bought, unsold',   'amber','kv-amber'),
        ('Pre-2025 Stock',        fmtN(K['pre_count'],0),    fmtA(K['pre_sv'],0)+' blank purchase date','amber','kv-amber'),
    ]
    kpi_html = ''.join(
        f'<div class="kcard kc-{b}"><div class="klbl">{l}</div>'
        f'<div class="kval {v}">{va}</div><div class="ksub">{s}</div></div>'
        for l, va, s, b, v in kpi_defs)

    action_rows = [
        ('b-red',  'P1 — Critical','OOS Active Demand',    f"{K['oos_count']:,} items", 'c-red',  'Daily revenue loss',     'Emergency POs — top items by last-3M velocity',       'Buying Team',   'Today'),
        ('b-red',  'P1 — Critical','Unwanted Repurchases', f"{K['uw_count']:,} items",  'c-red',  fmtA(K['uw_sv'],0),       'Supplier return + credit note negotiations',          'Category Mgmt', 'This Week'),
        ('b-amber','P2 — High',    'Negative Stock',       f"{K['neg_count']:,} items", 'c-amber','Reporting inaccuracy',   'Physical count + GRN audit + corrections',            'Operations',    '5 Business Days'),
        ('b-amber','P2 — High',    'Slow Mover Capital',   f"{K['risk_count']:,} items",'c-amber',fmtA(K['risk_sv'],0),     'Tiered markdown + bundle offers',                     'Merchandising', '14 Days'),
        ('b-blue', 'P3 — Medium',  'Zero Sale Dead Stock', f"{K['zero_count']:,} items",'c-blue', fmtA(K['zero_sv'],0),     'Range review — discontinue or clearance',             'Category Mgmt', '30 Days'),
        ('b-amber','P3 — Medium',  'Pre-2025 Dead Stock',  f"{K['pre_zero_count']:,} items",'c-amber',fmtA(K['pre_zero_sv'],0),'Write-off or targeted clearance — oldest dead stock','Category Mgmt','60 Days'),
        ('b-slate','P4 — Standard','New Listing Policy',   'Process change',            '',       'Future prevention',      'Mandatory 90-day sell-through KPI on new listings',   'Commercial Dir.','Next Quarter'),
    ]
    action_html = ''.join(
        f'<tr><td><span class="{b}">{p}</span></td><td>{i}</td><td>{sc}</td>'
        f'<td class="{ec}">{ex}</td><td>{ac}</td><td>{ow}</td><td>{dl}</td></tr>'
        for b, p, i, sc, ec, ex, ac, ow, dl in action_rows)

    rdate    = datetime.now().strftime('%d %b %Y')
    dlabel   = f'Source: {esc(source_filename)} &nbsp;|&nbsp; Generated: {rdate}' if source_filename else f'Generated: {rdate}'
    oos_h    = [esc(last3_labels[0]), esc(last3_labels[1]), esc(last3_labels[2])]
    cat_opts = ''.join(f'<option value="{esc(c)}">{esc(c)}</option>' for c in ['All Categories'] + cats)

    # ── CSS — single triple-quoted string, no substitution needed ────────────
    CSS = r"""
:root{
  --bg:#0D1117;--surf:#161B22;--surf2:#21262D;--surf3:#2A3038;--bdr:#30363D;--bdr2:#484F58;
  --txt:#E6EDF3;--txt2:#C9D1D9;--mut:#8B949E;--accent:#388BFD;
  --brand:#1B2B4B;--brand-l:#2D4A7A;--brand-t:#0D1F3C;
  --red:#F85149;--red-t:#2D1117;--red-b:#6E3B3B;
  --amb:#D29922;--amb-t:#2D2000;--amb-b:#7B5E08;
  --grn:#3FB950;--grn-t:#0D2818;--grn-b:#1A7431;
  --blu:#79C0FF;--blu-t:#0D1F3C;--blu-b:#1B4B7A;
  --pur:#A371F7;--pur-t:#1F1230;--pur-b:#4B2E7A;
  --slt:#8B949E;--r:8px;--sh:0 1px 4px rgba(0,0,0,.4);
  --sh-md:0 4px 12px rgba(0,0,0,.5);
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%}
body{background:var(--bg);color:var(--txt);font-family:"Inter",-apple-system,BlinkMacSystemFont,sans-serif;font-size:13.5px;line-height:1.55;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
::-webkit-scrollbar{width:8px;height:8px}::-webkit-scrollbar-track{background:var(--surf)}
::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:4px}::-webkit-scrollbar-thumb:hover{background:#5d6671}

/* HEADER */
.hdr{background:linear-gradient(180deg,#161B22 0%,#13181F 100%);color:var(--txt);padding:0 32px;display:flex;align-items:center;justify-content:space-between;height:64px;position:sticky;top:0;z-index:200;border-bottom:1px solid #30363D;backdrop-filter:blur(8px)}
.hdr-l{display:flex;align-items:center;gap:18px}.logo{display:flex;align-items:center;gap:11px}
.lm{width:38px;height:38px;background:linear-gradient(135deg,#1B2B4B 0%,#2563EB 100%);border-radius:9px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px;color:#fff;box-shadow:0 2px 8px rgba(37,99,235,.3);letter-spacing:-.5px}
.lt{font-weight:700;font-size:15px;letter-spacing:-.2px;color:var(--txt);line-height:1.15}
.ls{font-size:10.5px;color:var(--mut);line-height:1.15;margin-top:1px}
.hsep{width:1px;height:28px;background:rgba(255,255,255,.1)}
.hlbl{font-size:12px;color:var(--mut);font-weight:500}
.hr-r{display:flex;align-items:center;gap:14px}
.hmeta{font-size:11px;color:var(--mut);font-family:"IBM Plex Mono",ui-monospace,monospace;text-align:right;line-height:1.4}
.hmeta b{color:var(--txt);font-weight:600;font-family:"Inter",sans-serif;font-size:11.5px}

/* GLOBAL SEARCH */
.gs-wrap{position:relative;flex:1;max-width:480px;margin:0 18px}
.gs-input{width:100%;background:rgba(255,255,255,.04);border:1px solid #30363D;border-radius:7px;color:var(--txt);font-family:"Inter",sans-serif;font-size:13px;padding:8px 36px 8px 36px;outline:none;transition:border-color .15s,background .15s}
.gs-input::placeholder{color:#6e7681}
.gs-input:focus{border-color:var(--accent);background:rgba(56,139,253,.08)}
.gs-ico{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--mut);pointer-events:none;font-size:14px}
.gs-kbd{position:absolute;right:10px;top:50%;transform:translateY(-50%);font-size:10px;color:var(--mut);background:rgba(255,255,255,.05);border:1px solid #30363D;border-radius:3px;padding:1px 6px;font-family:"IBM Plex Mono",monospace}
.gs-results{position:absolute;left:0;right:0;top:calc(100% + 6px);background:#161B22;border:1px solid #30363D;border-radius:8px;box-shadow:var(--sh-md);max-height:420px;overflow-y:auto;display:none;z-index:300}
.gs-results.open{display:block}
.gs-result{padding:10px 14px;border-bottom:1px solid #21262D;cursor:pointer;display:flex;align-items:center;gap:12px;transition:background .1s}
.gs-result:last-child{border-bottom:none}
.gs-result:hover,.gs-result.sel{background:#21262D}
.gs-result .gsn{font-size:13px;font-weight:500;color:var(--txt);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.gs-result .gsmeta{font-family:"IBM Plex Mono",monospace;font-size:11px;color:var(--mut);white-space:nowrap}
.gs-result .gsbadge{font-size:10px;font-weight:600;padding:1px 7px;border-radius:3px;white-space:nowrap}
.gs-empty{padding:20px;text-align:center;color:var(--mut);font-size:12.5px}
.gs-section{padding:7px 14px;font-size:10px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em;background:#13181F;border-bottom:1px solid #21262D}

/* NAV */
.nav{background:#161B22;border-bottom:1px solid #30363D;display:flex;padding:0 32px;overflow-x:auto;gap:2px;position:sticky;top:64px;z-index:150}
.nav::-webkit-scrollbar{height:0}
.nbtn{background:none;border:none;color:rgba(255,255,255,.42);font-family:"Inter",sans-serif;font-size:12.5px;font-weight:500;padding:11px 14px;cursor:pointer;white-space:nowrap;border-bottom:2px solid transparent;transition:color .15s,border-color .15s;display:flex;align-items:center;gap:7px}
.nbtn:hover{color:rgba(255,255,255,.75)}
.nbtn.active{color:var(--txt);border-bottom-color:var(--accent)}
.npill{display:inline-flex;align-items:center;justify-content:center;background:rgba(255,255,255,.08);border-radius:10px;font-size:10px;font-weight:600;min-width:20px;height:17px;padding:0 6px}
.npill.red{background:rgba(248,81,73,.22);color:#F85149}
.npill.amb{background:rgba(210,153,34,.22);color:#D29922}
.npill.blu{background:rgba(56,139,253,.22);color:#79C0FF}

/* PAGE */
.page{display:none;animation:fade .15s ease}
.page.active{display:block}
@keyframes fade{from{opacity:0;transform:translateY(2px)}to{opacity:1;transform:none}}
.pi{max-width:1480px;margin:0 auto;padding:26px 32px 48px}
.ptitle{font-size:20px;font-weight:700;letter-spacing:-.4px;margin-bottom:4px;color:var(--txt)}
.pdesc{font-size:13px;color:var(--mut);margin-bottom:22px;max-width:780px;line-height:1.55}

/* FILTER BAR */
.fbar{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:13px 18px;display:flex;align-items:center;gap:12px;margin-bottom:18px;flex-wrap:wrap;box-shadow:var(--sh)}
.flbl{font-size:10.5px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.07em}
.fsep{width:1px;height:22px;background:var(--bdr)}
.fg{display:flex;align-items:center;gap:8px}
.fg label{font-size:11.5px;color:var(--mut);font-weight:500}
select.fsel,input.fin{background:var(--surf2);border:1px solid var(--bdr);border-radius:6px;color:var(--txt2);font-family:"Inter",sans-serif;font-size:12.5px;font-weight:500;padding:7px 11px;cursor:pointer;outline:none;min-width:170px;transition:border-color .15s}
input.fin{min-width:200px;cursor:text}
select.fsel:focus,input.fin:focus{border-color:var(--accent)}
.freset{background:none;border:1px solid var(--bdr);border-radius:6px;color:var(--mut);font-family:"Inter",sans-serif;font-size:11.5px;font-weight:500;padding:7px 13px;cursor:pointer;transition:all .15s}
.freset:hover{background:var(--surf2);color:var(--txt2)}
.fcnt{font-size:11.5px;color:var(--mut);font-family:"IBM Plex Mono",monospace;font-weight:500}
.fcnt b{color:var(--txt);font-weight:700}
.csv-btn{background:#238636;color:#fff;border:none;border-radius:6px;font-family:"Inter",sans-serif;font-size:11.5px;font-weight:600;padding:7px 14px;cursor:pointer;margin-left:auto;display:flex;align-items:center;gap:6px;transition:background .15s}
.csv-btn:hover{background:#2EA043}

/* KPI CARDS */
.kgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:13px;margin-bottom:26px}
.kcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:17px 18px;box-shadow:var(--sh);position:relative;overflow:hidden;transition:transform .15s,border-color .15s}
.kcard:hover{transform:translateY(-1px);border-color:var(--bdr2)}
.kcard::after{content:"";position:absolute;top:0;left:0;right:0;height:3px;border-radius:8px 8px 0 0}
.kc-brand::after{background:linear-gradient(90deg,#388BFD,#2563EB)}
.kc-green::after{background:linear-gradient(90deg,#3FB950,#1A7431)}
.kc-red::after{background:linear-gradient(90deg,#F85149,#DA3633)}
.kc-amber::after{background:linear-gradient(90deg,#D29922,#9E6A03)}
.kc-blue::after{background:linear-gradient(90deg,#79C0FF,#388BFD)}
.klbl{font-size:10.5px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.07em;margin-bottom:9px}
.kval{font-size:25px;font-weight:700;letter-spacing:-.6px;line-height:1;margin-bottom:5px}
.kv-brand{color:#79C0FF}.kv-green{color:var(--grn)}.kv-red{color:var(--red)}.kv-amber{color:var(--amb)}.kv-blue{color:var(--blu)}
.ksub{font-size:11.5px;color:var(--mut)}

/* CHARTS */
.crow{display:grid;gap:18px;margin-bottom:22px}
.c1{grid-template-columns:1fr}
.c21{grid-template-columns:2fr 1fr}
.c11{grid-template-columns:1fr 1fr}
.cbox{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:20px;box-shadow:var(--sh)}
.chead{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:8px}
.ctitle{font-size:13.5px;font-weight:600;color:var(--txt)}
.cmeta{font-size:11px;color:var(--mut)}
.ch{position:relative}.ch260{height:280px}.ch320{height:340px}

/* TABLE */
.tcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);box-shadow:var(--sh);overflow:hidden;margin-bottom:22px}
.thd{padding:14px 18px;border-bottom:1px solid var(--bdr);display:flex;align-items:center;justify-content:space-between;background:var(--surf2);flex-wrap:wrap;gap:8px}
.ttitle{font-size:13.5px;font-weight:600;color:var(--txt)}
.tinfo{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.tsc{overflow-x:auto}.tmh{max-height:600px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead tr{background:var(--surf2);position:sticky;top:0;z-index:2}
th{text-align:left;padding:10px 12px;font-size:10.5px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em;white-space:nowrap;border-bottom:1px solid var(--bdr);user-select:none}
th.srt{cursor:pointer;transition:color .1s}
th.srt:hover{color:var(--txt2)}
th.srt::after{content:" ⇅";opacity:.3;font-size:9px}
th.srt.asc::after{content:" ▲";opacity:1;color:var(--accent)}
th.srt.desc::after{content:" ▼";opacity:1;color:var(--accent)}
tbody tr{border-bottom:1px solid var(--bdr);transition:background .08s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#21262D}
tbody tr.hl{background:rgba(56,139,253,.08);box-shadow:inset 3px 0 0 var(--accent)}
td{padding:9px 12px;color:var(--txt2);vertical-align:middle}
.tn{font-weight:500;color:var(--txt);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}
.bc{font-family:"IBM Plex Mono",monospace;font-size:10.5px;color:var(--mut);margin-top:1px}
.mono{font-family:"IBM Plex Mono",monospace;font-size:12px}
.tr{text-align:right;font-family:"IBM Plex Mono",monospace;font-size:12px;white-space:nowrap}
.c-red{color:var(--red)}.c-amber{color:var(--amb)}.c-green{color:var(--grn)}.c-blue{color:var(--blu)}
.muted{color:var(--mut)}.fw6{font-weight:600}
.bold-green{color:var(--grn);font-weight:700}
.sup{max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}

/* BADGES */
.b-red,.b-amber,.b-green,.b-blue,.b-slate,.b-brand,.b-pur{display:inline-flex;font-size:10.5px;font-weight:600;padding:2px 8px;border-radius:4px;white-space:nowrap;letter-spacing:.01em}
.b-red{background:var(--red-t);color:var(--red);border:1px solid var(--red-b)}
.b-amber{background:var(--amb-t);color:var(--amb);border:1px solid var(--amb-b)}
.b-green{background:var(--grn-t);color:var(--grn);border:1px solid var(--grn-b)}
.b-blue{background:var(--blu-t);color:var(--blu);border:1px solid var(--blu-b)}
.b-slate{background:var(--surf2);color:var(--slt);border:1px solid var(--bdr)}
.b-brand{background:var(--brand-t);color:#79C0FF;border:1px solid #1B4B7A}
.b-pur{background:var(--pur-t);color:var(--pur);border:1px solid var(--pur-b)}

/* ALERT BOXES */
.arow{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:13px;margin-bottom:22px}
.abox{border-radius:var(--r);padding:15px 17px;border:1px solid;backdrop-filter:blur(4px)}
.abox.red{background:var(--red-t);border-color:var(--red-b)}
.abox.amb{background:var(--amb-t);border-color:var(--amb-b)}
.abox.grn{background:var(--grn-t);border-color:var(--grn-b)}
.abox.blu{background:var(--blu-t);border-color:var(--blu-b)}
.ah{display:flex;align-items:flex-start;gap:9px;margin-bottom:6px}
.ai{font-size:14px;flex-shrink:0;font-style:normal;font-weight:700;min-width:18px}
.at{font-size:12.5px;font-weight:600;color:var(--txt);line-height:1.35}
.ab{font-size:12px;color:var(--txt2);line-height:1.55;padding-left:27px}

/* STAT STRIP */
.strip{background:linear-gradient(180deg,#161B22 0%,#13181F 100%);border:1px solid #30363D;color:var(--txt);border-radius:var(--r);padding:14px 22px;display:flex;align-items:center;gap:26px;margin-bottom:22px;flex-wrap:wrap}
.si{display:flex;flex-direction:column;gap:2px}
.sl{font-size:10px;color:var(--mut);text-transform:uppercase;letter-spacing:.07em;font-weight:600}
.sv{font-size:17px;font-weight:700;font-family:"IBM Plex Mono",monospace;letter-spacing:-.5px;line-height:1}
.sv.red{color:var(--red)}.sv.amb{color:var(--amb)}.sv.grn{color:var(--grn)}.sv.blu{color:var(--blu)}
.ss{width:1px;height:34px;background:rgba(255,255,255,.1)}

/* RANK LIST */
.rank-list{display:flex;flex-direction:column;gap:8px}
.rank-row{display:grid;grid-template-columns:30px 1fr 140px 80px;align-items:center;gap:11px}
.rn{font-size:11px;color:var(--mut);font-weight:700;text-align:right;font-family:"IBM Plex Mono",monospace}
.rnm{font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--txt2)}
.rb{background:var(--surf2);border-radius:3px;height:9px;overflow:hidden}
.rf{height:100%;border-radius:3px;background:linear-gradient(90deg,#2563EB,#388BFD);transition:width .3s}
.rv{font-size:11.5px;font-family:"IBM Plex Mono",monospace;text-align:right;color:var(--mut);font-weight:500}

/* AGEING BADGES */
.age-fresh,.age-ok,.age-warn,.age-alert,.age-old{display:inline-flex;font-size:10px;font-weight:600;padding:2px 7px;border-radius:4px;white-space:nowrap}
.age-fresh{background:rgba(63,185,80,.13);color:#3FB950;border:1px solid rgba(63,185,80,.28)}
.age-ok{background:rgba(88,166,255,.13);color:#79C0FF;border:1px solid rgba(88,166,255,.28)}
.age-warn{background:rgba(210,153,34,.13);color:#D29922;border:1px solid rgba(210,153,34,.28)}
.age-alert{background:rgba(248,81,73,.13);color:#F85149;border:1px solid rgba(248,81,73,.28)}
.age-old{background:rgba(163,113,247,.13);color:#A371F7;border:1px solid rgba(163,113,247,.28)}

/* PAGINATION */
.pgn{display:flex;align-items:center;gap:8px;padding:13px 18px;background:var(--surf2);border-top:1px solid var(--bdr);font-size:12px;color:var(--mut);justify-content:space-between;flex-wrap:wrap}
.pgn-info{font-family:"IBM Plex Mono",monospace}
.pgn-ctl{display:flex;align-items:center;gap:6px}
.pgn-btn{background:none;border:1px solid var(--bdr);border-radius:5px;color:var(--txt2);padding:5px 11px;font-size:11.5px;cursor:pointer;font-family:"Inter",sans-serif;font-weight:500;transition:all .15s;min-width:34px}
.pgn-btn:hover:not(:disabled){background:var(--surf3);border-color:var(--bdr2)}
.pgn-btn:disabled{opacity:.4;cursor:not-allowed}
.pgn-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.pgn-sel{background:var(--surf);border:1px solid var(--bdr);border-radius:5px;color:var(--txt2);padding:5px 10px;font-size:11.5px;font-family:"Inter",sans-serif;cursor:pointer}

/* FOOTER */
.footer{text-align:center;padding:22px 32px;border-top:1px solid var(--bdr);font-size:11.5px;color:var(--mut);margin-top:30px}
.footer a{color:var(--blu);text-decoration:none}

/* RESPONSIVE */
@media(max-width:1100px){.crow.c21,.crow.c11{grid-template-columns:1fr}}
@media(max-width:900px){.pi{padding:18px}.hdr,.nav{padding-left:16px;padding-right:16px}.kgrid{grid-template-columns:repeat(2,1fr)}.gs-wrap{margin:0 8px}.hmeta{display:none}}
@media(max-width:640px){.kgrid{grid-template-columns:1fr}.fbar{padding:10px}}

/* PRINT */
@media print{.hdr,.nav,.fbar,.csv-btn,.pgn,.gs-wrap{display:none!important}.page{display:block!important}.tmh{max-height:none!important;overflow:visible!important}}
"""

    # ── JS — built using % substitution, NOT f-string {{}} doubling ──────────
    import json as _json
    _data_blob = {
        'CATS':   cats,
        'CATCLS': catcls,
        'TOP':    _json.loads(top_json),
        'ZERO':   _json.loads(zero_json),
        'NEG':    _json.loads(neg_json),
        'OOS':    _json.loads(oos_json),
        'RISK':   _json.loads(risk_json),
        'UW':     _json.loads(uw_json),
        'PRE':    _json.loads(pre_json),
        'ML':     ml,
        'MV':     mv,
        'CL':     cl,
        'CV':     cv,
        'HV':     hv,
        'OOS_LABELS': oos_h,
        'K':      K,
        'TOP_SUP': top_sup,
    }
    DATA_JSON = _json.dumps(_data_blob, separators=(',', ':'))

    # JS is a plain string — placeholders are %(name)s style, only one substitution at the end.
    # No f-string brace-doubling. Cannot mismatch.
    JS = r"""
// ═══ AL MADINA INVENTORY DASHBOARD — Client-side runtime ═══

var __D = JSON.parse(document.getElementById('__data__').textContent);
var CATS = __D.CATS, CATCLS = __D.CATCLS, K = __D.K;
var DATASETS = {
  top:  __D.TOP,  zero: __D.ZERO, neg: __D.NEG,
  oos:  __D.OOS,  risk: __D.RISK, uw:  __D.UW, pre: __D.PRE
};
var ML = __D.ML, MV = __D.MV, CL = __D.CL, CV = __D.CV, HV = __D.HV;
var OOS_LABELS = __D.OOS_LABELS, TOP_SUP = __D.TOP_SUP;

// ─── UTILITIES ───────────────────────────────────────────────────────────
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function fmtN(v,d){ if(d==null) d=1; v=parseFloat(v)||0; var a=Math.abs(v); if(a>=1e6) return (v/1e6).toFixed(d)+'M'; if(a>=1e3) return (v/1e3).toFixed(d)+'K'; return v.toLocaleString('en',{maximumFractionDigits:d}); }
function fmtT(v){ v=parseFloat(v)||0; if(v>=1e6) return (v/1e6).toFixed(1)+'M'; if(v>=1e3) return (v/1e3).toFixed(0)+'K'; return Number(v).toLocaleString('en'); }
function fmtS(n){ n=parseFloat(n)||0; var a=Math.abs(n); if(a>=1e6) return (n/1e6).toFixed(1)+'M'; if(a>=1e3) return (n/1e3).toFixed(1)+'K'; return n.toLocaleString('en',{maximumFractionDigits:0}); }
function ageBadge(r){
  var b = r.ageb || '';
  if(b==='0–30 days')    return '<span class="age-fresh">'+b+'</span>';
  if(b==='31–90 days')   return '<span class="age-ok">'+b+'</span>';
  if(b==='91–180 days')  return '<span class="age-warn">'+b+'</span>';
  if(b==='181–360 days') return '<span class="age-alert">'+b+'</span>';
  if(b==='360+ days')    return '<span class="age-old">'+b+'</span>';
  return '<span class="muted">—</span>';
}

// ─── PAGE SWITCHING ──────────────────────────────────────────────────────
function SP(id){
  document.querySelectorAll('.page').forEach(function(p){ p.classList.remove('active'); });
  document.querySelectorAll('.nbtn').forEach(function(b){ b.classList.remove('active'); });
  var pg = document.getElementById('pg-'+id);
  var nb = document.getElementById('nb-'+id);
  if(pg) pg.classList.add('active');
  if(nb) nb.classList.add('active');
  window.scrollTo(0,0);
  // Resize charts that may have been hidden
  if(id==='ov'){
    setTimeout(function(){
      if(window.Chart && Chart.instances){
        Object.values(Chart.instances).forEach(function(c){ if(c.resize) c.resize(); });
      }
    }, 80);
  }
}
window.SP = SP;

// ─── CHARTS ──────────────────────────────────────────────────────────────
function initCharts(){
  if(!window.Chart) return;
  var CF = { family: "'Inter',sans-serif", size: 11, weight: '500' };
  Chart.defaults.font = CF;
  Chart.defaults.color = '#8B949E';
  Chart.defaults.borderColor = '#21262D';

  var mc = []; for(var i=0;i<ML.length;i++) mc.push(i===ML.length-1 ? '#3B82F6' : '#1e3a5f');

  if(ML.length && document.getElementById('chM')){
    new Chart(document.getElementById('chM'), {
      type:'bar',
      data:{ labels: ML, datasets:[{ data: MV, backgroundColor: mc, borderRadius: 5, borderSkipped: false }] },
      options:{
        responsive: true, maintainAspectRatio: false,
        plugins:{
          legend:{ display: false },
          tooltip:{ backgroundColor:'#161B22', borderColor:'#30363D', borderWidth:1, titleColor:'#E6EDF3', bodyColor:'#C9D1D9', padding:10, callbacks:{ label: function(c){ return ' '+Number(c.parsed.y).toLocaleString()+' units'; } } }
        },
        scales:{
          x:{ grid:{ display: false }, ticks:{ color:'#8B949E' } },
          y:{ grid:{ color:'#21262D' }, ticks:{ color:'#8B949E', callback: function(v){ return fmtT(v); } }, border:{ display: false } }
        }
      }
    });
  }
  if(CL.length && document.getElementById('chC')){
    var cc = ['#1B2B4B','#1e3a5f','#1e40af','#1d4ed8','#2563eb','#3b82f6','#60a5fa','#0f766e','#0d9488','#14b8a6'];
    new Chart(document.getElementById('chC'), {
      type:'bar',
      data:{ labels: CL, datasets:[{ data: CV, backgroundColor: cc.slice(0, CL.length), borderRadius: 5 }] },
      options:{
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins:{
          legend:{ display: false },
          tooltip:{ backgroundColor:'#161B22', borderColor:'#30363D', borderWidth:1, padding:10, callbacks:{ label: function(c){ return ' '+Number(c.parsed.x).toLocaleString()+' units'; } } }
        },
        scales:{
          x:{ grid:{ color:'#21262D' }, ticks:{ color:'#8B949E', callback: function(v){ return fmtT(v); } }, border:{ display: false } },
          y:{ grid:{ display: false }, ticks:{ color:'#8B949E' } }
        }
      }
    });
  }
  if(document.getElementById('chH')){
    new Chart(document.getElementById('chH'), {
      type:'doughnut',
      data:{ labels:['Healthy','Zero Sale','Negative','OOS','High Risk'], datasets:[{ data: HV, backgroundColor:['#3FB950','#D29922','#F85149','#A371F7','#388BFD'], borderWidth: 2, borderColor:'#161B22' }] },
      options:{
        responsive: true, maintainAspectRatio: false, cutout: '64%',
        plugins:{
          legend:{ position:'bottom', labels:{ color:'#8B949E', padding:10, boxWidth:10, font:{ size: 11 } } },
          tooltip:{ backgroundColor:'#161B22', borderColor:'#30363D', borderWidth:1, padding:10 }
        }
      }
    });
  }
}

// ─── TABLE ENGINE (filtering + sorting + pagination) ─────────────────────
var STATE = {};  // per-section state

function ageRank(b){
  if(b==='0–30 days') return 1; if(b==='31–90 days') return 2;
  if(b==='91–180 days') return 3; if(b==='181–360 days') return 4;
  if(b==='360+ days') return 5; return 9;
}

function applyFilters(secId){
  var s = STATE[secId];
  if(!s) return [];
  var data = s.all;
  var cat = s.cat, cls = s.cls, q = (s.q||'').toLowerCase().trim();
  var out = [];
  for(var i=0; i<data.length; i++){
    var r = data[i];
    if(cat && cat!=='All Categories' && r.cat !== cat) continue;
    if(cls && cls!=='All Classes' && r.cls !== cls) continue;
    if(q){
      var hay = ((r.n||'')+' '+(r.bc||'')+' '+(r.sup||'')+' '+(r.brand||'')+' '+(r.grp||'')).toLowerCase();
      if(hay.indexOf(q) === -1) continue;
    }
    out.push(r);
  }
  // sort
  if(s.sortKey){
    var k = s.sortKey, dir = s.sortDir;
    out.sort(function(a, b){
      var va = a[k], vb = b[k];
      if(k === 'ageb'){ va = ageRank(va); vb = ageRank(vb); }
      if(typeof va === 'number' && typeof vb === 'number') return dir==='asc' ? va-vb : vb-va;
      va = String(va==null?'':va).toLowerCase(); vb = String(vb==null?'':vb).toLowerCase();
      return dir==='asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
  }
  return out;
}

function renderSection(secId){
  var s = STATE[secId];
  if(!s) return;
  var d = applyFilters(secId);
  s.filtered = d;

  // pagination
  var pageSize = s.pageSize || 50;
  var totalPages = Math.max(1, Math.ceil(d.length / pageSize));
  if(s.page > totalPages) s.page = totalPages;
  if(s.page < 1) s.page = 1;
  var start = (s.page - 1) * pageSize;
  var slice = d.slice(start, start + pageSize);

  // render rows
  var tb = document.getElementById(secId+'-tb');
  if(tb){
    var html = '';
    for(var i=0; i<slice.length; i++){
      html += s.rowFn(slice[i], start + i);
    }
    tb.innerHTML = html || '<tr><td colspan="20" style="padding:30px;text-align:center;color:#8B949E">No items match the current filters.</td></tr>';
  }

  // counter
  var cnt = document.getElementById(secId+'-cnt');
  if(cnt){
    cnt.innerHTML = 'Showing <b>'+Math.min(start+1, d.length)+'</b>–<b>'+Math.min(start+pageSize, d.length)+'</b> of <b>'+d.length.toLocaleString()+'</b>';
  }
  // top10 rank (top section only)
  if(secId === 'top'){
    var rk = document.getElementById('top-rank');
    if(rk){
      var top10 = d.slice(0, 10);
      var mx = top10.length ? top10[0].ts : 1;
      rk.innerHTML = top10.map(function(r,i){
        return '<div class="rank-row"><div class="rn">'+(i+1)+'</div>'+
          '<div class="rnm" title="'+esc(r.n)+'">'+esc(r.n)+'</div>'+
          '<div class="rb"><div class="rf" style="width:'+(r.ts/mx*100).toFixed(1)+'%"></div></div>'+
          '<div class="rv">'+fmtN(r.ts,0)+'</div></div>';
      }).join('') || '<div class="muted" style="text-align:center;padding:14px">No items</div>';
    }
  }
  // strip
  if(s.stripFn){
    var sv = 0, l3 = 0;
    for(var j=0; j<d.length; j++){ sv += (d[j].sv||0); l3 += (d[j].l3||0); }
    var items = s.stripFn(d, sv, l3);
    var stEl = document.getElementById(secId+'-strip');
    if(stEl){
      stEl.innerHTML = '<div class="strip">' + items.map(function(it, i){
        return (i ? '<div class="ss"></div>' : '') +
               '<div class="si"><div class="sl">'+it[0]+'</div><div class="sv '+it[2]+'">'+it[1]+'</div></div>';
      }).join('') + '</div>';
    }
  }
  // pagination UI
  renderPagination(secId, d.length, pageSize);
  // sort indicator
  var thead = document.querySelector('#tbl-'+secId+' thead');
  if(thead){
    thead.querySelectorAll('th.srt').forEach(function(th){
      th.classList.remove('asc','desc');
      if(th.getAttribute('data-key') === s.sortKey){
        th.classList.add(s.sortDir === 'asc' ? 'asc' : 'desc');
      }
    });
  }
}

function renderPagination(secId, total, pageSize){
  var el = document.getElementById(secId+'-pgn');
  if(!el) return;
  var s = STATE[secId];
  var totalPages = Math.max(1, Math.ceil(total / pageSize));
  var p = s.page;
  // build page numbers (compact)
  var pages = [];
  var add = function(n){ pages.push(n); };
  if(totalPages <= 7){
    for(var i=1; i<=totalPages; i++) add(i);
  } else {
    add(1);
    if(p > 4) add('…');
    var lo = Math.max(2, p-1), hi = Math.min(totalPages-1, p+1);
    for(var i=lo; i<=hi; i++) add(i);
    if(p < totalPages-3) add('…');
    add(totalPages);
  }
  var btns = pages.map(function(n){
    if(n === '…') return '<span class="pgn-info" style="padding:0 4px">…</span>';
    return '<button class="pgn-btn'+(n===p?' active':'')+'" onclick="gotoPage(\''+secId+'\','+n+')">'+n+'</button>';
  }).join('');
  el.innerHTML =
    '<div class="pgn-info">Page <b>'+p+'</b> of <b>'+totalPages+'</b></div>'+
    '<div class="pgn-ctl">'+
      '<button class="pgn-btn" onclick="gotoPage(\''+secId+'\',1)" '+(p===1?'disabled':'')+'>« First</button>'+
      '<button class="pgn-btn" onclick="gotoPage(\''+secId+'\','+(p-1)+')" '+(p===1?'disabled':'')+'>‹ Prev</button>'+
      btns +
      '<button class="pgn-btn" onclick="gotoPage(\''+secId+'\','+(p+1)+')" '+(p===totalPages?'disabled':'')+'>Next ›</button>'+
      '<button class="pgn-btn" onclick="gotoPage(\''+secId+'\','+totalPages+')" '+(p===totalPages?'disabled':'')+'>Last »</button>'+
    '</div>'+
    '<div class="pgn-ctl"><span style="font-size:11.5px">Rows:</span>'+
      '<select class="pgn-sel" onchange="changePageSize(\''+secId+'\',this.value)">'+
        ['25','50','100','200','500'].map(function(n){ return '<option'+(parseInt(n)===pageSize?' selected':'')+'>'+n+'</option>'; }).join('')+
      '</select></div>';
}
function gotoPage(secId, n){ if(STATE[secId]){ STATE[secId].page = n; renderSection(secId); var pg=document.querySelector('.page.active .tcard'); if(pg) pg.scrollIntoView({behavior:'smooth',block:'start'}); } }
function changePageSize(secId, n){ if(STATE[secId]){ STATE[secId].pageSize = parseInt(n); STATE[secId].page = 1; renderSection(secId); } }
window.gotoPage = gotoPage; window.changePageSize = changePageSize;

// ─── ROW BUILDERS ────────────────────────────────────────────────────────
function rowTop(r,i){
  var mc = r.mg < 10 ? 'c-red' : (r.mg > 30 ? 'c-green fw6' : '');
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="muted">'+esc(r.brand||'—')+'</td>'+
    '<td class="tr bold-green">'+fmtN(r.ts,0)+'</td>'+
    '<td class="tr c-green">'+fmtN(r.l3,0)+'</td>'+
    '<td class="tr">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr">'+fmtN(r.sv,0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td class="tr '+mc+'">'+r.mg.toFixed(1)+'%</td></tr>';
}
function rowZero(r,i){
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="muted">'+esc(r.grp||'—')+'</td>'+
    '<td class="tr c-amber">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr c-red fw6">'+fmtN(r.sv,0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td>'+esc(r.lpd||'—')+'</td>'+
    '<td class="tr">'+(r.lpq||'—')+'</td>'+
    '<td>'+ageBadge(r)+'</td>'+
    '<td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';
}
function rowNeg(r,i){
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="tr c-red fw6">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr c-red">'+fmtN(r.sv,0)+'</td>'+
    '<td class="tr">'+fmtN(r.ts,0)+'</td>'+
    '<td class="tr c-green">'+fmtN(r.l3,0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td>'+esc(r.lpd||'—')+'</td>'+
    '<td>'+ageBadge(r)+'</td>'+
    '<td class="tr">'+(r.lpq||'—')+'</td></tr>';
}
function rowOOS(r,i){
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="tr c-green fw6">'+fmtN(r.l3,0)+'</td>'+
    '<td class="tr">'+fmtN(r.m0,1)+'</td>'+
    '<td class="tr">'+fmtN(r.m1,1)+'</td>'+
    '<td class="tr">'+fmtN(r.m2,1)+'</td>'+
    '<td class="tr c-red fw6">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr">'+fmtN(r.ts,0)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';
}
function rowRisk(r,i){
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="tr c-amber fw6">'+fmtN(r.sv,0)+'</td>'+
    '<td class="tr">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr c-red">'+fmtN(r.ts,0)+'</td>'+
    '<td class="tr">'+(r.l3||0).toFixed(0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td class="tr">'+r.mg.toFixed(1)+'%</td>'+
    '<td>'+esc(r.lpd||'—')+'</td>'+
    '<td>'+ageBadge(r)+'</td>'+
    '<td class="tr">'+(r.lpq||'—')+'</td></tr>';
}
function rowUW(r,i){
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="tr c-amber">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr c-red fw6">'+fmtN(r.sv,0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td class="tr">'+(r.lpq||'—')+'</td>'+
    '<td>'+esc(r.lpd||'—')+'</td>'+
    '<td>'+ageBadge(r)+'</td>'+
    '<td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';
}
function rowPre(r,i){
  var mc = r.mg < 10 ? 'c-red' : (r.mg > 30 ? 'c-green fw6' : '');
  var tsBadge = r.ts <= 0 ? '<span class="b-red">Zero Sale</span>' : '<span class="b-green">'+fmtN(r.ts,0)+'</span>';
  return '<tr data-bc="'+esc(r.bc)+'"><td class="mono">'+(i+1)+'</td>'+
    '<td><span class="tn" title="'+esc(r.n)+'">'+esc(r.n)+'</span><span class="bc">'+esc(r.bc)+'</span></td>'+
    '<td><span class="b-slate">'+esc(r.cat)+'</span></td>'+
    '<td><span class="b-brand">'+esc(r.cls)+'</span></td>'+
    '<td class="muted">'+esc(r.grp||'—')+'</td>'+
    '<td class="muted">'+esc(r.brand||'—')+'</td>'+
    '<td class="tr c-amber">'+fmtN(r.stock,1)+'</td>'+
    '<td class="tr c-amber fw6">'+fmtN(r.sv,0)+'</td>'+
    '<td>'+tsBadge+'</td>'+
    '<td class="tr">'+fmtN(r.l3,0)+'</td>'+
    '<td class="tr">'+r.cost.toFixed(2)+'</td>'+
    '<td class="tr">'+r.sell.toFixed(2)+'</td>'+
    '<td class="tr '+mc+'">'+r.mg.toFixed(1)+'%</td>'+
    '<td><span class="sup">'+esc(r.sup||'—')+'</span></td></tr>';
}

// ─── CLASS OPTIONS ──────────────────────────────────────────────────────
function updClsOpts(selId, cat){
  var el = document.getElementById(selId);
  if(!el) return;
  var list = (cat && cat !== 'All Categories' && CATCLS[cat]) || [];
  el.innerHTML = '<option value="All Classes">All Classes</option>' +
    list.map(function(c){ return '<option value="'+esc(c)+'">'+esc(c)+'</option>'; }).join('');
}

// ─── CSV DOWNLOAD ───────────────────────────────────────────────────────
function downloadCSV(secId){
  var s = STATE[secId]; if(!s) return;
  var data = s.filtered || s.all;
  var headers = s.csvHeaders;
  var rows = [headers.join(',')];
  for(var i=0; i<data.length; i++){
    var r = data[i];
    rows.push(headers.map(function(h){
      var v = r[h]==null ? '' : r[h];
      var sv = String(v).replace(/"/g, '""');
      return '"' + sv + '"';
    }).join(','));
  }
  var blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = secId + '_export_' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}
window.downloadCSV = downloadCSV;

// ─── SECTION INITIALISER ────────────────────────────────────────────────
function initSection(cfg){
  var id = cfg.id;
  STATE[id] = {
    all: DATASETS[id], filtered: DATASETS[id],
    cat: 'All Categories', cls: 'All Classes', q: '',
    sortKey: null, sortDir: 'desc',
    page: 1, pageSize: 50,
    rowFn: cfg.rowFn, stripFn: cfg.stripFn, csvHeaders: cfg.csvHeaders
  };

  var catEl = document.getElementById(id+'-cat');
  var clsEl = document.getElementById(id+'-cls');
  var qEl   = document.getElementById(id+'-q');

  if(catEl){
    catEl.addEventListener('change', function(){
      STATE[id].cat = catEl.value;
      STATE[id].cls = 'All Classes';
      updClsOpts(id+'-cls', catEl.value);
      STATE[id].page = 1;
      renderSection(id);
    });
  }
  if(clsEl){
    clsEl.addEventListener('change', function(){
      STATE[id].cls = clsEl.value;
      STATE[id].page = 1;
      renderSection(id);
    });
  }
  if(qEl){
    var t;
    qEl.addEventListener('input', function(){
      clearTimeout(t);
      t = setTimeout(function(){
        STATE[id].q = qEl.value;
        STATE[id].page = 1;
        renderSection(id);
      }, 180);
    });
  }

  window[id+'Reset'] = function(){
    if(catEl) catEl.value = 'All Categories';
    if(qEl)   qEl.value = '';
    STATE[id].cat = 'All Categories';
    STATE[id].cls = 'All Classes';
    STATE[id].q   = '';
    STATE[id].page = 1;
    STATE[id].sortKey = null;
    updClsOpts(id+'-cls', 'All Categories');
    renderSection(id);
  };

  // sortable header clicks
  var thead = document.querySelector('#tbl-'+id+' thead');
  if(thead){
    thead.querySelectorAll('th.srt').forEach(function(th){
      th.addEventListener('click', function(){
        var key = th.getAttribute('data-key');
        var s = STATE[id];
        if(s.sortKey === key){
          s.sortDir = s.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          s.sortKey = key;
          s.sortDir = (key === 'n' || key === 'cat' || key === 'cls' || key === 'sup' || key === 'lpd') ? 'asc' : 'desc';
        }
        s.page = 1;
        renderSection(id);
      });
    });
  }

  if(clsEl) updClsOpts(id+'-cls', 'All Categories');
  renderSection(id);
}

// ─── GLOBAL SEARCH ──────────────────────────────────────────────────────
var GS_MAP = null;  // bc -> {section, record}
function buildSearchIndex(){
  GS_MAP = [];
  var sections = [
    ['top','Top Movers','b-green'],['zero','Zero Sale','b-red'],['neg','Negative Stock','b-red'],
    ['oos','OOS Active','b-red'],['risk','High Risk','b-amber'],['uw','Unwanted','b-red'],['pre','Pre-2025','b-amber']
  ];
  // build flat list with dedup by barcode keeping first appearance
  var seen = {};
  sections.forEach(function(sec){
    var sid = sec[0], slabel = sec[1], sbadge = sec[2];
    var arr = DATASETS[sid];
    for(var i=0; i<arr.length; i++){
      var r = arr[i];
      var key = r.bc || r.n;
      if(seen[key]) continue;
      seen[key] = true;
      GS_MAP.push({ bc: r.bc, n: r.n, cat: r.cat, cls: r.cls, sup: r.sup||'', brand: r.brand||'', sv: r.sv||0, stock: r.stock||0, ts: r.ts||0, sec: sid, secLabel: slabel, secBadge: sbadge });
    }
  });
}
var gsSel = -1, gsList = [];
function gsRender(q){
  var rEl = document.getElementById('gs-results');
  if(!q || q.length < 2){ rEl.classList.remove('open'); rEl.innerHTML = ''; gsList = []; return; }
  if(!GS_MAP) buildSearchIndex();
  q = q.toLowerCase();
  var hits = [];
  for(var i=0; i<GS_MAP.length && hits.length < 30; i++){
    var r = GS_MAP[i];
    var hay = ((r.n||'')+' '+(r.bc||'')+' '+(r.sup||'')+' '+(r.brand||'')).toLowerCase();
    if(hay.indexOf(q) !== -1) hits.push(r);
  }
  gsList = hits;
  gsSel = hits.length ? 0 : -1;
  if(!hits.length){
    rEl.innerHTML = '<div class="gs-empty">No items match <b>"'+esc(q)+'"</b></div>';
  } else {
    rEl.innerHTML = '<div class="gs-section">'+hits.length+' result'+(hits.length===1?'':'s')+' — press Enter to open</div>' +
      hits.map(function(r, idx){
        return '<div class="gs-result'+(idx===0?' sel':'')+'" data-idx="'+idx+'">' +
          '<span class="'+r.secBadge+' gsbadge">'+r.secLabel+'</span>' +
          '<div class="gsn" title="'+esc(r.n)+'">'+esc(r.n)+'</div>' +
          '<span class="gsmeta">'+esc(r.bc)+' • Stk '+fmtN(r.stock,0)+' • SV '+fmtN(r.sv,0)+'</span>' +
        '</div>';
      }).join('');
    rEl.querySelectorAll('.gs-result').forEach(function(el){
      el.addEventListener('mouseenter', function(){
        gsSel = parseInt(el.getAttribute('data-idx'));
        rEl.querySelectorAll('.gs-result').forEach(function(x){ x.classList.remove('sel'); });
        el.classList.add('sel');
      });
      el.addEventListener('click', function(){ gsOpen(parseInt(el.getAttribute('data-idx'))); });
    });
  }
  rEl.classList.add('open');
}
function gsOpen(idx){
  if(idx < 0 || idx >= gsList.length) return;
  var r = gsList[idx];
  document.getElementById('gs-input').value = '';
  document.getElementById('gs-results').classList.remove('open');
  // jump to that section, set query filter to the barcode, scroll to row
  SP(r.sec);
  setTimeout(function(){
    var qEl = document.getElementById(r.sec+'-q');
    if(qEl){
      qEl.value = r.bc;
      STATE[r.sec].q = r.bc;
      STATE[r.sec].page = 1;
      renderSection(r.sec);
      setTimeout(function(){
        var row = document.querySelector('#'+r.sec+'-tb tr[data-bc="'+CSS.escape(r.bc)+'"]');
        if(row){
          row.classList.add('hl');
          row.scrollIntoView({ behavior:'smooth', block:'center' });
          setTimeout(function(){ row.classList.remove('hl'); }, 2400);
        }
      }, 120);
    }
  }, 100);
}
function initGlobalSearch(){
  var inp = document.getElementById('gs-input');
  var rEl = document.getElementById('gs-results');
  if(!inp || !rEl) return;
  var t;
  inp.addEventListener('input', function(){
    clearTimeout(t);
    t = setTimeout(function(){ gsRender(inp.value); }, 120);
  });
  inp.addEventListener('keydown', function(e){
    if(!rEl.classList.contains('open')) return;
    if(e.key === 'ArrowDown'){
      e.preventDefault();
      gsSel = Math.min(gsList.length-1, gsSel+1);
      rEl.querySelectorAll('.gs-result').forEach(function(el, i){ el.classList.toggle('sel', i === gsSel); });
      var sel = rEl.querySelector('.gs-result.sel'); if(sel) sel.scrollIntoView({block:'nearest'});
    } else if(e.key === 'ArrowUp'){
      e.preventDefault();
      gsSel = Math.max(0, gsSel-1);
      rEl.querySelectorAll('.gs-result').forEach(function(el, i){ el.classList.toggle('sel', i === gsSel); });
      var sel = rEl.querySelector('.gs-result.sel'); if(sel) sel.scrollIntoView({block:'nearest'});
    } else if(e.key === 'Enter'){
      e.preventDefault();
      gsOpen(gsSel);
    } else if(e.key === 'Escape'){
      inp.value = ''; rEl.classList.remove('open'); inp.blur();
    }
  });
  document.addEventListener('click', function(e){
    if(!inp.contains(e.target) && !rEl.contains(e.target)) rEl.classList.remove('open');
  });
  // Ctrl/Cmd+K
  document.addEventListener('keydown', function(e){
    if((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k'){
      e.preventDefault();
      inp.focus();
      inp.select();
    }
    if(e.key === '/' && document.activeElement && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA' && document.activeElement.tagName !== 'SELECT'){
      e.preventDefault();
      inp.focus();
    }
  });
}

// ─── STRIP CONFIGS ──────────────────────────────────────────────────────
function stripZero(d, sv){ return [['Items Shown',fmtS(d.length),''],['Stock Value at Risk','AED '+fmtS(sv),'amb'],['Avg / Item',d.length?'AED '+fmtS(sv/d.length):'—',''],['Total Zero-Sale',fmtS(K.zero_count),'red']]; }
function stripNeg(d, sv){ return [['Items Shown',fmtS(d.length),''],['Neg. Value Exposure','AED '+fmtS(Math.abs(sv)),'red'],['Total Negative',fmtS(K.neg_count),'red'],['Priority','CRITICAL','amb']]; }
function stripOOS(d, sv, l3){ return [['OOS Items',fmtS(d.length),'red'],['Last 3M Sales',fmtS(l3)+' units','grn'],['Total OOS Portfolio',fmtS(K.oos_count),'red'],['Revenue Risk','HIGH','amb']]; }
function stripRisk(d, sv){ return [['Items Shown',fmtS(d.length),''],['Value at Risk','AED '+fmtS(sv),'amb'],['Portfolio Exposure','AED '+fmtS(K.risk_sv),'red'],['Total High Risk',fmtS(K.risk_count),'amb']]; }
function stripUW(d, sv){ return [['Items Shown',fmtS(d.length),''],['Stock Value','AED '+fmtS(sv),'red'],['Portfolio Total','AED '+fmtS(K.uw_sv),'red'],['Total Unwanted',fmtS(K.uw_count),'red']]; }
function stripPre(d, sv){
  var zi = d.filter(function(r){ return r.ts<=0; }), si = d.filter(function(r){ return r.ts>0; });
  var zsv = zi.reduce(function(s,r){ return s+(r.sv||0); }, 0);
  var ssv = si.reduce(function(s,r){ return s+(r.sv||0); }, 0);
  return [['Total Shown',fmtS(d.length),'amb'],['Total Stock Value','AED '+fmtS(sv),'amb'],['Zero Sale — '+fmtS(zi.length),'AED '+fmtS(zsv),'red'],['Still Selling — '+fmtS(si.length),'AED '+fmtS(ssv),'grn']];
}

// ─── BOOTSTRAP ──────────────────────────────────────────────────────────
function initAll(){
  // nav button click handlers
  document.querySelectorAll('.nbtn').forEach(function(b){
    var pid = b.id.replace('nb-','');
    b.addEventListener('click', function(){ SP(pid); });
  });

  initSection({ id:'top',  rowFn: rowTop,  stripFn: null,       csvHeaders:['bc','n','cat','cls','brand','ts','l3','stock','sv','cost','sell','mg'] });
  initSection({ id:'zero', rowFn: rowZero, stripFn: stripZero,  csvHeaders:['bc','n','cat','cls','grp','sup','stock','sv','cost','sell','lpd','lpq','ageb'] });
  initSection({ id:'neg',  rowFn: rowNeg,  stripFn: stripNeg,   csvHeaders:['bc','n','cat','cls','cost','sell','stock','sv','ts','l3','lpd','lpq','ageb'] });
  initSection({ id:'oos',  rowFn: rowOOS,  stripFn: stripOOS,   csvHeaders:['bc','n','cat','cls','sup','sell','stock','ts','l3','m0','m1','m2'] });
  initSection({ id:'risk', rowFn: rowRisk, stripFn: stripRisk,  csvHeaders:['bc','n','cat','cls','cost','sell','stock','sv','ts','mg','l3','lpd','lpq','ageb'] });
  initSection({ id:'uw',   rowFn: rowUW,   stripFn: stripUW,    csvHeaders:['bc','n','cat','cls','cost','sell','stock','sv','lpq','lpd','sup','ageb'] });
  initSection({ id:'pre',  rowFn: rowPre,  stripFn: stripPre,   csvHeaders:['bc','n','cat','cls','grp','brand','sup','stock','sv','ts','l3','cost','sell','mg'] });

  initCharts();
  initGlobalSearch();
  renderTopSuppliers();

  SP('ov');
}
function renderTopSuppliers(){
  var el = document.getElementById('top-sup');
  if(!el || !TOP_SUP || !TOP_SUP.length) return;
  var mx = TOP_SUP[0].sales || 1;
  el.innerHTML = TOP_SUP.map(function(s, i){
    return '<div class="rank-row"><div class="rn">'+(i+1)+'</div>'+
      '<div class="rnm" title="'+esc(s.name)+'">'+esc(s.name)+' <span class="muted" style="font-size:10.5px">('+s.items+' SKUs)</span></div>'+
      '<div class="rb"><div class="rf" style="width:'+(s.sales/mx*100).toFixed(1)+'%"></div></div>'+
      '<div class="rv">'+fmtN(s.sales,0)+'</div></div>';
  }).join('');
}

if(document.readyState === 'loading'){
  document.addEventListener('DOMContentLoaded', initAll);
} else {
  initAll();
}
"""

    # ── Sortable header helper (used in HTML below) ──────────────────────────
    def sth(label, key, num=False):
        cls = 'srt tr' if num else 'srt'
        return f'<th class="{cls}" data-key="{key}">{label}</th>'

    def fbar(prefix):
        # Each section has Cat + Class + Search box + Reset + Count + CSV
        return (f'<div class="fbar">'
                f'<span class="flbl">Filter</span><div class="fsep"></div>'
                f'<div class="fg"><label>Category</label><select class="fsel" id="{prefix}-cat">{cat_opts}</select></div>'
                f'<div class="fg"><label>Class</label><select class="fsel" id="{prefix}-cls"></select></div>'
                f'<div class="fg"><label>Search</label><input class="fin" id="{prefix}-q" type="text" placeholder="Name, barcode, supplier…"></div>'
                f'<button class="freset" onclick="{prefix}Reset()">Reset</button>'
                f'<span class="fcnt" id="{prefix}-cnt"></span>'
                f'<button class="csv-btn" onclick="downloadCSV(\'{prefix}\')">↓ Export CSV</button>'
                f'</div>')

    # ── Headers for each table (sortable) ────────────────────────────────────
    top_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                 + sth('Category', 'cat') + sth('Class', 'cls') + sth('Brand', 'brand')
                 + sth('Total Sales', 'ts', True) + sth('Last 3M', 'l3', True)
                 + sth('Stock', 'stock', True) + sth('Stock Value (AED)', 'sv', True)
                 + sth('Cost', 'cost', True) + sth('Selling', 'sell', True) + sth('Margin%', 'mg', True)
                 + '</tr>')
    zero_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                  + sth('Category', 'cat') + sth('Class', 'cls') + sth('Group', 'grp')
                  + sth('Stock', 'stock', True) + sth('Stock Value (AED)', 'sv', True)
                  + sth('Cost', 'cost', True) + sth('Selling', 'sell', True)
                  + sth('Last Purchase', 'lpd') + sth('LP Qty', 'lpq', True)
                  + sth('Ageing', 'ageb') + '<th>Supplier</th></tr>')
    neg_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                 + sth('Category', 'cat') + sth('Class', 'cls')
                 + sth('Stock Qty', 'stock', True) + sth('Stock Value (AED)', 'sv', True)
                 + sth('Total Sales', 'ts', True) + sth('Last 3M', 'l3', True)
                 + sth('Cost', 'cost', True) + sth('Selling', 'sell', True)
                 + sth('Last Purchase', 'lpd') + sth('Ageing', 'ageb')
                 + sth('LP Qty', 'lpq', True) + '</tr>')
    oos_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                 + sth('Category', 'cat') + sth('Class', 'cls')
                 + sth('Last 3M', 'l3', True)
                 + f'<th class="srt tr" data-key="m0">{oos_h[0]}</th>'
                 + f'<th class="srt tr" data-key="m1">{oos_h[1]}</th>'
                 + f'<th class="srt tr" data-key="m2">{oos_h[2]}</th>'
                 + sth('Current Stock', 'stock', True) + sth('Total Sales', 'ts', True)
                 + sth('Selling', 'sell', True) + '<th>Supplier</th></tr>')
    risk_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                  + sth('Category', 'cat') + sth('Class', 'cls')
                  + sth('Stock Value (AED)', 'sv', True) + sth('Stock', 'stock', True)
                  + sth('Total Sales', 'ts', True) + sth('Last 3M', 'l3', True)
                  + sth('Cost', 'cost', True) + sth('Selling', 'sell', True) + sth('Margin%', 'mg', True)
                  + sth('Last Purchase', 'lpd') + sth('Ageing', 'ageb') + sth('LP Qty', 'lpq', True) + '</tr>')
    uw_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                + sth('Category', 'cat') + sth('Class', 'cls')
                + sth('Current Stock', 'stock', True) + sth('Stock Value (AED)', 'sv', True)
                + sth('Cost', 'cost', True) + sth('Selling', 'sell', True)
                + sth('Last PO Qty', 'lpq', True) + sth('Last Purchase Date', 'lpd')
                + sth('Ageing', 'ageb') + '<th>Last Purchase Supplier</th></tr>')
    pre_thead = ('<tr><th>#</th><th>Item Name / Barcode</th>'
                 + sth('Category', 'cat') + sth('Class', 'cls') + sth('Group', 'grp') + sth('Brand', 'brand')
                 + sth('Stock Qty', 'stock', True) + sth('Stock Value (AED)', 'sv', True)
                 + sth('Total Sales', 'ts', True) + sth('Last 3M', 'l3', True)
                 + sth('Cost', 'cost', True) + sth('Selling', 'sell', True) + sth('Margin%', 'mg', True)
                 + '<th>Supplier</th></tr>')

    HTML_BODY = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>AL MADINA GROUP — Inventory Intelligence Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="hdr">
  <div class="hdr-l">
    <div class="logo">
      <div class="lm">AM</div>
      <div><div class="lt">AL MADINA GROUP</div><div class="ls">Inventory Intelligence Platform</div></div>
    </div>
    <div class="hsep"></div>
    <div class="hlbl">Stock Health Report</div>
  </div>
  <div class="gs-wrap">
    <span class="gs-ico">⌕</span>
    <input class="gs-input" id="gs-input" type="text" placeholder="Search all items by name, barcode, or supplier…" autocomplete="off">
    <span class="gs-kbd">⌘K</span>
    <div class="gs-results" id="gs-results"></div>
  </div>
  <div class="hr-r">
    <div class="hmeta"><b>{fmtN(K['total_items'],0)}</b> SKUs<br>{dlabel}</div>
  </div>
</header>

<nav class="nav">
  <button class="nbtn active" id="nb-ov">Overview</button>
  <button class="nbtn"        id="nb-top">Top Movers <span class="npill blu">{fmtN(len(_data_blob['TOP']),0)}</span></button>
  <button class="nbtn"        id="nb-zero">Zero Sale <span class="npill amb">{fmtN(K['zero_count'],0)}</span></button>
  <button class="nbtn"        id="nb-neg">Negative Stock <span class="npill red">{fmtN(K['neg_count'],0)}</span></button>
  <button class="nbtn"        id="nb-oos">OOS / Active Demand <span class="npill red">{fmtN(K['oos_count'],0)}</span></button>
  <button class="nbtn"        id="nb-risk">High Risk Slow Movers <span class="npill amb">{fmtN(K['risk_count'],0)}</span></button>
  <button class="nbtn"        id="nb-uw">Unwanted Repurchases <span class="npill red">{fmtN(K['uw_count'],0)}</span></button>
  <button class="nbtn"        id="nb-pre">Pre-2025 Stock <span class="npill amb">{fmtN(K['pre_count'],0)}</span></button>
  <button class="nbtn"        id="nb-ins">Insights &amp; Actions</button>
</nav>

<!-- ═══ OVERVIEW ═══ -->
<div id="pg-ov" class="page active"><div class="pi">
  <div class="ptitle">Inventory Overview</div>
  <div class="pdesc">Portfolio of <b>{fmtN(K['total_items'],0)} SKUs</b> — stock value <b>{fmtA(K['total_sv'],0)}</b>, cumulative sales <b>{fmtN(K['total_ts'],0)} units</b>. Tap any KPI category in the nav to drill in. Press <b>⌘K</b> or <b>/</b> to search across everything.</div>
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

<!-- ═══ TOP MOVERS ═══ -->
<div id="pg-top" class="page"><div class="pi">
  <div class="ptitle">Top Moving Items</div>
  <div class="pdesc">All items with recorded sales, ranked by total units sold. Click any column header to re-sort.</div>
  {fbar('top')}
  <div class="cbox" style="margin-bottom:20px">
    <div class="chead"><span class="ctitle">Top 10 — Sales Volume (Filtered View)</span></div>
    <div class="rank-list" id="top-rank"></div>
  </div>
  <div class="tcard" id="tbl-top"><div class="thd"><span class="ttitle">All Items by Sales Volume</span><div class="tinfo"><span class="b-brand">Click headers to sort</span></div></div>
    <div class="tsc tmh"><table><thead>{top_thead}</thead><tbody id="top-tb"></tbody></table></div>
    <div class="pgn" id="top-pgn"></div>
  </div>
</div></div>

<!-- ═══ ZERO SALE ═══ -->
<div id="pg-zero" class="page"><div class="pi">
  <div class="ptitle">Zero Sale Items &mdash; In Stock</div>
  <div class="pdesc">Items with no sales recorded, currently holding positive stock. Pre-2025 items (blank LP Date) have a dedicated page.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">⚠</i><span class="at">Dead Stock — {fmtN(K['zero_count'],0)} Items</span></div><div class="ab">Zero sales with no demonstrated demand. Capital tied up with no return.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">₹</i><span class="at">Capital at Risk — {fmtA(K['zero_sv'],0)}</span></div><div class="ab">Immediate action: supplier return, clearance promotion, or write-off.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">→</i><span class="at">Action Required</span></div><div class="ab">Classify: (1) Supplier return, (2) Promotional clearance, (3) Write-off. All within 30 days.</div></div>
  </div>
  {fbar('zero')}
  <div id="zero-strip"></div>
  <div class="tcard" id="tbl-zero"><div class="thd"><span class="ttitle">Zero Sale Items with Stock</span><div class="tinfo"><span class="b-red">Default: Stock Value (high → low)</span></div></div>
    <div class="tsc tmh"><table><thead>{zero_thead}</thead><tbody id="zero-tb"></tbody></table></div>
    <div class="pgn" id="zero-pgn"></div>
  </div>
</div></div>

<!-- ═══ NEGATIVE STOCK ═══ -->
<div id="pg-neg" class="page"><div class="pi">
  <div class="ptitle">Negative Stock Items</div>
  <div class="pdesc">Items with negative on-hand quantities — unposted GRNs, missing receipts, or reconciliation issues.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">⚠</i><span class="at">{fmtN(K['neg_count'],0)} Items — Negative Quantity</span></div><div class="ab">Distorts financial reports and reorder calculations. Root cause required immediately.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">☉</i><span class="at">Probable Causes</span></div><div class="ab">Unposted GRNs, manual sales errors, missing supplier invoices.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">→</i><span class="at">Resolution Steps</span></div><div class="ab">1. Physical count. 2. Match GRNs. 3. Post verified receipts. 4. Raise adjustment journals.</div></div>
  </div>
  {fbar('neg')}
  <div id="neg-strip"></div>
  <div class="tcard" id="tbl-neg"><div class="thd"><span class="ttitle">Negative Stock Items</span><div class="tinfo"><span class="b-red">Default: Negative Qty (worst first)</span></div></div>
    <div class="tsc tmh"><table><thead>{neg_thead}</thead><tbody id="neg-tb"></tbody></table></div>
    <div class="pgn" id="neg-pgn"></div>
  </div>
</div></div>

<!-- ═══ OOS ═══ -->
<div id="pg-oos" class="page"><div class="pi">
  <div class="ptitle">Out-of-Stock — Active Demand</div>
  <div class="pdesc">Items that sold in the last 3 months but currently have zero or negative stock. Highest priority for reordering.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">⚑</i><span class="at">{fmtN(K['oos_count'],0)} Items Actively Out of Stock</span></div><div class="ab">Confirmed recent demand but unavailable. Every day without stock is direct revenue loss.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">⇧</i><span class="at">Prioritise by Last 3M Velocity</span></div><div class="ab">Highest recent sales = greatest daily revenue risk. Emergency POs immediately.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">✓</i><span class="at">Immediate Reorder Candidates</span></div><div class="ab">Proven demand. Predictable replenishment ROI. Buying team action list for today.</div></div>
  </div>
  {fbar('oos')}
  <div id="oos-strip"></div>
  <div class="tcard" id="tbl-oos"><div class="thd"><span class="ttitle">OOS Items with Recent Sales</span><div class="tinfo"><span class="b-red">Default: Last 3M sales (high → low)</span></div></div>
    <div class="tsc tmh"><table><thead>{oos_thead}</thead><tbody id="oos-tb"></tbody></table></div>
    <div class="pgn" id="oos-pgn"></div>
  </div>
</div></div>

<!-- ═══ HIGH RISK ═══ -->
<div id="pg-risk" class="page"><div class="pi">
  <div class="ptitle">High-Value Slow Movers</div>
  <div class="pdesc">Stock value &gt; AED 200 and total sales &lt; 10 units. High capital deployment, minimal turnover.</div>
  <div class="arow">
    <div class="abox amb"><div class="ah"><i class="ai">↓</i><span class="at">{fmtA(K['risk_sv'],0)} — Low Turnover Capital</span></div><div class="ab">{fmtN(K['risk_count'],0)} items with fewer than 10 units sold.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">⊘</i><span class="at">Holding Cost Alert</span></div><div class="ab">At 15% annual holding cost: ~{fmtA(K['risk_sv']*0.15,0)}/year. Redeploy to fast-moving lines.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">→</i><span class="at">Exit Strategy</span></div><div class="ab">Items &gt;6 months, &lt;3 units: immediate markdown. Return-eligible: negotiate credit.</div></div>
  </div>
  {fbar('risk')}
  <div id="risk-strip"></div>
  <div class="tcard" id="tbl-risk"><div class="thd"><span class="ttitle">High-Value Slow Moving Items</span><div class="tinfo"><span class="b-amber">Default: Stock Value (high → low)</span></div></div>
    <div class="tsc tmh"><table><thead>{risk_thead}</thead><tbody id="risk-tb"></tbody></table></div>
    <div class="pgn" id="risk-pgn"></div>
  </div>
</div></div>

<!-- ═══ UNWANTED REPURCHASES ═══ -->
<div id="pg-uw" class="page"><div class="pi">
  <div class="ptitle">Unwanted Repurchases</div>
  <div class="pdesc">Stock &gt; Last Purchase Qty AND Total Sales = 0 — re-purchased even though prior stock was unsold.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">⊘</i><span class="at">{fmtA(K['uw_sv'],0)} — Confirmed Re-purchase Waste</span></div><div class="ab">{fmtN(K['uw_count'],0)} items had prior stock, were re-purchased, and still haven't sold.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">☉</i><span class="at">Supplier Leverage</span></div><div class="ab">Prior stock + repurchase + zero sales = strong evidence for buy-back or credit notes.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">→</i><span class="at">Process Fix</span></div><div class="ab">Require buyers to check stock before any PO. System alert on zero-sale items with existing stock.</div></div>
  </div>
  {fbar('uw')}
  <div id="uw-strip"></div>
  <div class="tcard" id="tbl-uw"><div class="thd"><span class="ttitle">Re-purchased Items with Zero Sales</span><div class="tinfo"><span class="b-red">Default: Stock Value (high → low)</span></div></div>
    <div class="tsc tmh"><table><thead>{uw_thead}</thead><tbody id="uw-tb"></tbody></table></div>
    <div class="pgn" id="uw-pgn"></div>
  </div>
</div></div>

<!-- ═══ PRE-2025 ═══ -->
<div id="pg-pre" class="page"><div class="pi">
  <div class="ptitle">Pre-2025 Stock</div>
  <div class="pdesc">Items with stock on hand but no Last Purchase Date — purchased before the current data window. Includes selling and non-selling items.</div>
  <div class="arow">
    <div class="abox amb"><div class="ah"><i class="ai">◔</i><span class="at">{fmtN(K['pre_count'],0)} Items — No Purchase Record</span></div><div class="ab">Blank Last Purchase Date = stock entered before 2025. Total value: {fmtA(K['pre_sv'],0)}.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">⚠</i><span class="at">{fmtN(K['pre_zero_count'],0)} Items — Zero Sales + Pre-2025</span></div><div class="ab">Never sold, no purchase history. Most aged dead stock — {fmtA(K['pre_zero_sv'],0)} at highest write-off risk.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">✓</i><span class="at">{fmtN(K['pre_count']-K['pre_zero_count'],0)} Items — Still Selling</span></div><div class="ab">Pre-2025 stock with proven longevity. Continue to replenish when stock runs low.</div></div>
  </div>
  {fbar('pre')}
  <div id="pre-strip"></div>
  <div class="tcard" id="tbl-pre"><div class="thd"><span class="ttitle">Pre-2025 Stock Items (Blank LP Date)</span><div class="tinfo"><span class="b-amber">Default: Stock Value (high → low)</span></div></div>
    <div class="tsc tmh"><table><thead>{pre_thead}</thead><tbody id="pre-tb"></tbody></table></div>
    <div class="pgn" id="pre-pgn"></div>
  </div>
</div></div>

<!-- ═══ INSIGHTS ═══ -->
<div id="pg-ins" class="page"><div class="pi">
  <div class="ptitle">Strategic Analysis &amp; Action Plan</div>
  <div class="pdesc">Prioritised management recommendations from the full inventory dataset.</div>
  <div class="arow" style="grid-template-columns:1fr 1fr">
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">OOS Active Demand — {fmtN(K['oos_count'],0)} Items</span></div><div class="ab">Actively selling but unavailable. Emergency POs for top items by last-3M velocity immediately.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">Repurchase Waste — {fmtA(K['uw_sv'],0)}</span></div><div class="ab">{fmtN(K['uw_count'],0)} items re-purchased with existing stock and zero sales. Supplier buy-back or credit notes.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">Negative Stock — {fmtN(K['neg_count'],0)} Items</span></div><div class="ab">Compromises financial reporting. 5-business-day resolution. Physical count + GRN reconciliation.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">{fmtA(K['risk_sv'],0)} Slow Mover Capital</span></div><div class="ab">Tiered markdown: 30% at 90 days, 50% at 180 days, supplier credit at 365 days.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">P3</i><span class="at">Zero Sale Dead Stock — {fmtA(K['zero_sv'],0)}</span></div><div class="ab">{fmtN(K['zero_count'],0)} items — supplier return, clearance, or write-off within 30 days.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">✓</i><span class="at">Food &amp; Essentials Strength</span></div><div class="ab">Grocery, Fresh Produce, and Beverages drive footfall. Zero OOS tolerance on these lines.</div></div>
  </div>

  <div class="crow c21">
    <div class="cbox">
      <div class="chead"><span class="ctitle">Top 10 Suppliers by Sales</span><span class="cmeta">{len(top_sup)} suppliers shown</span></div>
      <div class="rank-list" id="top-sup"></div>
    </div>
    <div class="cbox">
      <div class="chead"><span class="ctitle">Stock-to-Sales Ratio</span></div>
      <div style="display:flex;flex-direction:column;gap:10px">
        <div><div class="klbl" style="margin-bottom:3px">Stock Value / Annual Sales</div><div style="font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:700;color:#79C0FF">{(K['total_sv']/max(1,K['total_ts'])):.2f}x</div></div>
        <div><div class="klbl" style="margin-bottom:3px">Dead Stock %</div><div style="font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:700;color:#F85149">{(K['zero_sv']/max(1,K['total_sv'])*100):.1f}%</div></div>
        <div><div class="klbl" style="margin-bottom:3px">Active SKU %</div><div style="font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:700;color:#3FB950">{((K['in_stock']-K['zero_count'])/max(1,K['total_items'])*100):.1f}%</div></div>
        <div><div class="klbl" style="margin-bottom:3px">Recovery Target (30d)</div><div style="font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:700;color:#D29922">AED {fmtN((K['zero_sv']+K['uw_sv'])*0.4,0)}</div></div>
      </div>
    </div>
  </div>

  <div class="tcard"><div class="thd"><span class="ttitle">Prioritised Action Plan</span><span class="b-brand">Management Decision Matrix</span></div>
  <div class="tsc"><table><thead><tr><th>Priority</th><th>Issue</th><th>Scale</th><th>Financial Exposure</th><th>Action</th><th>Owner</th><th>Deadline</th></tr></thead>
  <tbody>{action_html}</tbody></table></div></div>

  <div class="arow" style="grid-template-columns:1fr 1fr 1fr;margin-top:8px">
    <div class="abox grn"><div class="ah"><i class="ai">∑</i><span class="at">Supplier Leverage</span></div><div class="ab">Purchased-but-unsold evidence. Buy-back, credit, or exchange. Recovery: 40–60% of exposed capital.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">⊕</i><span class="at">Clearance Strategy</span></div><div class="ab">Bundle dead stock with fast movers. End-cap placements. Target: 25% moved in 60 days.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">☉</i><span class="at">Governance</span></div><div class="ab">Monthly slow-mover reviews. 3 consecutive zero-sale months = mandatory review. 6 months + &lt;5 units = flag.</div></div>
  </div>
</div></div>

<div class="footer">AL MADINA GROUP &nbsp;|&nbsp; Inventory Intelligence Report &nbsp;|&nbsp; {dlabel} &nbsp;|&nbsp; Generated locally — no data leaves your browser</div>

<script type="application/json" id="__data__">{DATA_JSON}</script>
<script>{JS}</script>
</body></html>"""

    return HTML_BODY


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — STREAMLIT PORTAL
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div style="display:flex;align-items:center;gap:16px;padding:8px 0 4px">
  <div style="background:linear-gradient(135deg,#1B2B4B,#3B82F6);width:48px;height:48px;border-radius:10px;
       display:flex;align-items:center;justify-content:center;font-weight:800;font-size:17px;
       color:white;flex-shrink:0;letter-spacing:-.5px">AM</div>
  <div style="min-width:0">
    <div style="font-size:20px;font-weight:700;color:#E6EDF3;letter-spacing:-.3px">AL MADINA GROUP</div>
    <div style="font-size:12px;color:#8B949E;margin-top:1px">Inventory Intelligence Dashboard Generator</div>
  </div>
</div>
<hr style="border-color:#30363D;margin:12px 0 20px">
""", unsafe_allow_html=True)

with st.expander("ℹ️  How to use", expanded=False):
    st.markdown("""
**Step 1** — Export your inventory report as `.xlsx`

**Required columns:** `Item Bar Code`, `Item Name`, `Stock`, `Total Sales`, `Cost`, `Selling`, `Stock Value`, `Category`, `Class`, `Supplier`, `Margin%`, `LP Date`, `LP Qty`

**Optional:** `Group`, `Brand`, `LP Supplier`, monthly columns e.g. `May, 2025` … `May, 2026`

**Blank LP Date** = purchased before 2025 → shown on the dedicated **Pre-2025 Stock** page.

**Step 2** — Upload &nbsp;→&nbsp; **Step 3** — Generate &nbsp;→&nbsp; **Step 4** — Download HTML, open in any browser.

Each dashboard page has a **↓ CSV** button to export the current filtered view.
    """)

st.markdown("### 📂 Upload Inventory Excel File")

uploaded_file = st.file_uploader(
    "Drop your Excel file here (.xlsx / .xls)",
    type=["xlsx", "xls"],
    label_visibility="visible",
)

if uploaded_file:
    sz = len(uploaded_file.getvalue()) / 1024 / 1024
    c1, c2, c3 = st.columns(3)
    c1.metric("📄 File",   uploaded_file.name[:28] + ("…" if len(uploaded_file.name) > 28 else ""))
    c2.metric("📦 Size",   f"{sz:.1f} MB")
    c3.metric("🔖 Format", uploaded_file.name.split('.')[-1].upper())

    st.markdown("<hr style='border-color:#30363D;margin:16px 0'>", unsafe_allow_html=True)

    # ── Column check: run once per file, show live step status ──
    file_key = f"col_check_{uploaded_file.name}_{sz:.1f}"
    if st.session_state.get('_col_check_key') != file_key:
        st.markdown("""
<div style="background:#161B22;border:1px solid #30363D;border-radius:8px;padding:14px 18px;margin:8px 0">
  <div style="font-size:12px;font-weight:600;color:#8B949E;letter-spacing:.06em;margin-bottom:10px">COLUMN VALIDATION</div>
  <div id="chk-steps" style="display:flex;flex-direction:column;gap:6px;font-size:13px;color:#C9D1D9"></div>
</div>""", unsafe_allow_html=True)

        chk_area = st.empty()

        def _chk_step(msg, status="running"):
            icon = {"running":"⏳","ok":"✅","warn":"⚠️","err":"❌"}.get(status,"⏳")
            col  = {"running":"#8B949E","ok":"#3FB950","warn":"#D29922","err":"#F85149"}.get(status,"#8B949E")
            chk_area.markdown(
                f'<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #21262D">' 
                f'<span>{icon}</span><span style="color:{col}">{msg}</span></div>',
                unsafe_allow_html=True)

        try:
            import pandas as pd
            import time

            _chk_step("Reading file headers…", "running")
            file_bytes_val = uploaded_file.getvalue()
            engine = None
            for eng in ('calamine', 'openpyxl', 'xlrd'):
                try:
                    bio = BytesIO(file_bytes_val)
                    test = pd.read_excel(bio, header=0, nrows=1, engine=eng)
                    if len(test.columns) > 0:
                        engine = eng; break
                except Exception:
                    pass

            if engine:
                _chk_step(f"Opened with engine: {engine}", "ok")
                df_h = pd.read_excel(BytesIO(file_bytes_val), header=0, nrows=1, engine=engine)
                raw_cols = list(df_h.columns)
            else:
                _chk_step("Using built-in XML reader…", "ok")
                df_h, _ = _read_excel_stdlib(file_bytes_val)
                raw_cols = list(df_h.columns)

            _chk_step(f"Found {len(raw_cols)} columns — checking required fields…", "running")
            norm_cols = [normalise_col(c) for c in raw_cols]
            missing_req_chk, missing_opt_chk = check_missing_columns(norm_cols)
            month_preview_chk = detect_months(norm_cols)

            if missing_req_chk:
                _chk_step(f"Missing {len(missing_req_chk)} required column(s) — see below", "err")
            else:
                _chk_step("All required columns present ✓", "ok")

            if missing_opt_chk:
                _chk_step(f"Optional columns not found (ok to proceed): {', '.join(missing_opt_chk)}", "warn")

            if month_preview_chk:
                _chk_step(f"{len(month_preview_chk)} monthly columns detected: {', '.join(month_preview_chk[:4])}{'…' if len(month_preview_chk)>4 else ''}", "ok")
            else:
                _chk_step("No monthly columns found — trend chart will be empty", "warn")

            st.session_state['_col_check_key']    = file_key
            st.session_state['_missing_req']       = missing_req_chk
            st.session_state['_missing_opt']       = missing_opt_chk
            st.session_state['_month_preview']     = month_preview_chk
            st.session_state['_col_check_ok']      = True

        except Exception as e:
            _chk_step(f"Validation error: {e}", "err")
            st.session_state['_col_check_key']  = file_key
            st.session_state['_missing_req']    = []
            st.session_state['_missing_opt']    = []
            st.session_state['_month_preview']  = []
            st.session_state['_col_check_ok']   = False
            st.session_state['_col_check_err']  = str(e)

    # ── Show persisted check results (on reruns after Generate) ──
    else:
        missing_req   = st.session_state.get('_missing_req', [])
        missing_opt   = st.session_state.get('_missing_opt', [])
        month_preview = st.session_state.get('_month_preview', [])
        col_check_ok  = st.session_state.get('_col_check_ok', False)
        col_check_err = st.session_state.get('_col_check_err', '')

        lines = []
        if not col_check_ok and col_check_err:
            lines.append(f'<div style="color:#F85149">❌ Validation error: {col_check_err}</div>')
        elif col_check_ok:
            if missing_req:
                lines.append(f'<div style="color:#F85149">❌ Missing {len(missing_req)} required column(s) — re-upload a corrected file</div>')
            else:
                lines.append('<div style="color:#3FB950">✅ All required columns present</div>')
            if missing_opt:
                lines.append(f'<div style="color:#D29922">⚠️ Optional missing: {', '.join(missing_opt)}</div>')
            if month_preview:
                lines.append(f'<div style="color:#79C0FF">📅 {len(month_preview)} monthly columns</div>')
            else:
                lines.append('<div style="color:#D29922">⚠️ No monthly columns</div>')
        if lines:
            st.markdown(
                '<div style="background:#161B22;border:1px solid #30363D;border-radius:8px;padding:12px 16px;margin:8px 0;font-size:13px">' +
                '<div style="font-size:12px;font-weight:600;color:#8B949E;letter-spacing:.06em;margin-bottom:8px">COLUMN VALIDATION</div>' +
                ''.join(lines) + '</div>',
                unsafe_allow_html=True)

    missing_req   = st.session_state.get('_missing_req', [])
    missing_opt   = st.session_state.get('_missing_opt', [])
    month_preview = st.session_state.get('_month_preview', [])

    if missing_req:
        st.error("❌ **Missing required columns** — fix before generating:\n\n" +
                 "\n".join(f"- `{c}`" for c in missing_req))
        st.markdown("""
<div style="background:#2D1117;border:1px solid #6E3B3B;border-radius:8px;
     padding:12px 16px;color:#F85149;font-size:13px;margin:8px 0">
  ⛔ Fix the missing required columns above before generating.
</div>""", unsafe_allow_html=True)
    else:
        if st.button("⚡  Generate Dashboard", type="primary", use_container_width=True):

            # ── Live step-by-step generation status ──
            gen_box = st.empty()

            def _gen_step(steps_done):
                """Render the generation steps list, marking each step as done/running/pending."""
                steps = [
                    ("📂", "Reading Excel file"),
                    ("🔍", "Analysing inventory data"),
                    ("🧮", "Computing KPIs & segments"),
                    ("🏗️", "Building HTML dashboard"),
                    ("✅", "Done"),
                ]
                rows = []
                for i, (ico, lbl) in enumerate(steps):
                    if i < steps_done:
                        style = "color:#3FB950"
                        tick  = "✓"
                    elif i == steps_done:
                        style = "color:#79C0FF"
                        tick  = "⏳"
                    else:
                        style = "color:#484F58"
                        tick  = "○"
                    rows.append(
                        f'<div style="display:flex;align-items:center;gap:10px;padding:7px 0;'
                        f'border-bottom:1px solid #21262D;{style}">' 
                        f'<span style="font-size:16px">{ico}</span>'
                        f'<span style="flex:1;font-size:13px">{lbl}</span>'
                        f'<span style="font-size:12px;opacity:.8">{tick}</span></div>'
                    )
                gen_box.markdown(
                    '<div style="background:#161B22;border:1px solid #30363D;border-radius:8px;padding:14px 18px;margin:8px 0">' +
                    '<div style="font-size:12px;font-weight:600;color:#8B949E;letter-spacing:.06em;margin-bottom:10px">GENERATING DASHBOARD</div>' +
                    ''.join(rows) + '</div>',
                    unsafe_allow_html=True)

            progress = st.progress(0, text="Starting…")
            try:
                _gen_step(0); progress.progress(10, text="Reading Excel file…")
                file_bytes = uploaded_file.getvalue()

                _gen_step(1); progress.progress(25, text="Analysing inventory data…")
                df_p, mc_p = read_excel_df(file_bytes)

                _gen_step(2); progress.progress(45, text="Computing KPIs & segments…")
                K_p, *_    = build_kpis(df_p, mc_p)
                del df_p; gc.collect()

                st.markdown(
                    f'<div style="background:#0D2818;border:1px solid #1A7431;border-radius:8px;'
                    f'padding:12px 18px;font-size:13px;color:#C9D1D9;margin:6px 0">'
                    f'<span style="color:#3FB950;font-weight:600">Data scan complete</span> &nbsp;—&nbsp;'
                    f'<b style="color:#E6EDF3">{K_p["total_items"]:,}</b> SKUs &nbsp;·&nbsp;'
                    f'<span style="color:#F85149">{K_p["neg_count"]:,} negative</span> &nbsp;·&nbsp;'
                    f'<span style="color:#D29922">{K_p["zero_count"]:,} zero-sale</span> &nbsp;·&nbsp;'
                    f'<span style="color:#F85149">{K_p["oos_count"]:,} OOS</span> &nbsp;·&nbsp;'
                    f'<span style="color:#D29922">{K_p["uw_count"]:,} unwanted repurchases</span> &nbsp;·&nbsp;'
                    f'<span style="color:#D29922">{K_p["pre_count"]:,} pre-2025</span></div>',
                    unsafe_allow_html=True)

                _gen_step(3); progress.progress(55, text="Building HTML dashboard… (may take 30–60s for large files)")
                html_content = generate_html(file_bytes, source_filename=uploaded_file.name)

                _gen_step(4); progress.progress(100, text="Done!")
                st.success("✅  Dashboard generated successfully!")

                out_name = uploaded_file.name.rsplit('.', 1)[0] + '_dashboard.html'
                st.download_button(
                    "⬇️  Download Dashboard HTML",
                    data=html_content.encode('utf-8'),
                    file_name=out_name,
                    mime="text/html",
                    use_container_width=True,
                    type="primary",
                )
                st.caption(f"Open `{out_name}` in any browser. No internet required after download.")

                st.markdown("<hr style='border-color:#30363D;margin:16px 0'>", unsafe_allow_html=True)
                st.markdown("### 📊 Dashboard Preview")
                st.caption("Full interactive dashboard — all pages, filters, charts and CSV downloads.")
                st.components.v1.html(html_content, height=820, scrolling=True)

                st.markdown("<hr style='border-color:#30363D;margin:16px 0'>", unsafe_allow_html=True)
                st.markdown("### 📈 Summary Metrics")
                r1 = st.columns(4); r2 = st.columns(4)
                r1[0].metric("Total SKUs",       f"{K_p['total_items']:,}")
                r1[1].metric("Stock Value",       f"AED {K_p['total_sv']:,.0f}")
                r1[2].metric("Total Units Sold",  f"{K_p['total_ts']:,.0f}")
                r1[3].metric("Items In Stock",    f"{K_p['in_stock']:,}")
                r2[0].metric("Negative Stock",    f"{K_p['neg_count']:,}",  delta=f"-AED {K_p['neg_sv']:,.0f}",  delta_color="inverse")
                r2[1].metric("Zero Sale Stock",   f"{K_p['zero_count']:,}", delta=f"-AED {K_p['zero_sv']:,.0f}", delta_color="inverse")
                r2[2].metric("Pre-2025 Stock",    f"{K_p['pre_count']:,}",  delta=f"-AED {K_p['pre_sv']:,.0f}",  delta_color="inverse")
                r2[3].metric("Unwanted Repurch.", f"{K_p['uw_count']:,}",   delta=f"-AED {K_p['uw_sv']:,.0f}",   delta_color="inverse")

            except Exception as e:
                progress.empty()
                st.error(f"❌  Error: {e}")
                with st.expander("Technical details"):
                    import traceback; st.code(traceback.format_exc())

else:
    st.markdown("""
<div style="background:#161B22;border:2px dashed #30363D;border-radius:12px;
     padding:48px 32px;text-align:center;margin:20px 0">
  <div style="font-size:48px;margin-bottom:12px">📊</div>
  <div style="font-size:16px;font-weight:600;color:#E6EDF3;margin-bottom:6px">Upload your inventory Excel file</div>
  <div style="font-size:13px;color:#8B949E">Supports .xlsx from any ERP or POS system</div>
  <div style="font-size:12px;color:#484F58;margin-top:10px">Missing columns detected automatically before generating</div>
</div>
""", unsafe_allow_html=True)

st.markdown("<hr style='border-color:#30363D;margin:20px 0'>", unsafe_allow_html=True)
st.caption("AL MADINA GROUP Inventory Intelligence Portal  |  All data processed locally")
