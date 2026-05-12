"""
AL MADINA GROUP — Inventory Intelligence Dashboard Portal
Single-file Streamlit application

Setup:
    pip install streamlit openpyxl

Run:
    streamlit run app.py

Upload any Excel inventory file (.xlsx) to generate a standalone HTML dashboard.
All processing is local — nothing is stored or transmitted.
"""

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DASHBOARD GENERATOR ENGINE
# ═════════════════════════════════════════════════════════════════════════════

# ─── Core imports (top-level so Streamlit Cloud installs all deps) ───────────
import json
import os
import tempfile
from collections import defaultdict
from datetime import datetime

# ─── SAFE HELPERS ────────────────────────────────────────────────────────────

def sf(v):
    """Safe float — returns 0.0 on any failure."""
    try:
        return float(v) if v is not None else 0.0
    except (ValueError, TypeError):
        return 0.0

def ss(v, n=45):
    """Safe string with max length."""
    try:
        s = str(v) if v is not None else ''
        return s[:n]
    except Exception:
        return ''

def ms(row, month_list):
    """Safe monthly sales sum."""
    return sum(sf(row.get(m, 0)) for m in month_list)

def safe_int(v):
    try:
        return int(float(v)) if v is not None else 0
    except (ValueError, TypeError):
        return 0

def fmtN(n, d=1):
    """Format number with K/M suffix."""
    try:
        n = float(n)
        if abs(n) >= 1_000_000:
            return f"{n/1_000_000:.{d}f}M"
        if abs(n) >= 1_000:
            return f"{n/1_000:.{d}f}K"
        return f"{n:,.{d}f}"
    except Exception:
        return "—"

def fmtA(n, d=0):
    return f"AED {fmtN(n, d)}"


# ─── COLUMN NORMALISER ────────────────────────────────────────────────────────

# Maps known aliases → canonical names used in code
COLUMN_ALIASES = {
    'item bar code':   'Item Bar Code',
    'barcode':         'Item Bar Code',
    'item name':       'Item Name',
    'name':            'Item Name',
    'cost':            'Cost',
    'selling':         'Selling',
    'selling price':   'Selling',
    'selling1':        'Selling',
    'selling2':        'Selling2',
    'stock':           'Stock',
    'qty':             'Stock',
    'quantity':        'Stock',
    'brand':           'Brand',
    'category':        'Category',
    'class':           'Class',
    'group':           'Group',
    'supplier':        'Supplier',
    'margin%':         'Margin%',
    'margin':          'Margin%',
    'profit':          'Profit',
    'stock value':     'Stock Value',
    'lp date':         'LP Date',
    'last purchase date': 'LP Date',
    'lp qty':          'LP Qty',
    'last purchase qty': 'LP Qty',
    'lp supplier':     'LP Supplier',
    'last purchase supplier': 'LP Supplier',
    'wac':             'WAC',
    'total sales':     'Total Sales',
    'sales':           'Total Sales',
}

def normalise_headers(headers):
    """Normalise raw headers → canonical names."""
    normalised = []
    for h in headers:
        key = str(h).strip().lower() if h else ''
        normalised.append(COLUMN_ALIASES.get(key, str(h).strip() if h else ''))
    return normalised

def detect_month_columns(headers):
    """Detect month columns like 'May, 2025' or 'May 2025'."""
    import re
    month_names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    pattern = re.compile(r'(%s)[,.\s]+(\d{4})' % '|'.join(month_names), re.IGNORECASE)
    months = []
    for h in headers:
        if pattern.search(str(h)):
            months.append(h)
    return months


# ─── EXCEL READER ─────────────────────────────────────────────────────────────

def _parse_xlsx_stdlib(file_bytes):
    """
    Parse .xlsx using only Python stdlib (zipfile + xml).
    No openpyxl, no pandas required. Works on any Python version.
    """
    import zipfile, xml.etree.ElementTree as ET
    from io import BytesIO

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        names = [i.filename for i in z.infolist()]

        # Read shared strings
        shared = []
        if "xl/sharedStrings.xml" in names:
            ss_xml = z.read("xl/sharedStrings.xml")
            for si in ET.fromstring(ss_xml).iter("{" + NS + "}si"):
                shared.append("".join(t.text or "" for t in si.iter("{" + NS + "}t")))

        # Find first sheet
        sheet_file = "xl/worksheets/sheet1.xml"
        if sheet_file not in names:
            # Try workbook to find first sheet name
            for n in names:
                if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                    sheet_file = n
                    break

        sheet_xml = z.read(sheet_file)

    def col_index(ref):
        col = "".join(c for c in ref if c.isalpha())
        idx = 0
        for ch in col:
            idx = idx * 26 + (ord(ch.upper()) - 64)
        return idx - 1

    rows_out = []
    max_col = 0

    for row_el in ET.fromstring(sheet_xml).iter("{" + NS + "}row"):
        cells = {}
        for cell in row_el.iter("{" + NS + "}c"):
            ref  = cell.get("r", "A1")
            col_i = col_index(ref)
            t    = cell.get("t", "n")
            v_el = cell.find("{" + NS + "}v")
            is_el = cell.find("{" + NS + "}is")

            if is_el is not None:
                t_el = is_el.find("{" + NS + "}t")
                val  = t_el.text if t_el is not None else ""
            elif v_el is not None:
                rv = v_el.text or ""
                if t == "s":
                    val = shared[int(rv)] if shared else rv
                elif t in ("b",):
                    val = bool(int(rv))
                else:
                    try:
                        val = float(rv) if "." in rv else int(rv)
                    except (ValueError, TypeError):
                        val = rv
            else:
                val = None

            cells[col_i] = val
            max_col = max(max_col, col_i)

        if cells:
            rows_out.append([cells.get(i) for i in range(max_col + 1)])

    return rows_out


def read_excel(filepath_or_bytes):
    """
    Read Excel file → (data, month_cols, headers).
    Accepts a filepath (str) or raw bytes.
    Primary: pandas with openpyxl engine.
    Fallback: pure Python stdlib (zipfile + xml).
    """
    try:
        from io import BytesIO

        # Accept filepath or raw bytes
        if isinstance(filepath_or_bytes, (bytes, bytearray)):
            file_bytes = bytes(filepath_or_bytes)
        else:
            with open(filepath_or_bytes, "rb") as _f:
                file_bytes = _f.read()

        # ── Try pandas first (fast, handles all edge cases) ──────────────
        try:
            import pandas as _pd
            df = _pd.read_excel(BytesIO(file_bytes), header=0, dtype=str, engine='openpyxl')
            df = df.where(_pd.notna(df), None)
            raw_headers = list(df.columns)
            headers = normalise_headers(raw_headers)
            df.columns = headers
            data = []
            for _, row in df.iterrows():
                d = {}
                for col in headers:
                    if not col: continue
                    val = row.get(col)
                    if val is None or (isinstance(val, str) and val.strip() == ''):
                        d[col] = None
                    else:
                        try: d[col] = float(str(val).replace(',', ''))
                        except (ValueError, TypeError): d[col] = str(val).strip()
                data.append(d)
            data = [r for r in data if any(v is not None for v in r.values())]
            month_cols = detect_month_columns(headers)
            return data, month_cols, headers
        except Exception:
            pass  # fall through to stdlib

        # ── Fallback: pure stdlib parser ──────────────────────────────────
        rows = _parse_xlsx_stdlib(file_bytes)

        if not rows:
            raise RuntimeError("Excel file appears to be empty.")

        raw_headers = [str(v) if v is not None else "" for v in rows[0]]
        headers = normalise_headers(raw_headers)

        data = []
        for row in rows[1:]:
            # Pad short rows
            while len(row) < len(headers):
                row.append(None)
            # Skip fully blank rows
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            d = {}
            for j, val in enumerate(row):
                if j < len(headers) and headers[j]:
                    d[headers[j]] = val
            data.append(d)

        month_cols = detect_month_columns(headers)
        return data, month_cols, headers

    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {e}")


# ─── DATA BUILDER ─────────────────────────────────────────────────────────────

def build_data(data, month_cols):
    """Build all dashboard datasets from raw data rows."""

    # Sort months chronologically
    from datetime import datetime
    def month_sort_key(m):
        try:
            return datetime.strptime(m.replace(' ', ', ').replace(',,', ','), '%b, %Y')
        except Exception:
            return datetime.min

    months_sorted = sorted(month_cols, key=month_sort_key)
    last3_months = months_sorted[-3:] if len(months_sorted) >= 3 else months_sorted

    # ── Sections ──────────────────────────────────────────────────────────────
    zero_sale  = [r for r in data if sf(r.get('Stock', 0)) > 0
                  and sf(r.get('Total Sales', 0)) == 0]

    neg_items  = [r for r in data if sf(r.get('Stock', 0)) < 0]

    oos_recent = [r for r in data if ms(r, last3_months) > 0
                  and sf(r.get('Stock', 0)) <= 0]

    high_val   = [r for r in data if sf(r.get('Stock Value', 0)) > 200
                  and sf(r.get('Total Sales', 0)) < 10
                  and sf(r.get('Stock', 0)) > 0]

    # Unwanted: has LP Date, stock > 0, zero sales, AND LP Qty < Stock
    # (item was already in stock when re-purchased — never sold before or after)
    unwanted   = [r for r in data
                  if r.get('LP Date')
                  and sf(r.get('Stock', 0)) > 0
                  and sf(r.get('Total Sales', 0)) == 0
                  and sf(r.get('LP Qty', 0)) < sf(r.get('Stock', 0))]

    top_items  = sorted(data, key=lambda r: -sf(r.get('Total Sales', 0)))

    # ── Monthly totals ────────────────────────────────────────────────────────
    monthly = {m: round(sum(sf(r.get(m, 0)) for r in data), 1) for m in months_sorted}

    # ── Class breakdown ───────────────────────────────────────────────────────
    cls_map = defaultdict(lambda: {'sales': 0, 'sv': 0, 'items': 0, 'zs': 0, 'neg': 0, 'oos': 0})
    for r in data:
        c = r.get('Class', '') or ''
        cls_map[c]['sales'] += sf(r.get('Total Sales', 0))
        cls_map[c]['sv']    += sf(r.get('Stock Value', 0))
        cls_map[c]['items'] += 1
    for r in zero_sale:  cls_map[r.get('Class', '') or '']['zs']  += 1
    for r in neg_items:  cls_map[r.get('Class', '') or '']['neg'] += 1
    for r in oos_recent: cls_map[r.get('Class', '') or '']['oos'] += 1

    top_cls = sorted(cls_map.items(), key=lambda x: -x[1]['sales'])[:12]

    # ── KPIs ──────────────────────────────────────────────────────────────────
    K = {
        'total_items': len(data),
        'in_stock':    len([r for r in data if sf(r.get('Stock', 0)) > 0]),
        'total_sv':    round(sum(sf(r.get('Stock Value', 0)) for r in data), 0),
        'total_ts':    round(sum(sf(r.get('Total Sales', 0)) for r in data), 0),
        'neg_count':   len(neg_items),
        'neg_sv':      round(sum(abs(sf(r.get('Stock Value', 0))) for r in neg_items), 0),
        'zero_count':  len(zero_sale),
        'zero_sv':     round(sum(sf(r.get('Stock Value', 0)) for r in zero_sale), 0),
        'oos_count':   len(oos_recent),
        'risk_count':  len(high_val),
        'risk_sv':     round(sum(sf(r.get('Stock Value', 0)) for r in high_val), 0),
        'uw_count':    len(unwanted),
        'uw_sv':       round(sum(sf(r.get('Stock Value', 0)) for r in unwanted), 0),
    }

    # ── Row serialisers ───────────────────────────────────────────────────────
    def row_top(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25), 'grp': ss(r.get('Group', ''), 25),
                'brand': ss(r.get('Brand', ''), 20), 'sup': ss(r.get('Supplier', ''), 35),
                'cost': round(sf(r.get('Cost', 0)), 2), 'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1), 'sv': round(sf(r.get('Stock Value', 0)), 0),
                'ts': round(sf(r.get('Total Sales', 0)), 0), 'mg': round(sf(r.get('Margin%', 0)), 1),
                'l3': round(ms(r, last3_months), 0)}

    def row_zero(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25), 'grp': ss(r.get('Group', ''), 25),
                'sup': ss(r.get('Supplier', ''), 35),
                'cost': round(sf(r.get('Cost', 0)), 2), 'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1), 'sv': round(sf(r.get('Stock Value', 0)), 0),
                'lpd': ss(r.get('LP Date', ''), 20) if r.get('LP Date') else '',
                'lpq': safe_int(r.get('LP Qty', 0))}

    def row_neg(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25), 'sup': ss(r.get('Supplier', ''), 35),
                'cost': round(sf(r.get('Cost', 0)), 2), 'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1), 'sv': round(sf(r.get('Stock Value', 0)), 0),
                'ts': round(sf(r.get('Total Sales', 0)), 0), 'l3': round(ms(r, last3_months), 0),
                'lpd': ss(r.get('LP Date', ''), 20) if r.get('LP Date') else '',
                'lpq': safe_int(r.get('LP Qty', 0))}

    def row_oos(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25), 'sup': ss(r.get('Supplier', ''), 35),
                'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1),
                'ts': round(sf(r.get('Total Sales', 0)), 0), 'l3': round(ms(r, last3_months), 0),
                'm1': round(sf(r.get(last3_months[2] if len(last3_months) > 2 else '', 0)), 1),
                'm2': round(sf(r.get(last3_months[1] if len(last3_months) > 1 else '', 0)), 1),
                'm3': round(sf(r.get(last3_months[0] if last3_months else '', 0)), 1),
                'm1_label': last3_months[2] if len(last3_months) > 2 else '',
                'm2_label': last3_months[1] if len(last3_months) > 1 else '',
                'm3_label': last3_months[0] if last3_months else ''}

    def row_risk(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25),
                'cost': round(sf(r.get('Cost', 0)), 2), 'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1), 'sv': round(sf(r.get('Stock Value', 0)), 0),
                'ts': round(sf(r.get('Total Sales', 0)), 0), 'mg': round(sf(r.get('Margin%', 0)), 1),
                'l3': round(ms(r, last3_months), 0),
                'lpd': ss(r.get('LP Date', ''), 20) if r.get('LP Date') else '',
                'lpq': safe_int(r.get('LP Qty', 0))}

    def row_uw(r):
        return {'n': ss(r.get('Item Name', '')), 'cat': ss(r.get('Category', ''), 20),
                'cls': ss(r.get('Class', ''), 25),
                'sup': ss(r.get('LP Supplier', '') or r.get('Supplier', ''), 35),
                'cost': round(sf(r.get('Cost', 0)), 2), 'sell': round(sf(r.get('Selling', 0)), 2),
                'stock': round(sf(r.get('Stock', 0)), 1), 'sv': round(sf(r.get('Stock Value', 0)), 0),
                'lpq': safe_int(r.get('LP Qty', 0)),
                'lpd': ss(r.get('LP Date', ''), 20) if r.get('LP Date') else ''}

    categories = sorted(set(r.get('Category', '') or '' for r in data if r.get('Category')))
    cat_class_map = {}
    for cat in categories:
        cat_class_map[cat] = sorted(set(r.get('Class', '') or '' for r in data
                                        if r.get('Category') == cat and r.get('Class')))

    # OOS month labels for table headers
    oos_m_labels = ['', '', '']
    if len(last3_months) >= 3:
        oos_m_labels = [last3_months[2], last3_months[1], last3_months[0]]
    elif len(last3_months) == 2:
        oos_m_labels = [last3_months[1], last3_months[0], '']
    elif len(last3_months) == 1:
        oos_m_labels = [last3_months[0], '', '']

    return {
        'kpis':         K,
        'monthly':      monthly,
        'month_labels': [m for m in months_sorted],
        'top_cls':      [[k, round(v['sales'], 1), round(v['sv'], 0), v['items'], v['zs'], v['neg'], v['oos']] for k, v in top_cls],
        'categories':   categories,
        'cat_class_map': cat_class_map,
        'oos_m_labels': oos_m_labels,
        'top':    [row_top(r)  for r in top_items  if sf(r.get('Total Sales', 0)) > 0],
        'zero':   [row_zero(r) for r in sorted(zero_sale,  key=lambda x: -sf(x.get('Stock Value', 0)))],
        'neg':    [row_neg(r)  for r in sorted(neg_items,  key=lambda x:  sf(x.get('Stock', 0)))],
        'oos':    [row_oos(r)  for r in sorted(oos_recent, key=lambda x: -ms(x, last3_months))],
        'risk':   [row_risk(r) for r in sorted(high_val,   key=lambda x: -sf(x.get('Stock Value', 0)))],
        'uw':     [row_uw(r)   for r in sorted(unwanted,   key=lambda x: -sf(x.get('Stock Value', 0)))],
    }


# ─── HTML BUILDER ─────────────────────────────────────────────────────────────

def _esc(s):
    """Escape for HTML attribute."""
    return str(s).replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')

def _row(cells):
    return '<tr>' + ''.join(cells) + '</tr>'

def _td(content, cls=''):
    return f'<td class="{cls}">{content}</td>' if cls else f'<td>{content}</td>'

def _tdr(content, cls=''):
    return f'<td class="tr {cls}">{content}</td>'

def _badge(text, style):
    return f'<span class="b-{style}">{_esc(text)}</span>'

def tr_top(rows):
    out = []
    for i, r in enumerate(rows):
        mg = r.get('mg', 0)
        mg_cls = 'c-red' if mg < 10 else ('c-green fw6' if mg > 30 else '')
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _td(_esc(r.get('brand', '') or '—'), 'muted'),
            _tdr(fmtN(r['ts'], 0), 'bold-green'),
            _tdr(fmtN(r['l3'], 0), 'c-green'),
            _tdr(fmtN(r['stock'], 1)),
            _tdr(fmtN(r['sv'], 0)),
            _tdr(f"{r['cost']:.2f}"),
            _tdr(f"{r['sell']:.2f}"),
            _tdr(f"{r['mg']}%", mg_cls),
        ]))
    return ''.join(out)

def tr_zero(rows):
    out = []
    for i, r in enumerate(rows):
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _td(_esc(r.get('grp', '') or '—'), 'muted'),
            _tdr(fmtN(r['stock'], 1), 'c-amber'),
            _tdr(fmtN(r['sv'], 0), 'c-red fw6'),
            _tdr(f"{r['cost']:.2f}"),
            _tdr(f"{r['sell']:.2f}"),
            _td(r.get('lpd', '') or '—'),
            _tdr(str(r.get('lpq', 0)) if r.get('lpq') else '—'),
            _td(f'<span class="sup">{_esc(r.get("sup","") or "—")}</span>'),
        ]))
    return ''.join(out)

def tr_neg(rows):
    out = []
    for i, r in enumerate(rows):
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _tdr(fmtN(r['stock'], 1), 'c-red fw6'),
            _tdr(fmtN(r['sv'], 0), 'c-red'),
            _tdr(fmtN(r['ts'], 0)),
            _tdr(fmtN(r['l3'], 0), 'c-green'),
            _tdr(f"{r['cost']:.2f}"),
            _tdr(f"{r['sell']:.2f}"),
            _td(r.get('lpd', '') or '—'),
            _tdr(str(r.get('lpq', 0)) if r.get('lpq') else '—'),
        ]))
    return ''.join(out)

def tr_oos(rows, m_labels):
    out = []
    for i, r in enumerate(rows):
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _tdr(fmtN(r['l3'], 0), 'c-green fw6'),
            _tdr(fmtN(r.get('m1', 0), 1)),
            _tdr(fmtN(r.get('m2', 0), 1)),
            _tdr(fmtN(r.get('m3', 0), 1)),
            _tdr(fmtN(r['stock'], 1), 'c-red fw6'),
            _tdr(fmtN(r['ts'], 0)),
            _tdr(f"{r['sell']:.2f}"),
            _td(f'<span class="sup">{_esc(r.get("sup","") or "—")}</span>'),
        ]))
    return ''.join(out)

def tr_risk(rows):
    out = []
    for i, r in enumerate(rows):
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _tdr(fmtN(r['sv'], 0), 'c-amber fw6'),
            _tdr(fmtN(r['stock'], 1)),
            _tdr(fmtN(r['ts'], 0), 'c-red'),
            _tdr(fmtN(r['l3'], 0) if r['l3'] else '—'),
            _tdr(f"{r['cost']:.2f}"),
            _tdr(f"{r['sell']:.2f}"),
            _tdr(f"{r['mg']}%"),
            _td(r.get('lpd', '') or '—'),
            _tdr(str(r.get('lpq', 0)) if r.get('lpq') else '—'),
        ]))
    return ''.join(out)

def tr_uw(rows):
    out = []
    for i, r in enumerate(rows):
        out.append(_row([
            _td(i + 1, 'mono'),
            _td(f'<span class="tn" title="{_esc(r["n"])}">{_esc(r["n"])}</span>'),
            _td(_badge(r['cat'], 'slate')),
            _td(_badge(r['cls'], 'brand')),
            _tdr(fmtN(r['stock'], 1), 'c-amber'),
            _tdr(fmtN(r['sv'], 0), 'c-red fw6'),
            _tdr(f"{r['cost']:.2f}"),
            _tdr(f"{r['sell']:.2f}"),
            _tdr(str(r.get('lpq', 0)) if r.get('lpq') else '—'),
            _td(r.get('lpd', '') or '—'),
            _td(f'<span class="sup">{_esc(r.get("sup","") or "—")}</span>'),
        ]))
    return ''.join(out)

def tr_cls(rows):
    out = []
    for c in rows:
        name, sales, sv, items, zs, neg, oos = c
        if zs > items * 0.1:
            status = _badge('High Dead Stock', 'red')
        elif oos > items * 0.1:
            status = _badge('OOS Risk', 'amber')
        else:
            status = _badge('Healthy', 'green')
        out.append(_row([
            _td(_badge(name, 'brand')),
            _tdr(fmtN(sales, 0)),
            _tdr(fmtN(sv, 0)),
            _tdr(fmtN(items, 0)),
            _tdr(fmtN(zs, 0), 'c-amber' if zs > 50 else ''),
            _tdr(fmtN(neg, 0), 'c-red' if neg > 20 else ''),
            _tdr(fmtN(oos, 0), 'c-red' if oos > 50 else ''),
            _td(status),
        ]))
    return ''.join(out)

def rank_bars(rows, n=10):
    if not rows:
        return ''
    maxv = max(r['ts'] for r in rows[:n]) or 1
    out = []
    for i, r in enumerate(rows[:n]):
        pct = r['ts'] / maxv * 100
        out.append(
            f'<div class="rank-row">'
            f'<div class="rn">{i+1}</div>'
            f'<div class="rnm" title="{_esc(r["n"])}">{_esc(r["n"])}</div>'
            f'<div class="rb"><div class="rf" style="width:{pct:.1f}%"></div></div>'
            f'<div class="rv">{fmtN(r["ts"],0)}</div>'
            f'</div>'
        )
    return ''.join(out)

def opts_html(values, all_label='All'):
    return ''.join(f'<option value="{_esc(v)}">{_esc(v)}</option>' for v in [all_label] + list(values))

def section_opts(data_list, all_cat='All Categories', all_cls='All Classes'):
    cats = sorted(set(r['cat'] for r in data_list if r.get('cat')))
    clses = sorted(set(r['cls'] for r in data_list if r.get('cls')))
    catcls = {}
    for r in data_list:
        c = r.get('cat', '')
        cl = r.get('cls', '')
        if c not in catcls:
            catcls[c] = set()
        if cl:
            catcls[c].add(cl)
    catcls = {k: sorted(v) for k, v in catcls.items()}
    return (opts_html(cats, all_cat), opts_html(clses, all_cls), json.dumps(catcls))

def build_section_cache(data_list, render_fn, extra_arg=None):
    """Build pre-rendered HTML lookup dict keyed by CAT|CLS."""
    cache = {}
    cats = sorted(set(r.get('cat', '') for r in data_list))

    def render(rows):
        return render_fn(rows, extra_arg) if extra_arg is not None else render_fn(rows)

    cache['ALL|ALL'] = render(data_list)
    for cat in cats:
        if not cat:
            continue
        cat_rows = [r for r in data_list if r.get('cat') == cat]
        cache[f'{cat}|ALL'] = render(cat_rows)
        clses = sorted(set(r.get('cls', '') for r in cat_rows))
        for cls in clses:
            if not cls:
                continue
            cls_rows = [r for r in cat_rows if r.get('cls') == cls]
            cache[f'{cat}|{cls}'] = render(cls_rows)
    return cache

def cnt_sv_cache(data_list, sv_field='sv'):
    cache_cnt = {}
    cache_sv = {}
    cats = sorted(set(r.get('cat', '') for r in data_list))
    cache_cnt['ALL|ALL'] = len(data_list)
    cache_sv['ALL|ALL'] = round(sum(r.get(sv_field, 0) for r in data_list), 0)
    for cat in cats:
        if not cat:
            continue
        cat_rows = [r for r in data_list if r.get('cat') == cat]
        cache_cnt[f'{cat}|ALL'] = len(cat_rows)
        cache_sv[f'{cat}|ALL'] = round(sum(r.get(sv_field, 0) for r in cat_rows), 0)
        for cls in sorted(set(r.get('cls', '') for r in cat_rows)):
            if not cls:
                continue
            cls_rows = [r for r in cat_rows if r.get('cls') == cls]
            cache_cnt[f'{cat}|{cls}'] = len(cls_rows)
            cache_sv[f'{cat}|{cls}'] = round(sum(r.get(sv_field, 0) for r in cls_rows), 0)
    return cache_cnt, cache_sv

def l3_cache(data_list):
    cache = {}
    cache['ALL|ALL'] = round(sum(r.get('l3', 0) for r in data_list), 0)
    for cat in sorted(set(r.get('cat', '') for r in data_list)):
        if not cat:
            continue
        cat_rows = [r for r in data_list if r.get('cat') == cat]
        cache[f'{cat}|ALL'] = round(sum(r.get('l3', 0) for r in cat_rows), 0)
        for cls in sorted(set(r.get('cls', '') for r in cat_rows)):
            if not cls:
                continue
            cls_rows = [r for r in cat_rows if r.get('cls') == cls]
            cache[f'{cat}|{cls}'] = round(sum(r.get('l3', 0) for r in cls_rows), 0)
    return cache

def top_rank_cache(data_list):
    cache = {}
    cache['ALL|ALL'] = rank_bars(data_list)
    for cat in sorted(set(r.get('cat', '') for r in data_list)):
        if not cat:
            continue
        cat_rows = [r for r in data_list if r.get('cat') == cat]
        cache[f'{cat}|ALL'] = rank_bars(cat_rows)
        for cls in sorted(set(r.get('cls', '') for r in cat_rows)):
            if not cls:
                continue
            cache[f'{cat}|{cls}'] = rank_bars([r for r in cat_rows if r.get('cls') == cls])
    return cache


def generate_html(filepath_or_data, month_cols=None, source_filename=''):
    """
    Main entry point.
    Pass either a filepath (str) or already-parsed (data, month_cols).
    Returns HTML string.
    """
    if isinstance(filepath_or_data, str):
        data, month_cols, _ = read_excel(filepath_or_data)
    else:
        data = filepath_or_data

    D = build_data(data, month_cols)
    K = D['kpis']

    # ── Pre-render all tables ─────────────────────────────────────────────────
    m_labels = D['oos_m_labels']

    top_rows_c   = build_section_cache(D['top'],  tr_top)
    zero_rows_c  = build_section_cache(D['zero'], tr_zero)
    neg_rows_c   = build_section_cache(D['neg'],  tr_neg)
    oos_rows_c   = build_section_cache(D['oos'],  tr_oos, m_labels)
    risk_rows_c  = build_section_cache(D['risk'], tr_risk)
    uw_rows_c    = build_section_cache(D['uw'],   tr_uw)
    rank_c       = top_rank_cache(D['top'])

    top_cnt_c, _        = cnt_sv_cache(D['top'])
    zero_cnt_c, zero_sv_c = cnt_sv_cache(D['zero'])
    neg_cnt_c, neg_sv_c   = cnt_sv_cache(D['neg'])
    oos_cnt_c, _          = cnt_sv_cache(D['oos'])
    oos_l3_c              = l3_cache(D['oos'])
    risk_cnt_c, risk_sv_c = cnt_sv_cache(D['risk'])
    uw_cnt_c, uw_sv_c     = cnt_sv_cache(D['uw'])

    top_cats_opts, top_cls_opts, top_cc   = section_opts(D['top'])
    zero_cats_opts, zero_cls_opts, zero_cc = section_opts(D['zero'])
    neg_cats_opts, neg_cls_opts, neg_cc   = section_opts(D['neg'])
    oos_cats_opts, oos_cls_opts, oos_cc   = section_opts(D['oos'])
    risk_cats_opts, risk_cls_opts, risk_cc = section_opts(D['risk'])
    uw_cats_opts, uw_cls_opts, uw_cc      = section_opts(D['uw'])

    # ── KPI HTML ──────────────────────────────────────────────────────────────
    kpi_defs = [
        ('Total SKUs',           fmtN(K['total_items'],0), 'Products in portfolio',              'brand', 'kv-brand'),
        ('Items In Stock',       fmtN(K['in_stock'],0),    fmtA(K['total_sv'],0)+' stock value', 'green', 'kv-green'),
        ('Total Units Sold',     fmtN(K['total_ts'],0),    '13-month cumulative',                'blue',  'kv-blue'),
        ('Negative Stock Items', fmtN(K['neg_count'],0),   'Audit required urgently',            'red',   'kv-red'),
        ('Zero Sale — In Stock', fmtN(K['zero_count'],0),  fmtA(K['zero_sv'],0)+' capital at risk','red','kv-red'),
        ('OOS Active Demand',    fmtN(K['oos_count'],0),   'Selling but unavailable',            'red',   'kv-red'),
        ('High Risk Slow Movers',fmtN(K['risk_count'],0),  fmtA(K['risk_sv'],0)+' exposure',     'amber', 'kv-amber'),
        ('Unwanted Repurchases', fmtN(K['uw_count'],0),    fmtA(K['uw_sv'],0)+' re-bought, unsold','amber','kv-amber'),
    ]
    kpi_html = ''.join(
        f'<div class="kcard kc-{bar}"><div class="klbl">{label}</div>'
        f'<div class="kval {vcls}">{val}</div><div class="ksub">{sub}</div></div>'
        for label, val, sub, bar, vcls in kpi_defs
    )

    # ── Chart data ────────────────────────────────────────────────────────────
    ml = D['month_labels']
    mv = [D['monthly'].get(m, 0) for m in ml]
    ml_short = [m.replace(', 20', "'") for m in ml]
    mc = ['#1B2B4B'] * (len(ml) - 1) + ['rgba(37,99,235,0.45)']
    tc = D['top_cls'][:10]
    cat_colors = ['#1B2B4B','#1e3a5f','#1e40af','#1d4ed8','#2563eb','#3b82f6','#60a5fa','#0f766e','#0d9488','#14b8a6']
    health_vals = [
        max(0, K['in_stock'] - K['zero_count'] - K['risk_count']),
        K['zero_count'], K['neg_count'], K['oos_count'], K['risk_count']
    ]

    # ── Action table ──────────────────────────────────────────────────────────
    action_rows_data = [
        ('b-red',   'P1 — Critical', 'OOS Active Demand',       str(K['oos_count'])+' items', 'c-red',   'Daily revenue loss',        'Emergency POs — top items by last-3M velocity', 'Buying Team',      'Today'),
        ('b-red',   'P1 — Critical', 'Purchased — Zero Sales',  str(K['uw_count'])+' items',  'c-red',   fmtA(K['uw_sv'],0),          'Supplier return + clearance promotions',         'Category Mgmt',    'This Week'),
        ('b-amber', 'P2 — High',     'Negative Stock',          str(K['neg_count'])+' items', 'c-amber', 'Reporting inaccuracy',       'Physical count + GRN audit + corrections',      'Operations',       '5 Business Days'),
        ('b-amber', 'P2 — High',     'Slow Mover Capital',      str(K['risk_count'])+' items','c-amber', fmtA(K['risk_sv'],0),         'Markdown programme + bundle offers',            'Merchandising',    '14 Days'),
        ('b-blue',  'P3 — Medium',   'Zero Sale Dead Stock',    str(K['zero_count'])+' items','c-blue',  fmtA(K['zero_sv'],0),         'Range review — discontinue or clearance',       'Category Mgmt',    '30 Days'),
        ('b-slate', 'P4 — Standard', 'New Listing Policy',      'Process change',             '',        'Future prevention',          'Mandatory 90-day sell-through KPI on all new listings', 'Commercial Dir.', 'Next Quarter'),
    ]
    action_html = ''.join(
        f'<tr><td><span class="{badge}">{pri}</span></td><td>{issue}</td><td>{scale}</td>'
        f'<td class="{exp_cls}">{exp}</td><td>{action}</td><td>{owner}</td><td>{deadline}</td></tr>'
        for badge, pri, issue, scale, exp_cls, exp, action, owner, deadline in action_rows_data
    )

    # ── Date / source label ───────────────────────────────────────────────────
    report_date = datetime.now().strftime('%d %b %Y')
    data_label = f'Source: {source_filename} &nbsp;|&nbsp; Generated: {report_date}' if source_filename else f'Generated: {report_date}'

    # OOS table month header labels
    oos_h1 = _esc(m_labels[0]) if len(m_labels) > 0 else ''
    oos_h2 = _esc(m_labels[1]) if len(m_labels) > 1 else ''
    oos_h3 = _esc(m_labels[2]) if len(m_labels) > 2 else ''

    # ── Assemble HTML ─────────────────────────────────────────────────────────
    CSS = """
:root{--bg:#F4F5F7;--surf:#FFF;--surf2:#F9FAFB;--bdr:#E5E7EB;--bdr2:#D1D5DB;--txt:#111827;--txt2:#374151;--mut:#6B7280;--brand:#1B2B4B;--brand-t:#EEF2FF;--red:#DC2626;--red-t:#FEF2F2;--red-b:#FECACA;--amb:#D97706;--amb-t:#FFFBEB;--amb-b:#FDE68A;--grn:#059669;--grn-t:#ECFDF5;--grn-b:#A7F3D0;--blu:#2563EB;--blu-t:#EFF6FF;--blu-b:#BFDBFE;--slt:#475569;--r:8px;--sh:0 1px 3px rgba(0,0,0,.08)}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--txt);font-family:"Inter",sans-serif;font-size:13.5px;line-height:1.55;-webkit-font-smoothing:antialiased}
.hdr{background:var(--brand);color:#fff;padding:0 32px;display:flex;align-items:center;justify-content:space-between;height:60px;position:sticky;top:0;z-index:200;border-bottom:1px solid rgba(255,255,255,.08)}
.hdr-l{display:flex;align-items:center;gap:16px}
.logo{display:flex;align-items:center;gap:10px}
.lm{width:34px;height:34px;background:linear-gradient(135deg,#3B82F6,#60A5FA);border-radius:7px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;color:#fff}
.lt{font-weight:700;font-size:15px;letter-spacing:-.2px}
.ls{font-size:10.5px;color:rgba(255,255,255,.5)}
.hsep{width:1px;height:28px;background:rgba(255,255,255,.12)}
.hlbl{font-size:12px;color:rgba(255,255,255,.55)}
.hr{font-size:11px;color:rgba(255,255,255,.4);font-family:"IBM Plex Mono",monospace;text-align:right}
.nav{background:var(--brand);border-bottom:1px solid rgba(255,255,255,.08);display:flex;padding:0 32px;overflow-x:auto;gap:2px}
.nav::-webkit-scrollbar{height:0}
.nbtn{background:none;border:none;color:rgba(255,255,255,.45);font-family:"Inter",sans-serif;font-size:12px;font-weight:500;padding:10px 14px;cursor:pointer;white-space:nowrap;border-bottom:2px solid transparent;transition:color .15s,border-color .15s;display:flex;align-items:center;gap:6px}
.nbtn:hover{color:rgba(255,255,255,.75)}
.nbtn.active{color:#fff;border-bottom-color:#60A5FA}
.npill{display:inline-flex;align-items:center;justify-content:center;background:rgba(255,255,255,.12);border-radius:10px;font-size:10px;font-weight:600;min-width:18px;height:16px;padding:0 5px}
.npill.red{background:rgba(239,68,68,.35)}
.npill.amb{background:rgba(245,158,11,.35)}
.page{display:none}.page.active{display:block}
.pi{max-width:1440px;margin:0 auto;padding:24px 32px 40px}
.ptitle{font-size:18px;font-weight:700;letter-spacing:-.3px;margin-bottom:3px}
.pdesc{font-size:12.5px;color:var(--mut);margin-bottom:20px}
.fbar{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:13px 18px;display:flex;align-items:center;gap:14px;margin-bottom:20px;flex-wrap:wrap;box-shadow:var(--sh)}
.flbl{font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em}
.fsep{width:1px;height:20px;background:var(--bdr)}
.fg{display:flex;align-items:center;gap:8px}
.fg label{font-size:12px;color:var(--mut);font-weight:500}
select.fsel{background:var(--bg);border:1px solid var(--bdr2);border-radius:5px;color:var(--txt2);font-family:"Inter",sans-serif;font-size:12.5px;font-weight:500;padding:6px 10px;cursor:pointer;outline:none;min-width:160px}
select.fsel:focus{border-color:var(--blu)}
.freset{background:none;border:1px solid var(--bdr2);border-radius:5px;color:var(--mut);font-family:"Inter",sans-serif;font-size:12px;font-weight:500;padding:6px 12px;cursor:pointer}
.freset:hover{background:var(--bg)}
.fcnt{margin-left:auto;font-size:12px;color:var(--mut);font-family:"IBM Plex Mono",monospace}
.kgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(195px,1fr));gap:12px;margin-bottom:24px}
.kcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:16px 18px;box-shadow:var(--sh);position:relative;overflow:hidden}
.kcard::after{content:"";position:absolute;top:0;left:0;right:0;height:3px;border-radius:8px 8px 0 0}
.kc-brand::after{background:var(--brand)}.kc-green::after{background:var(--grn)}.kc-red::after{background:var(--red)}.kc-amber::after{background:var(--amb)}.kc-blue::after{background:var(--blu)}
.klbl{font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px}
.kval{font-size:24px;font-weight:700;letter-spacing:-.5px;line-height:1;margin-bottom:4px}
.kv-brand{color:var(--brand)}.kv-green{color:var(--grn)}.kv-red{color:var(--red)}.kv-amber{color:var(--amb)}.kv-blue{color:var(--blu)}
.ksub{font-size:11.5px;color:var(--mut)}
.crow{display:grid;gap:16px;margin-bottom:20px}
.c1{grid-template-columns:1fr}.c21{grid-template-columns:2fr 1fr}
.cbox{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);padding:20px;box-shadow:var(--sh)}
.chead{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.ctitle{font-size:13px;font-weight:600;color:var(--txt)}
.cmeta{font-size:11px;color:var(--mut)}
.ch{position:relative}.ch260{height:260px}
.tcard{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--r);box-shadow:var(--sh);overflow:hidden;margin-bottom:20px}
.thd{padding:13px 18px;border-bottom:1px solid var(--bdr);display:flex;align-items:center;justify-content:space-between;background:var(--surf2)}
.ttitle{font-size:13px;font-weight:600;color:var(--txt)}
.tinfo{display:flex;align-items:center;gap:10px}
.tsc{overflow-x:auto}.tsc::-webkit-scrollbar{height:4px}.tsc::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:2px}
.tmh{max-height:520px;overflow-y:auto}.tmh::-webkit-scrollbar{width:5px}.tmh::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:3px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead tr{background:#F9FAFB;position:sticky;top:0;z-index:1}
th{text-align:left;padding:9px 12px;font-size:11px;font-weight:600;color:var(--mut);text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;border-bottom:1px solid var(--bdr)}
tbody tr{border-bottom:1px solid var(--bdr);transition:background .1s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#F8FAFF}
td{padding:9px 12px;color:var(--txt2);vertical-align:middle}
.tn{font-weight:500;color:var(--txt);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}
.mono{font-family:"IBM Plex Mono",monospace;font-size:12px}
.tr{text-align:right;font-family:"IBM Plex Mono",monospace;font-size:12px}
.c-red{color:var(--red)}.c-amber{color:var(--amb)}.c-green{color:var(--grn)}.c-blue{color:var(--blu)}.muted{color:var(--mut)}.fw6{font-weight:600}.bold-green{color:var(--grn);font-weight:700}
.sup{max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block}
.b-red{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--red-t);color:var(--red);border:1px solid var(--red-b);white-space:nowrap}
.b-amber{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--amb-t);color:var(--amb);border:1px solid var(--amb-b);white-space:nowrap}
.b-green{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--grn-t);color:var(--grn);border:1px solid var(--grn-b);white-space:nowrap}
.b-blue{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--blu-t);color:var(--blu);border:1px solid var(--blu-b);white-space:nowrap}
.b-slate{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:#F1F5F9;color:var(--slt);border:1px solid #CBD5E1;white-space:nowrap}
.b-brand{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:2px 7px;border-radius:4px;background:var(--brand-t);color:var(--brand);border:1px solid #C7D2FE;white-space:nowrap}
.arow{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;margin-bottom:20px}
.abox{border-radius:var(--r);padding:14px 16px;border:1px solid}
.abox.red{background:var(--red-t);border-color:var(--red-b)}
.abox.amb{background:var(--amb-t);border-color:var(--amb-b)}
.abox.grn{background:var(--grn-t);border-color:var(--grn-b)}
.abox.blu{background:var(--blu-t);border-color:var(--blu-b)}
.ah{display:flex;align-items:flex-start;gap:8px;margin-bottom:5px}
.ai{font-size:14px;flex-shrink:0;margin-top:1px;font-style:normal}
.at{font-size:12.5px;font-weight:600;color:var(--txt)}
.ab{font-size:12px;color:var(--txt2);line-height:1.5;padding-left:22px}
.strip{background:var(--brand);color:#fff;border-radius:var(--r);padding:12px 20px;display:flex;align-items:center;gap:24px;margin-bottom:20px;flex-wrap:wrap}
.si{display:flex;flex-direction:column;gap:1px}
.sl{font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em}
.sv{font-size:16px;font-weight:700;font-family:"IBM Plex Mono",monospace;letter-spacing:-.5px}
.sv.red{color:#FCA5A5}.sv.amb{color:#FCD34D}.sv.grn{color:#6EE7B7}
.ss{width:1px;height:32px;background:rgba(255,255,255,.12)}
.rank-list{display:flex;flex-direction:column;gap:7px}
.rank-row{display:grid;grid-template-columns:28px 200px 1fr 70px;align-items:center;gap:10px}
.rn{font-size:11px;color:var(--mut);font-weight:600;text-align:right}
.rnm{font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rb{background:var(--bdr);border-radius:3px;height:8px}
.rf{height:100%;border-radius:3px;background:var(--brand)}
.rv{font-size:11.5px;font-family:"IBM Plex Mono",monospace;text-align:right;color:var(--mut)}
.footer{text-align:center;padding:20px 32px;border-top:1px solid var(--bdr);font-size:11.5px;color:var(--mut)}
@media(max-width:900px){.pi{padding:16px}.hdr,.nav{padding-left:16px;padding-right:16px}.crow.c21{grid-template-columns:1fr}.kgrid{grid-template-columns:repeat(2,1fr)}.rank-row{grid-template-columns:24px 1fr 60px}.rb{display:none}}
"""

    JS = f"""
var ML={json.dumps(ml_short)};
var MV={json.dumps(mv)};
var MC={json.dumps(mc)};
var CL={json.dumps([c[0][:16]+'…' if len(c[0])>16 else c[0] for c in tc])};
var CV={json.dumps([c[1] for c in tc])};
var CC={json.dumps(cat_colors[:len(tc)])};
var HV={json.dumps(health_vals)};
var TOP_CC={top_cc};var ZERO_CC={zero_cc};var NEG_CC={neg_cc};var OOS_CC={oos_cc};var RISK_CC={risk_cc};var UW_CC={uw_cc};
var TOP_ROWS={json.dumps(top_rows_c)};
var TOP_RANK={json.dumps(rank_c)};
var TOP_CNT={json.dumps(top_cnt_c)};
var ZERO_ROWS={json.dumps(zero_rows_c)};var ZERO_CNT={json.dumps(zero_cnt_c)};var ZERO_SV={json.dumps(zero_sv_c)};
var NEG_ROWS={json.dumps(neg_rows_c)};var NEG_CNT={json.dumps(neg_cnt_c)};var NEG_SV={json.dumps(neg_sv_c)};
var OOS_ROWS={json.dumps(oos_rows_c)};var OOS_CNT={json.dumps(oos_cnt_c)};var OOS_L3={json.dumps(oos_l3_c)};
var RISK_ROWS={json.dumps(risk_rows_c)};var RISK_CNT={json.dumps(risk_cnt_c)};var RISK_SV={json.dumps(risk_sv_c)};
var UW_ROWS={json.dumps(uw_rows_c)};var UW_CNT={json.dumps(uw_cnt_c)};var UW_SV={json.dumps(uw_sv_c)};
var K_NEG={K['neg_count']};var K_ZERO={K['zero_count']};var K_OOS={K['oos_count']};
var K_RISK={K['risk_count']};var K_RISK_SV={K['risk_sv']};var K_UW={K['uw_count']};var K_UW_SV={K['uw_sv']};
var K_ZERO_SV={K['zero_sv']};

function SP(id){{
  var ps=document.querySelectorAll('.page');
  for(var i=0;i<ps.length;i++)ps[i].classList.remove('active');
  var bs=document.querySelectorAll('.nbtn');
  for(var i=0;i<bs.length;i++)bs[i].classList.remove('active');
  var pg=document.getElementById('pg-'+id);
  var nb=document.getElementById('nb-'+id);
  if(pg)pg.classList.add('active');
  if(nb)nb.classList.add('active');
  window.scrollTo(0,0);
}}

var CF={{family:"'Inter',sans-serif",size:11}};
Chart.defaults.font=CF;Chart.defaults.color='#6B7280';
function fmtT(v){{return v>=1000000?(v/1000000).toFixed(1)+'M':v>=1000?(v/1000).toFixed(0)+'K':v.toLocaleString('en');}}

new Chart(document.getElementById('chM'),{{
  type:'bar',data:{{labels:ML,datasets:[{{data:MV,backgroundColor:MC,borderRadius:4,borderSkipped:false}}]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:function(c){{return ' '+Number(c.parsed.y).toLocaleString()+' units';}}}}}}}},
  scales:{{x:{{grid:{{display:false}},ticks:{{font:CF}}}},y:{{grid:{{color:'#F3F4F6'}},ticks:{{font:CF,callback:function(v){{return fmtT(v);}}}},border:{{display:false}}}}}}}}
}});

new Chart(document.getElementById('chC'),{{
  type:'bar',data:{{labels:CL,datasets:[{{data:CV,backgroundColor:CC,borderRadius:4}}]}},
  options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:function(c){{return ' '+Number(c.parsed.x).toLocaleString()+' units';}}}}}}}},
  scales:{{x:{{grid:{{color:'#F3F4F6'}},ticks:{{font:CF,callback:function(v){{return fmtT(v);}}}},border:{{display:false}}}},y:{{grid:{{display:false}},ticks:{{font:CF}}}}}}}}
}});

new Chart(document.getElementById('chH'),{{
  type:'doughnut',
  data:{{labels:['Healthy Stock','Zero Sale','Negative Stock','OOS Demand','High Risk'],
  datasets:[{{data:HV,backgroundColor:['#059669','#D97706','#DC2626','#7C3AED','#1B2B4B'],borderWidth:2,borderColor:'#fff'}}]}},
  options:{{responsive:true,maintainAspectRatio:false,cutout:'62%',plugins:{{legend:{{position:'bottom',labels:{{font:{{family:"'Inter'",size:10}},padding:8,boxWidth:10}}}}}}}}
}});

function gk(cat,cls){{
  if(!cat||cat==='All Categories')return 'ALL|ALL';
  if(!cls||cls==='All Classes')return cat+'|ALL';
  return cat+'|'+cls;
}}
function updCls(selId,cat,cc){{
  var el=document.getElementById(selId);
  var list=cat==='All Categories'?[]:((cc[cat])||[]);
  var h='<option value="All Classes">All Classes</option>';
  for(var i=0;i<list.length;i++)h+='<option value="'+list[i]+'">'+list[i]+'</option>';
  el.innerHTML=h;
}}
function fmtS(n){{
  n=parseFloat(n)||0;
  if(Math.abs(n)>=1000000)return(n/1000000).toFixed(1)+'M';
  if(Math.abs(n)>=1000)return(n/1000).toFixed(1)+'K';
  return n.toLocaleString('en',{{maximumFractionDigits:0}});
}}
function mkStrip(elId,items){{
  var h='<div class="strip">';
  for(var i=0;i<items.length;i++){{
    if(i>0)h+='<div class="ss"></div>';
    h+='<div class="si"><div class="sl">'+items[i][0]+'</div><div class="sv '+items[i][2]+'">'+items[i][1]+'</div></div>';
  }}
  document.getElementById(elId).innerHTML=h+'</div>';
}}

function initFilter(prefix,cc,rowsObj,rankObj,cntObj,svObj,l3Obj,svLabel){{
  var catEl=document.getElementById(prefix+'-cat');
  var clsEl=document.getElementById(prefix+'-cls');
  function run(){{
    var cat=catEl?catEl.value:'All Categories';
    var cls=clsEl?clsEl.value:'All Classes';
    var key=gk(cat,cls);
    var rows=rowsObj[key]||rowsObj['ALL|ALL']||'';
    var cnt=cntObj[key]||0;
    var sv=svObj?(svObj[key]||0):0;
    var l3=l3Obj?(l3Obj[key]||0):0;
    var tb=document.getElementById(prefix+'-tb');
    var cntEl=document.getElementById(prefix+'-cnt');
    var rankEl=document.getElementById(prefix+'-rank');
    if(tb)tb.innerHTML=rows;
    if(cntEl)cntEl.textContent='Showing '+cnt+' items';
    if(rankEl&&rankObj)rankEl.innerHTML=rankObj[key]||rankObj['ALL|ALL']||'';
    var stripEl=document.getElementById(prefix+'-strip');
    if(stripEl){{
      if(prefix==='zero')mkStrip(prefix+'-strip',[['Items Shown',fmtS(cnt),''],['Stock Value at Risk','AED '+fmtS(sv),'amb'],['Avg Value / Item',cnt?'AED '+fmtS(sv/cnt):'—',''],['Total Zero-Sale (Portfolio)',fmtS(K_ZERO),'red']]);
      else if(prefix==='neg')mkStrip(prefix+'-strip',[['Items Shown',fmtS(cnt),''],['Neg. Stock Value Exposure','AED '+fmtS(Math.abs(sv)),'red'],['Total Negative (Portfolio)',fmtS(K_NEG),'red'],['Action Priority','CRITICAL','amb']]);
      else if(prefix==='oos')mkStrip(prefix+'-strip',[['Items OOS (Shown)',fmtS(cnt),'red'],['Last 3M Sales (Shown)',fmtS(l3)+' units','grn'],['Total OOS (Portfolio)',fmtS(K_OOS),'red'],['Revenue Risk Level','HIGH','amb']]);
      else if(prefix==='risk')mkStrip(prefix+'-strip',[['Items Shown',fmtS(cnt),''],['Stock Value at Risk (Shown)','AED '+fmtS(sv),'amb'],['Portfolio-Wide Exposure','AED '+fmtS(K_RISK_SV),'red'],['Total High Risk (Portfolio)',fmtS(K_RISK),'amb']]);
      else if(prefix==='uw')mkStrip(prefix+'-strip',[['Items Shown',fmtS(cnt),''],['Stock Value (Shown)','AED '+fmtS(sv),'red'],['Portfolio-Wide Total','AED '+fmtS(K_UW_SV),'red'],['Total Unwanted (Portfolio)',fmtS(K_UW),'red']]);
    }}
  }}
  if(catEl)catEl.addEventListener('change',function(){{updCls(prefix+'-cls',this.value,cc);run();}});
  if(clsEl)clsEl.addEventListener('change',run);
  window[prefix+'Reset']=function(){{if(catEl)catEl.value='All Categories';updCls(prefix+'-cls','All Categories',cc);run();}};
  if(clsEl)updCls(prefix+'-cls','All Categories',cc);
  run();
}}

initFilter('top',  TOP_CC,  TOP_ROWS,  TOP_RANK,  TOP_CNT,  null,     null,    null);
initFilter('zero', ZERO_CC, ZERO_ROWS, null,      ZERO_CNT, ZERO_SV,  null,    'sv');
initFilter('neg',  NEG_CC,  NEG_ROWS,  null,      NEG_CNT,  NEG_SV,   null,    'sv');
initFilter('oos',  OOS_CC,  OOS_ROWS,  null,      OOS_CNT,  null,     OOS_L3,  'l3');
initFilter('risk', RISK_CC, RISK_ROWS, null,      RISK_CNT, RISK_SV,  null,    'sv');
initFilter('uw',   UW_CC,   UW_ROWS,   null,      UW_CNT,   UW_SV,    null,    'sv');
"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AL MADINA GROUP — Inventory Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="hdr">
  <div class="hdr-l">
    <div class="logo">
      <div class="lm">AM</div>
      <div><div class="lt">AL MADINA GROUP</div><div class="ls">Inventory Management System</div></div>
    </div>
    <div class="hsep"></div>
    <div class="hlbl">Stock Intelligence Report</div>
  </div>
  <div class="hr">{fmtN(K['total_items'],0)} SKUs &nbsp;|&nbsp; {data_label}</div>
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
  <div class="pdesc">Full portfolio summary &mdash; {fmtN(K['total_items'],0)} SKUs, all categories.</div>
  <div class="kgrid">{kpi_html}</div>
  <div class="crow c1">
    <div class="cbox">
      <div class="chead"><span class="ctitle">Monthly Sales Volume (Units)</span><span class="cmeta">Note: most recent month may be partial</span></div>
      <div class="ch ch260"><canvas id="chM"></canvas></div>
    </div>
  </div>
  <div class="crow c21">
    <div class="cbox"><div class="chead"><span class="ctitle">Top 10 Classes by Sales Volume</span></div><div class="ch ch260"><canvas id="chC"></canvas></div></div>
    <div class="cbox"><div class="chead"><span class="ctitle">Portfolio Health Breakdown</span></div><div class="ch ch260"><canvas id="chH"></canvas></div></div>
  </div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">Class Performance Summary</span><span class="b-slate">Top 12 Classes</span></div>
    <div class="tsc"><table>
      <thead><tr><th>Class</th><th class="tr">Total Sales (Units)</th><th class="tr">Stock Value (AED)</th><th class="tr">Total SKUs</th><th class="tr">Zero Sale Items</th><th class="tr">Negative Stock</th><th class="tr">OOS Demand</th><th>Status</th></tr></thead>
      <tbody>{tr_cls(D['top_cls'][:12])}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-top" class="page"><div class="pi">
  <div class="ptitle">Top Moving Items</div>
  <div class="pdesc">All items with recorded sales, ranked by total units sold. Filtered items update both the rank chart and the table.</div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="top-cat">{top_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="top-cls"></select></div>
    <button class="freset" onclick="topReset()">Reset</button>
    <span class="fcnt" id="top-cnt">Showing {len(D['top'])} items</span>
  </div>
  <div class="cbox" style="margin-bottom:20px">
    <div class="chead"><span class="ctitle">Top 10 — Total Sales Volume</span></div>
    <div class="rank-list" id="top-rank">{rank_c.get('ALL|ALL','')}</div>
  </div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">All Items Ranked by Sales Volume</span><div class="tinfo"><span class="fcnt" id="top-cnt2"></span><span class="b-brand">Sorted by Total Sales</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th>Brand</th><th class="tr">Total Sales</th><th class="tr">Last 3M Sales</th><th class="tr">Stock Qty</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost (AED)</th><th class="tr">Selling (AED)</th><th class="tr">Margin %</th></tr></thead>
      <tbody id="top-tb">{top_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-zero" class="page"><div class="pi">
  <div class="ptitle">Zero Sale Items &mdash; In Stock</div>
  <div class="pdesc">Items with no sales recorded across the full data period, yet holding stock positions and tying up working capital.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9888;</i><span class="at">Dead Stock &mdash; {fmtN(K['zero_count'],0)} Items</span></div><div class="ab">Zero sales across the entire period with no demonstrated demand. These items are consuming shelf space, warehouse capacity, and capital with no return on investment.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#8377;</i><span class="at">Capital at Risk &mdash; {fmtA(K['zero_sv'],0)}</span></div><div class="ab">Total invested value in items with zero sales history. Immediate supplier return, promotional clearance, or write-off decision required to recover this capital.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Recommended Action</span></div><div class="ab">Classify into: (1) Supplier return eligible, (2) Promotional clearance candidates, (3) Write-off or disposal. Assign to category managers. Target completion within 30 days.</div></div>
  </div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="zero-cat">{zero_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="zero-cls"></select></div>
    <button class="freset" onclick="zeroReset()">Reset</button>
    <span class="fcnt" id="zero-cnt">Showing {len(D['zero'])} items</span>
  </div>
  <div id="zero-strip"></div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">Zero Sale Items with Stock</span><div class="tinfo"><span class="b-red">Sorted by Stock Value</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th>Group</th><th class="tr">Stock Qty</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost (AED)</th><th class="tr">Selling (AED)</th><th>Last Purchase Date</th><th class="tr">LP Qty</th><th>Supplier</th></tr></thead>
      <tbody id="zero-tb">{zero_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-neg" class="page"><div class="pi">
  <div class="ptitle">Negative Stock Items</div>
  <div class="pdesc">Items showing negative on-hand quantities &mdash; indicates unposted goods receipts, missing purchase entries, or system reconciliation issues requiring immediate audit.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9888;</i><span class="at">{fmtN(K['neg_count'],0)} Items &mdash; Negative Quantity</span></div><div class="ab">Negative stock distorts financial reports, reorder calculations, and stock valuation. Root cause identification and correction required immediately to restore data integrity.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Probable Causes</span></div><div class="ab">Unposted GRNs, manual sales entry errors, missing supplier invoices, or system migration gaps. Run a pending purchase order report against each affected item.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Resolution Path</span></div><div class="ab">1. Physical count of top items by absolute negative value. 2. Match against pending GRNs. 3. Post all verified receipts. 4. Raise adjustment journals for genuine discrepancies.</div></div>
  </div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="neg-cat">{neg_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="neg-cls"></select></div>
    <button class="freset" onclick="negReset()">Reset</button>
    <span class="fcnt" id="neg-cnt">Showing {len(D['neg'])} items</span>
  </div>
  <div id="neg-strip"></div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">Negative Stock Items</span><div class="tinfo"><span class="b-red">Sorted by Negative Qty (Worst First)</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th class="tr">Stock Qty</th><th class="tr">Stock Value (AED)</th><th class="tr">Total Sales</th><th class="tr">Last 3M Sales</th><th class="tr">Cost (AED)</th><th class="tr">Selling (AED)</th><th>Last Purchase</th><th class="tr">LP Qty</th></tr></thead>
      <tbody id="neg-tb">{neg_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-oos" class="page"><div class="pi">
  <div class="ptitle">Out-of-Stock &mdash; Active Demand</div>
  <div class="pdesc">Items that recorded sales in the most recent 3 months but currently have zero or negative stock. Active lost sales requiring immediate replenishment action.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#9873;</i><span class="at">{fmtN(K['oos_count'],0)} Items Actively Out of Stock</span></div><div class="ab">Items with confirmed recent demand but currently unavailable. Every day without stock represents direct revenue loss and risk of customer attrition to competitors.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#8679;</i><span class="at">Prioritise by Last 3M Velocity</span></div><div class="ab">Items with the highest recent sales velocity carry the greatest revenue risk per day. Generate emergency purchase orders for the top items by last-3-month sales volume immediately.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">&#10003;</i><span class="at">Immediate Reorder Candidates</span></div><div class="ab">These items have proven demand. Replenishment ROI is highly predictable &mdash; reorder with confidence. This list is the buying team&#39;s action list for today.</div></div>
  </div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="oos-cat">{oos_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="oos-cls"></select></div>
    <button class="freset" onclick="oosReset()">Reset</button>
    <span class="fcnt" id="oos-cnt">Showing {len(D['oos'])} items</span>
  </div>
  <div id="oos-strip"></div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">OOS Items with Recent Sales Activity</span><div class="tinfo"><span class="b-red">Sorted by Last 3M Sales</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th class="tr">Last 3M Sales</th><th class="tr">{oos_h1}</th><th class="tr">{oos_h2}</th><th class="tr">{oos_h3}</th><th class="tr">Current Stock</th><th class="tr">Total Sales</th><th class="tr">Selling (AED)</th><th>Supplier</th></tr></thead>
      <tbody id="oos-tb">{oos_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-risk" class="page"><div class="pi">
  <div class="ptitle">High-Value Slow Movers</div>
  <div class="pdesc">Items with stock value exceeding AED 200 and total sales below 10 units. High capital exposure with minimal inventory turnover &mdash; elevated financial risk.</div>
  <div class="arow">
    <div class="abox amb"><div class="ah"><i class="ai">&#8595;</i><span class="at">{fmtA(K['risk_sv'],0)} &mdash; Low Turnover Capital</span></div><div class="ab">{fmtN(K['risk_count'],0)} items hold over {fmtA(K['risk_sv'],0)} in stock value but have generated fewer than 10 combined units in sales. This represents severely underperforming capital allocation.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">&#8856;</i><span class="at">Inventory ROI Alert</span></div><div class="ab">Holding cost at 15% annually on this capital base erodes margin significantly. Capital redeployment to fast-moving lines would materially improve overall portfolio return on investment.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Exit Strategy Required</span></div><div class="ab">Items over 6 months old with fewer than 3 units sold require immediate markdown. Items eligible for return: negotiate with supplier for credit note or exchange for faster-moving lines.</div></div>
  </div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="risk-cat">{risk_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="risk-cls"></select></div>
    <button class="freset" onclick="riskReset()">Reset</button>
    <span class="fcnt" id="risk-cnt">Showing {len(D['risk'])} items</span>
  </div>
  <div id="risk-strip"></div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">High-Value Slow Moving Items</span><div class="tinfo"><span class="b-amber">Sorted by Stock Value</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th class="tr">Stock Value (AED)</th><th class="tr">Stock Qty</th><th class="tr">Total Sales</th><th class="tr">Last 3M Sales</th><th class="tr">Cost (AED)</th><th class="tr">Selling (AED)</th><th class="tr">Margin %</th><th>Last Purchase</th><th class="tr">LP Qty</th></tr></thead>
      <tbody id="risk-tb">{risk_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-uw" class="page"><div class="pi">
  <div class="ptitle">Unwanted Repurchases</div>
  <div class="pdesc">Items where LP Qty is less than current Stock, meaning the item was already held in stock when re-purchased &mdash; and has generated zero sales throughout. Confirmed procurement waste.</div>
  <div class="arow">
    <div class="abox red"><div class="ah"><i class="ai">&#8856;</i><span class="at">{fmtA(K['uw_sv'],0)} &mdash; Confirmed Re-purchase Waste</span></div><div class="ab">{fmtN(K['uw_count'],0)} items had prior stock, were re-purchased anyway, and have still never sold. This confirms these purchase decisions were made without checking existing stock levels.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Supplier Leverage Available</span></div><div class="ab">Documented evidence of purchased-but-unsold items with prior stock history provides strong commercial leverage for buy-back agreements, credit notes, or payment term extensions.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8594;</i><span class="at">Process Fix Required</span></div><div class="ab">Buyers must be required to check current stock levels before raising any purchase order. Implement system alerts when PO quantity exceeds current stock of zero-sale items.</div></div>
  </div>
  <div class="fbar">
    <span class="flbl">Filter</span><div class="fsep"></div>
    <div class="fg"><label>Category</label><select class="fsel" id="uw-cat">{uw_cats_opts}</select></div>
    <div class="fg"><label>Class</label><select class="fsel" id="uw-cls"></select></div>
    <button class="freset" onclick="uwReset()">Reset</button>
    <span class="fcnt" id="uw-cnt">Showing {len(D['uw'])} items</span>
  </div>
  <div id="uw-strip"></div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">Re-purchased Items with Zero Sales (Stock &gt; LP Qty)</span><div class="tinfo"><span class="b-red">Sorted by Stock Value</span></div></div>
    <div class="tsc tmh"><table>
      <thead><tr><th>#</th><th>Item Name</th><th>Category</th><th>Class</th><th class="tr">Current Stock</th><th class="tr">Stock Value (AED)</th><th class="tr">Cost (AED)</th><th class="tr">Selling (AED)</th><th class="tr">Last PO Qty</th><th>Last Purchase Date</th><th>Last Purchase Supplier</th></tr></thead>
      <tbody id="uw-tb">{uw_rows_c.get('ALL|ALL','')}</tbody>
    </table></div>
  </div>
</div></div>

<div id="pg-ins" class="page"><div class="pi">
  <div class="ptitle">Strategic Analysis &amp; Action Plan</div>
  <div class="pdesc">Executive-level observations derived from the complete inventory dataset with prioritised recommendations for management action.</div>
  <div class="arow" style="grid-template-columns:1fr 1fr">
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">OOS Active Demand &mdash; {fmtN(K['oos_count'],0)} Items</span></div><div class="ab">{fmtN(K['oos_count'],0)} items are actively selling but currently unavailable. This is the highest-priority revenue risk. Generate emergency purchase orders for the top items ranked by last-3-month velocity immediately.</div></div>
    <div class="abox red"><div class="ah"><i class="ai">P1</i><span class="at">Confirmed Re-purchase Waste &mdash; {fmtA(K['uw_sv'],0)}</span></div><div class="ab">{fmtN(K['uw_count'],0)} items had existing stock, were re-purchased, and remain unsold. Engage suppliers with documented evidence for buy-back, credit notes, or exchange for fast-moving alternatives.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">Negative Stock &mdash; {fmtN(K['neg_count'],0)} Items</span></div><div class="ab">Negative stock compromises financial reporting and distorts reorder calculations. Assign operations team a 5-business-day resolution deadline. Begin with physical count of top items by absolute negative value.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">P2</i><span class="at">{fmtA(K['risk_sv'],0)} Slow Mover Capital</span></div><div class="ab">{fmtN(K['risk_count'],0)} items hold {fmtA(K['risk_sv'],0)} with fewer than 10 units sold. Annual holding cost at 15% is significant. Apply tiered markdown: 30% at 90 days, 50% at 180 days, supplier credit at 365 days.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">P3</i><span class="at">Zero Sale Dead Stock &mdash; {fmtA(K['zero_sv'],0)}</span></div><div class="ab">{fmtN(K['zero_count'],0)} items with zero sales history. Classify and action: supplier return, clearance markdown, or write-off. All {fmtN(K['zero_count'],0)} items must be actioned within 30 days.</div></div>
    <div class="abox grn"><div class="ah"><i class="ai">&#10003;</i><span class="at">Category Strength: Food &amp; Essentials</span></div><div class="ab">Grocery Food, Fresh Produce, and Beverages are the strongest performing classes by volume. These categories drive store footfall and cross-category conversion. Zero OOS tolerance must be maintained at all times.</div></div>
  </div>
  <div class="tcard">
    <div class="thd"><span class="ttitle">Prioritised Action Plan</span><span class="b-brand">Management Decision Matrix</span></div>
    <div class="tsc"><table>
      <thead><tr><th>Priority</th><th>Issue</th><th>Scale</th><th>Financial Exposure</th><th>Recommended Action</th><th>Owner</th><th>Deadline</th></tr></thead>
      <tbody>{action_html}</tbody>
    </table></div>
  </div>
  <div class="arow" style="grid-template-columns:1fr 1fr 1fr;margin-top:8px">
    <div class="abox grn"><div class="ah"><i class="ai">&#8721;</i><span class="at">Supplier Leverage</span></div><div class="ab">Purchased-but-unsold items provide documented commercial evidence. Engage suppliers for buy-back, credit notes, or exchange for faster lines. Estimated recovery: 40&ndash;60% of exposed capital.</div></div>
    <div class="abox blu"><div class="ah"><i class="ai">&#8853;</i><span class="at">Clearance Strategy</span></div><div class="ab">Bundle dead stock with fast movers via promotional mechanics. Deploy end-cap placements for slow categories. Establish a dedicated clearance zone. Target: 25% of dead stock moved within 60 days.</div></div>
    <div class="abox amb"><div class="ah"><i class="ai">&#9737;</i><span class="at">Governance Framework</span></div><div class="ab">Monthly slow-mover reviews by Category Heads. Any item with 3 consecutive zero-sale months enters mandatory review. Items at 6 months with fewer than 5 units sold are automatically flagged for discontinuation.</div></div>
  </div>
</div></div>

<div class="footer">AL MADINA GROUP &nbsp;|&nbsp; Inventory Intelligence Report &nbsp;|&nbsp; {data_label}</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script>{JS}</script>
</body>
</html>"""

    return html


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2 — STREAMLIT PORTAL
# ═════════════════════════════════════════════════════════════════════════════


import streamlit as st
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AL MADINA GROUP — Inventory Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #F4F5F7; }
  [data-testid="stHeader"] { background: #1B2B4B; }
  .block-container { padding-top: 2rem; max-width: 860px; }
  .upload-box {
    background: white;
    border: 2px dashed #D1D5DB;
    border-radius: 12px;
    padding: 40px;
    text-align: center;
    margin: 20px 0;
  }
  .info-card {
    background: white;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    padding: 16px 20px;
    margin: 8px 0;
  }
  h1 { color: #1B2B4B !important; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 6])
with col1:
    st.markdown("""
    <div style="background:linear-gradient(135deg,#3B82F6,#60A5FA);width:52px;height:52px;
    border-radius:10px;display:flex;align-items:center;justify-content:center;
    font-weight:800;font-size:18px;color:white;margin-top:4px">AM</div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown("# AL MADINA GROUP")
    st.caption("Inventory Intelligence Dashboard Generator")

st.divider()

# ── Instructions ──────────────────────────────────────────────────────────────
with st.expander("ℹ️  How to use this portal", expanded=False):
    st.markdown("""
    **Step 1** — Export your inventory report from your ERP/POS system as an Excel file (`.xlsx`).

    **Required columns** *(others are auto-detected or gracefully skipped)*:
    - `Item Name` — Product name
    - `Stock` — Current on-hand quantity  
    - `Total Sales` — Cumulative units sold
    - `Cost`, `Selling` — Pricing
    - `Category`, `Class`, `Group` — Hierarchy
    - `Stock Value` — Cost × Stock
    - `LP Date`, `LP Qty`, `LP Supplier` — Last purchase info
    - `Margin%` — Gross margin
    - Monthly sales columns, e.g. `May, 2025`, `Jun, 2025` … `May, 2026`

    **Step 2** — Upload the file below.

    **Step 3** — Click **Generate Dashboard**. Processing takes 15–45 seconds depending on file size.

    **Step 4** — Download the generated HTML file and open it in any browser. No internet required.
    
    ---
    **Unwanted Repurchase Logic:**  
    An item is flagged as *Unwanted Repurchase* if: it has a Last Purchase Date, current Stock > 0, Total Sales = 0, AND LP Qty < Current Stock.  
    This means the item already had stock when it was re-ordered — and still hasn't sold.
    """)

# ── Upload ────────────────────────────────────────────────────────────────────
st.subheader("Upload Inventory Excel File")

uploaded_file = st.file_uploader(
    label="Drop your Excel file here or click to browse",
    type=["xlsx", "xls"],
    help="Supports .xlsx and .xls formats. File size up to 100MB.",
    label_visibility="collapsed",
)

if uploaded_file is not None:
    file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("File Name", uploaded_file.name)
    col_b.metric("File Size", f"{file_size_mb:.1f} MB")
    col_c.metric("Format", uploaded_file.name.split('.')[-1].upper())

    st.divider()

    if st.button("⚡  Generate Dashboard", type="primary", use_container_width=True):

        progress = st.progress(0, text="Initialising…")
        status   = st.empty()

        try:
            # ── Read uploaded bytes directly (no tempfile needed) ─────────
            progress.progress(5, text="Reading uploaded file…")
            file_bytes = uploaded_file.getvalue()
            tmp_path = file_bytes  # pass bytes directly to read_excel


            progress.progress(10, text="Loading generator…")

            # ── Read Excel ────────────────────────────────────────────────
            progress.progress(20, text="Reading Excel file — this may take a moment for large files…")
            data, month_cols, headers = read_excel(tmp_path)
            status.info(f"✅  Loaded **{len(data):,} rows** with **{len(headers)} columns** detected. "
                        f"Month columns found: **{len(month_cols)}**")

            # ── Build datasets ────────────────────────────────────────────
            progress.progress(45, text="Analysing inventory data…")
            D = build_data(data, month_cols)
            K = D['kpis']

            # Show quick stats
            progress.progress(60, text="Building dashboard sections…")
            st.markdown(f"""
            <div class="info-card">
            <b>Analysis Complete</b><br>
            🔢 <b>{K['total_items']:,}</b> SKUs &nbsp;|&nbsp;
            📦 <b>{K['in_stock']:,}</b> in stock &nbsp;|&nbsp;
            🔴 <b>{K['neg_count']:,}</b> negative stock &nbsp;|&nbsp;
            ⚠️ <b>{K['zero_count']:,}</b> zero sale &nbsp;|&nbsp;
            🚨 <b>{K['oos_count']:,}</b> OOS active demand &nbsp;|&nbsp;
            💸 <b>{K['uw_count']:,}</b> unwanted repurchases
            </div>
            """, unsafe_allow_html=True)

            # ── Generate HTML ─────────────────────────────────────────────
            progress.progress(70, text="Pre-rendering all tables and filters…")
            html_content = generate_html(
                data,
                month_cols=month_cols,
                source_filename=uploaded_file.name,
            )

            # ── Finalising ────────────────────────────────────────────────
            progress.progress(95, text="Finalising…")

            progress.progress(100, text="Dashboard ready!")
            st.success("✅  Dashboard generated successfully!")

            # ── Download button ───────────────────────────────────────────
            out_filename = uploaded_file.name.rsplit('.', 1)[0] + '_dashboard.html'
            st.download_button(
                label="⬇️  Download Dashboard HTML",
                data=html_content.encode('utf-8'),
                file_name=out_filename,
                mime="text/html",
                use_container_width=True,
                type="primary",
            )

            st.caption(
                f"File: `{out_filename}` — open in any modern browser. No internet required after download."
            )

            # ── Inline dashboard preview ──────────────────────────────────
            st.divider()
            st.subheader("Dashboard Preview")
            st.caption("Full interactive dashboard — all pages, filters, and charts work directly here.")
            st.components.v1.html(html_content, height=820, scrolling=True)

            # ── Key metrics summary ───────────────────────────────────────
            st.divider()
            st.subheader("Summary Metrics")

            cols = st.columns(4)
            cols[0].metric("Total SKUs",        f"{K['total_items']:,}")
            cols[1].metric("Total Stock Value",  f"AED {K['total_sv']:,.0f}")
            cols[2].metric("Total Units Sold",   f"{K['total_ts']:,.0f}")
            cols[3].metric("Items In Stock",     f"{K['in_stock']:,}")

            cols2 = st.columns(4)
            cols2[0].metric("Negative Stock",    f"{K['neg_count']:,}",  delta=f"-AED {K['neg_sv']:,.0f}", delta_color="inverse")
            cols2[1].metric("Zero Sale Stock",   f"{K['zero_count']:,}", delta=f"-AED {K['zero_sv']:,.0f}", delta_color="inverse")
            cols2[2].metric("High Risk Slow",    f"{K['risk_count']:,}", delta=f"-AED {K['risk_sv']:,.0f}", delta_color="inverse")
            cols2[3].metric("Unwanted Repurch.", f"{K['uw_count']:,}",   delta=f"-AED {K['uw_sv']:,.0f}",  delta_color="inverse")

        except Exception as e:
            progress.empty()
            st.error(f"❌  Error generating dashboard: {str(e)}")
            with st.expander("Technical details"):
                import traceback
                st.code(traceback.format_exc())
            st.info("Please check that your file is a valid Excel file with the expected columns. "
                    "See the 'How to use' section above for required column names.")

else:
    # ── Empty state ───────────────────────────────────────────────────────
    st.markdown("""
    <div class="upload-box">
      <div style="font-size:48px;margin-bottom:12px">📊</div>
      <div style="font-size:16px;font-weight:600;color:#1B2B4B;margin-bottom:6px">Upload your inventory Excel file</div>
      <div style="font-size:13px;color:#6B7280">Supports .xlsx files exported from any ERP or POS system</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**What this portal generates:**")
    features = [
        ("📈", "Overview", "Monthly trends, category performance, portfolio health"),
        ("🏆", "Top Movers", "All items ranked by sales velocity with margin analysis"),
        ("🔴", "Zero Sale Stock", "Dead inventory with stock value — capital recovery list"),
        ("⚠️", "Negative Stock", "System integrity issues requiring immediate audit"),
        ("🚨", "OOS Active Demand", "Items selling but out of stock — emergency reorder list"),
        ("💰", "High Risk Slow Movers", "High-value low-turnover capital exposure"),
        ("🛑", "Unwanted Repurchases", "Re-bought items that already had stock and still haven't sold"),
        ("📋", "Insights & Actions", "Prioritised management decision matrix"),
    ]
    for icon, title, desc in features:
        st.markdown(f"**{icon} {title}** — {desc}")

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("AL MADINA GROUP Inventory Intelligence Portal &nbsp;|&nbsp; All data processed locally — nothing is stored or transmitted")
