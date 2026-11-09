# -*- coding: utf-8 -*-
# ============================================================================
# REORDER MANAGER - single interactive script
# ============================================================================
# One script replaces the 4-step chain:
#   make_carry -> build_core -> build_workbook -> brand_tool
# It asks which data files to use, then lets you build the FULL ANALYSIS
# workbook and/or per-brand REORDER lists from one clean menu.
#
#   Invoices   : MANDATORY (units actually sold)
#   Orders     : OPTIONAL - if missing, order figures = 0; every brand is
#                treated as "orders not captured" (judged on sales only),
#                no PUSHED / no FULFILL logic.
#   Stock      : OPTIONAL - if missing, every article is treated as OUT OF
#                STOCK (dispo = 0) and reorder eligibility is judged on
#                HISTORY (avg_all >= 1.0 + months_with_sales >= 1).
#
# Carry (LAST BUY / DELIVERED SINCE 2025 / PURCHASED 12M) is loaded from the
# existing carry_21_09.pkl, else rebuilt from the old master workbook, else
# empty.
#
# REF_DATE is taken from the Invoices filename (..._YYYY-MM-DD.xlsx), default
# 2026-09-21.
# ============================================================================
import sys
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass
import os, pickle, re, glob, math, datetime as dt
import json
import queue
import subprocess
import threading
import time
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from statistics import stdev
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles.fills import PatternFill as _PF

# ---------------------------------------------------------------------------
# default locations
# ---------------------------------------------------------------------------
X = r'C:\Users\ABSoft\Desktop\EXCEL\STOCK ROTATION\X TEST'
TMP = r'C:\Users\ABSoft\AppData\Local\Temp\opencode'
MASTER_OLD = r'C:\Users\ABSoft\Desktop\EXCEL\STOCK ROTATION\X TEST\STOCK ROTATION - FULL ANALYSIS 15-09-2026.xlsx'
CARRY_PKL = os.path.join(TMP, 'carry_21_09.pkl')
CORE_PKL = os.path.join(TMP, 'core_manager.pkl')

RES = None        # global per-ref metrics
GRID = None       # list of 'YYYY-MM'
REF_DATE = dt.date(2026, 9, 21)
HAS_ORDERS = True
HAS_STOCK = True
INV_N = 0
ORD_N = 0
OUTDIR = X       # where results are saved (can be changed in the menu)

# ---------------------------------------------------------------------------
# styles
# ---------------------------------------------------------------------------
NAVY = '001F4E78'; BLUE24 = '002E75B6'; GREEN25 = '00548235'; GOLD26 = '00BF8F00'
RED = '00FF9B9B'; ORANGE = '00FFC7A0'; YELLOW = '00FFEB9C'; AMBER = '00FFC000'
GREEN = '00C6EFCE'; GREY = '00D9D9D9'; LIGHTGREY = '00E7E6E6'
LIGHTBLUE = '00DDEBF7'; SALMON = '00F4B183'; TAN = '00F8CBAD'; SKY = '00BDD7EE'; MIST = '00D6CECE'

def fill(rgb):
    return PatternFill('solid', fgColor=rgb)

WHITE_B = Font(color='FFFFFF', bold=True)
NAVY_H = PatternFill('solid', fgColor=NAVY)

BRANDS = ['CAUTEX','MEYLE','GRAF','DAYCO','EYQUEM','REFLEX FRANCE','FPA','WOLIE','RENAULT',
          'FISPA','TORSON','HOFFER','MEAT DOREA','MOTRIO','ALFA','NURAL','HYDREX',
          'CLEAN FILTERS','KRIOS','BENDIX']

def to_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, str):
        try:
            return dt.datetime.strptime(v[:10], '%Y-%m-%d').date()
        except Exception:
            return None
    return v

def age_days(last):
    d = to_date(last)
    return None if d is None else (REF_DATE - d).days

def cov_fill(cov):
    if cov is None:
        return fill(GREY)
    if cov <= 3:
        return fill(RED)
    if cov <= 6:
        return fill(ORANGE)
    if cov <= 12:
        return fill(GREEN)
    return fill(GREY)

def last_fill(last):
    a = age_days(last)
    if a is None:
        return None
    if a > 180:
        return fill(RED)
    if a > 90:
        return fill(ORANGE)
    return None

def cat_fill(cat):
    return {'FAST': fill(RED), 'PUSHED': fill(AMBER), 'MEDIUM': fill(ORANGE),
            'SLOW': fill(GREEN), 'OVERSTOCK': fill(GREY), 'DEAD-STUCK': fill(RED),
            'DEAD': fill(GREY)}.get(cat)

def dts(v):
    return to_date(v)

LINK_FONT = Font(color='0000FF', underline='single')
def link_cell(cell, sheet):
    cell.hyperlink = "#'%s'!A1" % sheet
    cell.font = LINK_FONT

# ---------------------------------------------------------------------------
# shared ranking logic (identical to brand_tool / build_workbook)
# ---------------------------------------------------------------------------
M26 = 264.0 / 30.44  # months elapsed in 2026 by REF_DATE
COVER_MONTHS = 3     # order cover horizon: 2 months lead time + 1 month review period
QTY_ALPHA = 0.30
QTY_SAFETY_Z = 1.28
QTY_SAFETY_CAP = 0.5
QTY_MAX_COVER = 2.0
QTY_BACKLOG_LOOKBACK = 36
QTY_BACKLOG_ROUND = 500
QTY_STOCKOUT_COVER_MONTHS = 6  # out of stock all current year, no current-year orders -> refill 6 months
def demand_pace(v):
    return max(
        v['ord_y'].get(2026, 0) / M26,
        v['sold_y'].get(2026, 0) / M26,
        v['ord_y'].get(2025, 0) / 12.0,
        v['sold_y'].get(2025, 0) / 12.0,
        v['ord_y'].get(2024, 0) / 24.0,
        v['sold_y'].get(2024, 0) / 24.0,
        v['avg'],
        v['avg_all'])

def prio_key(v):
    cov = v['cov'] if v['cov'] is not None else 9999
    return (-demand_pace(v),
            0 if v['dispo'] == 0 else 1,
            cov if v['dispo'] > 0 else 0)

_brand_fast_cache = {}
def brand_captured(brand):
    if brand not in _brand_fast_cache:
        _brand_fast_cache[brand] = [x for x in RES.values() if x['brand'] == brand and x['cat'] == 'FAST']
    f = _brand_fast_cache[brand]
    if not f:
        return True
    nord = sum(1 for v in f if v['ord12'] > 0)
    return not (nord == 0 or nord / len(f) < 0.10)

def reorder_ok(v):
    if v['cat'] != 'FAST' or v['pushed']:
        return False
    if v['dispo'] == 0:
        # OUT OF STOCK -> not emailed, so both sold & orders stop. Judge by HISTORY.
        return v['avg_all'] >= 1.0 and v['months_with_sales'] >= 1 and v['suggested'] > 0
    # IN STOCK -> current 12M sales/orders reflect real demand.
    if v['m12'] <= 0:
        return False
    if brand_captured(v['brand']) and v['ord12'] <= 0:
        return False
    return v['suggested'] > 0

def _ses_forecast(series, alpha=QTY_ALPHA):
    vals = [float(x) for x in series if x is not None]
    if not vals:
        return 0.0
    level = vals[0]
    for x in vals[1:]:
        level = alpha * x + (1.0 - alpha) * level
    return max(0.0, level)

def _demand_sigma(series, level):
    vals = [float(x) for x in series if x is not None]
    if len(vals) < 2:
        return 0.0
    mean = sum(vals) / len(vals)
    var = sum((x - mean) ** 2 for x in vals) / (len(vals) - 1)
    sd = math.sqrt(var)
    if sd != sd:
        return 0.0
    return min(sd, level)

def _cover_months_for(v, strong_year):
    """3-month cover normally; 6 months when the item was unavailable all
    current year (0 stock, 0 current-year orders) but last year was strong."""
    if not v.get('stock_known', True):
        return COVER_MONTHS
    if float(v.get('dispo') or 0) > 0:
        return COVER_MONTHS
    ord_y = v.get('ord_y') or {}
    if ord_y.get(REF_DATE.year, 0) > 0:
        return COVER_MONTHS
    if ord_y.get(REF_DATE.year - 1, 0) <= 0:
        return COVER_MONTHS
    if strong_year != REF_DATE.year - 1:
        return COVER_MONTHS
    return QTY_STOCKOUT_COVER_MONTHS

def realistic_qty(v):
    sales = v.get('sales_full') or []
    orders = v.get('orders_full') or []
    stock = max(0.0, float(v.get('dispo') or 0))
    has_sales = any(x > 0 for x in sales)
    has_orders = any(x > 0 for x in orders)
    demand = orders if has_orders else sales
    if not demand:
        return 0, 'ORDERS' if has_orders else 'SALES', 'NO DEMAND HISTORY', 0

    lead = _ses_forecast(demand)
    sigma = _demand_sigma(demand[-24:], lead)
    safety = min(QTY_SAFETY_Z * sigma, QTY_SAFETY_CAP * lead)
    target = lead * COVER_MONTHS + safety
    cap = lead * COVER_MONTHS * QTY_MAX_COVER
    target = min(target, cap)
    forecast = max(0.0, target - stock)

    hist = list(zip(sales[-QTY_BACKLOG_LOOKBACK:], orders[-QTY_BACKLOG_LOOKBACK:]))
    tot_sales = sum(s for s, o in hist)
    tot_orders = sum(o for s, o in hist)
    gap = max(0.0, tot_orders - tot_sales - stock)

    mnow = float(M26) if M26 and M26 > 0 else 12.0
    paces = {}
    for y, q in (v.get('ord_y') or {}).items():
        if q and q > 0:
            months = mnow if y == REF_DATE.year else 12.0
            paces[y] = q / months
    cover_months = COVER_MONTHS
    if paces:
        top = max(paces.values())
        strong = {y: p for y, p in paces.items() if p >= 0.25 * top}
        y = max(strong)
        pace = strong[y]
        cover_months = _cover_months_for(v, y)
        if cover_months > COVER_MONTHS:
            target6 = min(lead * cover_months + safety, lead * cover_months * QTY_MAX_COVER)
            forecast = max(forecast, max(0.0, target6 - stock))
        backlog = min(gap, cover_months * pace)
        if backlog >= 500:
            backlog = math.ceil(backlog / QTY_BACKLOG_ROUND) * QTY_BACKLOG_ROUND
        elif backlog > 0:
            backlog = math.ceil(backlog / 100.0) * 100
    else:
        backlog = gap

    qty = int(round(max(forecast, backlog)))
    basis = 'ORDERS' if has_orders else 'SALES'
    flags = []
    if not v.get('stock_known', True):
        flags.append('STOCK UNKNOWN')
    if cover_months > COVER_MONTHS:
        flags.append('STOCKOUT %d - %d-MONTH COVER' % (REF_DATE.year, cover_months))
    if gap > 0 and backlog > 0:
        flags.append('UNSERVED DEMAND: ordered %d vs sold %d (last 3Y) - backlog covered'
                     % (tot_orders, tot_sales))
    elif lead > 0 and has_sales and not has_orders:
        flags.append('NO ORDERS - forecast from sales')
    return qty, basis, '; '.join(flags), int(round(backlog))

# ---------------------------------------------------------------------------
# DATE + FILES helpers
# ---------------------------------------------------------------------------
def find_default(path, patterns, what):
    for p in patterns:
        g = glob.glob(os.path.join(path, p))
        if g:
            return max(g, key=os.path.getmtime)
    return ''

def parse_ref_date(fname):
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', os.path.basename(fname))
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

def ask(prompt, default=''):
    if default:
        print('%s [%s]: ' % (prompt, default), end='', flush=True)
    else:
        print('%s: ' % prompt, end='', flush=True)
    v = input().strip()
    return v if v else default

def ask_path(prompt, default, required):
    # returns None when the (optional) file is skipped
    if default and os.path.exists(default):
        show = default
    else:
        show = default or '(none found - type a path or SKIP)'
    print('%s [%s]: ' % (prompt, show), end='', flush=True)
    v = input().strip()
    if not v:
        v = default
    elif v.upper() in ('SKIP', 'NONE', 'NO', '-'):
        return None
    if not v:
        return None
    if not os.path.exists(v):
        print('  ! path not found: %s' % v)
        return None
    return v

# ---------------------------------------------------------------------------
# 1) CARRY  (LAST BUY / DELIVERED SINCE 2025 / PURCHASED 12M)
# ---------------------------------------------------------------------------
def build_carry(carry_path=CARRY_PKL, master_path=MASTER_OLD):
    if os.path.exists(carry_path):
        with open(carry_path, 'rb') as f:
            carry = pickle.load(f)
        n_lb = sum(1 for d in carry.values() if d.get('last_buy'))
        print('carry: loaded %d refs from %s (last buy on %d)' % (len(carry), os.path.basename(carry_path), n_lb))
        return carry
    if master_path and os.path.exists(master_path):
        from openpyxl import load_workbook
        carry = {}
        wb = load_workbook(master_path, read_only=True, data_only=True)
        ws = wb['AZ ROTATION']
        for row in ws.iter_rows(min_row=3, values_only=True):
            ref = row[0]
            if not ref or str(ref).strip() in ('TOTAL',):
                continue
            ref = str(ref).strip()
            d = carry.setdefault(ref, {'del_since_2025': 0.0, 'last_buy': None, 'purchased_12m': 0.0})
            d['del_since_2025'] = row[50] or 0.0   # AY
            lb = row[51]                           # AZ
            if lb is not None:
                d['last_buy'] = lb
        for ws in wb.worksheets:
            if not ws.title.startswith('REORDER - '):
                continue
            for row in ws.iter_rows(min_row=6, values_only=True):
                ref = row[0]
                if not ref:
                    continue
                ref = str(ref).strip()
                if ref.startswith('SUBTOTAL'):
                    break
                d = carry.setdefault(ref, {'del_since_2025': 0.0, 'last_buy': None, 'purchased_12m': 0.0})
                d['purchased_12m'] = row[7] or 0.0   # H
                if not d['del_since_2025']:
                    d['del_since_2025'] = row[8] or 0.0  # I
        wb.close()
        n_lb = sum(1 for d in carry.values() if d['last_buy'])
        print('carry: rebuilt %d refs from old master workbook (last buy on %d)' % (len(carry), n_lb))
        return carry
    print('carry: NONE (no pkl, no old master) - LAST BUY / DELIVERED / PURCHASED 12M will be blank')
    return {}

# ---------------------------------------------------------------------------
# 2) CORE  (per-REF metrics from invoices + optional orders/stock)
# ---------------------------------------------------------------------------
def _num(v):
    if v is None:
        return None
    try:
        f = float(v)
        return None if f != f else f   # skip NaN
    except Exception:
        return None

def _read_stock(path):
    d = os.path.dirname(path) or '.'
    f = path
    if f.lower().endswith('.xls'):
        import xlrd
        bk = xlrd.open_workbook(f)
        sh = bk.sheet_by_index(0)
        rows = []
        for r in range(1, sh.nrows):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
        return rows
    df = pd.read_excel(f, header=None, sheet_name=0)
    rows = [list(r) for r in df.itertuples(index=False)][1:]
    return rows

def _norm_cols(df, wanted):
    ren = {}
    for c in df.columns:
        key = str(c).strip().casefold()
        for w in wanted:
            if key == w.casefold():
                ren[c] = w
                break
    return df.rename(columns=ren)


def build_core(inv_path, ord_path=None, stock_path=None, ref_date=REF_DATE):
    global RES, GRID, REF_DATE, HAS_ORDERS, HAS_STOCK, INV_N, ORD_N, M26, _brand_fast_cache
    REF_DATE = ref_date
    HAS_ORDERS = ord_path is not None and os.path.exists(ord_path)
    HAS_STOCK = stock_path is not None and os.path.exists(stock_path)
    M26 = ((REF_DATE - dt.date(REF_DATE.year, 1, 1)).days + 1) / 30.44
    _brand_fast_cache = {}

    inv = _norm_cols(pd.read_excel(inv_path),
                     ['Date', 'Reference', 'Designation', 'Brand', 'Quantity SOLD', 'Unit Price'])
    inv['Date'] = pd.to_datetime(inv['Date']).dt.date
    inv['Q'] = inv['Quantity SOLD'].astype(float)
    UNIT_COL = 'Unit Price'
    INV_N = len(inv)

    if HAS_ORDERS:
        ords = _norm_cols(pd.read_excel(ord_path),
                          ['Date', 'Reference', 'Designation', 'Brand', 'Quantity ORDERED'])
        ords['Date'] = pd.to_datetime(ords['Date']).dt.date
        ords['Q'] = ords['Quantity ORDERED'].astype(float)
        ORD_N = len(ords)
    else:
        ords = pd.DataFrame(columns=['Date', 'Reference', 'Quantity ORDERED', 'Designation', 'Brand'])
        ORD_N = 0

    stock = {}
    if HAS_STOCK:
        for row in _read_stock(stock_path):
            ref = str(row[0]).strip() if row and row[0] else ''
            if not ref:
                continue
            if not (len(row) >= 4):
                continue
            qty = _num(row[2]) if len(row) >= 3 else 0.0
            if qty is None:
                qty = 0.0
            prix = _num(row[3]) if len(row) >= 4 else None
            stock[ref] = {'qty': qty, 'prix': prix,
                          'des': str(row[1]).strip() if len(row) >= 2 and row[1] else '',
                          'brand': str(row[4]).strip() if len(row) >= 5 and row[4] else ''}
    stock_brands = {s['brand'] for s in stock.values() if s.get('brand')}

    # per-ref aggregates
    R = {}
    def rec(ref):
        if ref not in R:
            R[ref] = {'items': [], 'ords': [], 'des': None, 'brand': None, 'last_price': (None, None), 'last_qty': 0.0}
        return R[ref]

    for _, row in inv.iterrows():
        ref = str(row['Reference']).strip()
        d = row['Date']; q = row['Q']
        u = rec(ref)
        u['items'].append((d, q, float(row[UNIT_COL])))
        u['last_price'] = (d, float(row[UNIT_COL]))
        u['des'] = u['des'] or str(row['Designation'])
        u['brand'] = u['brand'] or str(row['Brand'])
    if HAS_ORDERS:
        for _, row in ords.iterrows():
            ref = str(row['Reference']).strip()
            u = rec(ref)
            u['ords'].append((row['Date'], row['Q']))
            if u['des'] is None:
                u['des'] = str(row['Designation'])
            if u['brand'] is None:
                u['brand'] = str(row['Brand'])

    # add stock-only refs
    for ref in stock:
        rec(ref)

    # grid: (ref_year-2)-01 .. ref_year/ref_month  (last entry = current partial)
    grid = []
    for y in (ref_date.year - 2, ref_date.year - 1, ref_date.year):
        for m in range(1, 13):
            if y == ref_date.year and m > ref_date.month:
                break
            grid.append(dt.date(y, m, 1))
    if not grid or not (grid[-1] <= ref_date):
        pass
    grididx = {g: i for i, g in enumerate(grid)}
    FULL = grid[:-1]  # full months (exclude current partial from stats)
    FULLIDX = [i for i, g in enumerate(grid) if g in set(FULL)]

    res = {}
    for ref, u in R.items():
        items = u['items']
        ords_l = u['ords']
        it_dates = [x[0] for x in items]
        price = None
        if u['last_price'][0] is not None:
            price = u['last_price'][1]
        elif ref in stock and stock[ref]['prix']:
            price = stock[ref]['prix']
        if ref in stock and stock[ref]['des']:
            des = stock[ref]['des']
        else:
            des = u['des'] or ('' if ref not in stock else stock[ref]['des'])
        if ref in stock and stock[ref]['brand']:
            brand = stock[ref]['brand']
        else:
            brand = u['brand'] or ''
        dispo = stock[ref]['qty'] if ref in stock else 0.0

        def wsum(src, days, pdate=REF_DATE):
            lo = pdate - dt.timedelta(days=days)
            return sum(q for (d, q, *r) in src if lo <= d <= pdate) if src else 0.0

        m4 = wsum(items, 122); m8 = wsum(items, 244); m12 = wsum(items, 365)
        ta = wsum(items, 61); tb = wsum(items, 122) - wsum(items, 61)
        sold30 = wsum(items, 30)
        ord12 = wsum(ords_l, 365); ord30 = wsum(ords_l, 30)

        sold_before_2024 = sum(q for (d, q, *r) in items if d < dt.date(2024, 1, 1)) if items else 0.0

        qty_monthly = [0.0] * len(grid)
        for (d, q, *r) in items:
            key = dt.date(d.year, d.month, 1)
            if key in grididx:
                qty_monthly[grididx[key]] += q
        grand = float(sum(qty_monthly))

        sold_y = {}; ord_y = {}
        for y in (2024, 2025, 2026):
            sold_y[y] = sum(q for (d, q, *r) in items if d.year == y) if items else 0.0
            ord_y[y] = sum(q for (d, q) in ords_l if d.year == y) if ords_l else 0.0

        last_date = max(it_dates) if it_dates else None
        avg = m12 / 12.0 if m12 > 0 else 0.0
        avg_all = (grand + sold_before_2024) / 36.0

        sf = [1.0] * 13
        ok_sf = True
        monthly_map = {}
        for i in FULLIDX:
            mth = grid[i].month
            monthly_map.setdefault(mth, []).append(qty_monthly[i])
        vals = [qty_monthly[i] for i in FULLIDX]
        overall = (sum(vals) / len(vals)) if vals else 0.0
        if overall > 0 and len(vals) >= 24 and all(len(v) >= 2 for v in monthly_map.values()):
            for mth in range(1, 13):
                mm = sum(monthly_map[mth]) / len(monthly_map[mth])
                sf[mth] = max(0.2, min(3.0, mm / overall))
        else:
            ok_sf = False
        if ok_sf:
            # cover horizon = the next COVER_MONTHS after REF_DATE (not a fixed quarter)
            fc3 = sum(sf[(REF_DATE.month + k - 1) % 12 + 1] for k in range(1, COVER_MONTHS + 1))
        else:
            fc3 = float(COVER_MONTHS)

        last12 = [qty_monthly[i] for i in FULLIDX[-12:]]
        safety = (1.65 * stdev(last12)) if (len(last12) >= 2 and sum(last12) > 0) else 0.0
        if safety is None or (isinstance(safety, float) and safety != safety):
            safety = 0.0

        rate = max(avg, avg_all, ta / 2.0)
        fulfill = (m12 / ord12) if ord12 > 0 else 1.0
        if ord12 > 0 and fulfill < 0.8:
            rate = max(rate, ord12 / 12.0)

        om = [0.0] * len(grid)
        for (d, q) in ords_l:
            key = dt.date(d.year, d.month, 1)
            if key in grididx:
                om[grididx[key]] += q
        srccol = 'o' if (ord12 > 0 and fulfill < 0.8) else 's'
        demand = [om[i] for i in FULLIDX] if srccol == 'o' else [qty_monthly[i] for i in FULLIDX]
        peak3 = 0.0
        for i in range(len(demand) - 2):
            peak3 = max(peak3, sum(demand[i:i+3]))
        peak3m = peak3 / 3.0

        months_with_sales = sum(1 for q in demand if q > 0)
        burst = 0.0
        if 6 <= months_with_sales <= 23:
            tot = float(sum(demand))
            if tot > 0:
                avg_event = tot / months_with_sales
                interval = len(demand) / months_with_sales
                look = max(1, round(3.0 / interval))
                burst = avg_event * look

        cov = (dispo / avg) if avg > 0 else None

        if dispo <= 0 or avg <= 0:
            stockout = 'NOW' if dispo <= 0 else None
        else:
            rem = dispo; mth = (REF_DATE.month % 12) + 1
            cnt = 0
            while rem > 0 and cnt < 60:
                f = sf[mth] if ok_sf else 1.0
                rem -= rate * f
                cnt += 1
                mth = (mth % 12) + 1
            stockout = '%dM' % cnt

        suggested = max(0, round(rate * fc3 + safety - dispo))
        if burst > 0 and burst > suggested:
            suggested = round(burst)
        if dispo <= 0 or (cov is not None and cov <= 1):
            floor2 = round(2.0 * peak3m)
            if floor2 > suggested:
                suggested = floor2
        backlog = max(0.0, ord30 - sold30)
        suggested = int(max(0, round(suggested + backlog)))

        pushed = (m12 > 0 and ord12 > 0 and m12 >= 1.2 * ord12 and ord12 < 300)
        stalepushed = (ord12 > 0 and m12 > 0 and m12 >= 1.2 * ord12 and (REF_DATE - last_date).days > 60) if last_date else False
        pushed = pushed or stalepushed
        has_hist = (grand > 0 or sold_before_2024 > 0 or len(it_dates) > 0)
        if m12 > 0:
            if pushed:
                cat = 'PUSHED'
            elif cov is None:
                cat = 'FAST'
            elif cov <= 3:
                cat = 'FAST'
            elif cov <= 6:
                cat = 'MEDIUM'
            elif cov <= 12:
                cat = 'SLOW'
            else:
                cat = 'OVERSTOCK'
        else:
            if dispo > 0:
                cat = 'IN STOCK ZERO SALES' if not has_hist else 'DEAD-STUCK'
            else:
                if not has_hist:
                    cat = 'DEAD'
                else:
                    cat = 'FAST' if avg_all >= 1 else ('MEDIUM' if (grand + sold_before_2024) > 0 else 'DEAD')

        gate_pass = all(ord_y[y] >= sold_y[y] for y in (2026, 2025, 2024))
        ratio = (ta / tb) if tb > 0 else (1.2 if ta > 0 else 1.0)
        trend = 'UP' if ratio >= 1.15 else 'DOWN' if ratio <= 0.85 else 'STABLE'

        res[ref] = dict(des=des or ref, brand=brand, price=price or 0.0, dispo=int(dispo),
                        m4=round(m4), m8=round(m8), m12=round(m12), avg=avg, cov=cov,
                        last=last_date, trend=trend, cat=cat, grand=grand,
                        sold_before=sold_before_2024, price_sales12=m12 * (price or 0.0),
                        grid=qty_monthly, gridkeys=grid, avg_all=avg_all,
                        ta=ta, tb=tb, peak3m=peak3m, burst=burst, rate=rate,
                        ord12=round(ord12), ord30=round(ord30), sold30=round(sold30),
                        backlog=round(backlog), fulfill=fulfill, fc3=fc3, safety=safety,
                        suggested=suggested, stockout=stockout,
                        sold_y=sold_y, ord_y=ord_y, gate_pass=gate_pass, pushed=pushed,
                        months_with_sales=months_with_sales, forecast12=round(rate * 12),
                        sales_full=[qty_monthly[i] for i in FULLIDX],
                        orders_full=[om[i] for i in FULLIDX],
                        stock_known=(ref in stock) or (brand in stock_brands),
                        cov_disp=(round(cov) if cov is not None else None))

    for _v in res.values():
        _old = _v['suggested']
        _q, _basis, _flag, _backlog = realistic_qty(_v)
        _v['suggested_old'] = _old
        _v['suggested'] = _q
        _v['qty_basis'] = _basis
        _v['qty_flag'] = _flag
        _v['backlog'] = _backlog
        _dem = _v['orders_full'] if _basis == 'ORDERS' else _v['sales_full']
        _rate = _ses_forecast(_dem)
        _v['rate'] = _rate
        _v['forecast12'] = int(round(_rate * 12))

    RES = res
    GRID = [g.strftime('%Y-%m') for g in grid]
    mode = []
    mode.append('invoices%s' % ('+orders' if HAS_ORDERS else ''))
    mode.append('+stock' if HAS_STOCK else ' NO-STOCK(all OOS, history-judged)')
    print('MODE: %s' % ' '.join(mode))
    print('refs=%d grid_months=%d' % (len(res), len(grid)))
    cats = {}
    for r in res.values():
        cats[r['cat']] = cats.get(r['cat'], 0) + 1
    print('categories:', cats)
    return res

# ---------------------------------------------------------------------------
# 3) FULL ANALYSIS WORKBOOK
# ---------------------------------------------------------------------------
def _nmon():
    return len(GRID)

def _now_tag():
    return dt.date.today().strftime('%d-%m-%Y')


def unique_path(folder, base, ext, stamp=True):
    """Never overwrite an older result: dates it, then adds (2), (3)... if needed."""
    name = '%s - %s' % (base, _now_tag()) if stamp else base
    suffix = ('.' + ext) if ext else ''
    path = os.path.join(folder, '%s%s' % (name, suffix))
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, '%s (%d)%s' % (name, i, suffix))
        i += 1
    return path


def today_name(base):
    return '%s - %s' % (base, _now_tag())


def build_workbook(carry, out_path=None):
    global RES, GRID, REF_DATE
    nmon = _nmon()
    tailA, tailB, tailC, tailD = 16 + nmon, 16 + nmon + 1, 16 + nmon + 2, 16 + nmon + 3  # GRAND / SOLD BEFORE / DELIVERED / LAST BUY
    if out_path is None:
        out_path = unique_path(OUTDIR, 'FULL STOCK ROTATION', 'xlsx')
    wb = Workbook()
    wb.remove(wb.active)

    # ================= INDEX =================
    ws = wb.create_sheet('INDEX')
    ws.column_dimensions['A'].width = 40; ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16; ws.column_dimensions['D'].width = 16
    ws['A1'] = 'REORDER MENU - one sheet per supplier'
    ws['A2'] = 'Click a brand to open its reorder list. STOP BUYING = do not order these again.'
    hdr = ['SHEET', 'BRAND', 'REFS TO REORDER', 'SUGGESTED UNITS']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(4, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    idx = 5
    brand_counts = {}; brand_units = {}
    main_rows = RES
    wb_brands = list(BRANDS)
    for _b in available_brands():
        if _b and _b != 'nan' and _b not in wb_brands:
            if any(v['brand'] == _b and reorder_ok(v) for v in main_rows.values()):
                wb_brands.append(_b)
    for b in wb_brands:
        refs = [r for r, v in main_rows.items() if v['brand'] == b and reorder_ok(v)]
        brand_counts[b] = len(refs)
        brand_units[b] = sum(main_rows[r]['suggested'] for r in refs)
    sb = [(r, v) for r, v in main_rows.items() if v['cat'] == 'OVERSTOCK'
          and carry.get(r, {}).get('last_buy') is not None
          and age_days(carry[r]['last_buy']) is not None and age_days(carry[r]['last_buy']) <= 90
          and v['dispo'] > 0]
    ws.cell(idx, 1, 'STOP BUYING (repurchased overstock)').fill = fill(TAN)
    link_cell(ws.cell(idx, 1), 'STOP BUYING')
    ws.cell(idx, 2, 'OVERSTOCK')
    ws.cell(idx, 3, len(sb))
    ws.cell(idx, 4, sum(carry[r].get('del_since_2025') or 0 for r, _ in sb)).number_format = '#,##0'
    ws.cell(idx, 3).fill = fill(TAN); ws.cell(idx, 4).fill = fill(TAN)
    idx += 1
    zs = [(r, v) for r, v in main_rows.items() if v['dispo'] > 0 and v['last'] is None]
    ws.cell(idx, 1, 'IN STOCK - ZERO SALES (dead money on shelf)').fill = fill(TAN)
    link_cell(ws.cell(idx, 1), 'IN STOCK - ZERO SALES')
    ws.cell(idx, 2, 'DEAD-STUCK')
    ws.cell(idx, 3, len(zs))
    ws.cell(idx, 4, sum(v['dispo'] * v['price'] for _, v in zs)).number_format = '#,##0'
    ws.cell(idx, 3).fill = fill(TAN); ws.cell(idx, 4).fill = fill(TAN)
    idx += 1
    for b in wb_brands:
        bc1 = ws.cell(idx, 1, 'REORDER - %s' % b)
        link_cell(bc1, 'REORDER - %s' % b)
        ws.cell(idx, 2, b)
        ws.cell(idx, 3, brand_counts[b])
        ws.cell(idx, 4, brand_units[b]).number_format = '#,##0'
        idx += 1
    ws.cell(idx, 1, 'GRAND TOTAL').fill = fill(GREY)
    ws.cell(idx, 3, sum(brand_counts.values())).fill = fill(GREY)
    ws.cell(idx, 4, sum(brand_units.values())).fill = fill(GREY)
    ws.cell(idx, 4).number_format = '#,##0'

    # ================= AZ ROTATION =================
    ws = wb.create_sheet('AZ ROTATION')
    ws.freeze_panes = 'D3'
    widths = {'A':16,'B':52,'C':13,'D':10,'E':11,'F':11,'G':8,'H':8,'I':9,'J':9,'K':11,'L':9,
              'M':8,'N':11,'O':12}
    for i in range(16, tailD + 7):
        widths[get_column_letter(i)] = 13
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    headers = ['REF','DESIGNATION','BRAND','QTY DISPO','UNIT PRICE','STOCK VALUE','M4','M8',
               'M12 SOLD','AVG/MONTH','COUVERTURE','LAST SALE','TREND 3M','CATEGORY',
               'SALES VALUE 12M']
    for c, t in enumerate(headers, 1):
        cell = ws.cell(2, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    mon = {'2024': BLUE24, '2025': GREEN25, '2026': GOLD26}
    MTH = {1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC'}
    for i, g in enumerate(GRID):
        y = g[:4]
        mm = int(g[5:7]); yy = int(g[2:4])
        lbl = '%s-%02d' % (MTH[mm], yy)
        if g == GRID[-1]:
            lbl = '%s-%02d*' % (MTH[mm], yy)
        cell = ws.cell(2, 16 + i, lbl); cell.fill = fill(mon.get(y, GREEN25)); cell.font = WHITE_B
    tail = ['GRAND TOTAL', 'SOLD BEFORE 2024', 'DELIVERED SINCE 2025', 'LAST BUY',
            'ORDERED 2024', 'SOLD 2024', 'ORDERED 2025', 'SOLD 2025', 'ORDERED 2026', 'SOLD 2026']
    for j, t in enumerate(tail):
        cell = ws.cell(2, tailA + j, t); cell.fill = fill(GOLD26); cell.font = WHITE_B

    items = [(r, v) for r, v in RES.items() if v['grand'] > 0 or v['dispo'] > 0]
    items.sort(key=lambda x: x[1]['m12'], reverse=True)
    r = 3
    for ref, v in items:
        ws.cell(r, 1, ref); ws.cell(r, 2, v['des']); ws.cell(r, 3, v['brand'])
        dcell = ws.cell(r, 4, v['dispo']); dcell.number_format = '#,##0'
        dcell.fill = fill(RED) if v['dispo'] == 0 else fill(LIGHTBLUE)
        pc = ws.cell(r, 5, v['price']); pc.number_format = '#,##0'
        sc = ws.cell(r, 6, v['dispo'] * v['price']); sc.number_format = '#,##0'
        for col, key, fmt in ((7,'m4','#,##0'),(8,'m8','#,##0'),(9,'m12','#,##0')):
            cc = ws.cell(r, col, v[key]); cc.number_format = fmt
        ac = ws.cell(r, 10, v['avg']); ac.number_format = '0.0'
        kc = ws.cell(r, 11, round(v['cov']) if v['cov'] is not None else None)
        if kc.value is not None:
            kc.number_format = '#,##0'; kc.fill = cov_fill(v['cov'])
        lc = ws.cell(r, 12, v['last'])
        if v['last'] is not None:
            lc.number_format = 'DD/MM/YYYY'; lc.fill = last_fill(v['last'])
        tc = ws.cell(r, 13, v['trend'])
        tc.fill = {'UP': fill(GREEN), 'DOWN': fill(RED), 'STABLE': fill(LIGHTGREY)}[v['trend']]
        nc = ws.cell(r, 14, v['cat'])
        f = cat_fill(v['cat'])
        if f: nc.fill = f
        oc = ws.cell(r, 15, v['m12'] * v['price']); oc.number_format = '#,##0'
        for i, q in enumerate(v['grid']):
            cell = ws.cell(r, 16 + i, int(q) if q else 0); cell.number_format = '#,##0'
        aw = ws.cell(r, tailA, int(v['grand'])); aw.number_format = '#,##0'
        ax = ws.cell(r, tailB, int(v['sold_before'])); ax.number_format = '#,##0'
        cb = carry.get(ref, {})
        ay = ws.cell(r, tailC, int(cb.get('del_since_2025') or 0)); ay.number_format = '#,##0'
        az = ws.cell(r, tailD, dts(cb.get('last_buy')))
        if az.value is not None:
            az.number_format = 'DD/MM/YYYY'
        py_vals = (v['ord_y'].get(2024, 0), v['sold_y'].get(2024, 0),
                   v['ord_y'].get(2025, 0), v['sold_y'].get(2025, 0),
                   v['ord_y'].get(2026, 0), v['sold_y'].get(2026, 0))
        for k, val in enumerate(py_vals):
            cc = ws.cell(r, tailD + 1 + k, int(val)); cc.number_format = '#,##0'
        r += 1
    tot = ws.cell(r, 1, 'TOTAL'); tot.fill = fill(GREY)
    sums = {'dispo': 0, 'stockval': 0, 'm4': 0, 'm8': 0, 'm12': 0, 'sales12': 0}
    gridsum = [0] * nmon
    for ref, v in items:
        sums['dispo'] += v['dispo']; sums['stockval'] += v['dispo'] * v['price']
        sums['m4'] += v['m4']; sums['m8'] += v['m8']; sums['m12'] += v['m12']
        sums['sales12'] += v['m12'] * v['price']
        for i, q in enumerate(v['grid']):
            gridsum[i] += int(q)
    for col, val in ((4,sums['dispo']),(5,None),(6,sums['stockval']),(7,sums['m4']),(8,sums['m8']),
                     (9,sums['m12']),(15,sums['sales12'])):
        if val is not None:
            cc = ws.cell(r, col, val); cc.number_format = '#,##0'; cc.fill = fill(GREY)
    for i, val in enumerate(gridsum):
        cc = ws.cell(r, 16 + i, val); cc.number_format = '#,##0'; cc.fill = fill(GREY)
    awc = ws.cell(r, tailA, sum(gridsum)); awc.number_format = '#,##0'; awc.fill = fill(GREY)
    for k, (src, yr) in enumerate((('ord_y', 2024), ('sold_y', 2024),
                                   ('ord_y', 2025), ('sold_y', 2025),
                                   ('ord_y', 2026), ('sold_y', 2026))):
        cc = ws.cell(r, tailD + 1 + k, sum(v[src].get(yr, 0) for _, v in items))
        cc.number_format = '#,##0'; cc.fill = fill(GREY)
    AZ_LAST = r

    # ================= ACTIONS =================
    ws = wb.create_sheet('ACTIONS')
    ws.freeze_panes = 'A3'
    for col, w in {'A':16,'B':52,'C':12,'D':11,'E':8,'F':9,'G':13,'H':46,'I':9,'J':13}.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    hdr = ['REF','DESIGNATION','CATEGORY','COUVERTURE','DISPO','AVG/MONTH(12M)',
           'SUGGESTED ORDER','RECOMMENDED ACTION','PRIORITY','BRAND']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(2, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    CAT_ORD = {'FAST':0,'PUSHED':1,'MEDIUM':2,'SLOW':3,'OVERSTOCK':4,'DEAD-STUCK':5,'DEAD':6,'IN STOCK ZERO SALES':7}
    items = [(r, v) for r, v in RES.items() if v['grand'] > 0 or v['dispo'] > 0]
    items.sort(key=lambda x: (CAT_ORD[x[1]['cat']], -x[1]['m12']))
    r = 3
    for ref, v in items:
        ws.cell(r, 1, ref); ws.cell(r, 2, v['des'])
        cc = ws.cell(r, 3, v['cat']); f = cat_fill(v['cat'])
        if f: cc.fill = f
        kc = ws.cell(r, 4, round(v['cov']) if v['cov'] is not None else None)
        if kc.value is not None:
            kc.number_format = '#,##0'; kc.fill = cov_fill(v['cov'])
        dc = ws.cell(r, 5, v['dispo']); dc.number_format = '#,##0'
        if v['dispo'] == 0: dc.fill = fill(RED)
        fc = ws.cell(r, 6, v['avg']); fc.number_format = '0.0'
        gc = ws.cell(r, 7, v['suggested']); gc.number_format = '#,##0'
        if v['suggested'] >= 1000:
            gc.fill = fill(RED)
        elif v['suggested'] > 0:
            gc.fill = fill(YELLOW)
        cat = v['cat']
        if cat == 'FAST':
            if v['dispo'] == 0:
                act = 'URGENT - BEST SELLER OUT OF STOCK'; af = fill(RED); pri = 'HIGH'; pf = fill(RED)
            elif v['suggested'] > 0:
                act = 'REORDER NOW - low cover vs selling speed'; af = fill(YELLOW); pri = 'MEDIUM'; pf = fill(YELLOW)
            else:
                act = 'AT TARGET - enough stock, no order needed'; af = fill(GREEN); pri = 'LOW'; pf = None
        elif cat == 'PUSHED':
            act = 'VERIFY - SOLD > 120% OF ORDERS (2024-2026), pushed or stale - adjust or discount, no reorder'
            af = fill(YELLOW); pri = 'MEDIUM'; pf = fill(YELLOW)
        elif cat == 'MEDIUM':
            act = 'MAINTAIN - healthy, do nothing (4-6 months)'; af = fill(GREEN); pri = 'LOW'; pf = None
        elif cat == 'SLOW':
            act = 'WATCH - slow mover, order on customer demand'; af = fill(GREEN); pri = 'LOW'; pf = None
        elif cat == 'OVERSTOCK':
            act = 'STOP ORDERING - let drain / discount to clear'; af = fill(GREY); pri = 'LOW'; pf = None
        elif cat == 'DEAD-STUCK':
            act = 'LIQUIDATE - no sales 12M with stock, promo/discount/return'; af = fill(RED); pri = 'HIGH'; pf = fill(RED)
        else:
            act = 'ARCHIVE - no stock, no sales'; af = fill(GREY); pri = 'LOW'; pf = None
        ac = ws.cell(r, 8, act); ac.fill = af
        ic = ws.cell(r, 9, pri)
        if pf: ic.fill = pf
        ws.cell(r, 10, v['brand'])
        r += 1
    ACT_LAST = r - 1

    # ================= REORDER - <BRAND> =================
    def build_reorder(brand):
        ws = wb.create_sheet('REORDER - %s' % brand)
        ws.freeze_panes = 'A3'
        for col, w in {'A':15,'B':50,'C':9,'D':12,'E':10,'F':12,'G':14,'H':12,'I':16,'J':12,'K':12,
                       'L':9,'M':14,'N':11,'O':12,'P':13,'Q':13,'R':13,'S':11,'T':13,'U':11,'V':11,
                       'W':15,'X':40,'Y':11,'Z':13,'AA':13,'AB':13,'AC':13,'AD':13,'AE':13,'AF':55}.items():
            ws.column_dimensions[col].width = w
        ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
        refs = [(r, v) for r, v in RES.items() if v['brand'] == brand and reorder_ok(v)]
        tot = sum(v['suggested'] for _, v in refs)
        ws['A2'] = '%s - REORDER LIST (%d refs - suggested %d units) - ranked by DEMAND PACE: biggest client demand first (current-year orders & sales per month, then 2025, half-weighted 2024). Out-of-stock wins ties. Orders count as demand even where sales lag - you under-delivered.' % (brand, len(refs), tot)
        ws['A2'].font = Font(bold=True)
        ws['A3'] = ('SUGGESTED ORDER uses ORDERS as demand when available, otherwise SALES: 3 months of smoothed '
                    'demand + capped safety - current stock. If clients ordered more than you sold, a backlog floor '
                    'covers the unserved gap (orders - sales - stock). The floor is normally capped at 3 months of '
                    'the strongest year; it uses 6 months when the item was out of stock all current year with no '
                    'current-year orders, so a ref like CX020035 is not under-ordered.')
        hdr = ['RANK','REF','DESIGNATION','DISPO','AVG/MONTH(12M)','PEAK 3M/M','SOLD LAST 2M/M','SOLD PRIOR 2M/M',
               'PURCHASED 12M','DELIVERED SINCE 2025','FORECAST 12M','ORDERS 12M','FULFILL %',
               'BURST NEXT 3M','COUVERTURE','EST STOCK-OUT','SUGGESTED ORDER','CLIENT ORDERED (30D)',
               'SOLD (30D)','BACKLOG (3Y)','QTY TO ORDER','UNIT PRICE','LAST SALE','LOST/MO (if OOS)','STATUS',
               'ORDERED 2024','SOLD 2024','ORDERED 2025','SOLD 2025','ORDERED 2026','SOLD 2026','NOTE']
        for c, t in enumerate(hdr, 1):
            cell = ws.cell(5, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
        def loss(v):
            return round(v['rate'] * v['price'])
        refs.sort(key=lambda x: prio_key(x[1]))
        r = 6
        sumP = 0
        for rank, (ref, v) in enumerate(refs, 1):
            rk = ws.cell(r, 1, rank); rk.number_format = '0'; rk.font = Font(bold=True)
            ws.cell(r, 2, ref); ws.cell(r, 3, v['des'])
            dc = ws.cell(r, 4, v['dispo']); dc.number_format = '#,##0'
            dc.fill = fill(RED) if v['dispo'] == 0 else fill(LIGHTBLUE)
            for col, key in ((5,'avg'),(6,'peak3m'),(7,'ta'),(8,'tb')):
                if key == 'ta':
                    val = v['ta'] / 2.0
                elif key == 'tb':
                    val = v['tb'] / 2.0
                else:
                    val = v[key]
                cc = ws.cell(r, col, round(val, 1)); cc.number_format = '0.0'
            cb = carry.get(ref, {})
            pc = ws.cell(r, 9, int(cb.get('purchased_12m') or 0)); pc.number_format = '#,##0'
            dl = ws.cell(r, 10, int(cb.get('del_since_2025') or 0)); dl.number_format = '#,##0'
            fc = ws.cell(r, 11, v['forecast12']); fc.number_format = '#,##0'
            oc = ws.cell(r, 12, v['ord12']); oc.number_format = '#,##0'
            fl = ws.cell(r, 13, min(1.0, v['fulfill'])); fl.number_format = '0%'
            if v['fulfill'] < 0.8:
                fl.fill = fill(SALMON)
            bc = ws.cell(r, 14, v['burst']); bc.number_format = '#,##0'
            nc = ws.cell(r, 15, round(v['cov']) if v['cov'] is not None else None)
            if nc.value is not None:
                nc.number_format = '#,##0'; nc.fill = cov_fill(v['cov'])
            oc2 = ws.cell(r, 16, v['stockout'] if v['stockout'] else 'NOW' if v['dispo']==0 else '')
            if v['dispo'] == 0:
                oc2.fill = fill(RED)
            g2 = ws.cell(r, 17, v['suggested']); g2.number_format = '#,##0'
            sumP += v['suggested']
            q1 = ws.cell(r, 18, v['ord30']); q1.number_format = '#,##0'
            q2 = ws.cell(r, 19, v['sold30']); q2.number_format = '#,##0'
            bk = ws.cell(r, 20, v['backlog']); bk.number_format = '#,##0'
            if v['backlog'] > 0:
                bk.fill = fill(SALMON)
            ty = ws.cell(r, 21, None); ty.number_format = '#,##0'; ty.fill = fill('00FFF2CC')
            up = ws.cell(r, 22, v['price']); up.number_format = '#,##0'
            ls = ws.cell(r, 23, v['last'])
            if v['last'] is not None:
                ls.number_format = 'DD/MM/YYYY'
            if v['dispo'] == 0:
                lost = ws.cell(r, 24, round(v['rate'] * v['price'])); lost.number_format = '#,##0'
                if v.get('stock_known', True):
                    lost.fill = fill(RED)
                    st = 'URGENT - BEST SELLER OUT OF STOCK'; stf = fill(RED)
                else:
                    lost.fill = fill(YELLOW)
                    st = 'CHECK STOCK - BRAND NOT IN STOCK FILE'; stf = fill(YELLOW)
            elif v['suggested'] > 0:
                st = 'REORDER NOW - cover low vs selling speed'; stf = fill(YELLOW)
            else:
                st = 'AT TARGET - enough stock'; stf = None
            sc2 = ws.cell(r, 25, st)
            if stf: sc2.fill = stf
            py_vals = (v['ord_y'].get(2024, 0), v['sold_y'].get(2024, 0),
                       v['ord_y'].get(2025, 0), v['sold_y'].get(2025, 0),
                       v['ord_y'].get(2026, 0), v['sold_y'].get(2026, 0))
            for k, val in enumerate(py_vals):
                cc = ws.cell(r, 26 + k, int(val)); cc.number_format = '#,##0'
            ws.cell(r, 32, v.get('qty_flag') or '')
            r += 1
        stot = ws.cell(r, 2, 'TOTAL SUGGESTED - %s' % brand); stot.fill = fill(GREY)
        totcc = ws.cell(r, 17, sumP); totcc.number_format = '#,##0'; totcc.fill = fill(GREY)
        for k, (src, yr) in enumerate((('ord_y', 2024), ('sold_y', 2024),
                                       ('ord_y', 2025), ('sold_y', 2025),
                                       ('ord_y', 2026), ('sold_y', 2026))):
            cc = ws.cell(r, 26 + k, sum(v[src].get(yr, 0) for _, v in refs))
            cc.number_format = '#,##0'; cc.fill = fill(GREY)
        return len(refs), sumP

    brand_counts2 = {}; brand_units2 = {}
    for b in wb_brands:
        n, u = build_reorder(b)
        brand_counts2[b] = n; brand_units2[b] = u

    # ================= OVERSTOCK =================
    ws = wb.create_sheet('OVERSTOCK')
    ws.freeze_panes = 'A3'
    for col, w in {'A':22,'B':14,'C':52,'D':11,'E':18,'F':15,'G':11,'H':11,'I':12,'J':46}.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    hdr = ['TYPE','REF','DESIGNATION','BRAND','COUVERTURE (months)','UNITS SOLD 12M','QTY DISPO',
           'UNIT PRICE','STOCK VALUE','RECOMMENDED ACTION']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(2, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    ov = [(r, v) for r, v in RES.items() if v['cat'] == 'OVERSTOCK' and (v['grand'] > 0 or v['dispo'] > 0)]
    sell = [x for x in ov if x[1]['m12'] >= 50]
    nsell = [x for x in ov if x[1]['m12'] < 50]
    def ov_rows(lst, typ, afill):
        lst.sort(key=lambda x: x[1]['dispo'] * x[1]['price'], reverse=True)
        nonlocal r
        for ref, v in lst:
            ws.cell(r, 1, typ).fill = afill
            ws.cell(r, 2, ref); ws.cell(r, 3, v['des']); ws.cell(r, 4, v['brand'])
            ec = ws.cell(r, 5, round(v['cov']) if v['cov'] is not None else ''); ec.number_format = '#,##0'
            fc = ws.cell(r, 6, v['m12']); fc.number_format = '#,##0'
            gc = ws.cell(r, 7, v['dispo']); gc.number_format = '#,##0'
            hc = ws.cell(r, 8, v['price']); hc.number_format = '#,##0'
            ic = ws.cell(r, 9, v['dispo'] * v['price']); ic.number_format = '#,##0'
            if typ == 'OVERSTOCK-SELLING':
                ws.cell(r, 10, 'STOP ORDERING - strong seller, let stock drain naturally (%0.f months)' % (v['cov'] or 0))
            else:
                ws.cell(r, 10, 'DO NOT ORDER - overstocked, discount to clear (%0.f months)' % (v['cov'] or 0))
            r += 1
    r = 3
    ov_rows(sell, 'OVERSTOCK-SELLING', fill(SKY))
    ov_rows(nsell, 'OVERSTOCK-NOT SELLING', fill(MIST))
    for typ, lst, lab in (('SELLING', sell, 'SELLING subtotal'), ('NOT SELLING', nsell, 'NOT SELLING subtotal')):
        if lst:
            cc = ws.cell(r, 1, lab); cc.fill = fill(GREY)
            ic = ws.cell(r, 9, sum(x['dispo'] * x['price'] for _, x in lst)); ic.number_format = '#,##0'
            ic.fill = fill(GREY)
            r += 1
    OV_LAST = r - 1

    # ================= STOP BUYING =================
    ws = wb.create_sheet('STOP BUYING')
    ws.freeze_panes = 'A3'
    for col, w in {'A':15,'B':50,'C':11,'D':11,'E':9,'F':9,'G':11,'H':13}.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    sb = [(r, v) for r, v in RES.items() if v['cat'] == 'OVERSTOCK'
          and carry.get(r, {}).get('last_buy') is not None
          and age_days(carry[r]['last_buy']) is not None and age_days(carry[r]['last_buy']) <= 90
          and v['dispo'] > 0 and v['grand'] > 0]
    ws['A2'] = 'STOP BUYING - OVERSTOCK refs repurchased in last 90 days (%d refs)' % len(sb)
    ws['A3'] = 'These articles are OVERSTOCK (13+ months of stock) yet you bought them again recently. Do NOT reorder.'
    hdr = ['REF','DESIGNATION','BRAND','COUVERTURE','SOLD 12M','DISPO','LAST BUY','DELIVERED SINCE 2025']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(5, c, t); cell.fill = fill(RED); cell.font = WHITE_B
    sb.sort(key=lambda x: dts(carry[x[0]].get('last_buy')) or dt.date.min, reverse=True)
    r = 6
    for ref, v in sb:
        f = fill(TAN)
        for c, val in ((1, ref), (2, v['des']), (3, v['brand'])):
            ws.cell(r, c, val).fill = f
        dc = ws.cell(r, 4, round(v['cov']) if v['cov'] is not None else ''); dc.number_format = '#,##0'; dc.fill = f
        ec = ws.cell(r, 5, v['m12']); ec.number_format = '#,##0'; ec.fill = f
        fc = ws.cell(r, 6, v['dispo']); fc.number_format = '#,##0'; fc.fill = f
        gc = ws.cell(r, 7, dts(carry[ref]['last_buy'])); gc.number_format = 'DD/MM/YYYY'; gc.fill = f
        hc = ws.cell(r, 8, int(carry[ref].get('del_since_2025') or 0)); hc.number_format = '#,##0'; hc.fill = f
        r += 1
    stot = ws.cell(r, 1, 'TOTAL')
    for col in range(1, 9):
        ws.cell(r, col).fill = fill(GREY)
    for col, key in ((5,'m12'),(6,'dispo'),(8,'del')):
        val = sum((x[1]['m12'] if key=='m12' else x[1]['dispo'] if key=='dispo'
                   else carry[x[0]].get('del_since_2025') or 0) for x in sb)
        cc = ws.cell(r, col, val); cc.number_format = '#,##0'; cc.fill = fill(GREY)
    SB_LAST = r

    # ================= IN STOCK - ZERO SALES =================
    ws = wb.create_sheet('IN STOCK - ZERO SALES')
    ws.freeze_panes = 'A3'
    for col, w in {'A':15,'B':50,'C':11,'D':10,'E':10,'F':12,'G':9,'H':16,'I':14,'J':16,'K':11,'L':11,'M':56}.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    zs = [(r, v) for r, v in RES.items() if v['dispo'] > 0 and v['last'] is None]
    ws['A2'] = 'IN STOCK - ZERO SALES - articles on the shelf with NO sales ever (%d refs)' % len(zs)
    ws['A3'] = 'This list = everything in stock that has NEVER SOLD (no last sale in the whole history). Bought, put on shelf, never sold - return/discount them.'
    hdr = ['REF','DESIGNATION','BRAND','QTY DISPO','UNIT PRICE','STOCK VALUE','M12 SOLD',
           'GRAND TOTAL (since 2024)','SOLD BEFORE 2024','DELIVERED SINCE 2025','LAST BUY','LAST SALE','STATUS']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(5, c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    r = 6
    for ref, v in zs:
        f = fill(SALMON)
        for c, val in ((1, ref), (2, v['des']), (3, v['brand'])):
            ws.cell(r, c, val).fill = f
        dc = ws.cell(r, 4, v['dispo']); dc.number_format = '#,##0'; dc.fill = f
        ec = ws.cell(r, 5, v['price']); ec.number_format = '#,##0'; ec.fill = f
        fc = ws.cell(r, 6, v['dispo'] * v['price']); fc.number_format = '#,##0'; fc.fill = f
        for col in (7,8,9,10):
            cc = ws.cell(r, col, 0); cc.number_format = '#,##0'; cc.fill = f
        cb = carry.get(ref, {})
        kc = ws.cell(r, 11, dts(cb.get('last_buy')))
        if kc.value is not None:
            kc.number_format = 'DD/MM/YYYY'; kc.fill = f
        lc = ws.cell(r, 12, None)
        mc = ws.cell(r, 13, 'NEVER MOVED - bought, put on shelf, never sold - return/discount'); mc.fill = fill(RED)
        r += 1
    stot = ws.cell(r, 1, 'TOTAL')
    for col in range(1, 14):
        ws.cell(r, col).fill = fill(GREY)
    cc = ws.cell(r, 4, sum(x[1]['dispo'] for x in zs)); cc.number_format = '#,##0'; cc.fill = fill(GREY)
    fc = ws.cell(r, 6, sum(x[1]['dispo'] * x[1]['price'] for x in zs)); fc.number_format = '#,##0'; fc.fill = fill(GREY)
    ZS_LAST = r

    # ================= DASHBOARD =================
    ws = wb.create_sheet('DASHBOARD')
    ws.column_dimensions['A'].width = 2
    for col, w in {'B':34,'C':56,'D':13,'E':12,'F':16,'G':13,'H':12,'I':12}.items():
        ws.column_dimensions[col].width = w
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    ws['B3'] = 'STOCK ROTATION - EXECUTIVE DASHBOARD'; ws['B3'].font = Font(bold=True, size=14)
    src = 'Invoices (%d lines)' % INV_N
    if HAS_ORDERS:
        src += ' + Orders (%d lines)' % ORD_N
    else:
        src += ' (NO Orders export -> judged on sales only)'
    stk = 'STOCK %s' % ('file present' if HAS_STOCK else 'MISSING -> all OOS, judged on history')
    ws['B4'] = ('Built %s from %s | REAL stock (QTY DISPO) from %s | ROLLING windows end %s '
                '(M4/M8/M12 = last 4/8/12 months) | Categories by MONTHS OF STOCK: FAST<=3, MEDIUM 4-6, SLOW 7-12, '
                'OVERSTOCK 13+, DEAD-STUCK = no sales 12M + stock | PUSHED = SOLD>=120%% ORDERS (base<300/yr), kept OUT '
                'of reorder | Reorder gate: ORDERS >= SOLD in each year 2026/2025/2024.'
                % (REF_DATE.strftime('%d-%m-%Y'), src, stk, REF_DATE.strftime('%d-%b-%Y')))
    ws['B6'] = 'STOCK HEALTH OVERVIEW'
    hdr = ['CATEGORY','ARTICLES','% ARTICLES','UNITS SOLD 12M','UNITS IN STOCK','STOCK VALUE']
    for c, t in enumerate(hdr, 0):
        cell = ws.cell(7, 2 + c, t); cell.fill = NAVY_H; cell.font = WHITE_B
    cats = ['FAST','PUSHED','MEDIUM','SLOW','OVERSTOCK','DEAD-STUCK (no sales, has stock)','DEAD']
    catfillmap = {'FAST': fill(RED), 'PUSHED': fill(AMBER), 'MEDIUM': fill(ORANGE), 'SLOW': fill(GREEN),
                  'OVERSTOCK': fill(GREY), 'DEAD-STUCK (no sales, has stock)': fill(RED), 'DEAD': fill(GREY)}
    keys = {'FAST':'FAST','PUSHED':'PUSHED','MEDIUM':'MEDIUM','SLOW':'SLOW','OVERSTOCK':'OVERSTOCK',
            'DEAD-STUCK (no sales, has stock)':'DEAD-STUCK','DEAD':'DEAD'}
    n_all = len([1 for r, v in RES.items() if v['grand'] > 0 or v['dispo'] > 0])
    r = 8
    for cat in cats:
        k = keys[cat]
        lst = [(rr, v) for rr, v in RES.items() if v['cat'] == k and (v['grand'] > 0 or v['dispo'] > 0)]
        n = len(lst); u = sum(v['m12'] for _, v in lst); d = sum(v['dispo'] for _, v in lst)
        val = sum(v['dispo'] * v['price'] for _, v in lst)
        bc = ws.cell(r, 2, cat); bc.fill = catfillmap[cat]
        ws.cell(r, 3, n)
        ws.cell(r, 4, round(n / n_all * 100, 1)) if n_all else ws.cell(r, 4, 0)
        ws.cell(r, 4).number_format = '0.0'
        cc = ws.cell(r, 5, u); cc.number_format = '#,##0'
        dc = ws.cell(r, 6, d); dc.number_format = '#,##0'
        gc = ws.cell(r, 7, val); gc.number_format = '#,##0'
        r += 1
    fast_oos = [1 for rr, v in RES.items() if v['cat'] == 'FAST' and v['dispo'] == 0 and v['m12'] > 0]
    deadstuck_val = sum(v['dispo'] * v['price'] for rr, v in RES.items() if v['cat'] == 'DEAD-STUCK')
    lost_mo = sum(round(v['rate'] * v['price']) for rr, v in RES.items()
                  if v['cat'] == 'FAST' and v['dispo'] == 0 and v['m12'] > 0)
    ws.cell(15, 2, 'ALERT: BEST-SELLERS OUT OF STOCK (sold in 12M, zero shelf)')
    ws.cell(15, 7, len(fast_oos)).number_format = '#,##0'
    ws.cell(16, 2, 'ALERT: MONEY TRAPPED IN DEAD-STUCK STOCK')
    ws.cell(16, 7, deadstuck_val).number_format = '#,##0'
    ws.cell(17, 2, 'ALERT: LOST SALES - %d best-sellers empty, you lose this DA every month' % len(fast_oos))
    ws.cell(17, 7, lost_mo).number_format = '#,##0'
    ws.cell(18, 6, 'COVERAGE HEALTH'); ws.cell(18, 3, 'MONEY YOU OWN BY MONTHS OF STOCK')
    bucketnames = ['0-3M (FAST - sell fast)', '4-6M (MEDIUM - healthy)', '7-12M (SLOW - watch)',
                   '13-24M (OVERSTOCK - clear/drain)', '25M+ (OVERSTOCK - money frozen)', 'NO SALES (dead money)']
    rowmap = {'0-3M (FAST - sell fast)': 19, '4-6M (MEDIUM - healthy)': 20, '7-12M (SLOW - watch)': 21,
              '13-24M (OVERSTOCK - clear/drain)': 22, '25M+ (OVERSTOCK - money frozen)': 23, 'NO SALES (dead money)': 24}
    covdist = {b: 0 for b in bucketnames}
    vdist = {b: 0 for b in bucketnames}
    for rr, v in RES.items():
        if not (v['grand'] > 0 or v['dispo'] > 0):
            continue
        if v['cov'] is None:
            b = 'NO SALES (dead money)'
        elif v['cov'] <= 3: b = '0-3M (FAST - sell fast)'
        elif v['cov'] <= 6: b = '4-6M (MEDIUM - healthy)'
        elif v['cov'] <= 12: b = '7-12M (SLOW - watch)'
        elif v['cov'] <= 24: b = '13-24M (OVERSTOCK - clear/drain)'
        else: b = '25M+ (OVERSTOCK - money frozen)'
        covdist[b] += 1
        vdist[b] += v['dispo'] * v['price']
    crit = covdist['0-3M (FAST - sell fast)']; low = covdist['4-6M (MEDIUM - healthy)']
    healthy = covdist['7-12M (SLOW - watch)']
    overall_over = sum(covdist[b] for b in ['13-24M (OVERSTOCK - clear/drain)', '25M+ (OVERSTOCK - money frozen)'] + ['NO SALES (dead money)'])
    ws.cell(19, 2, 'CRITICAL (< 3 months)'); ws.cell(19, 3, crit)
    ws.cell(20, 2, 'LOW (3-6 months)'); ws.cell(20, 3, low)
    ws.cell(21, 2, 'HEALTHY (6-12 months)'); ws.cell(21, 3, healthy)
    ws.cell(22, 2, 'OVERSTOCKED (>12) / NO SALES'); ws.cell(22, 3, overall_over)
    for b in bucketnames:
        rr_ = rowmap[b]
        ws.cell(rr_, 6, b)
        gc = ws.cell(rr_, 7, vdist[b]); gc.number_format = '#,##0'
    ws.cell(25, 2, 'TOP 15 BY UNITS SOLD 12M')
    top = [(rr, v) for rr, v in RES.items() if v['m12'] > 0]
    top.sort(key=lambda x: x[1]['m12'], reverse=True)
    for c, t in enumerate(['REF','DESIGNATION','BRAND','SOLD 12M','DISPO'], 2):
        ws.cell(26, c, t).fill = NAVY_H; ws.cell(26, c).font = WHITE_B
    r = 27
    for rr, v in top[:15]:
        ws.cell(r, 2, rr); ws.cell(r, 3, v['des']); ws.cell(r, 4, v['brand'])
        ws.cell(r, 5, v['m12']).number_format = '#,##0'
        ws.cell(r, 6, v['dispo']).number_format = '#,##0'
        r += 1
    ws.cell(43, 6, 'TOP 15 DEAD-STUCK RISK BY VALUE (zero sales 12M, money on shelf)')
    for c, t in enumerate(['REF','DESIGNATION','QTY STUCK','VALUE STUCK'], 6):
        ws.cell(44, c, t).fill = NAVY_H; ws.cell(44, c).font = WHITE_B
    ds = [(rr, v) for rr, v in RES.items() if v['cat'] == 'DEAD-STUCK']
    ds.sort(key=lambda x: x[1]['dispo'] * x[1]['price'], reverse=True)
    r = 45
    for rr, v in ds[:15]:
        ws.cell(r, 6, rr); ws.cell(r, 7, v['des'])
        ws.cell(r, 8, v['dispo']).number_format = '#,##0'
        ws.cell(r, 9, v['dispo'] * v['price']).number_format = '#,##0'
        r += 1
    ws.cell(60, 2, 'ROTATION BY BRAND (top 20 by stock value)')
    for c, t in enumerate(['BRAND','ARTICLES','UNITS 12M','STOCK VALUE'], 2):
        ws.cell(61, c, t).fill = NAVY_H; ws.cell(61, c).font = WHITE_B
    from collections import defaultdict
    barts = defaultdict(int); bunits = defaultdict(int); bval = defaultdict(int)
    for rr, v in RES.items():
        if not (v['grand'] > 0 or v['dispo'] > 0):
            continue
        b = v['brand']
        barts[b] += 1; bunits[b] += v['m12']; bval[b] += v['dispo'] * v['price']
    bl = sorted(barts, key=lambda b: bval[b], reverse=True)
    r = 62
    for b in bl[:20]:
        ws.cell(r, 2, b)
        ws.cell(r, 3, barts[b])
        ws.cell(r, 4, bunits[b]).number_format = '#,##0'
        ws.cell(r, 5, bval[b]).number_format = '#,##0'
        r += 1
    DASH_LAST = r

    # ================= HOW IT WORKS + FORMULAS =================
    HOW = [("A1", "Back to INDEX"),
     ("B2", "STOCK ROTATION ANALYSIS - HOW IT WORKS (plain language)"),
     ("B3", "Built %s from Invoices%s%s. Real stock quantities from the STOCK file included (0 = absent)."
            % (REF_DATE.strftime('%d-%m-%Y'), ' + Orders export' if HAS_ORDERS else ' only (no Orders export -> sales-only)', '; NO stock file -> every article treated OUT OF STOCK, judged on history' if not HAS_STOCK else '')),
     ("B4", "Supplier purchases from the previous workbook are carried over as context: DELIVERED SINCE 2025 and LAST BUY columns on the main sheet."),
     ("B5", "REAL STOCK: the STOCK file quantities OVERRIDE the invoice QTY DISPO (your physical shelf = the truth). Articles in that file that never appear in any invoice are added automatically and flagged NEVER MOVED."),
     ("B6", "INDEX sheet: menu of every supplier sheet (one REORDER per brand) + the STOP BUYING list. Click to jump."),
     ("B7", "Each \"REORDER - <BRAND>\" sheet: only FAST refs (3 months of stock or less, or out-of-stock best-sellers) that pass the gates (not PUSHED, ORDERS >= SOLD in 2026/2025/2024). SUGGESTED ORDER = 3-month smoothed demand (orders preferred) + capped safety - current stock, with a 3-year unserved-backlog floor. Coverage horizon = lead time (default 2 months) + 1 month review period = 3 months. Exception: 6 months when the item was out of stock all current year with no current-year orders."),
     ("B8", "STOP BUYING sheet: OVERSTOCK refs (13+ months) repurchased in the last 90 days - block them from the next PO."),
     ("B9", "IN STOCK - ZERO SALES sheet: every article on the shelf with NO sales date in the whole history (never moved). Delist/discount/return them; this is dead money."),
     ("B10", "OVERSTOCK sheet: every overstock ref split SELLING (>=50 units/12m, stop ordering, let drain) vs NOT SELLING (discount to clear)."),
     ("B12", "THE BIG IDEA"),
     ("B13", "Every article is judged by ONE question: at its real selling speed, how many MONTHS will the current shelf stock last?"),
     ("B14", "That number is COUVERTURE. It decides the category, the color and the recommended action."),
     ("B16", "WINDOWS - ROLLING FROM EXPORT DAY"),
     ("B17", "M4 / M8 / M12 = units sold in the last 4 / 8 / 12 months counted BACK FROM the export day."),
     ("B18", "Nothing is excluded, so articles delivered recently and sold right away ARE counted correctly."),
     ("B19", "AVG/MONTH = M12 / 12. COUVERTURE = QTY DISPO / AVG-MONTH. The current month is partial and excluded from stats."),
     ("B21", "CATEGORIES (by months of stock left)"),
     ("B22", "- FAST: 3 months of stock or less (includes out-of-stock best-sellers). Needs reordering NOW."),
     ("B23", "- PUSHED: looked FAST by the shelf, but you SOLD at least 120% of what clients ordered (order base under 300/yr) - the movement is pushed/discounted, not real demand. Kept OUT of every reorder list."),
     ("B24", "- MEDIUM: 4 to 6 months. Comfortable, no order needed."),
     ("B25", "- SLOW: 7 to 12 months. Order only on customer demand."),
     ("B26", "- OVERSTOCK: 13 months or more. Strong seller -> STOP ORDERING, let it drain. Weak seller -> DO NOT ORDER, discount to clear."),
     ("B27", "- DEAD-STUCK: zero sales in 12+ months BUT stock on the shelf. Money frozen -> liquidate."),
     ("B28", "- DEAD: zero sales in 12 months, zero stock, and no sales history at all. Forget it."),
     ("B29", "- Zero stock but HAS sales history: classified by its movement (FAST = reorder according to how it sold), never archived."),
     ("B31", "OTHER COLUMNS"),
     ("B32", "- UNIT PRICE = most recent selling price. STOCK VALUE = QTY DISPO x UNIT PRICE."),
     ("B33", "- LAST SALE: red = nothing sold for more than 6 months, orange = 3 to 6 months."),
     ("B34", "- TREND 3M: compares last 2 months vs the 2 months before: UP green, DOWN red, STABLE grey."),
     ("B35", "- DELIVERED SINCE 2025 = units purchased from suppliers since Jan-2025. LAST BUY = the most recent supplier purchase date (carried over)."),
     ("B36", "- Use DELIVERED / LAST BUY to double-check: are you still buying a ref that the analysis says is OVERSTOCK (13+ months)?"),
     ("B37", "- GRAND TOTAL = all units sold since Jan-24. SOLD BEFORE 2024 = older history before 2024."),
     ("B39", "REORDER SHEET STATS COLUMNS"),
     ("B40", "- AVG/MONTH(12M) = units sold per month in the last 12 months. PEAK 3M/M = best consecutive 3-month average in 2024-2026 (shows burst capacity)."),
     ("B41", "- SOLD LAST 2M/M = your selling pace in the LAST 2 months. SOLD PRIOR 2M/M = the 2 months BEFORE that - compare: is the article speeding up or slowing?"),
     ("B42", "- PURCHASED 12M = supplier quantity received in the last 12 months. DELIVERED SINCE 2025 = total received since Jan-2025 (carried over)."),
     ("B43", "- FORECAST 12M = expected units in the next 12 months = smoothed demand (orders preferred, else sales) x 12."),
     ("B44", "- EST STOCK-OUT = the month the article hits zero, stepping month-by-month with its seasonal profile. NOW = already empty."),
     ("B45", "- SUGGESTED = 3-month smoothed demand + safety - current stock, never more than 2x cover. If orders exceed sales, a backlog floor covers the unserved gap (last 3 years), capped at 3 months of the strongest recent year's pace. Exception: 6 months when the item was out of stock all current year with no current-year orders."),
     ("B46", "- SEASONALITY: each month gets a factor vs the article average (Jan..Dec). Low-data articles get no seasonality (flat factor = 1)."),
     ("B47", "- Monthly grid from the export (last month partial, shown but windows do not depend on it)."),
     ("B49", "COLORS"),
     ("B50", "- QTY DISPO: red = zero stock, blue = has stock."),
     ("B51", "- COUVERTURE: red 3 or less - orange 4-6 - green 7-12 - grey 13 and above."),
     ("B52", "- CATEGORY: red FAST - amber PUSHED - orange MEDIUM - green SLOW - grey OVERSTOCK - red DEAD-STUCK."),
     ("B54", "ACTIONS SHEET - what to do, article by article:"),
     ("B55", "- URGENT - BEST SELLER OUT OF STOCK (red): sells well, zero shelf. Order TODAY."),
     ("B56", "- REORDER NOW (yellow): stock finishes in under 3 months. NO REORDER if the shelf already covers 4+ months at the TRUE pace."),
     ("B57", "- MAINTAIN (green): healthy, do nothing."),
     ("B58", "- WATCH (green): slow mover, buy only when a customer asks."),
     ("B59", "- STOP ORDERING / DO NOT ORDER (grey): 13+ months of stock."),
     ("B60", "- LIQUIDATE (red): no sales for a year but money sleeping on the shelf. Discount or return to supplier."),
     ("B61", "- ARCHIVE: no stock, no sales."),
     ("B62", "- SUGGESTED ORDER: yellow = small order - red = big order (1000+ units). BRAND column shows the supplier family."),
     ("B63", "- ALERT LOST SALES /MO: money you lose each month because best-sellers are empty (their speed x price)."),
     ("B64", "- MONEY YOU OWN ribbon (DASHBOARD): how much of your cash sits at 0-3, 4-6, 7-12, 13-24, 25+ and no-sales months of stock."),
     ("B66", "NOTE: everything is baked as static VALUES and colors (no formulas, no conditional formatting) so Excel never complains.")]
    ws = wb.create_sheet('HOW IT WORKS')
    ws.column_dimensions['A'].width = 13; ws.column_dimensions['B'].width = 130
    for coord, text in HOW:
        ws[coord] = text
    ws['A1'] = 'Back to INDEX'; ws['A1'].fill = fill('00EDEDED')
    for rr in (2, 12):
        cc = 'B%d' % rr
        ws[cc].font = Font(bold=True)

    FR = [
     ("SHEET","COLUMN","FORMULA EXACT","PLAIN MEANING"),
     ("A-Z ROTATION","QTY DISPO","REAL stock file quantity, absent = 0.","Units on the shelf today. Red = 0."),
     ("A-Z ROTATION","UNIT PRICE","MOST RECENT selling price per unit (last invoice price).","What clients pay for 1 unit."),
     ("A-Z ROTATION","STOCK VALUE","QTY DISPO x UNIT PRICE.","Value of the shelf stock in DA."),
     ("A-Z ROTATION","M4 / M8 / M12","SUM of units sold in the last 122 / 244 / 365 days rolling back from REF_DATE.","Sold in recent months (4 / 8 / 12)."),
     ("A-Z ROTATION","AVG/MONTH","M12 / 12.","Per-month average of the last 12 months."),
     ("A-Z ROTATION","COUVERTURE","QTY DISPO / AVG-MONTH (months). None when no 12M sales.","Months until the shelf empties. The key number."),
     ("A-Z ROTATION","LAST SALE","The most recent invoice date for the REF (real date, not month start).","Last time it sold."),
     ("A-Z ROTATION","TREND 3M","Compare last-2-month rate (ta/2) vs the 2 months before (tb/2): >=1.15x UP, <=0.85x DOWN, else STABLE.","Is it speeding up or slowing down?"),
     ("A-Z ROTATION","CATEGORY","FAST: cov<=3. MEDIUM: cov<=6. SLOW: cov<=12. OVERSTOCK: cov>12. DEAD-STUCK: no 12M sales + stock. DEAD: no sales + 0 stock. PUSHED: SOLD>=120%% of ORDERS with order base <300/yr -> OUT of reorder.","The verdict."),
     ("A-Z ROTATION","SALES VALUE 12M","M12 x UNIT PRICE.","Cash made on it in a year."),
     ("A-Z ROTATION","GRAND TOTAL","SUM of units sold since the grid start.","Everything sold in the window."),
     ("A-Z ROTATION","SOLD BEFORE 2024","SUM of units sold before JAN-24 (older history).","Old sales history."),
     ("A-Z ROTATION","DELIVERED SINCE 2025","Supplier quantities received since JAN-2025 (carried over).","What you bought since 2025."),
     ("A-Z ROTATION","LAST BUY","Most recent supplier purchase date (carried over).","Last time a supplier delivered it."),
     ("REORDER","AVG/MONTH(12M)","avg = M12 / 12.","Your normal selling speed (units/month)."),
     ("REORDER","PEAK 3M/M","best3 / 3 where best3 = MAX over every 3 consecutive months in the grid.","Best-ever 3-month run: burst capacity."),
     ("REORDER","SOLD LAST 2M/M","last2 = ta / 2 (units in the ~2 months before REF_DATE).","Most recent 2-month pace."),
     ("REORDER","SOLD PRIOR 2M/M","prior2 = tb / 2 (the 2 months BEFORE that window).","Compare = trend."),
     ("REORDER","PURCHASED 12M","Supplier quantities in last 12 months (carried over).","Bought recently vs sold recently."),
     ("REORDER","DELIVERED SINCE 2025","sum received since JAN-2025 (carried over).","Total bought since 2025."),
     ("REORDER","FORECAST 12M","round(rate x 12); rate = MAX(12M avg, long-run avg, last-2-month), raised to ORDERS 12M / 12 when FULFILL < 80%.","Expected units next 12 months - the REAL market pace."),
     ("REORDER","ORDERS 12M","Sum of client orders in the Orders export for the last 12 months.","How much clients WANTED in the last year."),
     ("REORDER","FULFILL %","Invoiced last 12M / ORDERS 12M. Orange when < 80% = supply-constrained.","How much demand you fulfilled. Low = genuinely wanted."),
     ("REORDER","BURST NEXT 3M","IF intermittent (6-23 months with sales): avg_event x look. ELSE 0.","Burst insurance on empty shelves."),
     ("REORDER","COUVERTURE","QTY DISPO / AVG-MONTH.","Months of runway left."),
     ("REORDER","EST STOCK-OUT","Step month by month (seasonal factor) until <=0. NOW = already 0.","When it empties if you do nothing."),
     ("REORDER","SUGGESTED ORDER","smoothed demand (orders preferred) x 3 + capped safety - DISPO, capped at 2x cover; raised to the unserved-backlog floor when orders exceed sales.","How many to buy next (3m cover + safety - stock, with backlog floor). Backlog floor uses 6 months for stockout items with no current-year orders."),
     ("REORDER","CLIENT ORDERED (30D)","Sum of client orders last 30 days.","How much clients asked for recently."),
     ("REORDER","SOLD (30D)","Sum of invoiced quantities last 30 days.","How much you actually delivered recently."),
     ("REORDER","BACKLOG (3Y)","max(0, ordered - sold - stock) over the last 3 years, capped at 3 months of the strongest recent year's pace (6 months for stockout items with no current-year orders) and rounded. Orange when >0.","Unserved demand used as the order floor."),
     ("REORDER","QTY TO ORDER","Typed by you (yellow). Final decision.","The quantity to send to the supplier."),
     ("REORDER","UNIT PRICE","Most recent selling price.","Price per unit."),
     ("REORDER","LAST SALE","Most recent invoice date (real).","Last time it sold."),
     ("REORDER","LOST/MO (if OOS)","round(rate x UNIT PRICE), only when QTY DISPO = 0.","DA lost per month while out of stock."),
     ("REORDER","STATUS","URGENT: FAST + 0 stock. REORDER NOW: FAST + suggested>0. AT TARGET: enough stock.","What to do."),
     ("REORDER","rows order","Sorted by DEMAND first: biggest client demand pace (units/month) on top; out-of-stock wins ties; stocked ties by lowest coverage.","Demand-first: biggest mover wins."),
     ("CONCEPTS","SOLID vs PUSHED","Reorder gate applied this build: ORDERS >= SOLD in each year (2026, 2025, 2024) AND not PUSHED (SOLD>=120%% ORDERS, base <300/yr). High-order articles (300+/yr) stay.","Pushed junk leaves the reorder lists; genuinely ordered products never disappear."),
    ]
    wsx = wb.create_sheet('FORMULAS')
    wsx.column_dimensions['A'].width = 16; wsx.column_dimensions['B'].width = 26
    wsx.column_dimensions['C'].width = 90; wsx.column_dimensions['D'].width = 60
    wsx['A1'] = 'Back to INDEX'; wsx['A1'].fill = fill('00EDEDED')
    wsx['A2'] = 'EVERY COLUMN AND THE EXACT FORMULA BEHIND IT - REORDER SHEETS + A-Z + DASHBOARD + ALERTS'
    wsx['A3'] = ('REF_DATE = export day (%s) | full months = %d months ending the previous month | '
                 'LEAD TIME default = 2 months + 1-month review = 3-month order cover. '
                 'Exception: 6 months for stockout items with no current-year orders.'
                 % (REF_DATE.strftime('%d-%m-%Y'), len(GRID) - 1))
    for rr, row in enumerate(FR, 5):
        for cc, t in enumerate(row, 1):
            cell = wsx.cell(rr, cc, t)
            if rr == 5:
                cell.fill = NAVY_H; cell.font = WHITE_B

    # ================= SAVE =================
    seen = set()
    for ws in wb.worksheets:
        if ws.title in seen:
            wb.remove(ws)
        else:
            seen.add(ws.title)
    for sht in wb.worksheets:
        c = sht['A1']
        if isinstance(c.value, str) and c.value.startswith('Back to INDEX'):
            link_cell(c, 'INDEX')
    wb._fills = [_PF() if f is None else f for f in wb._fills]
    wb.save(out_path)
    print('SAVED', out_path)
    print('AZ ROTATION rows:', AZ_LAST - 2)
    print('ACTIONS rows:', ACT_LAST - 2)
    print('brand counts:', brand_counts2)
    print('grand total refs:', sum(brand_counts2.values()), 'units:', sum(brand_units2.values()))
    print('stop buying:', len(sb), 'zero sales:', len(zs), 'overstock:', len(ov))
    return out_path

# ---------------------------------------------------------------------------
# 4) BRAND REORDER FILE  (from brand_tool, with STATUS + keyword)
# ---------------------------------------------------------------------------
def build_brand_file(brand, n, kw=None, outdir=None):
    global RES
    brand = brand.strip().upper()
    all_flag = str(n).strip().upper() == 'ALL'
    nn = int(n) if not all_flag else 10 ** 9
    kw = kw.strip().upper() if kw else None
    kw_toks = None
    if kw:
        kw_toks = [t for t in kw.replace("'", ' ').split() if t not in ('DE', 'DU', 'DES', 'LA', 'LE', 'LES')]
    cap = brand_captured(brand)

    cand = {}
    for r, v in RES.items():
        if v['brand'] != brand:
            continue
        if kw_toks and not all(t in v['des'].upper() for t in kw_toks):
            continue
        if v['cat'] != 'FAST' or v['pushed']:
            continue
        if v['dispo'] == 0:
            if v['avg_all'] < 1.0 or v['months_with_sales'] < 1:
                continue
        else:
            if v['m12'] <= 0:
                continue
            if cap and v['ord12'] <= 0:
                continue
            if v['suggested'] <= 0:
                continue
        cand[r] = v
    rows = sorted(cand.items(), key=lambda x: prio_key(x[1]))
    rows = rows[:nn]

    def reason(v):
        if v['pushed']:
            return 'PUSHED - reorder not recommended (sold > 120% of orders)'
        if v['cat'] != 'FAST':
            return 'NOT FAST (%s)' % v['cat']
        if v['dispo'] == 0:
            if v['avg_all'] < 1.0:
                return 'LOW HISTORY - barely moved before'
        else:
            if v['m12'] <= 0:
                return 'NO SALES LAST 12M - stale'
            if cap and v['ord12'] <= 0:
                return 'NO ORDERS LAST 12M'
            if v['suggested'] <= 0:
                return 'AT TARGET - enough stock'
        return ''

    other = []
    if kw_toks:
        for r, v in RES.items():
            if v['brand'] != brand or not all(t in v['des'].upper() for t in kw_toks):
                continue
            if r in cand:
                continue
            other.append((r, v, reason(v)))

    label = 'ALL %s' % len(rows) if (all_flag or len(rows) < nn) else 'TOP %s' % nn
    wb = Workbook()
    ws = wb.active
    ws.title = label + ' ' + brand
    ws.freeze_panes = 'A3'
    hdr = ['RANK', 'REF', 'DESIGNATION', 'QTY (STOCK)', 'PRIX', 'QTY SUGGESTED', 'AVG MONTH(12M)',
           'SOLD LAST 2M', 'COUVERTURE', 'ORDERED 2026', 'SOLD 2026', 'ORDERED 2025', 'SOLD 2025',
           'ORDERED 2024', 'SOLD 2024', 'LAST SALE', 'STATUS']
    for c, t in enumerate(hdr, 1):
        cell = ws.cell(3, c, t)
        cell.fill = fill(NAVY)
        cell.font = Font(color='FFFFFF', bold=True)
    ws['A1'] = ('%s - REORDER LIST (%s) - %s - ranked by DEMAND: biggest client demand pace (units/month) first, out-of-stock wins ties' % (brand, len(rows), kw if kw else 'all articles'))
    ws['A2'] = ('Rule: FAST movers, not pushed, needing stock. OOS (0 in stock) judged on history (moved >=1/mo before); in stock judged on 12M sales. RANK 1 = top demand: current-year client orders & sales per month, then 2025, half 2024. Orders count as demand even where sales lag. Top 100 = rows with RANK <= 100.' + (' Below the ranked list: matching articles NOT to reorder now, with reason.' if other else ''))

    def num(r2, c, v, fmt2='#,##0'):
        x = ws.cell(r2, c, v); x.number_format = fmt2
        return x

    r = 4
    for rank, (ref, v) in enumerate(rows, 1):
        sy = v['sold_y']; oy = v['ord_y']
        cov = v['cov']
        rk = ws.cell(r, 1, rank); rk.number_format = '0'; rk.font = Font(bold=True)
        ws.cell(r, 2, ref)
        ws.cell(r, 3, v['des'])
        num(r, 4, v['dispo'])
        num(r, 5, v['price'])
        sc = num(r, 6, v['suggested'])
        sc.fill = fill(YELLOW)
        num(r, 7, round(v['avg'], 1), '0.0')
        num(r, 8, sum(int(q) for q in v['grid'][-3:-1]))
        cc = num(r, 9, round(cov, 1) if cov is not None else '')
        cc.fill = cov_fill(cov)
        num(r, 10, oy.get(2026, 0))
        num(r, 11, sy.get(2026, 0))
        num(r, 12, oy.get(2025, 0))
        num(r, 13, sy.get(2025, 0))
        num(r, 14, oy.get(2024, 0))
        num(r, 15, sy.get(2024, 0))
        num(r, 16, v['last'], 'DD/MM/YYYY')
        ws.cell(r, 17, 'REORDER').font = Font(bold=True)
        if v['dispo'] == 0:
            ws.cell(r, 2).fill = fill(RED)
            ws.cell(r, 4).fill = fill(RED)
        r += 1

    if other:
        # separator
        sep = r
        r += 1
        h = ws.cell(r, 1, 'NOT FOR REORDER NOW (%d) - why:' % len(other))
        h.font = Font(bold=True, color='FFFFFF'); h.fill = fill(GREY)
        r += 1
        for ref, v, why in sorted(other, key=lambda x: (x[2].split('(')[0], x[1]['dispo'])):
            cov = v['cov']
            ws.cell(r, 1, '-').font = Font(color='999999')
            ws.cell(r, 2, ref)
            ws.cell(r, 3, v['des'])
            num(r, 4, v['dispo'])
            num(r, 5, v['price'])
            num(r, 6, v['suggested'])
            num(r, 7, round(v['avg'], 1), '0.0')
            num(r, 8, sum(int(q) for q in v['grid'][-3:-1]))
            cc = num(r, 9, round(cov, 1) if cov is not None else '')
            cc.fill = cov_fill(cov)
            num(r, 10, v['ord_y'].get(2026, 0))
            num(r, 11, v['sold_y'].get(2026, 0))
            num(r, 12, v['ord_y'].get(2025, 0))
            num(r, 13, v['sold_y'].get(2025, 0))
            num(r, 14, v['ord_y'].get(2024, 0))
            num(r, 15, v['sold_y'].get(2024, 0))
            num(r, 16, v['last'], 'DD/MM/YYYY')
            ws.cell(r, 17, why)
            r += 1

    tot = ws.cell(r, 1, 'TOTAL SUGGESTED'); tot.font = Font(bold=True)
    num(r, 6, sum(x['suggested'] for _, x in rows))
    ws.cell(r, 5, 'total value:').font = Font(bold=True)
    tc = ws.cell(r + 1, 1, 'units x price:'); tc.font = Font(bold=True)
    num(r + 1, 6, sum(x['suggested'] * x['price'] for _, x in rows))

    if outdir is None:
        outdir = OUTDIR
    brand_folder = unique_path(outdir, '%s REORDER' % brand, '', stamp=True)
    os.makedirs(brand_folder, exist_ok=True)
    fname = 'RECOMMANDATION - COMMANDE %s - %s%s.xlsx' % (brand, label, ' - %s' % kw if kw else '')
    out_file = os.path.join(brand_folder, fname)
    wb.save(out_file)
    print('saved: %s  (%d articles)' % (out_file, len(rows)))
    return out_file

# ---------------------------------------------------------------------------
# 5) INTERACTIVE UI
# ---------------------------------------------------------------------------
def _clean_brand(value):
    b = str(value).strip() if value is not None else ""
    if not b or b.lower() in ("nan", "none", "null"):
        return ""
    return b


def available_brands():
    brands = {_clean_brand(v.get('brand')) for v in RES.values()}
    brands.discard("")
    known = [b for b in BRANDS if b in brands]
    return known + sorted(brands - set(BRANDS))

def choose_brand():
    brands = available_brands()
    if not brands:
        print('no brands found in the invoices')
        return None
    print('\nBrands with data:')
    for i, b in enumerate(brands, 1):
        print('  %2d. %s' % (i, b))
    v = input('choose brand (number or name, or ALL = every brand): ').strip()
    if v.upper() == 'ALL':
        return 'ALL'
    for i, b in enumerate(brands, 1):
        if v == str(i):
            return b
    vv = v.upper()
    for b in brands:
        if b.upper() == vv or b.upper().startswith(vv):
            return b
    return None

def menu():
    print()
    print('=' * 60)
    print('STOCK ROTATION - REORDER MANAGER')
    print('=' * 60)
    print('1  Build FULL ANALYSIS workbook (all sheets)')
    print('2  Build REORDER list - one brand')
    print('3  Build REORDER lists - ALL brands')
    print('4  Build keyword/family list - one brand (e.g. DISQUE DE FREIN)')
    print('5  Show current data files')
    print('6  Change SAVE folder (current: %s)' % OUTDIR)
    print('0  Quit')

def main():
    global RES, CORE_PKL, OUTDIR
    print('=' * 60)
    print('REORDER MANAGER - single interactive pipeline')
    print('=' * 60)

    # default file discovery
    def auto_detect():
        inv = find_default(X, ['Invoices_Export_20*.xlsx', 'Invoices*.xlsx'], 'invoices')
        ord = find_default(X, ['Orders_Export_20*.xlsx', 'Orders*.xlsx'], 'orders')
        stoc = find_default(X, ['STOCK *.xls*', 'Stock *.xls*'], 'stock')
        return inv, ord, stoc

    inv_d, ord_d, stoc_d = auto_detect()

    print("\nSTEP 1 - DATA FILES (Enter = accept default, SKIP = don't use)")
    inv_path = ask_path('  Invoices (MANDATORY)', inv_d, True)
    if not inv_path:
        while True:
            p = input('  Invoices path must be provided: ').strip()
            if p and os.path.exists(p):
                inv_path = p
                break
            print('  not found')
    ord_path = ask_path('  Orders (optional)', ord_d, False)
    stoc_path = ask_path('  Stock (optional)', stoc_d, False)

    ref = parse_ref_date(inv_path) or REF_DATE
    print('  REF_DATE = %s' % ref)

    carry = build_carry()
    build_core(inv_path, ord_path, stoc_path, ref_date=ref)
    with open(CORE_PKL, 'wb') as f:
        pickle.dump({'ref_date': str(ref), 'grid': GRID, 'res': RES}, f)

    while True:
        menu()
        ch = input('choice: ').strip()
        if ch == '0':
            print('bye')
            return
        if ch == '5':
            print('  invoices:', inv_path)
            print('  orders:  ', ord_path, '-> SALES ONLY' if not ord_path else '')
            print('  stock:   ', stoc_path, '-> ALL OOS (history)' if not stoc_path else '')
            print('  save to: ', OUTDIR)
            continue
        if ch == '6':
            v = input('  save folder [%s]: ' % OUTDIR).strip()
            if v:
                v = os.path.abspath(v)
                os.makedirs(v, exist_ok=True)
                OUTDIR = v
                print('  results will be saved to: %s' % OUTDIR)
            continue
        if ch == '1':
            build_workbook(carry)
            continue
        if ch == '2':
            b = choose_brand()
            if not b:
                continue
            if b == 'ALL':
                for bb in available_brands():
                    build_brand_file(bb, 'ALL')
            else:
                build_brand_file(b, 'ALL')
            continue
        if ch == '3':
            for bb in available_brands():
                build_brand_file(bb, 'ALL')
            continue
        if ch == '4':
            b = choose_brand()
            if not b or b == 'ALL':
                print('choose a single brand for a keyword list')
                continue
            kw = input('keyword (e.g. DISQUE DE FREIN): ').strip()
            if not kw:
                continue
            lim = input('limit (number of rows or ALL) [ALL]: ').strip() or 'ALL'
            build_brand_file(b, lim, kw)
            continue
        print('unknown choice')


class _Child:
    def __init__(self, proc, q, jid):
        self.proc = proc
        self.q = q
        self.jid = jid

    def pump(self):
        threading.Thread(target=self._read, daemon=True).start()
        threading.Thread(target=self._wait, daemon=True).start()

    def _read(self):
        for raw in self.proc.stdout:
            line = raw.rstrip()
            if line:
                self.q.put(("log", line))

    def _wait(self):
        code = self.proc.wait()
        self.q.put(("exit", (self.jid, code)))


def run_job(spec_path):
    global OUTDIR, RES, GRID, REF_DATE, M26, HAS_ORDERS, HAS_STOCK, INV_N, ORD_N
    with open(spec_path, "r", encoding="utf-8") as f:
        job = json.load(f)
    OUTDIR = job.get("out") or OUTDIR
    mode = job["mode"]

    if mode == "brands":
        df = pd.read_excel(job["inv"], usecols=["Brand"])
        brands = sorted({b for b in (_clean_brand(x) for x in df["Brand"]) if b}, key=str.casefold)
        print("__BRANDS__ %d|%s" % (len(brands), "|".join(brands)))
        return 0

    if mode == "load":
        ref = dt.date.fromisoformat(job["ref"])
        old = job.get("old") or ""
        carry = build_carry(carry_path=("" if old else CARRY_PKL), master_path=old)
        with open(os.path.join(TMP, 'sr_carry.pkl'), "wb") as f:
            pickle.dump(carry, f)
        build_core(job["inv"], job.get("ord") or None, job.get("stk") or None, ref_date=ref)
        with open(os.path.join(TMP, 'core_manager.pkl'), "wb") as f:
            pickle.dump({
                "ref_date": str(ref), "grid": GRID, "res": RES,
                "has_orders": HAS_ORDERS, "has_stock": HAS_STOCK,
                "inv_n": INV_N, "ord_n": ORD_N,
            }, f)
        brands = available_brands()
        print('__BRANDS__ %d|%s' % (len(brands), '|'.join(brands)))
        return 0

    with open(job["core"], "rb") as f:
        blob = pickle.load(f)
    RES = blob["res"]
    GRID = blob["grid"]
    REF_DATE = dt.date.fromisoformat(blob["ref_date"])
    M26 = ((REF_DATE - dt.date(REF_DATE.year, 1, 1)).days + 1) / 30.44
    HAS_ORDERS = blob.get("has_orders", True)
    HAS_STOCK = blob.get("has_stock", True)
    INV_N = blob.get("inv_n", 0)
    ORD_N = blob.get("ord_n", 0)
    carry_path = job.get("carry") or os.path.join(TMP, 'sr_carry.pkl')
    carry = {}
    if os.path.exists(carry_path):
        with open(carry_path, "rb") as f:
            carry = pickle.load(f)

    if mode == "full":
        os.makedirs(OUTDIR, exist_ok=True)
        build_workbook(carry, out_path=unique_path(OUTDIR, 'FULL STOCK ROTATION', 'xlsx'))
    elif mode == "brand":
        build_brand_file(job["brand"], job["limit"], job.get("kw") or None, outdir=OUTDIR)
    elif mode == "all":
        brands = available_brands()
        for b in brands:
            build_brand_file(b, "ALL", None, outdir=OUTDIR)
        print('DONE all brands: %d file(s) in %s' % (len(brands), OUTDIR))
    return 0


class StockRotationApp:
    FILE_ROWS = [
        ("inv", "Invoices export *", "REQUIRED - units actually sold"),
        ("ord", "Orders export", "optional - client demand, much better results"),
        ("stk", "Current stock", "optional - real quantities on the shelf"),
        ("old", "Old Stock Rotation", "optional - LAST BUY / DELIVERED / PURCHASED 12M"),
    ]
    PATTERNS = {
        "inv": ["Invoices_Export_20*.xlsx", "Invoices*.xlsx"],
        "ord": ["Orders_Export_20*.xlsx", "Orders*.xlsx"],
        "stk": ["STOCK *.xls*", "Stock *.xls*"],
        "old": ["STOCK ROTATION - FULL ANALYSIS *.xlsx", "FULL ANALYSIS*.xlsx"],
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("STOCK ROTATION - Reorder Manager")
        self.root.geometry("820x960")
        self.root.minsize(780, 720)
        self.queue = queue.Queue()
        self.vars = {}
        self.child = None
        self.carry_path = ""
        self.core_path = os.path.join(TMP, 'core_manager.pkl')
        self.last_file = ""
        self.job_spec = ""
        self.job_on_done = None
        self.job_quiet = False
        self.job_id = 0
        self.brand_list = []
        self.loaded = False
        self.buttons = []
        self.init_styles()
        self.build_ui()
        self._out_dir()
        self.say("Nothing is selected yet - pick your files above, then press LOAD + ANALYSE THE DATA.")
        self.say("Default folder searched when you browse: %s" % X)
        self.set_loaded(False)
        self.root.after(100, self.poll)

    def init_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.style = style
        self.palette = {
            "load": ("#1f618d", "#154360"),
            "full": ("#c0392b", "#922b21"),
            "brand": ("#935116", "#6e3f0f"),
            "kw": ("#1a5276", "#11303d"),
        }
        for name, (bg, active) in self.palette.items():
            style.configure(
                "big_%s.TButton" % name,
                background=bg,
                foreground="white",
                font=("Segoe UI", 10, "bold"),
                padding=(12, 8),
                anchor="w",
            )
            style.map(
                "big_%s.TButton" % name,
                background=[("pressed", active), ("active", active), ("disabled", "#dfe3e6")],
                foreground=[("pressed", "white"), ("active", "white"), ("disabled", "#8e9aa0")],
            )

    def _big(self, parent, text, name, command):
        btn = ttk.Button(parent, text=text, style="big_%s.TButton" % name, command=command)
        btn.pack(fill="x", pady=3)
        self.buttons.append(btn)
        return btn

    def _file_row(self, parent, key, row, label, hint, pick, clear):
        ttk.Label(parent, text=label, width=22, anchor="w").grid(
            row=row, column=0, padx=(8, 4), pady=2, sticky="w"
        )
        var = tk.StringVar()
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, padx=4, pady=2, sticky="we")
        ttk.Button(parent, text="Browse...", width=9, command=pick).grid(
            row=row, column=2, padx=(4, 3), pady=2
        )
        ttk.Button(parent, text="Clear", width=7, command=clear).grid(
            row=row, column=3, padx=(0, 4), pady=2
        )
        ttk.Label(parent, text=hint, foreground="#777").grid(
            row=row, column=4, padx=(0, 8), pady=2, sticky="w"
        )
        self.vars[key] = var
        return var

    def build_ui(self):
        frame_files = ttk.LabelFrame(self.root, text="1.  DATA FILES      (* = required, the rest improve the result)")
        frame_files.pack(fill="x", padx=10, pady=(10, 4))
        for i, (key, label, hint) in enumerate(self.FILE_ROWS):
            self._file_row(
                frame_files, key, i, label, hint,
                lambda k=key: self.pick_file(k), lambda k=key: self.clear_file(k),
            )
        frame_files.columnconfigure(1, weight=1)
        self.ref_var = tk.StringVar(value="REF_DATE: -")
        ttk.Label(frame_files, textvariable=self.ref_var, foreground="#555").grid(
            row=4, column=0, columnspan=5, padx=8, pady=(4, 2), sticky="w"
        )
        ttk.Label(
            frame_files,
            text="Folder searched:",
            foreground="#555",
        ).grid(row=5, column=0, padx=(8, 4), pady=(0, 6), sticky="w")
        self.folder_var = tk.StringVar(value=X)
        ttk.Entry(frame_files, textvariable=self.folder_var, state="readonly").grid(
            row=5, column=1, columnspan=4, padx=4, pady=(0, 6), sticky="we"
        )
        frame_load = ttk.Frame(frame_files)
        frame_load.grid(row=6, column=0, columnspan=5, padx=8, pady=(0, 8), sticky="ew")
        self.btn_load = self._big(
            frame_load, "LOAD + ANALYSE THE DATA   (step 1 - do this first)", "load",
            self.load_data,
        )

        frame_out = ttk.LabelFrame(self.root, text="2.  SAVE THE RESULTS WHERE YOU WANT")
        frame_out.pack(fill="x", padx=10, pady=6)
        ttk.Label(frame_out, text="Results go to", width=14, anchor="w").grid(
            row=0, column=0, padx=(8, 4), pady=6, sticky="w"
        )
        self.out_var = tk.StringVar(value=OUTDIR)
        ttk.Entry(frame_out, textvariable=self.out_var, font=("Consolas", 9)).grid(
            row=0, column=1, padx=4, pady=6, sticky="we"
        )
        ttk.Button(frame_out, text="Change folder...", width=15, command=self.pick_out).grid(
            row=0, column=2, padx=(4, 3), pady=6
        )
        ttk.Button(frame_out, text="Open", width=7, command=self.open_out).grid(
            row=0, column=3, padx=(0, 8), pady=6
        )
        frame_out.columnconfigure(1, weight=1)
        self.out_hint = ttk.Label(
            frame_out,
            text="Every result is named with today's date, and a number if you run it twice - nothing is overwritten.",
            foreground="#0b6b2f",
        )
        self.out_hint.grid(row=1, column=0, columnspan=4, padx=8, pady=(0, 6), sticky="w")

        frame_full = ttk.LabelFrame(self.root, text="3.  STOCK ROTATION RESULTS  (the full workbook, every sheet)")
        frame_full.pack(fill="x", padx=10, pady=6)
        self.btn_full = self._big(
            frame_full, "GENERATE FULL STOCK ROTATION ANALYSIS", "full", self.build_full
        )

        frame_brand = ttk.LabelFrame(self.root, text="4.  REORDER LIST FOR ONE BRAND")
        frame_brand.pack(fill="x", padx=10, pady=6)
        rowb = ttk.Frame(frame_brand)
        rowb.pack(fill="x", padx=8, pady=(6, 0))
        ttk.Label(rowb, text="Brand", width=10, anchor="w").pack(side="left")
        self.brand_var = tk.StringVar()
        self.brand_combo = ttk.Combobox(rowb, textvariable=self.brand_var, state="readonly", width=26)
        self.brand_combo.pack(side="left", padx=4)
        ttk.Label(rowb, text="How many rows", anchor="w").pack(side="left", padx=(12, 0))
        self.limit_var = tk.StringVar(value="ALL")
        ttk.Entry(rowb, textvariable=self.limit_var, width=8).pack(side="left", padx=4)
        ttk.Label(rowb, text="(or ALL)", foreground="#777").pack(side="left")
        self.btn_brand = self._big(
            frame_brand, "GENERATE REORDER LIST FOR THIS BRAND", "brand", self.build_brand
        )

        frame_kw = ttk.LabelFrame(self.root, text="5.  KEYWORD / FAMILY LIST")
        frame_kw.pack(fill="x", padx=10, pady=6)
        rowk = ttk.Frame(frame_kw)
        rowk.pack(fill="x", padx=8, pady=(6, 0))
        ttk.Label(rowk, text="Brand", width=10, anchor="w").pack(side="left")
        self.kw_brand_var = tk.StringVar()
        self.kw_brand_combo = ttk.Combobox(rowk, textvariable=self.kw_brand_var, state="readonly", width=26)
        self.kw_brand_combo.pack(side="left", padx=4)
        ttk.Label(rowk, text="Keyword", anchor="w").pack(side="left", padx=(12, 0))
        self.kw_var = tk.StringVar()
        ttk.Entry(rowk, textvariable=self.kw_var, width=24).pack(side="left", padx=4)
        ttk.Label(rowk, text="(e.g. DISQUE DE FREIN)", foreground="#777").pack(side="left")
        self.btn_kw = self._big(
            frame_kw, "GENERATE KEYWORD LIST", "kw", self.build_kw
        )

        frame_prog = ttk.Frame(self.root)
        frame_prog.pack(fill="both", expand=True, padx=10, pady=6)
        self.progress = ttk.Progressbar(frame_prog, mode="indeterminate")
        self.progress.pack(fill="x")
        self.status = tk.StringVar(value="Choose your files, then press LOAD + ANALYSE THE DATA.")
        ttk.Label(frame_prog, textvariable=self.status, foreground="#444").pack(anchor="w", pady=(4, 4))
        log_frame = ttk.Frame(frame_prog)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=6, wrap="none", bg="#fbfbfb", relief="solid", borderwidth=1)
        sy = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        sx = ttk.Scrollbar(log_frame, orient="horizontal", command=self.log.xview)
        self.log.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log.tag_configure("err", foreground="#b00000")
        self.log.tag_configure("ok", foreground="#0b6b2f", font=("Consolas", 9, "bold"))
        self.log.tag_configure("hi", foreground="#0a4f8a", font=("Consolas", 9, "bold"))

        bottom = ttk.Frame(self.root)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bottom, text="Clear log", command=self.clear_log).pack(side="left")
        ttk.Button(bottom, text="Open output folder", command=self.open_out).pack(side="left", padx=(6, 0))
        ttk.Button(bottom, text="Close", command=self.root.destroy).pack(side="right", ipadx=10)

    def say(self, text, tag=None):
        self.log.insert("end", text + "\n", tag or ())
        self.log.see("end")

    def clear_log(self):
        self.log.delete("1.0", "end")

    def set_loaded(self, loaded):
        self.loaded = loaded
        for btn in (self.btn_full, self.btn_brand, self.btn_kw):
            btn.configure(state="normal" if loaded else "disabled")
        if loaded and self.brand_list:
            self._fill_brands()
    def _fill_brands(self, source=""):
        for combo, var in ((self.brand_combo, self.brand_var), (self.kw_brand_combo, self.kw_brand_var)):
            combo.configure(values=self.brand_list or ["no brands found"], state="readonly")
            if self.brand_list and not var.get():
                var.set(self.brand_list[0])
        if not self.brand_list:
            self.say("No brands found in the invoices file.", "err")
            return
        if self.loaded:
            self.say("%d brands available: %s" % (len(self.brand_list), ", ".join(self.brand_list)), "hi")
        else:
            self.say(
                "%d brands found in the invoices%s - now press LOAD + ANALYSE to generate."
                % (len(self.brand_list), (" (%s)" % source) if source else ""),
                "hi",
            )

    def set_busy(self, busy, text=None):
        for btn in self.buttons:
            btn.configure(state="disabled" if busy else "normal")
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.configure(value=0)
        if text:
            self.status.set(text)
        if not busy:
            self.set_loaded(self.loaded)

    def pick_file(self, key):
        path = filedialog.askopenfilename(
            parent=self.root,
            title=self._label(key),
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.vars[key].set(path)
            self.update_ref()
            if key == "inv":
                self.clear_brands()
                self.scan_brands()

    def clear_file(self, key):
        self.vars[key].set("")
        self.update_ref()
        if key == "inv":
            self.cancel_quiet()
            self.clear_brands()
            if self.loaded:
                self.say("Invoices file cleared - press LOAD + ANALYSE again.", "err")
            self.loaded = False
            self.set_loaded(False)

    def clear_brands(self):
        self.brand_list = []
        for combo, var in ((self.brand_combo, self.brand_var), (self.kw_brand_combo, self.kw_brand_var)):
            combo.configure(values=[], state="readonly")
            var.set("")

    def _label(self, key):
        for k, label, _hint in self.FILE_ROWS:
            if k == key:
                return label
        return key

    def pick_out(self):
        path = filedialog.askdirectory(parent=self.root, title="Where should the results be saved?")
        if path:
            self.out_var.set(os.path.abspath(path))
            self._out_dir()
            self.say("Results will be saved to %s" % self.out_var.get())

    def open_out(self):
        path = self.out_var.get().strip() or OUTDIR
        if not os.path.isdir(path):
            messagebox.showerror("Folder not found", path, parent=self.root)
            return
        try:
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("Could not open", str(exc), parent=self.root)

    def update_ref(self):
        path = self.vars["inv"].get().strip()
        ref = parse_ref_date(path) if path else None
        self.ref_var.set(
            "REF_DATE = %s   (read from the invoices filename - every rolling window ends here)"
            % (ref.strftime("%d-%m-%Y") if ref else "-")
        )

    def autodetect(self):
        folder = self.folder_var.get().strip() or X
        if not os.path.isdir(folder):
            self.say("Data folder not found: %s" % folder, "err")
            return
        for key, _label, _hint in self.FILE_ROWS:
            found = find_default(folder, self.PATTERNS[key], key)
            self.vars[key].set(found)
            base = os.path.basename(found) if found else "not found"
            tag = "" if found or key == "inv" else "err"
            self.say("%-22s %s" % (self._label(key) + ":", base), tag or None)
        self.update_ref()
        self.root.after(400, self.scan_brands)

    def _out_dir(self):
        path = os.path.abspath(self.out_var.get().strip() or OUTDIR)
        os.makedirs(path, exist_ok=True)
        self.out_var.set(path)
        self.out_hint.configure(
            text="Saving to: %s   -   every result gets its own name, nothing is overwritten." % path,
            foreground="#0b6b2f",
        )
        return path

    def load_data(self):
        self.cancel_quiet()
        inv = self.vars["inv"].get().strip()
        if not inv or not os.path.exists(inv):
            messagebox.showerror("Invoices file needed", "The Invoices export is REQUIRED - pick it first.",
                                 parent=self.root)
            return
        for key in ("ord", "stk", "old"):
            p = self.vars[key].get().strip()
            if p and not os.path.exists(p):
                messagebox.showerror("File not found", "%s file not found:\n%s" % (self._label(key), p),
                                     parent=self.root)
                return
        ref = parse_ref_date(inv) or REF_DATE
        self.say("")
        self.say("STEP 1 - ANALYSING   REF_DATE %s" % ref.strftime("%d-%m-%Y"), "ok")
        self.say("  invoices : %s" % os.path.basename(inv))
        self.say("  orders   : %s" % (os.path.basename(self.vars["ord"].get().strip()) or "SKIPPED - sales only"))
        self.say("  stock    : %s" % (os.path.basename(self.vars["stk"].get().strip()) or "SKIPPED - all treated out of stock"))
        self.say("  old file : %s" % (os.path.basename(self.vars["old"].get().strip()) or "SKIPPED - no last-buy data"))
        self.set_busy(True, "Analysing - the window stays usable, this runs in a separate process...")
        self._child_job(
            "load",
            inv=inv,
            ord=self.vars["ord"].get().strip(),
            stk=self.vars["stk"].get().strip(),
            old=self.vars["old"].get().strip(),
            ref=ref.isoformat(),
            on_done=self._after_load,
        )

    def _after_load(self, _out):
        self.loaded = True
        self.set_busy(False, "Data loaded - pick what you want to generate.")
        self.say("STEP 1 DONE. Now press one of the GENERATE buttons.", "ok")

    def build_full(self):
        if self._need_data():
            return
        out = self._out_dir()
        self.say("")
        self.say("GENERATING FULL STOCK ROTATION ANALYSIS -> %s" % out, "hi")
        self._child_job("full", out=out)

    def build_brand(self):
        if self._need_data():
            return
        brand = self.brand_var.get().strip()
        if not brand:
            messagebox.showerror("Pick a brand", "Choose a brand first.", parent=self.root)
            return
        limit = self.limit_var.get().strip() or "ALL"
        if limit.upper() != "ALL" and not limit.isdigit():
            messagebox.showerror("Invalid rows", "How many rows must be a number or ALL.", parent=self.root)
            return
        out = self._out_dir()
        self.say("")
        self.say("GENERATING REORDER LIST - %s (%s rows)" % (brand, limit), "hi")
        self._child_job("brand", out=out, brand=brand, limit=limit)

    def build_kw(self):
        if self._need_data():
            return
        brand = self.kw_brand_var.get().strip()
        kw = self.kw_var.get().strip()
        if not brand:
            messagebox.showerror("Pick a brand", "Choose a brand first.", parent=self.root)
            return
        if not kw:
            messagebox.showerror("Keyword needed", "Type the keyword, e.g. DISQUE DE FREIN.", parent=self.root)
            return
        out = self._out_dir()
        self.say("")
        self.say("GENERATING KEYWORD LIST - %s / %s" % (brand, kw), "hi")
        self._child_job("brand", out=out, brand=brand, limit="ALL", kw=kw)

    def _need_data(self):
        if not self.loaded:
            messagebox.showinfo("Load the data first", "Press 'LOAD + ANALYSE THE DATA' first.", parent=self.root)
            return True
        return False

    def _child_job(self, mode, on_done=None, quiet=False, **kw):
        if self.child is not None:
            return
        job = dict(kw)
        job["mode"] = mode
        if mode == "load":
            job["out"] = self.out_var.get().strip() or OUTDIR
        else:
            job["core"] = self.core_path
            job["carry"] = os.path.join(TMP, 'sr_carry.pkl')

        spec = os.path.join(TMP, 'sr_job_%d.json' % int(time.time() * 1000))
        with open(spec, "w", encoding="utf-8") as f:
            json.dump(job, f)
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUNBUFFERED"] = "1"
        proc = subprocess.Popen(
            [sys.executable, os.path.abspath(__file__), "--job", spec],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            env=env,
            bufsize=1,
            universal_newlines=True,
        )
        self.job_id += 1
        self.job_spec = spec
        self.job_on_done = on_done
        self.job_quiet = quiet
        if not quiet:
            self.set_busy(True, "Working...")
        self.child = _Child(proc, self.queue, self.job_id)
        self.child.pump()

    def cancel_quiet(self):
        if self.child is None or not self.job_quiet:
            return False
        try:
            self.child.proc.kill()
        except Exception:
            pass
        self.child = None
        self.job_quiet = False
        self._drop_spec()
        self.say("(brand list scan stopped - the full analysis brings the brands anyway)")
        return True

    def _drop_spec(self):
        try:
            if self.job_spec and os.path.exists(self.job_spec):
                os.remove(self.job_spec)
        except OSError:
            pass
        self.job_spec = ""

    def scan_brands(self):
        inv = self.vars["inv"].get().strip()
        if not inv or not os.path.exists(inv):
            return
        if self.child is not None:
            return
        self.say("Reading the brand list from %s ..." % os.path.basename(inv))
        for combo, var in ((self.brand_combo, self.brand_var), (self.kw_brand_combo, self.kw_brand_var)):
            combo.configure(values=["reading the brand list, please wait..."], state="disabled")
            var.set("")
        self._child_job("brands", quiet=True, inv=inv)

    def poll(self):
        try:
            while True:
                kind = self.queue.get_nowait()
                if kind[0] == "log":
                    line = kind[1]
                    if line.startswith("__BRANDS__ "):
                        _count, _, rest = line.split(" ", 1)[1].partition("|")
                        self.brand_list = [b for b in rest.split("|") if b]
                        self._fill_brands("from the file")
                        continue
                    self.say(line)
                    if line.startswith("SAVED ") or line.startswith("saved: "):
                        self.last_file = line.split(" ", 1)[-1].split("  (")[0].strip()
                    elif "DONE all brands" in line:
                        self.last_file = ""
                elif kind[0] == "exit":
                    jid, code = kind[1]
                    if jid == self.job_id:
                        self._on_exit(code)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def _on_exit(self, code):
        self.child = None
        quiet = self.job_quiet
        self._drop_spec()
        out = getattr(self, "last_file", "")
        self.last_file = ""
        if quiet:
            if code != 0:
                self.say("(brand list scan failed - you will get the full list after LOAD)", "err")
            return
        if code == 0:
            self.say("DONE.", "ok")
            if self.job_on_done:
                self.job_on_done(out)
            else:
                self.set_busy(False, "Finished - see the log above.")
                if out and os.path.isfile(out):
                    if messagebox.askyesno("Finished", "Saved:\n%s\n\nOpen the folder?"):
                        try:
                            os.startfile(os.path.dirname(out))
                        except Exception:
                            pass
                else:
                    try:
                        os.startfile(self._out_dir())
                    except Exception:
                        pass
        else:
            self.set_busy(False, "Failed - the reason is in the log above.")
            messagebox.showerror(
                "Failed", "The job stopped with an error (exit code %s).\nThe full message is in the log panel."
                % code, parent=self.root,
            )


def run_console():
    try:
        main()
    except Exception:
        traceback.print_exc()
    print("\nPress ENTER to close this window.")
    try:
        input()
    except Exception:
        pass


if __name__ == '__main__':
    if '--job' in sys.argv[1:]:
        spec = sys.argv[sys.argv.index('--job') + 1]
        try:
            sys.exit(run_job(spec))
        except SystemExit:
            raise
        except Exception:
            traceback.print_exc()
            sys.exit(1)
    elif '--console' in sys.argv[1:]:
        run_console()
    else:
        StockRotationApp().root.mainloop()
