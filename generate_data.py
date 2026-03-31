"""
Great Clips Weekly Dashboard Data Generator
============================================
Run this script each week to generate a fresh dashboard.

Usage:
    python generate_data.py

Required files in the same folder as this script:
    - Employee_Stats.xlsx         (weekly employee stats export from Great Clips)
    - Employee_Return_Stats.xlsx  (employee customer return stats export)
    - All_Salons.xlsx             (system-wide weekly report)

Output:
    - index.html          (owner view — all salons with tabs)
    - 3750.html           (County Line — manager view)
    - 3800.html           (Braden River — manager view)
    - 3826.html           (Kings Crossing — manager view)
    - 4216.html           (North River Ranch — manager view)
"""

import pandas as pd
import numpy as np
import json
import sys
import os
from datetime import datetime

# ── File names ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

EMP_STATS_FILE   = os.path.join(SCRIPT_DIR, 'Employee_Stats.xlsx')
RETURN_STATS_FILE= os.path.join(SCRIPT_DIR, 'Employee_Return_Stats.xlsx')
SYSTEM_FILE      = os.path.join(SCRIPT_DIR, 'All_Salons.xlsx')
TEMPLATE_FILE    = os.path.join(SCRIPT_DIR, 'template.html')

# ── Salon config ─────────────────────────────────────────────────────────────
SALONS = {
    3800: 'Publix At Braden River',
    3750: 'Publix At County Line Road',
    3826: 'Kings Crossing Publix',
    4216: 'North River Ranch',
}

SALON_SHORT = {
    3800: 'Braden River',
    3750: 'County Line',
    3826: 'Kings Crossing',
    4216: 'North River Ranch',
}

# ── Helpers ──────────────────────────────────────────────────────────────────
def safe(val, digits=1):
    """Return rounded float or None if NaN."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    return round(float(val), digits)

def safe_int(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return 0
    return int(val)

# ── Load files ───────────────────────────────────────────────────────────────
print("Loading files...")
for f in [EMP_STATS_FILE, RETURN_STATS_FILE, SYSTEM_FILE]:
    if not os.path.exists(f):
        print(f"  ERROR: File not found: {f}")
        print(f"  Make sure the file is in the same folder as this script and named correctly.")
        sys.exit(1)
    print(f"  ✓ {os.path.basename(f)}")

def smart_read(filepath, key_col):
    """Read Excel file, auto-detecting the header row by looking for key_col."""
    raw = pd.read_excel(filepath, header=None)
    for i, row in raw.iterrows():
        if key_col in str(row.values):
            return pd.read_excel(filepath, header=i)
    # Fallback to default
    return pd.read_excel(filepath)

df     = smart_read(EMP_STATS_FILE, 'Salon Number')
df2    = smart_read(RETURN_STATS_FILE, 'Salon Number')
sys_df = smart_read(SYSTEM_FILE, 'Salon')

# ── Parse dates ───────────────────────────────────────────────────────────────
if 'Date' in df.columns:
    df['DateParsed'] = pd.to_datetime(
        df['Date'].str.extract(r'Friday, (\d+/\d+/\d+)')[0], format='%m/%d/%Y'
    )
else:
    date_col = None
    for col in df.columns:
        sample = df[col].dropna().astype(str)
        if sample.str.contains('Friday').any():
            date_col = col
            break
    if date_col:
        df['DateParsed'] = pd.to_datetime(
            df[date_col].astype(str).str.extract(r'Friday, (\d+/\d+/\d+)')[0],
            format='%m/%d/%Y', errors='coerce'
        )
        if df['DateParsed'].isna().all():
            raw = pd.read_excel(EMP_STATS_FILE, header=None)
            date_str = raw.iloc[0, 0] if pd.notna(raw.iloc[0, 0]) else raw.iloc[1, 0]
            import re
            match = re.search(r'(\d+/\d+/\d+)\s*$', str(date_str))
            if match:
                parsed_date = pd.to_datetime(match.group(1), format='%m/%d/%Y')
                df['DateParsed'] = parsed_date
    else:
        import os as _os
        mtime = _os.path.getmtime(EMP_STATS_FILE)
        df['DateParsed'] = pd.Timestamp.fromtimestamp(mtime).normalize()

df2['WeekEnd'] = pd.to_datetime(df2['Week Ending Date'])

latest     = df['DateParsed'].max()
latest_ret = df2['WeekEnd'].max()
week_label = latest.strftime('%m/%d/%Y')

print(f"\nLatest employee stats week : {week_label}")
print(f"Latest return stats week   : {latest_ret.strftime('%m/%d/%Y')}")

# ── Current active stylists (worked in last 3 weeks) ─────────────────────────
hist_file = os.path.join(SCRIPT_DIR, 'Employee_Stats_History.xlsx')
if os.path.exists(hist_file):
    df_hist = pd.read_excel(hist_file)
    df_hist['DateParsed'] = pd.to_datetime(
        df_hist['Date'].str.extract(r'Friday, (\d+/\d+/\d+)')[0], format='%m/%d/%Y'
    )
    df_combined = pd.concat([df_hist, df], ignore_index=True).drop_duplicates(
        subset=['Salon Number','Employee','DateParsed']
    )
else:
    df_combined = df

recent = df_combined[df_combined['DateParsed'] >= latest - pd.Timedelta(weeks=2)]
current_stylists = {}
for snum in SALONS:
    active = recent[
        (recent['Salon Number'] == snum) & (recent['Cust Count'] > 0)
    ]['Employee'].unique().tolist()
    current_stylists[snum] = active

# ── Split system file: full history for trends, current week for KPIs ────────
sys_df['WeekDate'] = pd.to_datetime(sys_df['SalonWeekEndingDate'])
sys_hist_all = sys_df.copy()
current_week = sys_df['WeekDate'].max()
sys_df = sys_df[sys_df['WeekDate'] == current_week].copy()
print(f"System file current week: {current_week.strftime('%m/%d/%Y')} ({len(sys_df)} salons)")

# ── System benchmarks ─────────────────────────────────────────────────────────
bench = {
    'avg_cust'          : safe(sys_df['Cust Count'].mean(), 0),
    'avg_cph'           : safe(sys_df['CPH'].mean(), 3),
    'avg_hc_time'       : safe(sys_df['Avg HC Time'].mean(), 2),
    'avg_new_return'    : safe(sys_df['Salon New Cust Return %'].mean() * 100, 1),
    'avg_repeat_return' : safe(sys_df['Salon Repeat Cust Return %'].mean() * 100, 1),
    'avg_yoy'           : safe(sys_df['Cust Count % Change'].mean() * 100, 1),
    'avg_wait'          : safe(sys_df['Avg Wait Time'].mean(), 1),
    'total_salons'      : len(sys_df),
}

# ── KPIs from system file for our 4 salons ────────────────────────────────────
sys_sorted = sys_df.sort_values('Cust Count', ascending=False).reset_index(drop=True)
sys_sorted['Rank'] = sys_sorted.index + 1
total_salons = len(sys_df)

kpis = {}
for snum, sname in SALONS.items():
    row = sys_df[sys_df['Salon'].str.startswith(str(snum))]
    rank_row = sys_sorted[sys_sorted['Salon'].str.startswith(str(snum))]
    if row.empty:
        kpis[str(snum)] = {}
        continue
    r = row.iloc[0]
    rank = int(rank_row.iloc[0]['Rank']) if not rank_row.empty else None
    kpis[str(snum)] = {
        'cust'          : safe_int(r['Cust Count']),
        'new_cust'      : safe_int(r['New Cust']),
        'yoy'           : safe(r['Cust Count % Change'] * 100 if pd.notna(r['Cust Count % Change']) else None, 1),
        'new_return'    : safe(r['Salon New Cust Return %'] * 100 if pd.notna(r['Salon New Cust Return %']) else None, 1),
        'repeat_return' : safe(r['Salon Repeat Cust Return %'] * 100 if pd.notna(r['Salon Repeat Cust Return %']) else None, 1),
        'cph'           : safe(r['CPH'], 2),
        'avg_wait'      : safe(r['Avg Wait Time'], 1),
        'hc_time'       : safe(r['Avg HC Time'], 1),
        'new_cust_pct'  : safe(r['New Cust %'] * 100, 1),
        'rank'          : rank,
        'total_salons'  : total_salons,
    }

# ── Weekly trend (last 12 weeks) with YoY comparison ─────────────────────────
sys_hist = sys_hist_all.copy()
sys_hist['SalonNum'] = sys_hist['Salon'].str[:4].astype(int)

w25 = sys_hist[sys_hist['WeekDate'].dt.year==2025].copy()
w26 = sys_hist[sys_hist['WeekDate'].dt.year==2026].copy()
w25['week'] = w25['WeekDate'].dt.isocalendar().week.astype(int)
w26['week'] = w26['WeekDate'].dt.isocalendar().week.astype(int)

trend = {}
for snum in SALONS:
    s26 = w26[w26['SalonNum']==snum].drop_duplicates('week').sort_values('week').tail(12)
    s25 = w25[w25['SalonNum']==snum].drop_duplicates('week').set_index('week')['Cust Count']
    trend[str(snum)] = {
        'dates' : [d.strftime('%m/%d') for d in s26['WeekDate']],
        'curr'  : [safe_int(x) for x in s26['Cust Count']],
        'prior' : [safe_int(s25.get(w, 0)) for w in s26['week']],
    }

# ── Stylist return stats (last 8 weeks) ───────────────────────────────────────
last8_ret = df2[df2['WeekEnd'] >= latest_ret - pd.Timedelta(weeks=8)]
stylist_ret = last8_ret.groupby(['Salon Number', 'Employee']).agg(
    New_Cust    =('New Cust', 'sum'),
    New_Returns =('New Cust Returns', 'sum'),
    Repeat_Cust =('Repeat Cust', 'sum'),
    Total_Cust  =('Total Cust', 'sum'),
).reset_index()
stylist_ret['new_return_pct'] = (
    stylist_ret['New_Returns'] / stylist_ret['New_Cust'].replace(0, np.nan) * 100
).round(1)
stylist_ret['reg_cust_pct'] = (
    stylist_ret['Repeat_Cust'] / stylist_ret['Total_Cust'].replace(0, np.nan) * 100
).round(1)

# ── Latest week stylist performance ───────────────────────────────────────────
latest_perf = df[df['DateParsed'] == latest].copy()

# ── Build stylist data (current stylists only) ────────────────────────────────
stylists = {}
for snum in SALONS:
    active = current_stylists[snum]

    ret = stylist_ret[
        (stylist_ret['Salon Number'] == snum) &
        (stylist_ret['Employee'].isin(active))
    ][['Employee','new_return_pct','reg_cust_pct']].copy()

    perf = latest_perf[
        (latest_perf['Salon Number'] == snum) &
        (latest_perf['Employee'].isin(active))
    ][['Employee', 'Cuts Per Hour', 'Avg HC Time', 'Cust Count']].copy()

    base = pd.DataFrame({'Employee': active})
    merged = base.merge(ret, on='Employee', how='left').merge(perf, on='Employee', how='left')

    records = []
    for _, r in merged.iterrows():
        records.append({
            'name'       : r['Employee'],
            'new_return' : safe(r.get('new_return_pct')),
            'reg_pct'    : safe(r.get('reg_cust_pct')),
            'cph'        : safe(r.get('Cuts Per Hour')),
            'hc_time'    : safe(r.get('Avg HC Time')),
            'cust_count' : safe_int(r.get('Cust Count', 0)),
        })

    records.sort(key=lambda x: (x['reg_pct'] is None, -(x['reg_pct'] or 0), -(x['cph'] or 0)))
    stylists[str(snum)] = records

# ── Assemble final output ─────────────────────────────────────────────────────
output = {
    'generated'  : datetime.now().strftime('%Y-%m-%d %H:%M'),
    'week_label' : week_label,
    'bench'      : bench,
    'kpis'       : kpis,
    'trend'      : trend,
    'stylists'   : stylists,
    'salon_names': {str(k): v for k, v in SALON_SHORT.items()},
}

# ── Read template ─────────────────────────────────────────────────────────────
if not os.path.exists(TEMPLATE_FILE):
    print(f"  ERROR: template.html not found at {TEMPLATE_FILE}")
    sys.exit(1)

with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
    template = f.read()

data_json = json.dumps(output, separators=(',', ':'))

# ── Owner view: index.html (all salons with tabs) ─────────────────────────────
owner_html = template.replace('__DASHBOARD_DATA__', data_json).replace('__SALON_ID__', 'null')
owner_file = os.path.join(SCRIPT_DIR, 'index.html')
with open(owner_file, 'w', encoding='utf-8') as f:
    f.write(owner_html)
print(f"\n✅ index.html (owner view — all salons)")

# ── Per-location views: one file per salon ────────────────────────────────────
print(f"\nGenerating per-location reports...")
for snum, sname in SALONS.items():
    salon_id = str(snum)
    loc_html = template.replace('__DASHBOARD_DATA__', data_json).replace('__SALON_ID__', f"'{salon_id}'")
    loc_file = os.path.join(SCRIPT_DIR, f'{snum}.html')
    with open(loc_file, 'w', encoding='utf-8') as f:
        f.write(loc_html)
    print(f"   ✓ {snum}.html  ({SALON_SHORT[snum]})")

print(f"\nWeek: {week_label}")
print(f"Active stylists:")
for snum, active in current_stylists.items():
    print(f"  {snum} {SALON_SHORT[snum]}: {len(active)} stylists")
print(f"\nUpload all .html files to your GitHub repo to update the dashboard.")
