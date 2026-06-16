"""
generar_data.py — Proyección Post-Venta
Uso: python generar_data.py

Coloca este script en la misma carpeta que los archivos Excel y data.json.
Busca automáticamente:
  - Producción*.xlsx
  - Seguimiento*.xlsx
Genera data.json en la misma carpeta.

Método: ZIP+XML+regex para el archivo grande (~63MB) → ~10 segundos.
"""

import re, json, zipfile, time, warnings, sys, os, calendar
from pathlib import Path
from datetime import date

warnings.filterwarnings('ignore')

try:
    import pandas as pd
except ImportError:
    sys.exit("❌ Falta pandas. Instala con: pip install pandas openpyxl")

# ── Configuración ─────────────────────────────────────────────
VALID_TD = {'FACTURA S/T', 'FACTURA GARANTIA S/T', 'CARGO INTERNO S/T'}
VALID_TV = {'VTA MANTENCIONES', 'VTA MEC', 'VTA MOVIL',
            'VTA INTERNA-CURI', 'VTA GARANTIA', 'VTA MANT PREPAGADAS'}

HERE    = Path(__file__).parent
HOY     = date.today()
YEAR    = HOY.year
CURR_M  = HOY.month
PREV_M  = CURR_M - 1 if CURR_M > 1 else 12
PREV_Y  = YEAR if CURR_M > 1 else YEAR - 1

# ── Helpers ───────────────────────────────────────────────────
def norm_suc(name):
    if not name: return 'DESCONOCIDA'
    s = re.sub(r'^\d{1,2}\s+', '', str(name).strip().upper())
    s = re.sub(r'\s*\(\d+\)\s*$', '', s)
    if re.match(r'^0+\d{4,}$', s.strip()):
        return 'DESCONOCIDA'
    return s.strip() or 'DESCONOCIDA'

def wdays(year, month, through=None):
    _, dim = calendar.monthrange(year, month)
    top = min(through if through else dim, dim)
    return sum(1 for d in range(1, top + 1) if date(year, month, d).weekday() < 5)

def find_file(pattern):
    matches = sorted(HERE.glob(pattern))
    if not matches:
        sys.exit(f"❌ No se encontró archivo {pattern} en {HERE}")
    if len(matches) > 1:
        # Prefer most recently modified
        matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        print(f"  ⚠ Múltiples {pattern} encontrados, usando: {matches[0].name}")
    return matches[0]

# ── 1. Cargar Producción (pandas, rápido) ─────────────────────
def load_produccion():
    t = time.time()
    fp = find_file('Producción*.xlsx')
    print(f"  Leyendo {fp.name}...")
    raw = pd.read_excel(fp, sheet_name=0, header=None, nrows=15)
    hrow = next(
        i for i, r in raw.iterrows()
        if any(str(v).strip() == 'TIPODOCTO' for v in r)
    )
    df = pd.read_excel(fp, sheet_name=0, header=hrow)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df['TIPODOCTO'].isin(VALID_TD) & df['TIPO-VENTA'].isin(VALID_TV)].copy()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['_MES']  = df['FECHA'].dt.month
    df['_DIA']  = df['FECHA'].dt.day
    df['NETO']  = pd.to_numeric(df['NETO'], errors='coerce').fillna(0)
    df['_SUC']  = df['SUCURSAL'].apply(norm_suc)
    print(f"  → {len(df)} filas válidas ({time.time()-t:.1f}s)")
    return df

# ── 2. Cargar Seguimiento (ZIP+XML+regex, ~10s) ───────────────
def load_seguimiento():
    t = time.time()
    fp = find_file('Seguimiento*.xlsx')
    print(f"  Leyendo {fp.name} (método XML)...")

    with zipfile.ZipFile(fp) as z:
        # Shared strings
        ss_raw = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
        ss = re.findall(r'<t(?:\s[^>]*)?>([^<]*)</t>', ss_raw)
        print(f"  → {len(ss):,} shared strings ({time.time()-t:.1f}s)")

        # Sheet (large — read all into memory)
        t2 = time.time()
        raw = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='replace')
        print(f"  → {len(raw)/1e6:.0f}MB XML leído ({time.time()-t2:.1f}s)")

    def get_cell(row_xml, col, rn):
        m = re.search(
            rf'<c r="{col}{rn}"[^>]*>(?:<f>[^<]*</f>)?<v>([^<]*)</v>',
            row_xml
        )
        return m.group(1) if m else None

    def ss_val(idx):
        try:
            return ss[int(idx)] if idx is not None else None
        except Exception:
            return None

    # Detect header row (row r="9" by default, verify it)
    # Identify column letters for key fields from header
    def find_header():
        # Try rows 1-20 to find 'FOLIO OT'
        folio_idx_str = None
        for candidate_ss, candidate_val in enumerate(ss):
            if candidate_val == 'FOLIO OT':
                folio_idx_str = str(candidate_ss)
                break
        if folio_idx_str is None:
            sys.exit("❌ No se encontró 'FOLIO OT' en shared strings")

        # Find header row: a row containing a cell with value == folio_idx in shared strings
        # Search for <c r="?{rn}" t="s"><v>{folio_idx}</v></c> in first 25 rows
        for rn in range(1, 26):
            m = re.search(
                rf'<c r="([A-Z]+){rn}" t="s"><v>{re.escape(folio_idx_str)}</v></c>',
                raw
            )
            if m:
                col_folio = m.group(1)
                # Extract this row to find all headers
                row_m = re.search(rf'<row r="{rn}"[^>]*>(.*?)</row>', raw, re.DOTALL)
                if row_m:
                    return rn, row_m.group(1), col_folio
        sys.exit("❌ No se encontró fila de encabezado en Seguimiento")

    hrow_n, hrow_xml, _ = find_header()
    print(f"  → Encabezado en fila {hrow_n}")

    # Build column map: field_name → column_letter
    want = {'FOLIO OT', 'SUCURSAL', 'AÑO', 'MES', 'DÍA', 'TIPO VENTA', 'FECHA OT'}
    col_map = {}
    for cell_m in re.finditer(r'<c r="([A-Z]+)\d+" t="s"><v>(\d+)</v></c>', hrow_xml):
        col_letter, sidx = cell_m.group(1), int(cell_m.group(2))
        try:
            val = ss[sidx]
        except Exception:
            continue
        if val in want:
            col_map[val] = col_letter

    missing = want - set(col_map)
    if missing:
        print(f"  ⚠ Columnas no encontradas: {missing}")
    print(f"  → Columnas: {col_map}")

    # Find all 2026 rows via regex on AÑO column
    t3 = time.time()
    col_anio = col_map.get('AÑO')
    if not col_anio:
        sys.exit("❌ No se encontró columna AÑO")

    row_nums_2026 = re.findall(
        rf'<c r="{col_anio}(\d+)"[^>]*><v>2026</v></c>', raw
    )
    print(f"  → {len(row_nums_2026):,} filas 2026 ({time.time()-t3:.1f}s)")

    # Extract row XMLs for those row numbers
    t4 = time.time()
    all_row_xmls = {
        m.group(1): m.group(2)
        for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', raw, re.DOTALL)
        if m.group(1) in set(row_nums_2026)
    }
    print(f"  → {len(all_row_xmls):,} row XMLs extraídos ({time.time()-t4:.1f}s)")

    col_folio = col_map.get('FOLIO OT')
    col_suc   = col_map.get('SUCURSAL')
    col_mes   = col_map.get('MES')
    col_dia   = col_map.get('DÍA')
    col_tv    = col_map.get('TIPO VENTA')

    seen = set()
    may_ots, jun_ots = [], []

    for rn_s in row_nums_2026:
        row_xml = all_row_xmls.get(rn_s)
        if not row_xml:
            continue

        folio_raw = get_cell(row_xml, col_folio, rn_s) if col_folio else None
        try:
            folio = int(folio_raw) if folio_raw else None
        except Exception:
            folio = None
        if not folio or folio in seen:
            continue
        seen.add(folio)

        mes_raw = get_cell(row_xml, col_mes, rn_s) if col_mes else None
        try:
            mes = int(mes_raw) if mes_raw else None
        except Exception:
            mes = None
        if mes not in (PREV_M, CURR_M):
            continue

        dia_raw = get_cell(row_xml, col_dia, rn_s) if col_dia else None
        try:
            dia = int(dia_raw) if dia_raw else None
        except Exception:
            dia = None

        suc_raw = get_cell(row_xml, col_suc, rn_s) if col_suc else None
        suc = norm_suc(ss_val(suc_raw))

        tv_raw = get_cell(row_xml, col_tv, rn_s) if col_tv else None
        tv  = ss_val(tv_raw) or ''

        if tv not in VALID_TV:
            continue

        rec = {'folio': folio, 'mes': mes, 'dia': dia, 'suc': suc, 'tv': tv}
        if mes == PREV_M:
            may_ots.append(rec)
        else:
            jun_ots.append(rec)

    print(f"  → OTs Mayo: {len(may_ots)} | OTs Junio: {len(jun_ots)} ({time.time()-t:.1f}s total)")
    return may_ots, jun_ots

# ── 3. Calcular métricas y proyecciones ──────────────────────
def calcular(prod_df, may_ots, jun_ots):
    today_day = HOY.day

    # Mayo — Producción
    may_prod = prod_df[prod_df['_MES'] == PREV_M]
    may_billing = float(may_prod['NETO'].sum())
    may_billing_by_day = {
        str(int(k)): float(v)
        for k, v in may_prod.groupby('_DIA')['NETO'].sum().items()
    }
    may_by_suc_billing = {
        k: float(v) for k, v in may_prod.groupby('_SUC')['NETO'].sum().items()
    }
    may_mix = {
        k: float(v) for k, v in may_prod.groupby('TIPO-VENTA')['NETO'].sum().items()
    }

    # Mayo — Seguimiento
    may_n = len(may_ots)
    may_ots_by_day, may_by_suc_ots = {}, {}
    for o in may_ots:
        d = str(o['dia']) if o['dia'] else None
        if d: may_ots_by_day[d] = may_ots_by_day.get(d, 0) + 1
        may_by_suc_ots[o['suc']] = may_by_suc_ots.get(o['suc'], 0) + 1

    may_wdays   = wdays(PREV_Y, PREV_M)
    ticket_avg  = may_billing / may_n if may_n else 0
    daily_bill  = may_billing / may_wdays if may_wdays else 0

    # Junio — Seguimiento
    jun_n = len(jun_ots)
    jun_ots_by_day, jun_by_suc_ots = {}, {}
    for o in jun_ots:
        d = str(o['dia']) if o['dia'] else None
        if d: jun_ots_by_day[d] = jun_ots_by_day.get(d, 0) + 1
        jun_by_suc_ots[o['suc']] = jun_by_suc_ots.get(o['suc'], 0) + 1

    jun_so_far  = wdays(YEAR, CURR_M, today_day)
    jun_total   = wdays(YEAR, CURR_M)
    jun_remain  = jun_total - jun_so_far
    jun_rate    = jun_n / jun_so_far if jun_so_far else (may_n / may_wdays if may_wdays else 0)

    # Escenarios
    scenarios = {}
    for name, (fo, fb) in {
        'conservador': (0.88, 0.93),
        'probable':    (1.00, 1.00),
        'optimista':   (1.10, 1.07),
    }.items():
        proj_ots = round(jun_n + jun_remain * jun_rate * fo)
        scenarios[name] = {
            'ots':     proj_ots,
            'billing': float(proj_ots * ticket_avg * fb),
        }

    # Semanas de junio
    _, dim = calendar.monthrange(YEAR, CURR_M)

    def wk(d):
        return (d - 1 + date(YEAR, CURR_M, 1).weekday()) // 7 + 1

    week_days_map = {}
    for d in range(1, dim + 1):
        w = wk(d)
        week_days_map.setdefault(w, []).append(d)

    weekly = []
    for w, days in sorted(week_days_map.items()):
        wd = [d for d in days if date(YEAR, CURR_M, d).weekday() < 5]
        actual    = sum(jun_ots_by_day.get(str(d), 0) for d in wd)
        all_past  = max(days) < today_day
        partial   = min(days) <= today_day <= max(days)
        if all_past:
            proj = actual
        elif partial:
            proj = actual + len([d for d in wd if d > today_day]) * jun_rate
        else:
            proj = len(wd) * jun_rate
        weekly.append({
            'week':     w,
            'start':    min(days),
            'end':      max(days),
            'weekdays': len(wd),
            'actual':   int(actual),
            'projected': int(round(proj)),
            'billing':  float(round(proj) * ticket_avg),
            'all_past': bool(all_past),
            'partial':  bool(partial),
        })

    # Por sucursal
    all_s = sorted(
        set(list(may_by_suc_billing) + list(may_by_suc_ots) + list(jun_by_suc_ots))
    )
    suc_rows = []
    for suc in all_s:
        m_b = float(may_by_suc_billing.get(suc, 0))
        m_o = int(may_by_suc_ots.get(suc, 0))
        j_o = int(jun_by_suc_ots.get(suc, 0))
        if m_o == 0 and j_o == 0:
            continue
        tkt  = m_b / m_o if m_o else ticket_avg
        rate = j_o / jun_so_far if jun_so_far else 0
        j_p  = round(j_o + rate * jun_remain)
        suc_rows.append({
            'sucursal':       suc,
            'may_ots':        m_o,
            'may_billing':    m_b,
            'ticket_avg':     float(tkt),
            'jun_ots_actual': j_o,
            'jun_ots_proj':   j_p,
            'jun_billing_proj': float(j_p * tkt),
        })

    return {
        'generado':    HOY.isoformat(),
        'today_day':   today_day,
        'curr_month':  CURR_M,
        'curr_year':   YEAR,
        'prev_month':  PREV_M,
        'prev_year':   PREV_Y,
        'may': {
            'billing_total':    may_billing,
            'ot_count':         may_n,
            'ticket_avg':       float(ticket_avg),
            'working_days':     may_wdays,
            'daily_billing_avg': float(daily_bill),
            'billing_by_day':   may_billing_by_day,
            'ots_by_day':       may_ots_by_day,
            'mix':              may_mix,
        },
        'jun': {
            'ot_count':       jun_n,
            'ots_by_day':     jun_ots_by_day,
            'working_so_far': jun_so_far,
            'working_total':  jun_total,
            'working_remain': jun_remain,
            'daily_ot_rate':  float(jun_rate),
        },
        'scenarios': scenarios,
        'weekly':    weekly,
        'sucursales': suc_rows,
    }

# ── Main ──────────────────────────────────────────────────────
if __name__ == '__main__':
    t0 = time.time()
    print(f"\n📊 Generando proyección Post-Venta — {HOY}")
    print("─" * 50)

    print("\n[1/3] Producción")
    prod_df = load_produccion()

    print("\n[2/3] Seguimiento")
    may_ots, jun_ots = load_seguimiento()

    print("\n[3/3] Calculando métricas")
    data = calcular(prod_df, may_ots, jun_ots)

    out = HERE / 'data.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    sc = data['scenarios']
    print(f"""
✅ data.json generado en {out} ({time.time()-t0:.1f}s)

   Mayo billing    : ${data['may']['billing_total']:>15,.0f}
   Mayo OTs        : {data['may']['ot_count']:>6}
   Ticket promedio : ${data['may']['ticket_avg']:>15,.0f}
   ─────────────────────────────────────
   Junio OTs (real): {data['jun']['ot_count']:>6}  ({data['jun']['working_so_far']} días hábiles)
   Tasa diaria     : {data['jun']['daily_ot_rate']:>6.1f} OTs/día
   ─────────────────────────────────────
   Proyección cierre (Junio):
     Conservador   : {sc['conservador']['ots']:>6} OTs  ${sc['conservador']['billing']:>15,.0f}
     Probable      : {sc['probable']['ots']:>6} OTs  ${sc['probable']['billing']:>15,.0f}
     Optimista     : {sc['optimista']['ots']:>6} OTs  ${sc['optimista']['billing']:>15,.0f}
""")
