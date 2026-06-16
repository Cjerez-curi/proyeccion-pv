"""
generar_data.py â ProyecciÃ³n Post-Venta
Uso: python generar_data.py

Coloca este script en la misma carpeta que los archivos Excel y data.json.
Busca automÃ¡ticamente:
  - ProducciÃ³n*.xlsx
  - Seguimiento*.xlsx
Genera data.json en la misma carpeta.

MÃ©todo: ZIP+XML+regex para el archivo grande (~63MB) â ~10 segundos.
"""

import re, json, zipfile, time, warnings, sys, os, calendar
from pathlib import Path
from datetime import date

warnings.filterwarnings('ignore')

try:
    import pandas as pd
except ImportError:
    sys.exit("Falta pandas. Instala con: pip install pandas openpyxl")

# Configuracion
VALID_TD  = {'FACTURA S/T', 'FACTURA GARANTIA S/T', 'CARGO INTERNO S/T'}
DEDUCT_TD = {'CIERRE GERENCIA S/T', 'NC CLIENTE S/T'}
ALL_TD    = VALID_TD | DEDUCT_TD
VALID_TV  = {'VTA MANTENCIONES', 'VTA MEC', 'VTA MOVIL',
             'VTA INTERNA-CURI', 'VTA GARANTIA', 'VTA MANT PREPAGADAS'}

HERE    = Path(__file__).parent
HOY     = date.today()
YEAR    = HOY.year
CURR_M  = HOY.month
PREV_M  = CURR_M - 1 if CURR_M > 1 else 12
PREV_Y  = YEAR if CURR_M > 1 else YEAR - 1

# Mapa de codigos SUC â nombre normalizado
SUC_MAP = {
    'SUC010': 'LA FLORIDA',      'SUC020': 'CHILLAN',        'SUC030': 'CHILLAN VIEJO',
    'SUC040': 'COQUIMBO',        'SUC050': 'CURICO',         'SUC070': 'LINDEROS',
    'SUC080': 'LIRA',            'SUC090': 'OVALLE 3',       'SUC100': 'MALL PLAZA NORTE',
    'SUC110': 'LO BLANCO',       'SUC120': 'OVALLE MALL',    'SUC130': 'PLACILLA',
    'SUC140': 'RANCAGUA USADOS', 'SUC150': 'SAN FERNANDO',   'SUC160': 'TALCA',
    'SUC180': 'TALCA 3',         'SUC210': 'OVALLE 2',       'SUC230': 'MAIPU',
    'SUC240': 'DIEZ DE JULIO 2', 'SUC250': 'BRASIL',         'SUC260': 'CONCEPCION RPTO',
    'SUC270': 'GRAN AVENIDA',    'SUC280': 'CD REPUESTOS',   'SUC290': 'LA SERENA',
    'SUC300': 'LO BLANCO 2',     'SUC310': 'RANCAGUA',       'SUC320': 'TALCA 2',
}

def norm_suc(name):
    if not name: return 'DESCONOCIDA'
    s = str(name).strip().upper()
    # Codigo tipo "SUC070" directo
    if s in SUC_MAP:
        return SUC_MAP[s]
    # Codigo numerico puro: "70", "070", "0070" â SUC070
    m = re.match(r'^0*(\d+)$', s)
    if m:
        code = 'SUC' + m.group(1).zfill(3)
        return SUC_MAP.get(code, 'DESCONOCIDA')
    # Nombre de texto: quitar prefijo numerico tipo "5 LINDEROS"
    s = re.sub(r'^\d{1,2}\s+', '', s)
    s = re.sub(r'\s*\((\d+)\)\s*$', r' \1', s)
    return s.strip() or 'DESCONOCIDA'

def wdays(year, month, through=None):
    _, dim = calendar.monthrange(year, month)
    top = min(through if through else dim, dim)
    return sum(1 for d in range(1, top + 1) if date(year, month, d).weekday() < 5)

def find_file(pattern):
    matches = sorted(HERE.glob(pattern))
    if not matches:
        sys.exit(f"No se encontro archivo {pattern} en {HERE}")
    if len(matches) > 1:
        matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        print(f"  Multiples {pattern} encontrados, usando: {matches[0].name}")
    return matches[0]

def load_produccion():
    t = time.time()
    fp = find_file('Producci*.xlsx')
    print(f"  Leyendo {fp.name}...")
    raw = pd.read_excel(fp, sheet_name=0, header=None, nrows=15)
    hrow = next(
        i for i, r in raw.iterrows()
        if any(str(v).strip() == 'TIPODOCTO' for v in r)
    )
    df = pd.read_excel(fp, sheet_name=0, header=hrow)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df['TIPODOCTO'].isin(ALL_TD) & df['TIPO-VENTA'].isin(VALID_TV)].copy()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['_MES']  = df['FECHA'].dt.month
    df['_DIA']  = df['FECHA'].dt.day
    df['NETO']  = pd.to_numeric(df['NETO'], errors='coerce').fillna(0)
    # Excluir vehiculos pesados (camiones, buses, etc.)
    PESADOS = {'CAMION', 'BUS', 'AMBULANCIA', 'CHASIS CABINADO', 'RUBRO 101', 'RUBRO 25'}
    n_pesados = df[df['TIPO VEHICULO'].isin(PESADOS)]['NUMERO'].nunique()
    df = df[~df['TIPO VEHICULO'].isin(PESADOS)].copy()
    print(f"  Vehiculos pesados excluidos: {n_pesados} docs")
    # Excluir documentos con NETO total <= $2 (paso vehicular sin facturacion real)
    doc_totals = df.groupby('NUMERO')['NETO'].sum()
    valid_docs = doc_totals[doc_totals.abs() > 2].index
    before = len(df)
    n_docs_removed = len(doc_totals) - len(valid_docs)
    df = df[df['NUMERO'].isin(valid_docs)].copy()
    print(f"  Paso vehicular <=\$2 excluidos: {n_docs_removed} docs ({before - len(df)} filas)")
    df['_SUC']  = df['SUCURSAL'].apply(norm_suc)
    print(f"  {len(df)} filas validas ({time.time()-t:.1f}s)")
    return df

def load_seguimiento():
    t = time.time()
    fp = find_file('Seguimiento*.xlsx')
    print(f"  Leyendo {fp.name} (metodo XML)...")

    with zipfile.ZipFile(fp) as z:
        ss_raw = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
        ss = re.findall(r'<t(?:\s[^>]*)?>([^<]*)</t>', ss_raw)
        print(f"  {len(ss):,} shared strings ({time.time()-t:.1f}s)")

        t2 = time.time()
        raw = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='replace')
        print(f"  {len(raw)/1e6:.0f}MB XML leido ({time.time()-t2:.1f}s)")

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

    def find_header():
        folio_idx_str = None
        for i, v in enumerate(ss):
            if v == 'FOLIO OT':
                folio_idx_str = str(i)
                break
        if folio_idx_str is None:
            sys.exit("No se encontro 'FOLIO OT' en shared strings")
        for rn in range(1, 35):
            m = re.search(
                rf'<c r="([A-Z]+){rn}"[^>]*t="s"[^>]*><v>{re.escape(folio_idx_str)}</v></c>',
                raw
            )
            if m:
                col_folio = m.group(1)
                row_m = re.search(rf'<row r="{rn}"[^>]*>(.*?)</row>', raw, re.DOTALL)
                if row_m:
                    return rn, row_m.group(1), col_folio
        sys.exit("No se encontro fila de encabezado en Seguimiento")

    hrow_n, hrow_xml, _ = find_header()
    print(f"  Encabezado en fila {hrow_n}")

    want_orig = {'FOLIO OT': 'FOLIO OT', 'SUCURSAL': 'SUCURSAL',
                 'A\u00d1O': 'ANO', 'MES': 'MES', 'D\u00cdA': 'DIA',
                 'TIPO VENTA': 'TIPO VENTA', 'FECHA OT': 'FECHA OT'}
    col_map = {}
    for cell_m in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*t="s"[^>]*><v>(\d+)</v></c>', hrow_xml):
        col_letter, sidx = cell_m.group(1), int(cell_m.group(2))
        try:
            val = ss[sidx]
        except Exception:
            continue
        if val in want_orig:
            col_map[want_orig[val]] = col_letter

    print(f"  Columnas: {col_map}")

    t3 = time.time()
    col_anio = col_map.get('ANO')
    if not col_anio:
        sys.exit("No se encontro columna ANO")

    row_nums_2026 = re.findall(
        rf'<c r="{col_anio}(\d+)"[^>]*><v>2026</v></c>', raw
    )
    print(f"  {len(row_nums_2026):,} filas 2026 ({time.time()-t3:.1f}s)")

    t4 = time.time()
    row_set_2026 = set(row_nums_2026)
    all_row_xmls = {
        m.group(1): m.group(2)
        for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', raw, re.DOTALL)
        if m.group(1) in row_set_2026
    }
    print(f"  {len(all_row_xmls):,} row XMLs extraidos ({time.time()-t4:.1f}s)")

    col_folio = col_map.get('FOLIO OT')
    col_suc   = col_map.get('SUCURSAL')
    col_mes   = col_map.get('MES')
    col_dia   = col_map.get('DIA')
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

    print(f"  OTs Mayo: {len(may_ots)} | OTs Junio: {len(jun_ots)} ({time.time()-t:.1f}s total)")
    return may_ots, jun_ots

def calcular(prod_df, may_ots, jun_ots):
    today_day = HOY.day

    may_prod = prod_df[prod_df['_MES'] == PREV_M]
    may_billing        = float(may_prod['NETO'].sum())
    may_billing_gross  = float(may_prod[may_prod['TIPODOCTO'].isin(VALID_TD)]['NETO'].sum())
    may_billing_deduct = float(may_prod[may_prod['TIPODOCTO'].isin(DEDUCT_TD)]['NETO'].sum())

    may_billing_by_day = {
        str(int(k)): float(v)
        for k, v in may_prod.groupby('_DIA')['NETO'].sum().items()
    }
    may_by_suc_billing = {
        k: float(v) for k, v in may_prod.groupby('_SUC')['NETO'].sum().items()
    }
    may_mix = {
        k: float(v) for k, v in may_prod.groupby('TIPODOCTO')['NETO'].sum().items()
    }

    def _wk(d, year=PREV_Y, month=PREV_M):
        return (d - 1 + date(year, month, 1).weekday()) // 7 + 1

    may_prod_wk = may_prod.copy()
    may_prod_wk['_WK'] = may_prod_wk['_DIA'].apply(
        lambda d: _wk(int(d)) if pd.notna(d) else None
    )
    may_billing_by_week = {
        int(k): float(v)
        for k, v in may_prod_wk.groupby('_WK')['NETO'].sum().items()
    }
    total_may_wk = sum(may_billing_by_week.values())
    may_week_pct = {
        k: v / total_may_wk
        for k, v in may_billing_by_week.items()
        if total_may_wk
    }

    last_wk_key    = max(may_billing_by_week.keys())
    other_wks      = [v for k, v in may_billing_by_week.items() if k != last_wk_key]
    avg_other      = sum(other_wks) / len(other_wks) if other_wks else 1
    last_wk_factor = may_billing_by_week.get(last_wk_key, 0) / avg_other if avg_other else 1

    # OT count desde Seguimiento (tasa de ingreso de vehiculos)
    may_n_seg = len(may_ots)
    may_ots_by_day, may_by_suc_ots = {}, {}
    for o in may_ots:
        d = str(o['dia']) if o['dia'] else None
        if d:
            may_ots_by_day[d] = may_ots_by_day.get(d, 0) + 1
        may_by_suc_ots[o['suc']] = may_by_suc_ots.get(o['suc'], 0) + 1

    # OT count desde N°RECEPCION de Produccion (denominador real del ticket)
    may_n_prod = int(may_prod['N°RECEPCION'].nunique())
    conversion_factor = may_n_prod / may_n_seg if may_n_seg else 1.0

    may_wdays  = wdays(PREV_Y, PREV_M)
    ticket_avg = may_billing / may_n_prod if may_n_prod else 0
    daily_bill = may_billing / may_wdays if may_wdays else 0

    jun_n = len(jun_ots)
    jun_ots_by_day, jun_by_suc_ots = {}, {}
    for o in jun_ots:
        d = str(o['dia']) if o['dia'] else None
        if d:
            jun_ots_by_day[d] = jun_ots_by_day.get(d, 0) + 1
        jun_by_suc_ots[o['suc']] = jun_by_suc_ots.get(o['suc'], 0) + 1

    jun_so_far = wdays(YEAR, CURR_M, today_day)
    jun_total  = wdays(YEAR, CURR_M)
    jun_remain = jun_total - jun_so_far
    jun_rate_seg = jun_n / jun_so_far if jun_so_far else (may_n_seg / may_wdays if may_wdays else 0)
    jun_rate = jun_rate_seg * conversion_factor

    scenarios = {}
    for name, fo, fb in [('conservador', 0.88, 0.93), ('probable', 1.00, 1.00), ('optimista', 1.10, 1.07)]:
        proj_ots = round(jun_n + jun_remain * jun_rate * fo)
        scenarios[name] = {
            'ots': proj_ots,
            'billing': float(proj_ots * ticket_avg * fb),
        }

    _, dim = calendar.monthrange(YEAR, CURR_M)

    def wk(d):
        return (d - 1 + date(YEAR, CURR_M, 1).weekday()) // 7 + 1

    week_days_map = {}
    for d in range(1, dim + 1):
        w = wk(d)
        week_days_map.setdefault(w, []).append(d)

    total_weeks_jun  = max(week_days_map.keys())
    prob_billing_jun = scenarios['probable']['billing']

    def may_pct_for_jun_week(w_jun, total_w_jun, pct_map):
        total_w_may = max(pct_map.keys()) if pct_map else total_w_jun
        w_may = total_w_may - (total_w_jun - w_jun)
        return pct_map.get(w_may, pct_map.get(w_jun, 1.0 / total_w_jun))

    weekly = []
    for w, days in sorted(week_days_map.items()):
        wd       = [d for d in days if date(YEAR, CURR_M, d).weekday() < 5]
        actual   = sum(jun_ots_by_day.get(str(d), 0) for d in wd)
        all_past = max(days) < today_day
        partial  = min(days) <= today_day <= max(days)

        if all_past:
            proj_ots = actual
        elif partial:
            proj_ots = actual + len([d for d in wd if d > today_day]) * jun_rate
        else:
            proj_ots = round(len(wd) * jun_rate)

        if all_past or partial:
            billing_proj = float(round(proj_ots) * ticket_avg)
        else:
            pct = may_pct_for_jun_week(w, total_weeks_jun, may_week_pct)
            billing_proj = float(prob_billing_jun * pct)

        weekly.append({
            'week':         w,
            'start':        min(days),
            'end':          max(days),
            'weekdays':     len(wd),
            'actual':       int(actual),
            'projected':    int(round(proj_ots)),
            'billing':      billing_proj,
            'all_past':     bool(all_past),
            'partial':      bool(partial),
            'is_last_week': bool(w == total_weeks_jun),
        })

    all_s = sorted(
        set(list(may_by_suc_billing) + list(may_by_suc_ots) + list(jun_by_suc_ots))
    )
    suc_rows = []
    for suc in all_s:
        m_b = float(may_by_suc_billing.get(suc, 0))
        m_o = int(may_by_suc_ots.get(suc, 0))
        j_o = int(jun_by_suc_ots.get(suc, 0))
        if m_o == 0 and j_o == 0 and m_b == 0:
            continue
        tkt  = m_b / m_o if m_o else ticket_avg
        rate = j_o / jun_so_far if jun_so_far else 0
        j_p  = round(j_o + rate * jun_remain)
        suc_rows.append({
            'sucursal':         suc,
            'may_ots':          m_o,
            'may_billing':      m_b,
            'ticket_avg':       float(tkt),
            'jun_ots_actual':   j_o,
            'jun_ots_proj':     j_p,
            'jun_billing_proj': float(j_p * tkt),
        })

    suc_names_prod = sorted(may_by_suc_billing.keys())

    return {
        'generado':   HOY.isoformat(),
        'today_day':  today_day,
        'curr_month': CURR_M,
        'curr_year':  YEAR,
        'prev_month': PREV_M,
        'prev_year':  PREV_Y,
        '_debug_suc_prod': suc_names_prod,
        'may': {
            'billing_total':     may_billing,
            'billing_gross':     may_billing_gross,
            'billing_deduct':    may_billing_deduct,
            'ot_count':          may_n_prod,
            'ot_count_seg':      may_n_seg,
            'conversion_factor': float(conversion_factor),
            'ticket_avg':        float(ticket_avg),
            'working_days':      may_wdays,
            'daily_billing_avg': float(daily_bill),
            'billing_by_day':    may_billing_by_day,
            'billing_by_week':   {str(k): float(v) for k, v in may_billing_by_week.items()},
            'week_pct':          {str(k): float(v) for k, v in may_week_pct.items()},
            'last_week_factor':  float(last_wk_factor),
            'ots_by_day':        may_ots_by_day,
            'mix':               may_mix,
        },
        'jun': {
            'ot_count':       jun_n,
            'ots_by_day':     jun_ots_by_day,
            'working_so_far': jun_so_far,
            'working_total':  jun_total,
            'working_remain': jun_remain,
            'daily_ot_rate_seg': float(jun_rate_seg),
            'daily_ot_rate':     float(jun_rate),
        },
        'scenarios':  scenarios,
        'weekly':     weekly,
        'sucursales': suc_rows,
    }

if __name__ == '__main__':
    t0 = time.time()
    print(f"\nGenerando proyeccion Post-Venta - {HOY}")
    print("-" * 50)

    print("\n[1/3] Produccion")
    prod_df = load_produccion()

    print("\n[2/3] Seguimiento")
    may_ots, jun_ots = load_seguimiento()

    print("\n[3/3] Calculando metricas")
    data = calcular(prod_df, may_ots, jun_ots)

    out = HERE / 'data.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    sc = data['scenarios']
    print(f"\ndata.json generado en {out} ({time.time()-t0:.1f}s)")
    print(f"  Mayo billing NET  : {data['may']['billing_total']:,.0f}")
    print(f"  Mayo billing bruto: {data['may']['billing_gross']:,.0f}")
    print(f"  Deducciones NC/CG : {data['may']['billing_deduct']:,.0f}")
    print(f"  Mayo OTs          : {data['may']['ot_count']}")
    print(f"  Ticket promedio   : {data['may']['ticket_avg']:,.0f}")
    print(f"  Factor fin de mes : {data['may']['last_week_factor']:.2f}x")
    print(f"  Junio OTs (real)  : {data['jun']['ot_count']} ({data['jun']['working_so_far']} dias habiles)")
    print(f"  Tasa diaria       : {data['jun']['daily_ot_rate']:.1f} OTs/dia")
    print(f"  Conservador       : {sc['conservador']['ots']} OTs  {sc['conservador']['billing']:,.0f}")
    print(f"  Probable          : {sc['probable']['ots']} OTs  {sc['probable']['billing']:,.0f}")
    print(f"  Optimista         : {sc['optimista']['ots']} OTs  {sc['optimista']['billing']:,.0f}")
    print(f"  Sucursales Prod   : {data['_debug_suc_prod']}")
