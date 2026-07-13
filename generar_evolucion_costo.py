# -*- coding: utf-8 -*-
"""
Genera evolucion_costo.html con la evolución mensual de costo (USD/H) por persona.

Fuentes (prioridad de mayor a menor):
  1. BDD Bamboo Histórico: col 0=Month (YYYY-MM-DD o serial), col 2=email, col 23=Costo USD/H
  2. Compensaciones 2025 "Bamboo: Compensaciones -Sueldos MS" col Z:
     12 filas por persona (Jan-Dec 2025). El mes usa encoding especial:
     serial 45658 = Jan, 45659 = Feb, ..., 45669 = Dec 2025.
  3. Sheet histórico "Compensation Info - To Juli": col A=email agrupado,
     col D=Month ("Jan 2026"), col E=CE usd ("$28.00")
"""

import sys, json, re
sys.stdout.reconfigure(encoding='utf-8')

CREDS_FILE    = r"C:\Users\aprato\documents\Proyectos\Compensaciones\talentserviceproject-1ce2ed91696b.json"
SHEET_HIST    = '1G4ZkjoFYtATmkiVhxDRLanx5zD3dCtzNOHHWpNanKUw'
TAB_HIST_OLD  = 'Compensation Info - To Juli'
SHEET_BDD     = '11kTmD6YBpquXMlqI1QoY4xU8bmjKY0Rh30qX55rebsQ'
TAB_HIST_BDD  = 'Histórico'
SHEET_COMP20  = '17DPzFH0W3BOLHRaNte71uxpRLO5ZOY8glmecxGWCV6A'
SHEET_COMP21  = '14GyY0zAdZO1x9oBIwyoUKJv11EtG4gY9V3qNVUF1scM'
SHEET_COMP22  = '1lhZVM-OnEynU3jl7LJQSxTZ6IdbXPuYxjifElCNwk-A'
SHEET_COMP23  = '1HQ9ZX5Ac5JvbpBjJaMGidQsAu-0jHjCI2nbScAb8r7E'
SHEET_COMP24  = '1swqMg2bD4ZeOjVSW_5M-bz7stcUTlF0BbOKDGBIoJ0I'
SHEET_COMP25  = '1uPuYunVQb8Pi1AxDcg9BKftzau9e9DER8azoHC9CoWA'
TAB_COMP_MS   = 'Bamboo: Compensaciones -Sueldos MS'
TAB_LISTA     = 'LISTA PARA IMPORTAR Actual+Fcst'
OUTPUT_HTML   = r"C:\Users\aprato\documents\Proyectos\Compensaciones\evolucion_costo.html"

import datetime
_SERIAL_EPOCH = datetime.date(1899, 12, 30)

MONTH_ABBR = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
    'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
    'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12',
}

EMAIL_CORRECTIONS = {
    'mmaza@makignsense.com': 'mmaza@makingsense.com',
    'kromeroquijano@makingsense.com': 'kromero@makingsense.com',
}

_SENIORITY_EXCLUDE = {'Maestranza'}

# Solo normaliza dígitos sueltos sin cero (por si quedan en fuentes históricas)
_SINGLE_DIGIT_PAD = {str(i): f'{i:02d}' for i in range(1, 10)}

def _normalize_seniority(s):
    if not s:
        return s
    if s in _SENIORITY_EXCLUDE:
        return ''
    # Zero-pad dígito suelto (1→01, ..., 9→09)
    if s in _SINGLE_DIGIT_PAD:
        return _SINGLE_DIGIT_PAD[s]
    return s

AGREEMENT_CORRECTIONS = {
    'Plus fijo': 'Plus Fijo',
    'Plus fijo (sueldo en AR$ + fijo en USD)': 'Plus Fijo',
    'Mixto (Monto total acordado en USD)': 'Mixto',
}


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    try:
        import streamlit as st
        creds = service_account.Credentials.from_service_account_info(
            dict(st.secrets['gcp_service_account']), scopes=scopes)
    except Exception:
        creds = service_account.Credentials.from_service_account_file(
            CREDS_FILE, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)


def _parse_old_month(s):
    """'Jan 2026' → '2026-01'"""
    s = str(s).strip()
    m = re.match(r'^([A-Za-z]{3})\s+(\d{4})$', s)
    if not m:
        return None
    mon, yr = m.group(1).capitalize(), m.group(2)
    abbr = MONTH_ABBR.get(mon)
    return f'{yr}-{abbr}' if abbr else None


def _parse_ce(s):
    """'$28.00 ' → 28.0"""
    try:
        return round(float(re.sub(r'[^\d.]', '', str(s).strip())), 4)
    except Exception:
        return None


def _parse_bdd_month(s):
    """'2024-01-01' o serial 46143 → '2024-01'"""
    s = str(s).strip()
    if len(s) >= 7 and s[4] == '-':
        return s[:7]
    # Intentar como serial de Excel
    if s.isdigit():
        try:
            d = _SERIAL_EPOCH + datetime.timedelta(days=int(s))
            return d.strftime('%Y-%m')
        except Exception:
            pass
    return None


# ─── LECTURA SHEET VIEJO ──────────────────────────────────────────────────────

def read_old_sheet(svc):
    print('Leyendo sheet histórico...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_HIST,
        range=f"'{TAB_HIST_OLD}'!A6:G19200",
        valueRenderOption='FORMATTED_VALUE',
    ).execute()
    rows = r.get('values', [])

    data = {}   # email → {mes → {costo, code, agreement}}
    cur_email     = None
    cur_agreement = None

    for row in rows:
        while len(row) < 7:
            row.append('')
        email_raw = str(row[0]).strip().lower()
        agr_raw   = str(row[1]).strip()   # col B = Agreement H
        month_raw = str(row[3]).strip()
        ce_raw    = str(row[4]).strip()   # col E = CE usd
        bill_raw  = str(row[6]).strip()   # col G = Bill

        if email_raw:
            cur_email     = email_raw
            cur_agreement = agr_raw
        if not cur_email or '@' not in cur_email:
            continue

        mes = _parse_old_month(month_raw)
        if not mes:
            continue

        # Col E (CE usd) es siempre USD/H directamente para todos los acuerdos.
        costo = _parse_ce(ce_raw)

        if costo is None or costo == 0:
            continue

        if cur_email not in data:
            data[cur_email] = {}
        prev = data[cur_email].get(mes)
        if not prev or costo > prev['costo']:
            data[cur_email][mes] = {'costo': costo, 'code': '', 'agreement': cur_agreement}

    total_puntos = sum(len(v) for v in data.values())
    print(f'  {len(data)} personas, {total_puntos} puntos de datos (sheet viejo).')
    return data


# ─── LECTURA SHEETS "LISTA PARA IMPORTAR" (2020-2022) ────────────────────────

def _read_lista_sheet(svc, sheet_id, year, actual_only=False, skip_ars=False):
    """
    Lee un sheet 'LISTA PARA IMPORTAR Actual+Fcst' y genera datos mensuales.
    col 1=email, col 2=date_start, col 3=cost_rate (USD/H), col 6=periodo/moneda.
    actual_only: si True, filtra period in ('Actual','') (usado para 2020).
    skip_ars: si True, salta filas con moneda que contenga 'AR$'.
    Retorna email → {mes → {costo, code, agreement}}.
    """
    import calendar as _cal
    print(f'Leyendo lista {year}...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TAB_LISTA}'!A:J",
        valueRenderOption='FORMATTED_VALUE',
    ).execute()
    rows = r.get('values', [])[1:]

    by_email = {}
    for row in rows:
        while len(row) < 4:
            row.append('')
        email   = str(row[1]).strip().lower()
        if '@' not in email:
            continue
        col6 = str(row[6]).strip() if len(row) > 6 else ''
        if actual_only and col6 not in ('Actual', ''):
            continue
        if skip_ars and 'AR$' in col6:
            continue
        date_s = str(row[2]).strip()
        cost_s = str(row[3]).strip()
        d = None
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
            try:
                d = datetime.datetime.strptime(date_s, fmt).date()
                break
            except Exception:
                pass
        if d is None:
            continue
        try:
            cost = round(float(re.sub(r'[^\d.]', '', cost_s)), 4)
        except Exception:
            continue
        if cost <= 0:
            continue
        if email not in by_email:
            by_email[email] = []
        by_email[email].append((d, cost))

    for email in by_email:
        by_email[email].sort()

    data = {}
    for month in range(1, 13):
        last_day = _cal.monthrange(year, month)[1]
        m_end = datetime.date(year, month, last_day)
        mes   = f'{year}-{month:02d}'
        for email, entries in by_email.items():
            effective = None
            for d, cost in entries:
                if d <= m_end:
                    effective = cost
            if effective is None:
                continue
            if email not in data:
                data[email] = {}
            data[email][mes] = {'costo': effective, 'code': '', 'agreement': ''}

    total = sum(len(v) for v in data.values())
    print(f'  {len(data)} personas, {total} puntos de datos (lista {year}).')
    return data


def read_comp2020(svc):
    """
    Lee 2020 usando period='Actual' con prioridad.
    Para las 16 personas que solo tienen Fcst (sin ningún Actual), incluye
    sus entradas Fcst como fallback — sus costos están en USD/H y son válidos.
    """
    import calendar as _cal
    print('Leyendo lista 2020 (Actual + Fcst-fallback)...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_COMP20,
        range=f"'{TAB_LISTA}'!A:J",
        valueRenderOption='FORMATTED_VALUE',
    ).execute()
    rows = r.get('values', [])[1:]

    actual_entries = {}   # email → [(date, cost)]
    fcst_entries   = {}   # email → [(date, cost)]

    for row in rows:
        while len(row) < 4:
            row.append('')
        email  = str(row[1]).strip().lower()
        if '@' not in email:
            continue
        col6   = str(row[6]).strip() if len(row) > 6 else ''
        date_s = str(row[2]).strip()
        cost_s = str(row[3]).strip()
        d = None
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
            try:
                d = datetime.datetime.strptime(date_s, fmt).date()
                break
            except Exception:
                pass
        if d is None:
            continue
        try:
            cost = round(float(re.sub(r'[^\d.]', '', cost_s)), 4)
        except Exception:
            continue
        if cost <= 0:
            continue

        if col6 == 'Actual':
            actual_entries.setdefault(email, []).append((d, cost))
        elif col6 == 'Fcst':
            fcst_entries.setdefault(email, []).append((d, cost))

    # Personas sin ningún Actual → usar Fcst como fallback
    solo_fcst = {e: v for e, v in fcst_entries.items() if e not in actual_entries}
    print(f'  Personas con Actual: {len(actual_entries)}  |  solo Fcst (fallback): {len(solo_fcst)}')

    by_email = {**actual_entries}
    for email, entries in solo_fcst.items():
        by_email[email] = entries
    for email in by_email:
        by_email[email].sort()

    data = {}
    for month in range(1, 13):
        last_day = _cal.monthrange(2020, month)[1]
        m_end    = datetime.date(2020, month, last_day)
        mes      = f'2020-{month:02d}'
        for email, entries in by_email.items():
            effective = None
            for d, cost in entries:
                if d <= m_end:
                    effective = cost
            if effective is None:
                continue
            if email not in data:
                data[email] = {}
            data[email][mes] = {'costo': effective, 'code': '', 'agreement': ''}

    total = sum(len(v) for v in data.values())
    print(f'  {len(data)} personas, {total} puntos de datos (lista 2020).')
    return data

def read_comp2021(svc):
    return _read_lista_sheet(svc, SHEET_COMP21, 2021, skip_ars=True)

def read_comp2022(svc):
    return _read_lista_sheet(svc, SHEET_COMP22, 2022, skip_ars=True)


# ─── LECTURA COMPENSACIONES ANUALES ──────────────────────────────────────────

def _read_comp_sheet(svc, sheet_id, year, col_costo):
    """
    Lee una solapa 'Bamboo: Compensaciones -Sueldos MS' de un sheet anual.

    Encoding del mes (col A): los 12 seriales únicos se ordenan y mapean
    en orden a Jan–Dec del año dado. Funciona tanto con el encoding
    consecutivo (día 1-12 de enero) como con seriales reales de primer
    día de mes, o mezclas (caso 2023: Jan-Sep día, Oct-Dic real).

    col_costo: índice 0-based de la columna de Costo USD/H.
    """
    col_letter = chr(65 + col_costo)
    n_cols = col_costo + 1
    print(f'Leyendo Compensaciones {year} (col {col_letter})...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{TAB_COMP_MS}'!A2:{col_letter}",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = r.get('values', [])

    # Auto-detectar el mapeo serial → YYYY-MM
    serial_set = sorted({
        int(row[0]) for row in rows
        if row and str(row[0]).strip().isdigit()
    })
    if len(serial_set) != 12:
        print(f'  ADVERTENCIA: {len(serial_set)} seriales únicos (esperado 12). '
              f'Valores: {serial_set}')
    serial_to_mes = {s: f'{year}-{i+1:02d}' for i, s in enumerate(serial_set)}

    # Necesitamos al menos hasta col K (índice 10 = Agreement) además de costo
    n_cols = max(col_costo + 1, 11)

    data = {}   # email → {mes → {costo, code, agreement}}

    for row in rows:
        while len(row) < n_cols:
            row.append('')
        email     = str(row[2]).strip().lower()
        serial_s  = str(row[0]).strip()
        costo_raw = row[col_costo]
        code      = str(row[8]).strip()    # col I = Code
        agreement = str(row[10]).strip()   # col K = Agreement

        if '@' not in email or not serial_s.isdigit():
            continue

        mes = serial_to_mes.get(int(serial_s))
        if not mes:
            continue

        try:
            costo = round(float(costo_raw), 4)
        except Exception:
            continue
        if costo <= 0:
            continue

        # Normalizar #N/A
        code      = '' if code      in ('#N/A', '#VALUE!') else code
        agreement = '' if agreement in ('#N/A', '#VALUE!') else agreement

        if email not in data:
            data[email] = {}
        data[email][mes] = {'costo': costo, 'code': code, 'agreement': agreement}

    total_puntos = sum(len(v) for v in data.values())
    print(f'  {len(data)} personas, {total_puntos} puntos de datos (Comp {year}).')
    return data


def read_comp2023(svc):
    return _read_comp_sheet(svc, SHEET_COMP23, 2023, col_costo=20)

def read_comp2024(svc):
    return _read_comp_sheet(svc, SHEET_COMP24, 2024, col_costo=18)

def read_comp2025(svc):
    return _read_comp_sheet(svc, SHEET_COMP25, 2025, col_costo=25)


# ─── LECTURA BDD BAMBOO ───────────────────────────────────────────────────────

def read_bdd(svc):
    print('Leyendo BDD Bamboo...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_BDD,
        range=f"'{TAB_HIST_BDD}'!A:BC",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    all_rows = r.get('values', [])
    data_rows = all_rows[1:]

    nombres        = {}   # email → nombre
    costos         = {}   # email → {mes → {costo, code, agreement}}
    divisiones     = {}   # email → division (último registro)
    countries_map  = {}   # email → country (último registro)
    seniorities_map = {}  # email → seniority (último registro)
    departments_map = {}  # email → department (último registro)
    doppler_emails = set()

    for row in data_rows:
        while len(row) < 24:
            row.append('')
        month_raw = str(row[0]).strip()
        email     = str(row[2]).strip().lower()
        nombre    = str(row[1]).strip()
        costo_raw = row[23]
        code      = str(row[8]).strip()   if len(row) > 8  else ''
        agreement = str(row[10]).strip()  if len(row) > 10 else ''

        if '@' not in email:
            continue

        mes = _parse_bdd_month(month_raw)
        if not mes:
            continue

        try:
            costo = round(float(costo_raw), 4) if costo_raw != '' else None
        except Exception:
            costo = None

        if costo is None or costo == 0:
            continue

        code      = '' if code      in ('#N/A', '#VALUE!') else code
        agreement = '' if agreement in ('#N/A', '#VALUE!') else agreement

        if email not in costos:
            costos[email] = {}
        costos[email][mes] = {'costo': costo, 'code': code, 'agreement': agreement}

        division   = str(row[46]).strip() if len(row) > 46 else ''
        country    = str(row[44]).strip() if len(row) > 44 else ''
        seniority  = _normalize_seniority(str(row[9]).strip() if len(row) > 9 else '')
        department = str(row[4]).strip()  if len(row) > 4  else ''

        if 'doppler' in division.lower():
            doppler_emails.add(email)

        if division:   divisiones[email]     = division
        if country:    countries_map[email]  = country
        if seniority:  seniorities_map[email] = seniority
        if department: departments_map[email] = department
        if nombre and email not in nombres:
            nombres[email] = nombre

    total_puntos = sum(len(v) for v in costos.values())
    print(f'  {len(costos)} personas, {total_puntos} puntos de datos (BDD).')
    print(f'  División Doppler identificados: {len(doppler_emails)}')
    return costos, nombres, doppler_emails, divisiones, countries_map, seniorities_map, departments_map


# ─── COMBINAR ─────────────────────────────────────────────────────────────────

def _norm_emails(d):
    """Normaliza emails con typos conocidos en un dict {email: {mes: ...}}."""
    out = {}
    for k, v in d.items():
        k2 = EMAIL_CORRECTIONS.get(k, k)
        if k2 in out:
            # el entry existente tiene prioridad; solo agrego meses que falten
            for mes, val in v.items():
                out[k2].setdefault(mes, val)
        else:
            out[k2] = v
    return out


def combinar(old_data, comp21_data, comp22_data,
             comp23_data, comp24_data, comp25_data, bdd_data, nombres,
             doppler_emails=None, divisiones=None, countries_map=None,
             seniorities_map=None, departments_map=None):
    """
    Une siete fuentes. Prioridad (mayor sobreescribe):
      BDD > Comp2025 > Comp2024 > Comp2023 > sheet viejo > lista2022 > lista2021
    doppler_emails: set de emails a excluir (empleados Doppler).
    """
    doppler_emails  = doppler_emails  or set()
    divisiones      = divisiones      or {}
    countries_map   = countries_map   or {}
    seniorities_map = seniorities_map or {}
    departments_map = departments_map or {}

    # Normalizar typos de email en fuentes históricas
    old_data    = _norm_emails(old_data)
    comp21_data = _norm_emails(comp21_data)
    comp22_data = _norm_emails(comp22_data)
    comp23_data = _norm_emails(comp23_data)
    comp24_data = _norm_emails(comp24_data)
    comp25_data = _norm_emails(comp25_data)

    all_emails = (set(old_data.keys())
                  | set(comp21_data.keys()) | set(comp22_data.keys())
                  | set(comp23_data.keys()) | set(comp24_data.keys()) | set(comp25_data.keys())
                  | set(bdd_data.keys()))
    personas = []

    for email in sorted(all_emails):
        is_doppler = email in doppler_emails
        puntos_old = old_data.get(email, {})
        puntos_c21 = comp21_data.get(email, {})
        puntos_c22 = comp22_data.get(email, {})
        puntos_c23 = comp23_data.get(email, {})
        puntos_c24 = comp24_data.get(email, {})
        puntos_c25 = comp25_data.get(email, {})
        puntos_bdd = bdd_data.get(email, {})

        serie = {}
        # Prioridad ascendente (mayor prioridad sobreescribe)
        for mes, v in puntos_c21.items():
            serie[mes] = {**v, 'fuente': 'lista21'}
        for mes, v in puntos_c22.items():
            serie[mes] = {**v, 'fuente': 'lista22'}
        for mes, v in puntos_old.items():
            serie[mes] = {**v, 'fuente': 'hist'}
        for mes, v in puntos_c23.items():
            serie[mes] = {**v, 'fuente': 'comp23'}
        for mes, v in puntos_c24.items():
            serie[mes] = {**v, 'fuente': 'comp24'}
        for mes, v in puntos_c25.items():
            serie[mes] = {**v, 'fuente': 'comp25'}
        for mes, v in puntos_bdd.items():
            serie[mes] = {**v, 'fuente': 'bdd'}

        if not serie:
            continue

        nombre     = nombres.get(email, '')
        division   = divisiones.get(email, '')
        country    = countries_map.get(email, '')
        seniority  = seniorities_map.get(email, '')
        department = departments_map.get(email, '')

        # Propagar el código más reciente hacia atrás a meses sin código.
        # Permite filtrar por SEN 06 y ver toda la historia de esa persona.
        sorted_meses = sorted(serie.keys())
        current_code = ''
        for mes in reversed(sorted_meses):
            c = serie[mes].get('code', '')
            if c and c not in ('#N/A', '#VALUE!'):
                current_code = c
                break
        if current_code:
            for mes in sorted_meses:
                if not serie[mes].get('code'):
                    serie[mes]['code'] = current_code

        if is_doppler:
            continue   # excluir Doppler del HTML (se mantiene en BDD)

        datos_sorted = [
            {'mes': mes, 'costo': v['costo'], 'fuente': v['fuente'],
             'code': v.get('code', ''),
             'agreement': AGREEMENT_CORRECTIONS.get(v.get('agreement', ''), v.get('agreement', ''))}
            for mes, v in sorted(serie.items())
            if mes >= '2021-01'
        ]
        personas.append({
            'email':      email,
            'nombre':     nombre,
            'division':   division,
            'country':    country,
            'level':      seniority,
            'department': department,
            'datos':      datos_sorted,
        })

    print(f'Total combinado: {len(personas)} personas.')
    return personas


# ─── GENERAR HTML ─────────────────────────────────────────────────────────────

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Evolución de Costo Mensual</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>
  :root {
    --ms-green:       #33AD73;
    --ms-green-dark:  #27945E;
    --ms-green-light: #E8F5E9;
    --ms-dark:        #0F1923;
    --ms-text:        #1A2332;
    --ms-text-light:  #64748B;
    --ms-border:      #E2E8F0;
  }
  body { background: #F4F6F9; font-family: 'Segoe UI', system-ui, sans-serif; color: var(--ms-text); }
  .ms-header {
    background: var(--ms-dark); border-radius: 10px;
    padding: 16px 24px; margin-bottom: 24px;
  }
  .ms-header-title { color: #fff; font-size: 1.15rem; font-weight: 700; margin: 0 0 3px 0; }
  .ms-header-sub {
    color: var(--ms-green); font-size: 0.7rem; font-weight: 700;
    letter-spacing: 0.09em; text-transform: uppercase; margin: 0 0 4px 0;
  }
  .ms-header-desc { color: #94A3B8; font-size: 0.78rem; margin: 0; }
  .card { border-radius: 10px; box-shadow: 0 1px 8px rgba(0,0,0,.07); border: 1px solid var(--ms-border) !important; }
  .section-label {
    font-size: 0.68rem; font-weight: 700; color: var(--ms-text-light);
    border-left: 3px solid var(--ms-green); padding-left: 8px;
    letter-spacing: 0.07em; text-transform: uppercase; margin-bottom: 12px;
  }
  #autocomplete-list {
    position: absolute; z-index: 1000; width: 100%;
    background: #fff; border: 1px solid var(--ms-border); border-top: none;
    border-radius: 0 0 8px 8px; max-height: 260px; overflow-y: auto;
    box-shadow: 0 4px 12px rgba(0,0,0,.1);
  }
  .ac-item { padding: 9px 14px; cursor: pointer; font-size: .9rem; }
  .ac-item:hover, .ac-item.active { background: var(--ms-green-light); }
  .ac-item .email { color: var(--ms-text-light); font-size: .78rem; }
  #chart-wrap { position: relative; height: 360px; }
  .stat-card { border-left: 4px solid var(--ms-green) !important; }
  .stat-value { color: var(--ms-dark); font-size: 1.1rem; font-weight: 700; }
  th { background: var(--ms-dark) !important; color: #fff !important; font-weight: 500; font-size: .82rem; letter-spacing: .03em; }
  tr:nth-child(even) td { background: #F7FAFC; }
  td { font-size: .88rem; }
  .source-dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 4px; }
  .dot-hist   { background: #94A3B8; }
  .dot-comp23 { background: #E65100; }
  .dot-comp24 { background: #F59E0B; }
  .dot-comp25 { background: var(--ms-green); }
  .dot-bdd    { background: var(--ms-dark); }
  #no-data { display: none; }
  #empty-state { display: flex; flex-direction: column; align-items: center;
    justify-content: center; height: 280px; color: #CBD5E1; }
  .nav-pills .nav-link { color: var(--ms-text); font-weight: 500; border-radius: 6px; padding: 7px 18px; }
  .nav-pills .nav-link:hover { background: var(--ms-green-light); color: var(--ms-green-dark); }
  .nav-pills .nav-link.active { background-color: var(--ms-dark); color: #fff; }
  .form-select, .form-control { border-color: var(--ms-border); font-size: .88rem; }
  .form-select:focus, .form-control:focus { border-color: var(--ms-green); box-shadow: 0 0 0 .2rem rgba(51,173,115,.2); }
  .form-label { font-size: .72rem; font-weight: 700; color: var(--ms-text-light); letter-spacing: .05em; text-transform: uppercase; }
  .card-header { background: #F8FAFC !important; border-bottom: 1px solid var(--ms-border); }
</style>
</head>
<body>
<div class="container py-4" style="max-width:960px">

  <div class="ms-header">
    <p class="ms-header-sub">Making Sense &middot; Talent Care</p>
    <h4 class="ms-header-title">Evolución de Costo Mensual</h4>
    <p class="ms-header-desc">USD/H &nbsp;&middot;&nbsp; MS &middot; Viallion &middot; Holding (excluye Doppler)</p>
  </div>

  <!-- Navegación -->
  <ul class="nav nav-pills mb-4">
    <li class="nav-item">
      <button class="nav-link active" id="btn-agg" onclick="switchTab('agg')">Costo Mensual</button>
    </li>
    <li class="nav-item ms-2">
      <button class="nav-link" id="btn-level" onclick="switchTab('level')">Por Level</button>
    </li>
    <li class="nav-item ms-2">
      <button class="nav-link" id="btn-persona" onclick="switchTab('persona')">Por Persona</button>
    </li>
  </ul>

  <!-- ═══ SECCIÓN 1: ANÁLISIS AGREGADO ═══════════════════════════════════════ -->
  <div id="section-agg">

    <!-- Filtros -->
    <div class="card p-3 mb-3">
      <div class="row g-2 align-items-end">
        <div class="col-md-3">
          <label class="form-label small fw-semibold text-muted mb-1">División</label>
          <select id="sel-division" class="form-select" onchange="updateAggChart()">
            <option value="">Todas</option>
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label small fw-semibold text-muted mb-1">País</label>
          <select id="sel-country" class="form-select" onchange="updateAggChart()">
            <option value="">Todos</option>
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label small fw-semibold text-muted mb-1">Agreement</label>
          <select id="sel-agreement" class="form-select" onchange="updateAggChart()">
            <option value="">Todos</option>
          </select>
        </div>
        <div class="col-md-3">
          <label class="form-label small fw-semibold text-muted mb-1">Code</label>
          <select id="sel-code" class="form-select" onchange="updateAggChart()">
            <option value="">Todos</option>
          </select>
        </div>
      </div>
    </div>

    <!-- Cards de resumen -->
    <div class="row g-3 mb-3">
      <div class="col-md-4">
        <div class="card p-3 h-100 stat-card">
          <div class="form-label mb-1">Mediana último mes</div>
          <div id="agg-stat-med" class="stat-value">—</div>
          <div id="agg-stat-med-rango" class="text-muted small"></div>
          <div id="agg-stat-med-mes" class="text-muted small"></div>
        </div>
      </div>
      <div class="col-md-4">
        <div class="card p-3 h-100 stat-card">
          <div class="form-label mb-1">Variación anual</div>
          <div id="agg-stat-var" class="stat-value" style="color:var(--ms-text-light)">—</div>
          <div id="agg-stat-var-label" class="text-muted small"></div>
        </div>
      </div>
      <div class="col-md-4">
        <div class="card p-3 h-100 stat-card">
          <div class="form-label mb-1">Personas activas</div>
          <div id="agg-stat-active" class="stat-value">—</div>
          <div class="text-muted small">con datos en los últimos 3 meses</div>
        </div>
      </div>
    </div>

    <!-- Gráfico -->
    <div class="card p-3 mb-2">
      <div style="position:relative;height:380px">
        <canvas id="aggChart"></canvas>
      </div>
    </div>
    <div class="d-flex justify-content-end align-items-center gap-3 mb-0">
      <span class="small text-muted">
        <span style="display:inline-block;width:28px;height:10px;background:rgba(51,173,115,0.18);border-radius:3px;vertical-align:middle"></span>
        Rango P25–P75
      </span>
      <span class="small text-muted">
        <span style="display:inline-block;width:28px;height:2px;background:#0F1923;vertical-align:middle"></span>
        Mediana
      </span>
    </div>

  </div>

  <!-- ═══ SECCIÓN 2: POR LEVEL ══════════════════════════════════════════════ -->
  <div id="section-level" style="display:none">

    <!-- Filtro departamento / agreement -->
    <div class="card p-3 mb-3">
      <div class="row g-2 align-items-end">
        <div class="col-md-6">
          <label class="form-label">Departamento</label>
          <select id="sel-sen-dept" class="form-select" onchange="updateSeniorityView()">
            <option value="">Todos los departamentos</option>
          </select>
        </div>
        <div class="col-md-6">
          <label class="form-label">Agreement</label>
          <select id="sel-sen-agreement" class="form-select" onchange="updateSeniorityView()">
            <option value="">Todos</option>
          </select>
        </div>
      </div>
    </div>

    <!-- Gráfico -->
    <div class="card p-3 mb-3">
      <div class="section-label mb-3">Costo promedio anual por level (USD/H)</div>
      <div style="position:relative;height:380px">
        <canvas id="senChart"></canvas>
      </div>
    </div>

    <!-- Tabla -->
    <div class="card">
      <div class="card-header py-2">
        <span class="fw-semibold" style="font-size:.88rem">Costo promedio USD/H por año</span>
      </div>
      <div style="overflow-x:auto">
        <table class="table table-sm mb-0" id="sen-table"></table>
      </div>
    </div>

  </div>

  <!-- ═══ SECCIÓN 3: POR PERSONA ══════════════════════════════════════════════ -->
  <div id="section-persona" style="display:none">

    <!-- Buscador -->
    <div class="card p-3 mb-4">
      <div class="position-relative">
        <input id="search" type="text" class="form-control form-control-lg"
          placeholder="Buscar por nombre o email..." autocomplete="off">
        <div id="autocomplete-list"></div>
      </div>
    </div>

    <!-- Stats row -->
    <div id="stats-row" class="row g-3 mb-3" style="display:none!important">
      <div class="col-4">
        <div class="card p-3 stat-card">
          <div class="form-label mb-1">Último costo</div>
          <div id="stat-ultimo" class="stat-value">—</div>
          <div id="stat-ultimo-mes" class="text-muted small"></div>
        </div>
      </div>
      <div class="col-4">
        <div class="card p-3 stat-card">
          <div class="form-label mb-1">Primer registro</div>
          <div id="stat-primero" class="stat-value">—</div>
          <div id="stat-primero-mes" class="text-muted small"></div>
        </div>
      </div>
      <div class="col-4">
        <div class="card p-3 stat-card">
          <div class="form-label mb-1">Variación total</div>
          <div id="stat-variacion" class="stat-value">—</div>
          <div class="text-muted small">vs primer mes</div>
        </div>
      </div>
    </div>

    <!-- Gráfico -->
    <div id="chart-area" class="card p-3 mb-4">
      <div id="empty-state">
        <svg width="64" height="64" fill="none" stroke="#dee2e6" stroke-width="2"
          viewBox="0 0 24 24"><path d="M3 3v18h18"/><path d="m7 16 4-4 4 4 4-6"/></svg>
        <p class="mt-3">Buscá una persona para ver su evolución de costo</p>
      </div>
      <div id="chart-wrap" style="display:none">
        <canvas id="myChart"></canvas>
      </div>
    </div>

    <!-- Leyenda fuentes -->
    <div id="legend-sources" class="mb-3 small text-muted" style="display:none">
      <span class="source-dot dot-hist"></span> Histórico
      &nbsp;&nbsp;
      <span class="source-dot dot-comp23"></span> Comp 2023
      &nbsp;&nbsp;
      <span class="source-dot dot-comp24"></span> Comp 2024
      &nbsp;&nbsp;
      <span class="source-dot dot-comp25"></span> Comp 2025
      &nbsp;&nbsp;
      <span class="source-dot dot-bdd"></span> BDD Bamboo
    </div>

    <!-- Tabla -->
    <div id="tabla-area" class="card" style="display:none">
      <div class="card-header d-flex justify-content-between align-items-center py-2">
        <span class="fw-semibold" id="tabla-titulo">Detalle mensual</span>
        <span id="tabla-count" class="badge bg-secondary"></span>
      </div>
      <div style="max-height:340px;overflow-y:auto">
        <table class="table table-sm mb-0">
          <thead><tr>
            <th>Mes</th><th id="tabla-costo-header" class="text-end">Costo USD/H</th><th>Fuente</th>
          </tr></thead>
          <tbody id="tabla-body"></tbody>
        </table>
      </div>
    </div>

    <div id="no-data" class="alert alert-warning mt-3">
      Sin datos de costo para esta persona.
    </div>

  </div><!-- /section-persona -->
</div><!-- /container -->

<script>
const DATA = __DATA__;

// ── Helpers ────────────────────────────────────────────────────────────────
function fmtMes(s) {
  const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const parts = s.split('-');
  return MESES[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
}

function fColor(f) {
  if (f === 'bdd')    return '#0F1923';
  if (f === 'comp25') return '#33AD73';
  if (f === 'comp24') return '#F59E0B';
  if (f === 'comp23') return '#E65100';
  return '#94A3B8';
}

function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort(function(a, b) { return a - b; });
  const m = Math.floor(s.length / 2);
  return s.length % 2 !== 0 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function percentile(arr, p) {
  if (!arr.length) return null;
  const s = [...arr].sort(function(a, b) { return a - b; });
  const idx = (p / 100) * (s.length - 1);
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  return lo === hi ? s[lo] : s[lo] + (s[hi] - s[lo]) * (idx - lo);
}

// ── Navegación entre secciones ─────────────────────────────────────────────
function switchTab(tab) {
  document.getElementById('section-agg').style.display     = tab === 'agg'     ? '' : 'none';
  document.getElementById('section-level').style.display   = tab === 'level'   ? '' : 'none';
  document.getElementById('section-persona').style.display = tab === 'persona' ? '' : 'none';
  document.getElementById('btn-agg').classList.toggle('active',     tab === 'agg');
  document.getElementById('btn-level').classList.toggle('active',   tab === 'level');
  document.getElementById('btn-persona').classList.toggle('active', tab === 'persona');
  if (tab === 'level' && !senChart) updateSeniorityView();
}

// ── Sección 1: Costo Mensual ────────────────────────────────────────────────
var selDivision  = document.getElementById('sel-division');
var selCode      = document.getElementById('sel-code');
var selAgreement = document.getElementById('sel-agreement');

DATA.divisions.forEach(function(d) {
  var o = document.createElement('option');
  o.value = d; o.textContent = d;
  selDivision.appendChild(o);
});
DATA.codes.forEach(function(c) {
  var o = document.createElement('option');
  o.value = c; o.textContent = c;
  selCode.appendChild(o);
});
DATA.agreements.forEach(function(a) {
  var o = document.createElement('option');
  o.value = a; o.textContent = a;
  selAgreement.appendChild(o);
});
var selCountry = document.getElementById('sel-country');
DATA.countries.forEach(function(c) {
  var o = document.createElement('option');
  o.value = c; o.textContent = c;
  selCountry.appendChild(o);
});

var aggChart = null;

function updateAggChart() {
  var divFilter     = selDivision.value;
  var codeFilter    = selCode.value;
  var agrFilter     = selAgreement.value;
  var countryFilter = selCountry.value;

  var mesMap = {};
  DATA.personas.forEach(function(p) {
    var divOk     = !divFilter     || p.division === divFilter;
    var countryOk = !countryFilter || p.country  === countryFilter;
    if (!divOk || !countryOk) return;
    p.datos.forEach(function(d) {
      if (d.costo <= 0) return;
      var codeOk = !codeFilter || d.code === codeFilter;
      var agrOk  = !agrFilter  || d.agreement === agrFilter;
      if (codeOk && agrOk) {
        if (!mesMap[d.mes]) mesMap[d.mes] = [];
        mesMap[d.mes].push(d.costo);
      }
    });
  });

  var unidadAgg = 'USD/H';

  var sortedMeses = Object.keys(mesMap).sort();
  var labels  = sortedMeses.map(fmtMes);
  var medVals = sortedMeses.map(function(m) { return median(mesMap[m]); });
  var p25Vals = sortedMeses.map(function(m) { return percentile(mesMap[m], 25); });
  var p75Vals = sortedMeses.map(function(m) { return percentile(mesMap[m], 75); });
  var counts  = sortedMeses.map(function(m) { return mesMap[m].length; });

  // ── Cards de resumen ────────────────────────────────────────────────────
  if (sortedMeses.length) {
    var lastIdx = sortedMeses.length - 1;
    var lastMes = sortedMeses[lastIdx];
    var lastMed = medVals[lastIdx];
    var lastP25 = p25Vals[lastIdx];
    var lastP75 = p75Vals[lastIdx];

    document.getElementById('agg-stat-med').style.color = 'var(--ms-dark)';
  document.getElementById('agg-stat-active').style.color = 'var(--ms-green)';
  document.getElementById('agg-stat-med').textContent      = lastMed !== null ? '$' + lastMed.toFixed(2) + ' ' + unidadAgg : '—';
    document.getElementById('agg-stat-med-rango').textContent = (lastP25 !== null && lastP75 !== null)
      ? 'P25–P75: $' + lastP25.toFixed(2) + ' – $' + lastP75.toFixed(2) : '';
    document.getElementById('agg-stat-med-mes').textContent   = fmtMes(lastMes);

    // Variación anual: mismo mes del año anterior
    var parts = lastMes.split('-');
    var mes12 = (parseInt(parts[0], 10) - 1) + '-' + parts[1];
    var idx12 = sortedMeses.indexOf(mes12);
    var varEl    = document.getElementById('agg-stat-var');
    var varLabel = document.getElementById('agg-stat-var-label');
    if (idx12 >= 0 && medVals[idx12] !== null && lastMed !== null) {
      var varAbs = lastMed - medVals[idx12];
      var varPct = medVals[idx12] ? (varAbs / medVals[idx12] * 100) : 0;
      var sign   = varAbs >= 0 ? '+' : '';
      varEl.textContent = sign + '$' + varAbs.toFixed(2) + ' (' + sign + varPct.toFixed(1) + '%)';
      varEl.style.color = varAbs >= 0 ? 'var(--ms-green-dark)' : '#C62828';
      varEl.className   = 'stat-value';
      varLabel.textContent = 'vs ' + fmtMes(mes12);
    } else {
      varEl.textContent  = '—';
      varEl.className    = 'fs-5 fw-bold text-muted';
      varLabel.textContent = 'sin datos hace 12 meses';
    }

    // Personas activas: últimos 3 meses con datos
    var last3 = new Set(sortedMeses.slice(-3));
    var activeEmails = new Set();
    DATA.personas.forEach(function(p) {
      var divOk     = !divFilter     || p.division === divFilter;
      var countryOk = !countryFilter || p.country  === countryFilter;
      if (!divOk || !countryOk) return;
      p.datos.forEach(function(d) {
        if (!last3.has(d.mes) || d.costo <= 0) return;
        var codeOk = !codeFilter || d.code === codeFilter;
        var agrOk  = !agrFilter  || d.agreement === agrFilter;
        if (codeOk && agrOk) activeEmails.add(p.email);
      });
    });
    document.getElementById('agg-stat-active').textContent = activeEmails.size;
  } else {
    document.getElementById('agg-stat-med').textContent      = '—';
    document.getElementById('agg-stat-med-rango').textContent = '';
    document.getElementById('agg-stat-med-mes').textContent   = '';
    document.getElementById('agg-stat-var').textContent       = '—';
    document.getElementById('agg-stat-var').className         = 'fs-5 fw-bold text-muted';
    document.getElementById('agg-stat-var-label').textContent = '';
    document.getElementById('agg-stat-active').textContent    = '0';
  }

  // ── Gráfico ─────────────────────────────────────────────────────────────
  if (aggChart) aggChart.destroy();
  aggChart = new Chart(document.getElementById('aggChart'), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'P25',
          data: p25Vals,
          borderColor: 'rgba(51,173,115,0.25)',
          borderWidth: 1,
          pointRadius: 0,
          fill: false,
          tension: 0.3,
        },
        {
          label: 'P75',
          data: p75Vals,
          borderColor: 'rgba(51,173,115,0.25)',
          borderWidth: 1,
          pointRadius: 0,
          fill: '-1',
          backgroundColor: 'rgba(51,173,115,0.10)',
          tension: 0.3,
        },
        {
          label: 'Mediana',
          data: medVals,
          borderColor: '#0F1923',
          pointBackgroundColor: '#33AD73',
          pointBorderColor: '#0F1923',
          pointBorderWidth: 1.5,
          pointRadius: 3.5,
          pointHoverRadius: 6,
          borderWidth: 2.5,
          fill: false,
          tension: 0.3,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              if (ctx.datasetIndex === 2) return 'Mediana: $' + ctx.parsed.y.toFixed(2) + ' ' + unidadAgg;
              if (ctx.datasetIndex === 1) return 'P75: $' + ctx.parsed.y.toFixed(2) + ' ' + unidadAgg;
              return 'P25: $' + ctx.parsed.y.toFixed(2) + ' ' + unidadAgg;
            },
            afterBody: function(items) {
              var i = items.findIndex(function(x) { return x.datasetIndex === 2; });
              if (i >= 0) return counts[items[i].dataIndex] + ' personas en ese mes';
              return '';
            },
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxTicksLimit: 18, font: { size: 11 } } },
        y: { grid: { color: '#f0f0f0' }, ticks: { callback: function(v) { return '$' + v.toFixed(0); } } }
      }
    }
  });
}

updateAggChart();

// ── Sección 2: Por Level ───────────────────────────────────────────────────
var selSenDept = document.getElementById('sel-sen-dept');
DATA.departments.forEach(function(d) {
  var o = document.createElement('option');
  o.value = d; o.textContent = d;
  selSenDept.appendChild(o);
});

var selSenAgreement = document.getElementById('sel-sen-agreement');
DATA.agreements.forEach(function(a) {
  var o = document.createElement('option');
  o.value = a; o.textContent = a;
  selSenAgreement.appendChild(o);
});

var senChart = null;
var YEARS = ['2021','2022','2023','2024','2025','2026'];
var SEN_COLORS = [
  '#0F1923','#33AD73','#F59E0B','#E65100','#3B82F6',
  '#8B5CF6','#EC4899','#14B8A6','#64748B','#84CC16',
  '#0EA5E9','#F97316','#D946EF','#6366F1'
];
var SEN_ORDER = [
  '01','02','03','04','05','06','07','08','09','10','11','12','13',
  'Tr','Pas','Jr','Jr Adv','SSr','SSr Adv','Sr','Lead','PdM','Manager','HoD','C-Level','C-level'
];

function updateSeniorityView() {
  var deptFilter = selSenDept.value;
  var agrFilter  = selSenAgreement.value;

  // Acumular: level → year → [costos]
  var bySnYear = {};

  DATA.personas.forEach(function(p) {
    if (!p.level) return;
    if (deptFilter && p.department !== deptFilter) return;
    p.datos.forEach(function(d) {
      if (d.costo <= 0) return;
      if (agrFilter && d.agreement !== agrFilter) return;
      var yr = d.mes.slice(0, 4);
      if (YEARS.indexOf(yr) === -1) return;
      if (!bySnYear[p.level]) bySnYear[p.level] = {};
      if (!bySnYear[p.level][yr]) bySnYear[p.level][yr] = [];
      bySnYear[p.level][yr].push(d.costo);
    });
  });

  var allSen = Object.keys(bySnYear);
  var seniorities = SEN_ORDER.filter(function(s) { return allSen.indexOf(s) !== -1; })
    .concat(allSen.filter(function(s) { return SEN_ORDER.indexOf(s) === -1; }).sort());

  // ── Gráfico ────────────────────────────────────────────────────────────
  var datasets = seniorities.map(function(sn, i) {
    var color = SEN_COLORS[i % SEN_COLORS.length];
    return {
      label: sn,
      data: YEARS.map(function(yr) {
        var vals = (bySnYear[sn] && bySnYear[sn][yr]) || [];
        if (!vals.length) return null;
        return vals.reduce(function(a, b) { return a + b; }, 0) / vals.length;
      }),
      borderColor: color,
      pointBackgroundColor: color,
      backgroundColor: 'transparent',
      borderWidth: 2,
      pointRadius: 3.5,
      pointHoverRadius: 6,
      tension: 0.2,
      spanGaps: true,
      hidden: true,
    };
  });

  if (senChart) senChart.destroy();
  senChart = new Chart(document.getElementById('senChart'), {
    type: 'line',
    data: { labels: YEARS, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'right', labels: { boxWidth: 10, padding: 12, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              return ctx.dataset.label + ': $' + (ctx.parsed.y !== null ? ctx.parsed.y.toFixed(2) : '—');
            }
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 12 } } },
        y: {
          grid: { color: '#f0f0f0' },
          ticks: {
            callback: function(v) {
              var label = '$' + v.toFixed(0);
              if (label === this._lastSenLabel) return '';
              this._lastSenLabel = label;
              return label;
            }
          }
        }
      }
    }
  });

  // ── Tabla ──────────────────────────────────────────────────────────────
  var table = document.getElementById('sen-table');
  var html = '<thead><tr><th>Level</th>';
  YEARS.forEach(function(yr) { html += '<th class="text-end">' + yr + '</th>'; });
  html += '</tr></thead><tbody>';

  seniorities.forEach(function(sn) {
    var color = SEN_COLORS[seniorities.indexOf(sn) % SEN_COLORS.length];
    html += '<tr><td><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + color + ';margin-right:6px"></span><span class="fw-semibold">' + sn + '</span></td>';
    var prev = null;
    YEARS.forEach(function(yr) {
      var vals = (bySnYear[sn] && bySnYear[sn][yr]) || [];
      var avg = vals.length ? vals.reduce(function(a, b) { return a + b; }, 0) / vals.length : null;
      var cell = avg !== null ? '$' + avg.toFixed(2) : '—';
      var style = '';
      if (avg !== null && prev !== null) {
        style = avg > prev ? 'color:var(--ms-green-dark);font-weight:600'
                           : avg < prev ? 'color:#C62828;font-weight:600' : '';
      }
      html += '<td class="text-end num" style="' + style + '">' + cell + '</td>';
      if (avg !== null) prev = avg;
    });
    html += '</tr>';
  });
  html += '</tbody>';
  table.innerHTML = html;
}

// ── Sección 3: Por Persona ─────────────────────────────────────────────────
var idx = DATA.personas.map(function(p) {
  return {
    email:      p.email,
    nombre:     p.nombre,
    search: (p.nombre + ' ' + p.email).toLowerCase(),
    datos:  p.datos,
  };
});

var chart = null;
var input = document.getElementById('search');
var list  = document.getElementById('autocomplete-list');

function renderList(items) {
  list.innerHTML = '';
  items.slice(0, 12).forEach(function(p, i) {
    var d = document.createElement('div');
    d.className = 'ac-item' + (i === 0 ? ' active' : '');
    d.innerHTML = '<div class="fw-semibold">' + (p.nombre || p.email) + '</div>'
                + '<div class="email">' + p.email + '</div>';
    d.addEventListener('mousedown', function(e) { e.preventDefault(); selectPerson(p); });
    list.appendChild(d);
  });
}

input.addEventListener('input', function() {
  var q = input.value.trim().toLowerCase();
  if (!q) { list.innerHTML = ''; return; }
  renderList(idx.filter(function(p) { return p.search.includes(q); }));
});

input.addEventListener('keydown', function(e) {
  var items  = list.querySelectorAll('.ac-item');
  var active = Array.from(items).findIndex(function(el) { return el.classList.contains('active'); });
  if (e.key === 'ArrowDown') {
    if (active < items.length - 1) {
      if (items[active]) items[active].classList.remove('active');
      items[active + 1].classList.add('active');
    }
  } else if (e.key === 'ArrowUp') {
    if (active > 0) {
      items[active].classList.remove('active');
      items[active - 1].classList.add('active');
    }
  } else if (e.key === 'Enter') {
    var sel = list.querySelector('.ac-item.active');
    if (sel) sel.dispatchEvent(new MouseEvent('mousedown'));
  } else if (e.key === 'Escape') {
    list.innerHTML = '';
  }
});

document.addEventListener('click', function(e) {
  if (!input.contains(e.target)) list.innerHTML = '';
});

function selectPerson(p) {
  list.innerHTML = '';
  input.value = p.nombre || p.email;
  showDashboard(p);
}

function showDashboard(p) {
  var noData    = document.getElementById('no-data');
  var chartWrap = document.getElementById('chart-wrap');
  var emptyState= document.getElementById('empty-state');
  var tablaArea = document.getElementById('tabla-area');
  var statsRow  = document.getElementById('stats-row');
  var legend    = document.getElementById('legend-sources');

  noData.style.display    = 'none';
  tablaArea.style.display = 'none';
  statsRow.style.display  = 'none';
  legend.style.display    = 'none';

  var datos = p.datos.filter(function(d) { return d.costo > 0; });

  if (!datos.length) {
    emptyState.style.display = 'flex';
    chartWrap.style.display  = 'none';
    noData.style.display     = 'block';
    return;
  }

  emptyState.style.display = 'none';
  chartWrap.style.display  = 'block';
  tablaArea.style.display  = 'block';
  statsRow.style.removeProperty('display');
  legend.style.display     = 'block';

  var primero   = datos[0];
  var ultimo    = datos[datos.length - 1];
  var variacion = ultimo.costo - primero.costo;
  var variPct   = primero.costo ? (variacion / primero.costo * 100) : 0;
  document.getElementById('stat-ultimo').textContent      = '$' + ultimo.costo.toFixed(2) + ' USD/H';
  document.getElementById('stat-ultimo-mes').textContent  = fmtMes(ultimo.mes);
  document.getElementById('stat-primero').textContent     = '$' + primero.costo.toFixed(2) + ' USD/H';
  document.getElementById('stat-primero-mes').textContent = fmtMes(primero.mes);
  document.getElementById('tabla-costo-header').textContent = 'Costo USD/H';

  var varEl = document.getElementById('stat-variacion');
  var sign  = variacion >= 0 ? '+' : '';
  varEl.textContent = sign + '$' + variacion.toFixed(2) + ' (' + sign + variPct.toFixed(1) + '%)';
  varEl.className   = 'stat-value';
  varEl.style.color = variacion >= 0 ? 'var(--ms-green-dark)' : '#C62828';

  var labels = datos.map(function(d) { return fmtMes(d.mes); });
  var values = datos.map(function(d) { return d.costo; });
  var colors = datos.map(function(d) { return fColor(d.fuente); });

  if (chart) chart.destroy();
  chart = new Chart(document.getElementById('myChart'), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Costo USD/H',
        data: values,
        borderColor: colors,
        pointBackgroundColor: colors,
        pointRadius: 4,
        pointHoverRadius: 6,
        fill: false,
        tension: 0.2,
        segment: {
          borderColor: function(ctx) { return fColor(datos[ctx.p0DataIndex] && datos[ctx.p0DataIndex].fuente); }
        },
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) { return '$' + ctx.parsed.y.toFixed(2) + unidad; },
            afterLabel: function(ctx) {
              var f = datos[ctx.dataIndex].fuente;
              return f === 'bdd'    ? 'Fuente: BDD Bamboo'
                   : f === 'comp25' ? 'Fuente: Comp 2025'
                   : f === 'comp24' ? 'Fuente: Comp 2024'
                   : f === 'comp23' ? 'Fuente: Comp 2023'
                   : 'Fuente: Histórico';
            },
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { maxTicksLimit: 18, font: { size: 11 } } },
        y: { grid: { color: '#f0f0f0' }, ticks: { callback: function(v) { return '$' + v.toFixed(0); } } }
      }
    }
  });

  var tbody = document.getElementById('tabla-body');
  tbody.innerHTML = '';
  datos.slice().reverse().forEach(function(d) {
    var tr = document.createElement('tr');
    var fuente = d.fuente === 'bdd'
      ? '<span class="badge" style="background:#0F1923;color:#fff">BDD</span>'
      : d.fuente === 'comp25'
      ? '<span class="badge" style="background:#E8F5E9;color:#27945E">Comp 2025</span>'
      : d.fuente === 'comp24'
      ? '<span class="badge" style="background:#FFF8E1;color:#E65100">Comp 2024</span>'
      : d.fuente === 'comp23'
      ? '<span class="badge" style="background:#FFF0F0;color:#C62828">Comp 2023</span>'
      : '<span class="badge" style="background:#F1F5F9;color:#64748B">Histórico</span>';
    tr.innerHTML = '<td>' + fmtMes(d.mes) + '</td>'
                 + '<td class="text-end fw-semibold">$' + d.costo.toFixed(2) + '</td>'
                 + '<td>' + fuente + '</td>';
    tbody.appendChild(tr);
  });

  document.getElementById('tabla-titulo').textContent = (p.nombre || p.email) + ' — Detalle mensual';
  document.getElementById('tabla-count').textContent  = datos.length + ' meses';
}
</script>
</body>
</html>
"""


def generar_html(personas):
    codes = sorted({
        d['code']
        for p in personas
        for d in p['datos']
        if d.get('code')
    })
    agreements = sorted({
        d['agreement']
        for p in personas
        for d in p['datos']
        if d.get('agreement')
    })
    divisions = sorted({
        p['division']
        for p in personas
        if p.get('division')
    })
    countries = sorted({
        p['country'] for p in personas if p.get('country')
    })
    levels = sorted({
        p['level'] for p in personas if p.get('level')
    })
    departments = sorted({
        p['department'] for p in personas if p.get('department')
    })
    payload = {
        'personas':    personas,
        'codes':       codes,
        'agreements':  agreements,
        'divisions':   divisions,
        'countries':   countries,
        'levels':      levels,
        'departments': departments,
    }
    print(f'  Codes: {codes}')
    print(f'  Agreements: {agreements}')
    data_json = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    html = HTML_TEMPLATE.replace('__DATA__', data_json)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'HTML generado: {OUTPUT_HTML}')


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    svc = _service()
    old_data    = read_old_sheet(svc)
    comp21_data = read_comp2021(svc)
    comp22_data = read_comp2022(svc)
    comp23_data = read_comp2023(svc)
    comp24_data = read_comp2024(svc)
    comp25_data = read_comp2025(svc)
    bdd_data, nombres, doppler_emails, divisiones, countries_map, seniorities_map, departments_map = read_bdd(svc)
    personas = combinar(old_data, comp21_data, comp22_data,
                        comp23_data, comp24_data, comp25_data, bdd_data,
                        nombres, doppler_emails, divisiones, countries_map,
                        seniorities_map, departments_map)
    generar_html(personas)


if __name__ == '__main__':
    main()
