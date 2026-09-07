# -*- coding: utf-8 -*-
"""
Descarga datos de BambooHR para un mes y actualiza los 4 Google Sheets directamente.

Uso:
    python actualizar_mes_sheets.py Abril 2026
    python actualizar_mes_sheets.py Abril 2026 --dry-run   (no escribe nada)
"""

import sys
import os
import json
import datetime
import requests

sys.stdout.reconfigure(encoding='utf-8')

# ─── IMPORTS DE FUNCIONES DE CÁLCULO ─────────────────────────────────────────
# Reutilizamos las funciones puras de generar_compensaciones.py
sys.path.insert(0, os.path.dirname(__file__))
from generar_compensaciones import (
    to_float,
    calc_seniority,
    calc_new_code,
    calc_costo_usd_h,
    calc_gap_banda,
    calc_antiguedad,
    calc_revisar_seniority,
    calc_pct_gap_banda,
    calc_ajuste_ultimo_ano,
    calc_dispersion_en_banda,
    get_employee_type,
    MONTH_ABBR_ES,
    MONTHLY_TAB_COLS,
)

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

WORK_DIR = r"C:\Users\aprato\documents\Proyectos\Compensaciones"
CREDS_FILE = os.path.join(WORK_DIR, 'talentserviceproject-1ce2ed91696b.json')
BAMBOO_CONFIG = os.path.join(WORK_DIR, 'bamboo_config.json')

SPREADSHEET_IDS = {
    'General':            '15mTvxuwXGa8B8Cuc86qqDyb8eKhKrDlwRekx9I1dZ_0',
    'Development':        '1hGPGhCbzNi8yCa6lqXhQtNXYAJTUGP70Txj04CaTgK8',
    'Quality Assurance':  '1M832_stiu3khX4qxIU4MmdS6rK5e10zvDpPL4BPQFcw',
    'Product Experience': '1dis277RAhHaXmkFPpRO-eUXczqZfUrCwkjzRDB1jguc',
}

# "Variables para Rep" y "Bandas Div" se mantienen y actualizan mensualmente en
# este sheet separado ("Bandas + variables 2026"), no en General.
VARIABLES_BANDAS_SHEET_ID = '1vBmF6PdU9ZCeB-6pqX5hxQ-FM0G3yWgd7MqtDfFlt78'

DEPT_TO_SHEET = {
    'MS - Development':        'Development',
    'MS - Quality Assurance':  'Quality Assurance',
    'MS - Product Experience': 'Product Experience',
}

DEPARTMENTS = list(DEPT_TO_SHEET.keys())

# Personas que se incluyen en el sheet de un depto aunque su Department en
# BambooHR sea otro (ej. Facundo Amores es MS - Lab pero se lo sigue
# trackeando en el tablero de Product Experience).
DEPT_SHEET_INCLUDE_OVERRIDES = {
    'Product Experience': ['famores@makingsense.com'],
}

# Personas con jornada parcial: las Bandas Div están calibradas para 8hs/día,
# así que hay que escalarlas por (horas reales / 8) antes de compararlas contra
# el Costo USD/H real. No hay ningún campo en BambooHR que marque esto — es
# conocimiento manual confirmado por el usuario (ver memoria
# project_bandas_tiempo_parcial). Actualizar esta lista si cambia la jornada
# de alguien o se detecta un caso nuevo.
PARTTIME_FACTOR_BY_EMAIL = {
    'jmayoral@makingsense.com': 6 / 8,
    'mbedjan@makingsense.com':  6 / 8,
    'tbasualdo@makingsense.com': 6 / 8,
}

# calc_seniority() extrae lo que sigue al primer espacio del Código (ej. "SEN 08" -> "08"),
# lo que da texto sin sentido para códigos sin ese formato (ej. "Executive assistant" ->
# "assistant"). Para estas personas el usuario confirmó (04/09/2026) el nivel real; se
# fija por email + Código esperado, así que si algún día tienen una recat (cambia el
# Código), el override deja de aplicar solo y vuelve a calc_seniority() normal.
SENIORITY_OVERRIDE_BY_EMAIL = {
    'vmunoz@makingsense.com':     ('Executive assistant', '08'),
    'nlenkovich@makingsense.com': ('FinSsrAdv', '05'),
    'fflorez@makingsense.com':    ('MKTHoD', '10'),
    'sgavilan@makingsense.com':   ('DevSr2', '08'),
}

MONTH_FULL_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo',  6: 'Junio',   7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
}


# ─── GOOGLE SHEETS SERVICE ───────────────────────────────────────────────────

def _sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        CREDS_FILE,
        scopes=['https://www.googleapis.com/auth/spreadsheets'],
    )
    return build('sheets', 'v4', credentials=creds)


# ─── BAMBOOHR DOWNLOAD ────────────────────────────────────────────────────────

def _bamboo_auth():
    with open(BAMBOO_CONFIG) as f:
        cfg = json.load(f)
    return (
        f"https://api.bamboohr.com/api/gateway.php/{cfg['subdomain']}/v1",
        (cfg['api_key'], 'x'),
    )


def _parse_date(val):
    if not val or val == '0000-00-00':
        return None
    # Strip time component if present (e.g. "2026-04-01 0:00:00" → "2026-04-01")
    val = str(val).strip().split(' ')[0]
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.datetime.strptime(val, fmt)
        except ValueError:
            pass
    return None


def _serial_to_date(val):
    """Converts a Google Sheets date serial number to datetime (epoch 1899-12-30)."""
    if isinstance(val, (int, float)):
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=val)
    return _parse_date(val)


def _month_range(month_date):
    """Returns (first_day, last_day) as datetime objects for the given month."""
    first = month_date.replace(day=1)
    if month_date.month == 12:
        last = datetime.datetime(month_date.year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        last = datetime.datetime(month_date.year, month_date.month + 1, 1) - datetime.timedelta(days=1)
    return first, last


def download_employees(month_date):
    """
    Returns employees active in the target month:
    - Currently active employees (no termination date)
    - Employees terminated during the target month
    Uses saved report 701 for compensation data + a lightweight call for termination dates.
    """
    base_url, auth = _bamboo_auth()

    # Get all employees with a Code (active + terminated) from saved report
    r = requests.get(
        f'{base_url}/reports/701',
        auth=auth,
        headers={'Accept': 'application/json'},
        params={'format': 'json'},  # no onlyCurrent → includes terminated
    )
    r.raise_for_status()
    all_employees = [e for e in r.json().get('employees', []) if e.get('customCode')]

    # Get termination dates via a lightweight custom report — keyed by employee id
    # NOTE: employeeNumber in report 701 is a numeric string for terminated employees,
    # so it doesn't match workEmail. Matching by id is reliable across all employees.
    r_term = requests.post(
        f'{base_url}/reports/custom',
        auth=auth,
        headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
        params={'format': 'json'},
        json={'title': 'term dates', 'fields': ['id', 'terminationDate']},
    )
    r_term.raise_for_status()
    term_by_id = {}
    for emp in r_term.json().get('employees', []):
        eid = emp.get('id')
        if eid:
            term_by_id[eid] = _parse_date(emp.get('terminationDate'))

    # Filter: active (no term date) OR terminated within the target month
    first_day, last_day = _month_range(month_date)
    filtered, bajas = [], 0
    for e in all_employees:
        eid = e.get('id')
        term_date = term_by_id.get(eid)
        if term_date is None:
            filtered.append(e)
        elif first_day <= term_date <= last_day:
            filtered.append(e)
            bajas += 1

    print(f'  BambooHR: {len(filtered)} empleados ({bajas} bajas en el mes).')
    return filtered


def build_job_status_lookup():
    """Returns dict email → sorted list of (date, bill, payroll) tuples.

    Stores ALL historical Cambio Proactivo entries so that:
    - calc_ultimo_cambio_proactivo can filter the date by month
    - process_month can use the historically correct Bill/PayRoll per month
    """
    base_url, auth = _bamboo_auth()
    r = requests.get(
        f'{base_url}/reports/740',
        auth=auth,
        headers={'Accept': 'application/json'},
        params={'format': 'json'},
    )
    r.raise_for_status()
    lookup = {}
    for rec in r.json().get('employees', []):
        email = rec.get('employeeNumber')
        if not email:
            continue
        d = _parse_date(rec.get('customEffectiveDate3'))
        if not d:
            continue
        bill    = rec.get('customBill')
        payroll = rec.get('customPayRoll')
        try: bill    = float(bill)    if bill    else None
        except: bill = None
        try: payroll = float(payroll) if payroll else None
        except: payroll = None
        if email not in lookup:
            lookup[email] = []
        # Avoid duplicates on same date
        if not any(e[0] == d for e in lookup[email]):
            lookup[email].append((d, bill, payroll))
    # Sort each list chronologically
    for email in lookup:
        lookup[email].sort(key=lambda x: x[0])
    total = sum(len(v) for v in lookup.values())
    print(f'  Job Status: {len(lookup)} empleados, {total} cambios proactivos.')
    return lookup


# ─── REFERENCIA: VARIABLES Y BANDAS DESDE SHEETS ─────────────────────────────

def get_variables_for_month(service, month_date):
    """Reads Variables para Rep from the Bandas+Variables sheet and returns vars dict."""
    sid = VARIABLES_BANDAS_SHEET_ID
    r = service.spreadsheets().values().get(
        spreadsheetId=sid,
        range="'Variables para Rep'!A1:AL50",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = r.get('values', [])
    for row in rows[3:]:   # data starts at row 4 (index 3)
        if not row:
            continue
        # La col A es un serial de fecha (formato de celda "M/D", sin año visible)
        d = _serial_to_date(row[0]) if row[0] != '' else None
        if d and d.year == month_date.year and d.month == month_date.month:
            def _f(idx):
                try:
                    return float(str(row[idx]).replace(',', '.'))
                except (IndexError, ValueError):
                    return None
            return {
                'FX_ARS':        _f(1),
                'FX_COP':        _f(3),
                'COSTO_ARG':     _f(4),
                'COSTO_COL_MAY': _f(5),
                'COSTO_COL_MEN': _f(6),
                'HORAS':         _f(7),
                'COSTO_ARG_USD': _f(8),
                'COSTO_USA':     _f(37),
            }
    return None


def build_bandas_lookup(service, vars_dict):
    """Calcula H/I/J (Banda Min/Med/Max USD/h) a partir de D/E/F y las
    variables del mes (FX, cargas sociales, horas).

    Tipos detectados por sufijo del New Code:
      -C   → Contractor: bandas ya en USD/h (H=D, I=E, J=F)
      -COL → Colombia:   H=D×COSTO_COL_MEN/FX_COP/HORAS
      -ARG → Argentina:
               Si D < 5000 (componente USD = Plus Fijo):
                 H=(E×COSTO_ARG/FX_ARS + D×COSTO_ARG_USD)/HORAS
                 J=(E×COSTO_ARG/FX_ARS + F×COSTO_ARG_USD)/HORAS
                 I=(H+J)/2
               Si D ≥ 5000 (ARS puro = Empleado):
                 H=D×COSTO_ARG/FX_ARS/HORAS
    """
    sid = VARIABLES_BANDAS_SHEET_ID

    r_vals = service.spreadsheets().values().get(
        spreadsheetId=sid,
        range="'Bandas Div'!A4:K500",
    ).execute()

    val_rows = r_vals.get('values', [])

    fx_ars        = vars_dict.get('FX_ARS',        1)
    fx_cop        = vars_dict.get('FX_COP',         1)
    costo_arg     = vars_dict.get('COSTO_ARG',      1)
    costo_col_men = vars_dict.get('COSTO_COL_MEN',  1)
    horas         = vars_dict.get('HORAS',         168)
    costo_arg_usd = vars_dict.get('COSTO_ARG_USD',  1)

    lookup = {}
    for row in val_rows:
        if len(row) < 2:
            continue
        new_code = str(row[1]).strip() if row[1] else ''
        if not new_code:
            continue

        def _fv(idx):
            try:
                return float(str(row[idx]).replace(',', '.'))
            except (IndexError, ValueError):
                return None

        d = _fv(3)
        e = _fv(4)
        f = _fv(5)

        if None in (d, e, f):
            # Fallback: intentar leer H/I/J pre-calculados
            h, m_val, j = _fv(7), _fv(8), _fv(9)
            if None in (h, m_val, j):
                continue
            lookup[new_code] = (h, m_val, j)
            continue

        h = m_val = j = None

        if new_code.endswith('-C'):
            # Contractor: bandas ya expresadas en USD/h
            h, m_val, j = d, e, f

        elif new_code.endswith('-COL'):
            # Colombia (COP → USD/h)
            if fx_cop and horas:
                h     = d * costo_col_men / fx_cop / horas
                m_val = e * costo_col_men / fx_cop / horas
                j     = f * costo_col_men / fx_cop / horas

        elif new_code.endswith('-ARG'):
            if fx_ars and horas:
                if d < 5000:
                    # Plus Fijo: D/F son componente USD, E es nómina ARS
                    h     = (e * costo_arg / fx_ars + d * costo_arg_usd) / horas
                    j     = (e * costo_arg / fx_ars + f * costo_arg_usd) / horas
                    m_val = (h + j) / 2
                else:
                    # Empleado puro: D/E/F en ARS
                    h     = d * costo_arg / fx_ars / horas
                    m_val = e * costo_arg / fx_ars / horas
                    j     = f * costo_arg / fx_ars / horas

        if None in (h, m_val, j):
            continue

        lookup[new_code] = (round(h, 4), round(m_val, 4), round(j, 4))

    print(f'  Bandas Div: {len(lookup)} entradas.')
    return lookup


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

def _fmt_date(val, fmt='%m/%d/%Y'):
    if isinstance(val, datetime.datetime):
        return val.strftime(fmt)
    return ''


def process_month(employees, job_lookup, vars_dict, bandas_lookup, month_date, today):
    """Processes all employees and returns list of output rows (as lists for Sheets)."""
    rows = []
    for emp in employees:
        last_first   = emp.get('fullName2', '')
        email        = emp.get('employeeNumber')
        hire_date    = _parse_date(emp.get('hireDate'))
        department   = emp.get('department')
        location     = emp.get('location')
        xm           = emp.get('customxM') or ''
        eff_date1    = _parse_date(emp.get('customEffectiveDate'))
        level_code   = emp.get('customCode') or ''
        agreement    = emp.get('customAgreement') or ''
        bill         = emp.get('customBill')
        payroll      = emp.get('customPayRoll')
        # Nota: bill y payroll se sobreescriben más abajo con el valor histórico correcto
        per          = emp.get('customPer') or ''
        pay_currency = emp.get('customPayRate-CurrencyCode') or ''
        eff_date2    = _parse_date(emp.get('customEffectiveDate1'))
        client       = emp.get('customCliente') or ''
        project      = emp.get('customProyecto') or ''
        raw_pct      = emp.get('custom%Asignación')
        pct_assign   = (float(raw_pct) / 100) if raw_pct else None
        term_reason  = emp.get('customRazónSalida') or ''
        assigned_end = _parse_date(emp.get('customFechaFinalizaciónAsignación'))
        vac_raw      = emp.get('4471.2', '')
        pto_raw      = emp.get('4940.2', '')
        vac_policy   = vac_raw if vac_raw and vac_raw != 'None' else ''
        pto_policy   = pto_raw if pto_raw and pto_raw != 'None' else ''

        seniority  = calc_seniority(level_code)
        _sen_override = SENIORITY_OVERRIDE_BY_EMAIL.get(email)
        if _sen_override and level_code == _sen_override[0]:
            seniority = _sen_override[1]
        costo_usdh = calc_costo_usd_h(agreement, location, bill, payroll, per, vars_dict)
        new_code   = calc_new_code(level_code, agreement, location)

        banda_vals = bandas_lookup.get(new_code)
        banda_min  = banda_vals[0] if banda_vals else None
        banda_med  = banda_vals[1] if banda_vals else None
        banda_max  = banda_vals[2] if banda_vals else None

        parttime_factor = PARTTIME_FACTOR_BY_EMAIL.get(email)
        if parttime_factor:
            banda_min = banda_min * parttime_factor if banda_min is not None else None
            banda_med = banda_med * parttime_factor if banda_med is not None else None
            banda_max = banda_max * parttime_factor if banda_max is not None else None

        gap_banda  = calc_gap_banda(costo_usdh, banda_min, banda_max)
        antiguedad = calc_antiguedad(hire_date, today)

        # Último Cambio Proactivo + Bill/PayRoll histórico correcto
        import calendar as _cal
        _last_day = _cal.monthrange(month_date.year, month_date.month)[1]
        _month_end = datetime.datetime(month_date.year, month_date.month, _last_day, 23, 59, 59)
        _all_changes = job_lookup.get(email, [])  # list of (date, bill, payroll)
        _valid = [entry for entry in _all_changes if entry[0] <= _month_end]
        if _valid:
            _latest = max(_valid, key=lambda x: x[0])
            ult_cambio = MONTH_ABBR_ES[_latest[0].month] + '-' + str(_latest[0].year)
            # NO pisar bill/payroll con report 740: ese reporte guarda el valor
            # ANTERIOR al cambio, no el nuevo. Usar siempre el valor actual de
            # BambooHR (report 701) que ya está en bill/payroll al entrar aquí.
        else:
            ult_cambio = ''

        rev_sen    = calc_revisar_seniority(level_code, eff_date1, today)
        pct_gap    = calc_pct_gap_banda(gap_banda)
        last_change = _latest[0] if _valid else None  # extraer solo la fecha del tuple
        ajuste     = calc_ajuste_ultimo_ano(month_date, hire_date, last_change)
        dispersion = calc_dispersion_en_banda(costo_usdh, banda_min, banda_max)

        # Format values for Sheets (USER_ENTERED)
        def _n(v):
            return round(v, 6) if isinstance(v, float) else (v if v is not None else '')

        row = [
            month_date.strftime('%Y-%m-%d'),       # Month
            last_first,                             # Last name, First name
            email,                                  # Employee #
            _fmt_date(hire_date),                   # Hire Date
            department or '',                       # Department
            location or '',                         # Location
            xm,                                     # xM
            _fmt_date(eff_date1),                   # Effective Date
            level_code,                             # Code (Level)
            seniority,                              # Seniority
            agreement,                              # Agreement
            _n(to_float(bill)),                     # Bill
            _n(to_float(payroll)),                  # PayRoll
            per,                                    # Per
            pay_currency,                           # Pay Rate - Currency Code
            _fmt_date(eff_date2),                   # Effective Date (2)
            client,                                 # Client
            project,                                # Project
            _n(pct_assign),                         # % Assignment
            term_reason,                            # Termination Reason
            _fmt_date(assigned_end),                # Assigned End Date
            vac_policy,                             # Vacations Policy
            pto_policy,                             # PTO Policy
            _n(costo_usdh),                         # Costo USD/H
            new_code,                               # New Code
            _n(banda_min),                          # Banda Min
            _n(banda_med),                          # Banda Med
            _n(banda_max),                          # Banda Max
            gap_banda if gap_banda != '' else '',   # GAP BANDA
            _n(antiguedad),                         # Antigüedad
            ("'" + ult_cambio if ult_cambio else ult_cambio),  # Ultimo Cambio Proactivo (forzado a texto: evita que Sheets lo auto-parsee como fecha)
            rev_sen,                                # Revisar seniority
            pct_gap,                                # % GAP Banda
            ajuste,                                 # Ajuste ultimo año
            _n(dispersion) if dispersion != '' else '',  # Dispersión en Banda
        ]
        rows.append(row)
    return rows


def tag_movimientos(new_rows, prev_emails):
    """Adds Movimiento tag (Alta/Baja/Activo) to each row."""
    new_emails = {r[2] for r in new_rows}
    result = []
    for row in new_rows:
        email = row[2]
        if not prev_emails:
            mov = 'Activo'
        elif email not in prev_emails:
            mov = 'Alta'
        else:
            mov = 'Activo'
    # Mark Bajas: emails in prev not in new — these don't appear as rows in current month,
    # so we only tag Activo/Alta here; Bajas are inferred from absence.
        result.append(row + [mov])
    return result


# ─── LECTURA DE EMAILS DEL MES ANTERIOR ──────────────────────────────────────

def get_prev_month_emails(service, month_date):
    """Returns set of employee emails present in the previous month in General sheet."""
    prev = month_date.replace(day=1) - datetime.timedelta(days=1)
    prev_str = prev.replace(day=1).strftime('%Y-%m-%d')

    sid = SPREADSHEET_IDS['General']
    r = service.spreadsheets().values().get(
        spreadsheetId=sid,
        range="'Bamboo Compensaciones - Sueldos'!A:C",
    ).execute()
    rows = r.get('values', [])
    emails = set()
    for row in rows[1:]:
        if len(row) >= 3 and row[0] == prev_str:
            emails.add(row[2])
    return emails


# ─── VERIFICAR QUE EL MES NO EXISTE YA ───────────────────────────────────────

def month_already_exists(service, spreadsheet_id, month_str):
    r = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Bamboo Compensaciones - Sueldos'!A:A",
    ).execute()
    return any(row and row[0] == month_str for row in r.get('values', []))


# ─── ESCRITURA A SHEETS ───────────────────────────────────────────────────────

def append_rows(service, spreadsheet_id, rows, dry_run=False):
    if dry_run:
        print(f'    [DRY-RUN] Se agregarían {len(rows)} filas.')
        return
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range="'Bamboo Compensaciones - Sueldos'!A1",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': rows},
    ).execute()
    print(f'    {len(rows)} filas agregadas.')


def fix_seniority_overrides(service, spreadsheet_id, dry_run=False):
    """SENIORITY_OVERRIDE_BY_EMAIL escribe texto (ej. '08'), pero append_rows() usa
    valueInputOption=USER_ENTERED para toda la fila (lo necesitan las fechas, que deben
    auto-parsearse a Date) — Sheets interpreta un string puramente numérico como si el
    usuario lo hubiera tipeado, y lo convierte a número, perdiendo el cero a la izquierda.
    Corre después de escribir el mes: relee la columna Seniority de TODAS las filas de las
    personas con override cuyo Código coincida con el esperado, y las re-escribe con
    valueInputOption=RAW (no parsea, queda texto literal) si no están ya correctas. Barre
    todo el histórico (no solo el mes nuevo) para que también quede prolijo el texto viejo
    que haya calculado mal calc_seniority() antes de que existiera el override."""
    if not SENIORITY_OVERRIDE_BY_EMAIL:
        return
    resp = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'Bamboo Compensaciones - Sueldos'!A1:AJ20000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])
    if not rows:
        return
    headers = rows[0]
    i_email = headers.index('Employee #')
    i_code  = headers.index('Code (Level)')
    i_sen   = headers.index('Seniority')

    updates = []
    for idx, row in enumerate(rows[1:], start=2):
        email = str(row[i_email]).lower().strip() if i_email < len(row) else ''
        override = SENIORITY_OVERRIDE_BY_EMAIL.get(email)
        if not override:
            continue
        expected_code, fixed_level = override
        code = row[i_code] if i_code < len(row) else ''
        if code != expected_code:
            continue
        current = row[i_sen] if i_sen < len(row) else ''
        if current == fixed_level and isinstance(current, str):
            continue  # ya está correcto (texto, mismo valor)
        updates.append({'range': f"'Bamboo Compensaciones - Sueldos'!J{idx}", 'values': [[fixed_level]]})

    if not updates:
        print('    Seniority overrides: sin cambios (ya estaba todo correcto).')
        return
    print(f'    Seniority overrides: corrigiendo {len(updates)} celdas...')
    if not dry_run:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'valueInputOption': 'RAW', 'data': updates},
        ).execute()


def add_pivot(service, spreadsheet_id, month_name, month_date_str, dry_run=False):
    """Creates a pivot table sourcing from Bamboo Compensaciones - Sueldos, filtered by month."""
    if dry_run:
        print(f'    [DRY-RUN] Se crearía Pivot {month_name}.')
        return

    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {s['properties']['title']: s['properties']['sheetId']
                for s in spreadsheet['sheets']}

    pivot_name = f'Pivot {month_name}'
    source_sheet_id = existing.get('Bamboo Compensaciones - Sueldos')
    if source_sheet_id is None:
        print(f'    ERROR: No se encontró "Bamboo Compensaciones - Sueldos".')
        return

    if pivot_name in existing:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [{'deleteSheet': {'sheetId': existing[pivot_name]}}]},
        ).execute()

    pivot_result = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{'addSheet': {'properties': {'title': pivot_name}}}]},
    ).execute()
    pivot_sheet_id = pivot_result['replies'][0]['addSheet']['properties']['sheetId']

    # Column offsets in Bamboo Compensaciones - Sueldos:
    # 0=Fecha, 1=Last name, 3=Hire Date, 6=xM, 10=Agreement, 11=Bill, 12=PayRoll,
    # 23=Costo USD/H, 24=New Code, 25=Banda Min, 26=Banda Med, 27=Banda Max,
    # 28=GAP BANDA, 30=Ultimo Cambio Proactivo
    pivot_table = {
        'source': {
            'sheetId': source_sheet_id,
            'startRowIndex': 0, 'startColumnIndex': 0,
            'endRowIndex': 5000, 'endColumnIndex': 36,
        },
        'filterSpecs': [{
            'columnOffsetIndex': 0,
            'filterCriteria': {
                # visibleValues matches the formatted display value of the cell,
                # which works for date-type cells (stored as serial numbers) formatted as YYYY-MM-DD.
                'visibleValues': [month_date_str],
            }
        }],
        'rows': [
            {'sourceColumnOffset': 6,  'showTotals': True,  'sortOrder': 'ASCENDING'},  # xM
            {'sourceColumnOffset': 10, 'showTotals': False, 'sortOrder': 'ASCENDING'},  # Agreement
            {'sourceColumnOffset': 24, 'showTotals': False, 'sortOrder': 'ASCENDING'},  # New Code
            {'sourceColumnOffset': 1,  'showTotals': False, 'sortOrder': 'ASCENDING'},  # Last name
            {'sourceColumnOffset': 3,  'showTotals': False, 'sortOrder': 'ASCENDING'},  # Hire Date
            {'sourceColumnOffset': 28, 'showTotals': False, 'sortOrder': 'ASCENDING'},  # GAP BANDA
            {'sourceColumnOffset': 30, 'showTotals': False, 'sortOrder': 'ASCENDING'},  # Ultimo Cambio
        ],
        'values': [
            {'summarizeFunction': 'SUM',     'sourceColumnOffset': 11, 'name': 'Bill'},
            {'summarizeFunction': 'SUM',     'sourceColumnOffset': 12, 'name': 'PayRoll'},
            {'summarizeFunction': 'AVERAGE', 'sourceColumnOffset': 23, 'name': 'Costo USD/H'},
            {'summarizeFunction': 'AVERAGE', 'sourceColumnOffset': 25, 'name': 'Banda Min'},
            {'summarizeFunction': 'AVERAGE', 'sourceColumnOffset': 26, 'name': 'Banda Med'},
            {'summarizeFunction': 'AVERAGE', 'sourceColumnOffset': 27, 'name': 'Banda Max'},
        ],
        'valueLayout': 'HORIZONTAL',
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [{
            'updateCells': {
                'rows': [{'values': [{'pivotTable': pivot_table}]}],
                'start': {'sheetId': pivot_sheet_id, 'rowIndex': 0, 'columnIndex': 0},
                'fields': 'pivotTable',
            }
        }]},
    ).execute()
    print(f'    Pivot table "{pivot_name}" creada.')


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    args     = [a for a in sys.argv[1:] if not a.startswith('--')]
    dry_run  = '--dry-run' in sys.argv

    if len(args) < 2:
        print('Uso: python actualizar_mes_sheets.py <Mes> <Año> [--dry-run]')
        print('Ej:  python actualizar_mes_sheets.py Abril 2026')
        sys.exit(1)

    month_name = args[0]
    year       = int(args[1])

    # Map month name to number
    month_num = next(
        (k for k, v in MONTH_FULL_ES.items() if v.lower() == month_name.lower()), None
    )
    if not month_num:
        print(f'Mes no reconocido: {month_name}')
        sys.exit(1)

    month_date = datetime.datetime(year, month_num, 1)
    month_str  = month_date.strftime('%Y-%m-%d')
    today      = datetime.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

    print(f'\n{"[DRY-RUN] " if dry_run else ""}Actualizando sheets para {month_name} {year}...\n')

    # ── 1. Google Sheets service ──────────────────────────────────────────────
    service = _sheets_service()

    # ── 2. Verificar que el mes no exista ya en General ───────────────────────
    if not dry_run and month_already_exists(service, SPREADSHEET_IDS['General'], month_str):
        print(f'El mes {month_name} {year} ya existe en el sheet General. Abortando.')
        print('Usá --dry-run para verificar sin escribir.')
        sys.exit(1)

    # ── 3. Variables para Rep ─────────────────────────────────────────────────
    print('Leyendo Variables para Rep...')
    vars_dict = get_variables_for_month(service, month_date)
    if not vars_dict:
        print(f'No se encontraron variables para {month_name} {year} en el sheet.')
        print('Asegurate de agregar la fila correspondiente en la tab "Variables para Rep".')
        sys.exit(1)
    print(f'  FX_ARS={vars_dict["FX_ARS"]}, FX_COP={vars_dict["FX_COP"]}, HORAS={vars_dict["HORAS"]}')

    # ── 4. Bandas Div ─────────────────────────────────────────────────────────
    print('Leyendo Bandas Div...')
    bandas_lookup = build_bandas_lookup(service, vars_dict)

    # ── 5. BambooHR employees ─────────────────────────────────────────────────
    print('Descargando empleados de BambooHR...')
    employees = download_employees(month_date)

    # ── 6. Job Status (Cambio Proactivo) ──────────────────────────────────────
    print('Descargando Job Status...')
    job_lookup = build_job_status_lookup()

    # ── 7. Procesar datos ─────────────────────────────────────────────────────
    print('Procesando columnas calculadas...')
    all_rows = process_month(employees, job_lookup, vars_dict, bandas_lookup, month_date, today)
    print(f'  {len(all_rows)} filas procesadas.')

    # ── 8. Movimiento (Alta/Baja) ─────────────────────────────────────────────
    print('Calculando Movimiento...')
    prev_emails = get_prev_month_emails(service, month_date)
    print(f'  Mes anterior: {len(prev_emails)} empleados.')
    all_rows_with_mov = tag_movimientos(all_rows, prev_emails)

    # ── 9. Escribir en General (todas las filas + Movimiento) ─────────────────
    print('\nActualizando General...')
    append_rows(service, SPREADSHEET_IDS['General'], all_rows_with_mov, dry_run)
    fix_seniority_overrides(service, SPREADSHEET_IDS['General'], dry_run)

    # ── 10. Escribir en sheets de departamento ────────────────────────────────
    for dept, sheet_name in DEPT_TO_SHEET.items():
        sid = SPREADSHEET_IDS[sheet_name]
        print(f'\nActualizando {sheet_name}...')

        # Filter: only this dept + exclude seniority >= 10
        def _excl(row):
            lc = row[8]   # Code (Level) = index 8 in our out_row
            s = str(lc or '')
            idx = s.find(' ')
            num = s[idx + 1:] if idx >= 0 else s
            try:
                return int(num) >= 10
            except (ValueError, TypeError):
                return False

        include_emails = DEPT_SHEET_INCLUDE_OVERRIDES.get(sheet_name, [])
        dept_rows = [
            r for r in all_rows
            if (r[4] == dept or r[2] in include_emails) and not _excl(r)
        ]
        print(f'  {len(dept_rows)} filas (sin seniority 10+).')

        # Append to Bamboo Compensaciones - Sueldos (without Movimiento col)
        if not dry_run and month_already_exists(service, sid, month_str):
            print(f'  Ya existe en {sheet_name}. Saltando.')
        else:
            append_rows(service, sid, dept_rows, dry_run)

        # Add pivot table (sourcing directly from Bamboo Compensaciones - Sueldos)
        add_pivot(service, sid, month_name, month_str, dry_run)

    print(f'\n{"[DRY-RUN] " if dry_run else ""}Listo! {month_name} {year} actualizado en todos los sheets.')


if __name__ == '__main__':
    main()
