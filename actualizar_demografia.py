# -*- coding: utf-8 -*-
"""
Actualización diaria de la Base de Datos Bamboo.

Proceso:
  1. Elimina columna AK (Movimiento) si aún existe.
  2. Actualiza Histórico:
     - Filas del mes actual ya existen → actualiza cols diarias (B, D–W, Y, AK–AZ).
     - Sin filas del mes actual + hoy es día hábil → crea filas nuevas para
       personas activas en BambooHR (1er día hábil del mes).
  3. Si NO se crearon filas nuevas: agrega activos nuevos en BambooHR que no figuren.
  4. Reconstruye Actualizado: fila más reciente por persona (activos + inactivos),
     con cols diarias frescas desde BambooHR.

Columnas diarias: B(1), D–W(3–22) [R/Project: Talent Discussion > Bamboo], Y(24), AK–AZ(36–51).
Columnas no modificadas: X(23), Z–AC(25–28), AD–AJ(29–35), BA–BC(52–54).

Uso:
    python actualizar_demografia.py [--dry-run]
"""

import sys, os, json, requests, datetime
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

CREDS_FILE    = r"C:\Users\aprato\documents\Proyectos\Compensaciones\talentserviceproject-1ce2ed91696b.json"
BAMBOO_CONFIG = r"C:\Users\aprato\documents\Proyectos\Compensaciones\bamboo_config.json"

SHEET_BDD   = '11kTmD6YBpquXMlqI1QoY4xU8bmjKY0Rh30qX55rebsQ'
TAB_HIST    = 'Histórico'
TAB_ACT     = 'Actualizado'
DEMO_SSID   = '1Cm-bSBOiwmSeWpI9CwGy4-UMrI9k7ZgUuPBaAWDJaco'
DEMO_TAB    = 'demografia - claude'
EXCLUDE_DIV = {'Externos'}

EMAIL_COL  = 2
MONTH_COL  = 0
TOTAL_COLS = 54

# Emails excluidos permanentemente (cuentas de prueba, etc.)
EXCLUDE_EMAIL = {'pruebadomicilio@yopmail.com'}

# Typos conocidos en el email del empleado (col C) → email correcto.
# Se aplican al leer Histórico para evitar filas duplicadas en Actualizado.
EMAIL_CORRECTIONS = {
    'mmaza@makignsense.com':        'mmaza@makingsense.com',
    'kromeroquijano@makingsense.com': 'kromero@makingsense.com',
}

# Overrides de seniority: email → valor fijo (ignora calc_seniority)
SENIORITY_OVERRIDES = {
    'lrey@makingsense.com':     '10',
    'sgavilan@makingsense.com': '08',
}

# Overrides de Termination Type: emails que Bamboo tiene mal clasificados
# (comentario "Mutuo acuerdo" pero Type = Voluntary) y no se puede corregir en
# Bamboo. Se fuerza el valor acá, ignorando lo que traiga la API. Ver
# termination_type_overrides.json y memoria [[project-bdd-employment-status-fix]].
def _load_termination_type_overrides():
    path = os.path.join(os.path.dirname(__file__), 'termination_type_overrides.json')
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        emails = set(e.lower().strip() for e in data.get('emails', []))
        return emails, data.get('termination_type', 'Termination (Involuntary)')
    except FileNotFoundError:
        return set(), ''
    except Exception as e:
        print(f'Advertencia: no se pudo leer termination_type_overrides.json: {e}')
        return set(), ''


TERMINATION_TYPE_OVERRIDE_EMAILS, TERMINATION_TYPE_OVERRIDE_VALUE = _load_termination_type_overrides()

# Seniority para Doppler: calc_seniority() asume códigos "SEN 04" (con espacio) y
# devuelve todo lo que sigue al espacio; los códigos Doppler no tienen espacio
# ("DevSSr", "TASr1") así que devuelve el código entero — de ahí que Seniority
# quedara igual al Code. Se usa en cambio el mapeo code→seniority de la solapa
# "Seniority Doppler" (fuente de verdad mantenida a mano por el usuario), leída
# en vivo una vez por corrida vía _load_doppler_seniority_map().
TAB_SENIORITY_DOPPLER = 'Seniority Doppler'
DOPPLER_SENIORITY_MAP = {}

# Mismo problema que Doppler, pero para MS: algunos Code de BambooHR no tienen
# espacio ("DevSr1", "QCSSrAdv", "PdMSr1", ...) así que calc_seniority() también
# devuelve el código entero en vez del nivel numérico. Mapa derivado de los
# valores limpios ya existentes en Histórico (ver conversación 13/08/2026);
# 3 conflictos resueltos a mano (DevSr2, PMSr3, PdMSr1). No incluye "HoD" ni
# "Maestranza": para esos el código y el nivel son legítimamente el mismo valor.
MS_LETTER_CODE_SENIORITY = {
    'BASr1': '06', 'BIManager': '10', 'COO': 'C-level', 'CPO': 'C-level',
    'CTO': 'C-level', 'DESr1': '06', 'DevArc1': '12', 'DevJr': '02',
    'DevJrAdv': '03', 'DevOpsSr1': '06', 'DevOpsSsr': '04', 'DevSSr': '04',
    'DevSSrAdv': '05', 'DevSpSSr': '04', 'DevSpSSrAdv': '05', 'DevSpSr1': '06',
    'DevSr1': '06', 'DevSr2': '07', 'DevTecExpSr': '8', 'DevTr': '01',
    'FESr1': '06', 'FESsr': '04', 'FinJr': '02', 'FinSr': '06',
    'FinSsr': '04', 'FinSsrAdv': '05', 'HRJrAdv': '03', 'HrAdmJr': '02',
    'HrAdmSsrAdv': '05', 'HrAdminSp': '06', 'HrL&ODLead': '10', 'HrL&ODSr': '08',
    'HrPDSsr': '04', 'ITJR': '02', 'MKTHoD': 'HoD', 'MktJr': '02',
    'MktManager': '10', 'MktSr': '06', 'MktSr1': '06', 'MktSsr': '04',
    'PASr1': '06', 'PASr2': '07', 'PASr3': '08', 'PASsr': '04',
    'PASsrAdv': '05', 'PML1': '1', 'PML2': '2', 'PML3': '3',
    'PMSSrAdv': '05', 'PMSr1': '6', 'PMSr3': '08', 'PdMSr1': '06',
    'QCJr': '02', 'QCSSrAdv': '05', 'QCSr1': '06', 'QCSr2': '07',
    'QCSsr': '04', 'RecJrAdv': '03', 'RecPas': 'Pas', 'RecSr1': '06',
    'RecSsr': '04', 'Sales': '10', 'SalesManager': '10', 'SalesSr1': '06',
    'SalesTr': '01', 'TASr1': '06', 'TASr2': '07', 'TASsr': '04',
    'TASsrAdv': '05', 'UXDevJr': '02', 'UXResearcherSr': '06', 'UxDesJr': '02',
    'UxDesJrAdv': '03', 'UxDesSSr': '04', 'UxDesSSrAdv': '05', 'UxDesSr1': '06',
    'UxDesSr2': '07', 'UxDevJrAdv': '03', 'UxDevSSr': '04', 'UxDevSSrAdv': '05',
}


def _load_doppler_seniority_map(svc):
    global DOPPLER_SENIORITY_MAP
    try:
        r = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_BDD, range=f"'{TAB_SENIORITY_DOPPLER}'!A:Z",
            valueRenderOption='FORMATTED_VALUE').execute()
        rows = r.get('values', [])[1:]
        mapping = {}
        for row in rows:
            code   = str(row[8]).strip() if len(row) > 8 else ''
            senior = str(row[9]).strip() if len(row) > 9 else ''
            if code and senior:
                mapping[code] = senior
        DOPPLER_SENIORITY_MAP = mapping
        print(f'  Mapa Seniority Doppler: {len(mapping)} códigos.')
    except Exception as e:
        print(f'Advertencia: no se pudo leer "{TAB_SENIORITY_DOPPLER}": {e}')
        DOPPLER_SENIORITY_MAP = {}

# Project (col R): la fuente principal es la col A ("Account / Proyecto") de la
# solapa "Base Consolidada" de Talent Discussion, cruzando por Mail (col C).
# Si la persona no figura ahí, queda el customProyecto de BambooHR.
TD_SSID     = '1goxKg8LylMSZcVzkPVBPczknKlFbB8WBYge_XpAQxTw'
TD_TAB      = 'Base Consolidada'
TD_PROJECT_MAP = {}


def _load_td_project_map(svc):
    global TD_PROJECT_MAP
    try:
        r = svc.spreadsheets().values().get(
            spreadsheetId=TD_SSID, range=f"'{TD_TAB}'!A:C",
            valueRenderOption='FORMATTED_VALUE').execute()
        rows = r.get('values', [])[1:]
        mapping = {}
        for row in rows:
            project = str(row[0]).strip() if len(row) > 0 else ''
            email   = str(row[2]).strip().lower() if len(row) > 2 else ''
            if email and project:
                mapping[email] = project
        TD_PROJECT_MAP = mapping
        print(f'  Mapa Project Talent Discussion: {len(mapping)} emails.')
    except Exception as e:
        print(f'Advertencia: no se pudo leer "{TD_TAB}" (Talent Discussion): {e}')
        TD_PROJECT_MAP = {}

# Formatos de celda de la BDD (índices 0-based)
_DATE_COLS    = {3, 7, 15, 20, 36, 37, 45, 49}
_NUMBER_COLS  = {11, 12, 23, 25, 26, 27, 29, 34, 41}
_PERCENT_COLS = {18}

# Bloque demográfico diario: AJ(35)–AY(50) — tras eliminar col AJ(Estado)
# (índice en demo sheet, índice en BDD) — ambos 0-based
DEMO_MAP_DAILY = [
    (2,  35), (4,  36), (5,  37), (6,  38), (7,  39),
    (8,  40), (9,  41), (10, 42), (11, 43), (12, 44),
    (13, 45), (16, 46), (18, 47), (19, 48),
    (22, 49), (23, 50),
]
# BA(51)–BC(53): Termination Type, Eligible For Re-hire, Regrettable — se
# completan en _apply_daily_cols desde BambooHR en vivo, no desde la demo.
_DEMO_MAX_DAILY = max(src for src, _ in DEMO_MAP_DAILY) + 1  # = 24

# Importaciones opcionales desde generar_compensaciones
try:
    from generar_compensaciones import calc_new_code as _ext_new_code
except Exception:
    _ext_new_code = None

try:
    from generar_compensaciones import calc_seniority as _ext_seniority
except Exception:
    _ext_seniority = None


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _today_snapshot():
    return datetime.date.today().strftime('%Y-%m-01')


def _is_weekday():
    return datetime.date.today().weekday() < 5


def _service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        CREDS_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return build('sheets', 'v4', credentials=creds)


def _bamboo():
    with open(BAMBOO_CONFIG) as f:
        cfg = json.load(f)
    base = f"https://api.bamboohr.com/api/gateway.php/{cfg['subdomain']}/v1"
    return base, (cfg['api_key'], 'x')


def _parse_date(val):
    if not val or str(val).strip() in ('', '0000-00-00'):
        return None
    s = str(val).strip().split(' ')[0]
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


_SERIAL_EPOCH = datetime.date(1899, 12, 30)

def _fmt_d(d):
    if not d:
        return ''
    dd = d.date() if isinstance(d, datetime.datetime) else d
    return (dd - _SERIAL_EPOCH).days


def _num(v):
    try:
        return float(str(v).replace(',', '').strip())
    except Exception:
        return ''


def _is_cosmolex(emp):
    client  = str(emp.get('customCliente') or '').strip().lower()
    project = str(emp.get('customProyecto') or '').strip().lower()
    return 'cosmolex' in client or 'cosmolex' in project


def _calc_new_code(code, agreement, location):
    if _ext_new_code:
        try:
            return _ext_new_code(code, agreement, location) or ''
        except Exception:
            pass
    return ''


def _calc_seniority(code):
    if _ext_seniority:
        try:
            return _ext_seniority(code) or ''
        except Exception:
            pass
    return ''


def _apply_demo_daily(row, demo_row):
    """Aplica entradas DEMO_MAP_DAILY a row (AK–AZ, índices 36–51)."""
    padded = list(demo_row)
    while len(padded) < _DEMO_MAX_DAILY:
        padded.append('')
    for src_i, dst_i in DEMO_MAP_DAILY:
        row[dst_i] = padded[src_i]


def _apply_sheet_formats(svc, sheet_id, max_rows=30000):
    """Aplica formatos de fecha, número y porcentaje a las columnas de la BDD."""
    date_fmt = {'type': 'DATE',    'pattern': 'MM/DD/YYYY'}
    num_fmt  = {'type': 'NUMBER',  'pattern': '0.00'}
    pct_fmt  = {'type': 'PERCENT', 'pattern': '0.00%'}

    def _req(col, fmt):
        return {'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 1,
                'endRowIndex': max_rows,
                'startColumnIndex': col,
                'endColumnIndex': col + 1,
            },
            'cell': {'userEnteredFormat': {'numberFormat': fmt}},
            'fields': 'userEnteredFormat.numberFormat',
        }}

    reqs = ([_req(c, date_fmt) for c in _DATE_COLS]
            + [_req(c, num_fmt)  for c in _NUMBER_COLS]
            + [_req(c, pct_fmt)  for c in _PERCENT_COLS])
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_BDD, body={'requests': reqs}).execute()


def _apply_daily_cols(row, emp, demo_row):
    """
    Aplica columnas diarias a row desde BambooHR + demografía.
    Escribe: B(1), D–W(3–22), Y(24), AK–AZ(36–51).
    NO toca: A(0/Month), C(2/email), X(23), Z–AJ(25–35), BA–BC(52–54).
    """
    hire_date = _parse_date(emp.get('hireDate'))
    term_date = _parse_date(emp.get('terminationDate'))
    eff_date1 = _parse_date(emp.get('customEffectiveDate'))
    eff_date2 = _parse_date(emp.get('customEffectiveDate1'))
    asig_end  = _parse_date(emp.get('customFechaFinalizaciónAsignación'))
    raw_pct   = emp.get('custom%Asignación')
    pct       = round(float(raw_pct) / 100, 6) if raw_pct else ''
    vac       = emp.get('4471.2') or ''
    pto       = emp.get('4940.2') or ''
    code      = emp.get('customCode') or ''

    email = str(emp.get('workEmail') or '').strip().lower()

    row[1]  = f"{emp.get('lastName', '')}, {emp.get('firstName', '')}".strip(', ')
    row[3]  = _fmt_d(hire_date)
    row[4]  = emp.get('department') or ''
    row[5]  = emp.get('location') or ''
    row[6]  = emp.get('customxM') or ''
    row[7]  = _fmt_d(eff_date1)
    row[8]  = code
    division_raw = str(emp.get('division') or '').strip().lower()
    if 'doppler' in division_raw and code in DOPPLER_SENIORITY_MAP:
        senior_val = DOPPLER_SENIORITY_MAP[code]
    elif email in SENIORITY_OVERRIDES:
        senior_val = SENIORITY_OVERRIDES[email]
    elif code in MS_LETTER_CODE_SENIORITY:
        senior_val = MS_LETTER_CODE_SENIORITY[code]
    else:
        senior_val = _calc_seniority(code)
    # Apóstrofe fuerza texto bajo USER_ENTERED — sin esto Sheets interpreta
    # valores como "04" como número y se come el cero adelante.
    row[9] = f"'{senior_val}" if senior_val else ''
    row[10] = emp.get('customAgreement') or ''
    row[11] = _num(emp.get('customBill'))
    row[12] = _num(emp.get('customPayRoll'))
    row[13] = emp.get('customPer') or ''
    row[14] = emp.get('customPayRate-CurrencyCode') or ''
    row[15] = _fmt_d(eff_date2)
    row[16] = emp.get('customCliente') or ''
    row[17] = TD_PROJECT_MAP.get(email) or emp.get('customProyecto') or ''
    row[18] = pct
    row[19] = emp.get('4314') or ''  # Termination Reason (customRazónSalida está siempre vacío en Bamboo)
    row[20] = _fmt_d(asig_end)
    row[21] = vac if vac and vac != 'None' else ''
    row[22] = pto if pto and pto != 'None' else ''
    row[24] = _calc_new_code(code, emp.get('customAgreement') or '', emp.get('location') or '')

    if demo_row:
        _apply_demo_daily(row, demo_row)

    # BambooHR tiene prioridad sobre demo para Status y End Date
    row[35] = emp.get('status') or ''
    row[36] = _fmt_d(term_date)

    # BambooHR tiene prioridad sobre demo para Employment Status y datos de
    # terminación (la hoja "demografia - claude" puede quedar desactualizada
    # cuando una baja no dispara un evento de "Job Information" en Bamboo)
    row[50] = emp.get('employmentHistoryStatus') or ''
    row[51] = emp.get('4313') or ''  # Termination Type
    row[52] = emp.get('4315') or ''  # Eligible For Re-hire
    row[53] = emp.get('4964') or ''  # Regrettable or Non-Regrettable

    if email in TERMINATION_TYPE_OVERRIDE_EMAILS:
        row[51] = TERMINATION_TYPE_OVERRIDE_VALUE

    if _is_cosmolex(emp):
        row[46] = 'Cosmolex'

    for i, v in enumerate(row):
        row[i] = _clean_placeholder(v)


# ─── ELIMINAR COLUMNA AK ──────────────────────────────────────────────────────

def delete_ak_if_exists(svc, dry_run):
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_BDD,
        range=f"'{TAB_HIST}'!A1:BF1",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    header = r.get('values', [[]])[0]
    if len(header) <= 54:
        print('Columna AK ya eliminada, nada que hacer.')
        return

    print(f'Eliminando columna AK de "{TAB_HIST}" y "{TAB_ACT}"...')
    if dry_run:
        print('  [DRY-RUN] Sin escritura.')
        return

    meta      = svc.spreadsheets().get(spreadsheetId=SHEET_BDD).execute()
    sheet_ids = {s['properties']['title']: s['properties']['sheetId']
                 for s in meta['sheets']}

    reqs = []
    for tab in [TAB_HIST, TAB_ACT]:
        sid = sheet_ids.get(tab)
        if sid is not None:
            reqs.append({'deleteDimension': {'range': {
                'sheetId': sid, 'dimension': 'COLUMNS',
                'startIndex': 36, 'endIndex': 37,
            }}})

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_BDD, body={'requests': reqs}).execute()
    print(f'  Columna AK eliminada de {len(reqs)} solapa(s).')


# ─── CARGA DE DATOS ───────────────────────────────────────────────────────────

# Campos demográficos de BambooHR → índice en la "fila demo" que consume
# DEMO_MAP_DAILY. Antes venían de la planilla "demografia - claude" (DEMO_SSID),
# que dejó de existir (oct-2026); ahora la fuente es directamente BambooHR.
_DEMO_BAMBOO_FIELDS = {
    2:  'status',
    4:  'terminationDate',
    5:  'dateOfBirth',
    6:  'customGender',
    7:  'customEstadoCivil',
    8:  'customPareja',
    9:  'customHijos',
    10: 'city',
    11: 'state',
    12: 'country',
    13: 'originalHireDate',
    16: 'division',
    18: 'jobTitle',
    19: 'reportsTo',
    22: 'employeeStatusDate',
    23: 'employmentHistoryStatus',
}
_DEMO_BAMBOO_DATES = {4, 5, 13, 22}


# Valores "placeholder" de Bamboo que deben quedar vacíos en la BDD.
_BLANK_VALUES = {'to complete'}


def _clean_placeholder(v):
    return '' if str(v).strip().lower() in _BLANK_VALUES else v


def load_demo_lookup(svc=None):
    """Arma {email: fila demo} desde BambooHR (mismo formato que la ex planilla demo)."""
    print('Leyendo datos demográficos desde BambooHR...')
    base, auth = _bamboo()
    r = requests.post(
        f'{base}/reports/custom', auth=auth,
        headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
        params={'format': 'json'},
        json={'title': 'daily_demo_fields',
              'fields': ['workEmail'] + list(_DEMO_BAMBOO_FIELDS.values())},
    )
    r.raise_for_status()
    lookup = {}
    for emp in r.json().get('employees', []):
        email = str(emp.get('workEmail') or '').strip().lower()
        if not email:
            continue
        row = [''] * _DEMO_MAX_DAILY
        row[1] = email
        for i, field in _DEMO_BAMBOO_FIELDS.items():
            val = emp.get(field)
            if i in _DEMO_BAMBOO_DATES:
                row[i] = _fmt_d(_parse_date(val))
            elif i == 9:   # Children: número
                v = str(val or '').strip()
                row[i] = int(v) if v.isdigit() else v
            else:
                row[i] = str(val).strip() if val is not None else ''
        lookup[email] = [_clean_placeholder(x) for x in row]
    print(f'  {len(lookup)} emails con datos demográficos.')
    return lookup


def load_historico(svc):
    print('Leyendo Histórico...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_BDD,
        range=f"'{TAB_HIST}'!A:BC",
        valueRenderOption='UNFORMATTED_VALUE',
        dateTimeRenderOption='FORMATTED_STRING',
    ).execute()
    all_rows = r.get('values', [])
    header   = all_rows[0] if all_rows else []
    data     = [list(row) for row in all_rows[1:]]
    print(f'  {len(data)} filas de datos.')
    return header, data


def fetch_all_bamboo():
    base, auth = _bamboo()
    print('Descargando directorio de BambooHR...')
    r = requests.post(
        f'{base}/reports/custom', auth=auth,
        headers={'Accept': 'application/json', 'Content-Type': 'application/json'},
        params={'format': 'json'},
        json={'title': 'daily_demo_update', 'fields': [
            'id', 'firstName', 'lastName', 'workEmail',
            'hireDate', 'terminationDate', 'status', 'division',
            'department', 'location', 'employmentHistoryStatus',
            '4614', '4656', '4657', '4658',   # Code, Agreement, Bill, PayRoll
            '4659', '4660',                    # Per, Pay Rate - Currency Code
            '4500', '4505',                    # EffectiveDate, EffectiveDate1
            '4506', '4507', '4776',            # Cliente, Proyecto, %Asignación
            '4314', '4777',                    # Termination Reason (real), FechaFinalizaciónAsignación
            '4839', '4471.2', '4940.2',       # xM, Vacaciones, PTO
            '4313', '4315', '4964',           # Termination Type, Eligible For Re-hire, Regrettable
        ]},
    )
    r.raise_for_status()
    emps = r.json().get('employees', [])
    print(f'  {len(emps)} empleados en directorio.')
    return emps


# ─── HISTÓRICO ────────────────────────────────────────────────────────────────

def _write_daily_cols_to_hist(svc, data, indexed_rows):
    """Escribe cols diarias (B, D–W, Y, AK–AZ) en las filas indicadas."""
    BATCH = 100
    written = 0
    for start in range(0, len(indexed_rows), BATCH):
        chunk = indexed_rows[start:start + BATCH]
        vr = []
        for i, _ in chunk:
            sr  = i + 2  # fila en sheet (header en fila 1, datos desde fila 2)
            row = data[i]
            vr.extend([
                {'range': f"'{TAB_HIST}'!B{sr}",          'values': [[row[1]]]},
                {'range': f"'{TAB_HIST}'!D{sr}:W{sr}",   'values': [row[3:23]]},
                {'range': f"'{TAB_HIST}'!Y{sr}",          'values': [[row[24]]]},
                {'range': f"'{TAB_HIST}'!AJ{sr}:AY{sr}", 'values': [row[35:51]]},
            ])
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_BDD,
            body={'valueInputOption': 'USER_ENTERED', 'data': vr},
        ).execute()
        written += len(chunk)
    print(f'  {written} filas escritas en Histórico.')


def _latest_by_email(data):
    """Devuelve dict email → (month_str, list(row)) con la fila más reciente."""
    latest = {}
    for row in data:
        if len(row) <= EMAIL_COL:
            continue
        email = str(row[EMAIL_COL]).strip().lower()
        email = EMAIL_CORRECTIONS.get(email, email)   # normalizar typos
        if '@' not in email:
            # Filas con legajo en lugar de email — ignorar para no duplicar en Actualizado
            continue
        month = str(row[MONTH_COL]).strip() if row else ''
        if not email or not month:
            continue
        prev = latest.get(email)
        if not prev or month > prev[0]:
            latest[email] = (month, list(row))
    return latest


def _create_new_month_rows(svc, data, bamboo_lookup, demo_lookup, snapshot_month, dry_run,
                           skip_emails=None):
    """Crea filas del mes nuevo para activos en BambooHR que no tengan fila del mes."""
    latest     = _latest_by_email(data)
    skip_emails = skip_emails or set()

    new_rows = []
    for email, emp in bamboo_lookup.items():
        if email in skip_emails:
            continue  # ya tiene fila del mes
        if email in EXCLUDE_EMAIL:
            continue
        if _parse_date(emp.get('terminationDate')):
            continue  # solo activos
        if emp.get('division') in EXCLUDE_DIV:
            continue

        base    = latest.get(email)
        new_row = list(base[1]) if base else [''] * TOTAL_COLS
        while len(new_row) < TOTAL_COLS:
            new_row.append('')
        new_row[MONTH_COL] = snapshot_month
        new_row[EMAIL_COL] = emp.get('workEmail') or email

        _apply_daily_cols(new_row, emp, demo_lookup.get(email))
        new_rows.append(new_row)

    new_rows.sort(key=lambda r: str(r[1]).lower())
    print(f'  {len(new_rows)} filas nuevas para {snapshot_month}.')

    if dry_run or not new_rows:
        return

    BATCH = 500
    for start in range(0, len(new_rows), BATCH):
        chunk = new_rows[start:start + BATCH]
        svc.spreadsheets().values().append(
            spreadsheetId=SHEET_BDD,
            range=f"'{TAB_HIST}'!A1",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': chunk},
        ).execute()
        print(f'  Lote {start+1}–{start+len(chunk)} agregado.')


def update_historico(svc, header, data, bamboo_lookup, demo_lookup, snapshot_month, dry_run):
    """
    Actualiza Histórico. Devuelve True si se crearon filas del mes nuevo.

    - Siempre actualiza las filas del mes actual que ya existen (cols diarias).
    - Si hay activos en BambooHR sin fila del mes actual + hoy es día hábil →
      crea las filas faltantes.
    """
    # Filas y emails que ya tienen fila del mes actual
    current_emails = set()
    current_rows   = []
    for i in range(len(data)):
        row = data[i]
        if len(row) > MONTH_COL and str(row[MONTH_COL]).strip() == snapshot_month:
            email = str(row[EMAIL_COL]).strip().lower() if len(row) > EMAIL_COL else ''
            current_emails.add(email)
            current_rows.append((i, row))

    # Activos en BambooHR que aún no tienen fila del mes
    missing_active = [
        (email, emp) for email, emp in bamboo_lookup.items()
        if not _parse_date(emp.get('terminationDate'))
        and emp.get('division') not in EXCLUDE_DIV
        and email not in EXCLUDE_EMAIL
        and email not in current_emails
    ]

    created_new = False
    if missing_active:
        if _is_weekday():
            print(f'{len(missing_active)} activos sin fila para {snapshot_month}. Creando...')
            _create_new_month_rows(svc, data, bamboo_lookup, demo_lookup,
                                   snapshot_month, dry_run, skip_emails=current_emails)
            created_new = True
        else:
            print(f'{len(missing_active)} activos sin fila para {snapshot_month} '
                  f'y hoy es fin de semana. Sin nuevas filas.')

    if current_rows:
        print(f'Actualizando {len(current_rows)} filas de {snapshot_month} en Histórico...')
        updated = 0
        for i, row in current_rows:
            while len(data[i]) < TOTAL_COLS:
                data[i].append('')
            email    = str(row[EMAIL_COL]).strip().lower() if len(row) > EMAIL_COL else ''
            emp      = bamboo_lookup.get(email)
            demo_row = demo_lookup.get(email)
            if emp:
                _apply_daily_cols(data[i], emp, demo_row)
                updated += 1
            elif demo_row:
                _apply_demo_daily(data[i], demo_row)
        print(f'  {updated} de {len(current_rows)} filas actualizadas desde BambooHR.')
        if dry_run:
            print('[DRY-RUN] Sin escritura al Histórico.')
        else:
            _write_daily_cols_to_hist(svc, data, current_rows)
    elif not missing_active:
        print(f'Sin filas para {snapshot_month} y hoy es fin de semana. Sin cambios en Histórico.')

    return created_new


# ─── PERSONAS NUEVAS ──────────────────────────────────────────────────────────

def add_missing_employees(svc, data, demo_lookup, bamboo_lookup, all_emps, snapshot_month, dry_run):
    """Agrega personas activas en BambooHR que no figuren en Histórico."""
    existing = {
        str(row[EMAIL_COL]).strip().lower()
        for row in data
        if len(row) > EMAIL_COL and row[EMAIL_COL]
    }
    missing = [
        e for e in all_emps
        if not _parse_date(e.get('terminationDate'))          # solo activos
        and e.get('division') not in EXCLUDE_DIV
        and str(e.get('workEmail') or '').strip().lower()
        and str(e.get('workEmail') or '').strip().lower() not in EXCLUDE_EMAIL
        and str(e.get('workEmail') or '').strip().lower() not in existing
    ]
    print(f'Personas activas nuevas en BambooHR: {len(missing)}')
    if not missing:
        return

    new_rows = []
    for emp in missing:
        email = str(emp.get('workEmail') or '').strip().lower()
        row   = [''] * TOTAL_COLS
        row[MONTH_COL] = snapshot_month
        row[EMAIL_COL] = email
        _apply_daily_cols(row, emp, demo_lookup.get(email))
        new_rows.append(row)

    print(f'  Filas a agregar: {len(new_rows)}')
    if dry_run or not new_rows:
        return

    BATCH = 500
    for start in range(0, len(new_rows), BATCH):
        chunk = new_rows[start:start + BATCH]
        svc.spreadsheets().values().append(
            spreadsheetId=SHEET_BDD,
            range=f"'{TAB_HIST}'!A1",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': chunk},
        ).execute()
        print(f'  Lote {start+1}–{start+len(chunk)} agregado.')


# ─── RECONSTRUIR ACTUALIZADO ──────────────────────────────────────────────────

def rebuild_actualizado(svc, bamboo_lookup, demo_lookup, dry_run):
    """
    Reconstruye Actualizado desde Histórico.
    Incluye TODAS las personas (activos e inactivos).
    Aplica cols diarias frescas desde BambooHR para todos.
    """
    print('Leyendo Histórico para reconstruir Actualizado...')
    r = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_BDD,
        range=f"'{TAB_HIST}'!A:BC",
        valueRenderOption='UNFORMATTED_VALUE',
        dateTimeRenderOption='FORMATTED_STRING',
    ).execute()
    all_rows = r.get('values', [])
    if not all_rows:
        print('  Histórico vacío.')
        return
    header = all_rows[0]
    data   = [list(row) for row in all_rows[1:]]

    # Fila más reciente por persona (TODAS, incluyendo inactivos)
    latest = _latest_by_email(data)

    # Índice de costos por email: email → lista (mes, costo) desc, para fallback
    _latest_costo = {}
    for row in data:
        if len(row) <= EMAIL_COL: continue
        email = EMAIL_CORRECTIONS.get(str(row[EMAIL_COL]).strip().lower(),
                                      str(row[EMAIL_COL]).strip().lower())
        if '@' not in email: continue
        try:
            c = float(row[23]) if len(row) > 23 and str(row[23]).strip() not in ('', '0') else None
        except (ValueError, TypeError):
            c = None
        if c and c > 0:
            mes = str(row[MONTH_COL]).strip() if row else ''
            _latest_costo.setdefault(email, []).append((mes, c))
    for email in _latest_costo:
        _latest_costo[email].sort(key=lambda x: x[0], reverse=True)

    # Aplicar cols diarias frescas para todos
    act_data = []
    for email, (_, row) in latest.items():
        while len(row) < TOTAL_COLS:
            row.append('')
        emp      = bamboo_lookup.get(email)
        demo_row = demo_lookup.get(email)
        if emp:
            _apply_daily_cols(row, emp, demo_row)
        elif demo_row:
            _apply_demo_daily(row, demo_row)
        # Si el costo de la fila más reciente está vacío, usar el más reciente disponible
        try:
            costo_actual = float(row[23]) if str(row[23]).strip() not in ('', '0') else None
        except (ValueError, TypeError):
            costo_actual = None
        if not costo_actual:
            entradas = _latest_costo.get(email, [])
            if entradas:
                row[23] = entradas[0][1]
        act_data.append(row)

    act_data.sort(key=lambda r: str(r[1]).lower())
    act_rows = [header] + act_data
    print(f'  {len(act_data)} personas en Actualizado.')

    if dry_run:
        print('[DRY-RUN] Sin escritura al Actualizado.')
        return

    meta      = svc.spreadsheets().get(spreadsheetId=SHEET_BDD).execute()
    sheet_ids = {s['properties']['title']: s['properties']['sheetId']
                 for s in meta['sheets']}

    reqs = []
    if TAB_ACT in sheet_ids:
        reqs.append({'deleteSheet': {'sheetId': sheet_ids[TAB_ACT]}})
    reqs.append({'addSheet': {'properties': {'title': TAB_ACT, 'index': 0}}})
    resp    = svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_BDD, body={'requests': reqs}).execute()
    new_sid = resp['replies'][-1]['addSheet']['properties']['sheetId']

    BATCH = 1000
    for start in range(0, len(act_rows), BATCH):
        svc.spreadsheets().values().update(
            spreadsheetId=SHEET_BDD,
            range=f"'{TAB_ACT}'!A{start+1}",
            valueInputOption='USER_ENTERED',
            body={'values': act_rows[start:start+BATCH]},
        ).execute()

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_BDD,
        body={'requests': [
            {'updateSheetProperties': {
                'properties': {'sheetId': new_sid,
                               'gridProperties': {'frozenRowCount': 1}},
                'fields': 'gridProperties.frozenRowCount',
            }},
            {'repeatCell': {
                'range': {'sheetId': new_sid, 'startRowIndex': 0, 'endRowIndex': 1},
                'cell': {'userEnteredFormat': {
                    'backgroundColor': {'red': 0.122, 'green': 0.286, 'blue': 0.490},
                    'textFormat': {'bold': True,
                                   'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}},
                }},
                'fields': 'userEnteredFormat(backgroundColor,textFormat)',
            }},
        ]},
    ).execute()
    _apply_sheet_formats(svc, new_sid, len(act_rows) + 500)
    print('  Actualizado reconstruido.')


# ─── COMPENSACIONES (3er día hábil) ──────────────────────────────────────────

GENERAL_SID = '15mTvxuwXGa8B8Cuc86qqDyb8eKhKrDlwRekx9I1dZ_0'


def _business_day_of_month() -> int:
    """Retorna el número de día hábil del mes hasta hoy (inclusive)."""
    today = datetime.date.today()
    count = 0
    d = today.replace(day=1)
    while d <= today:
        if d.weekday() < 5:
            count += 1
        d += datetime.timedelta(days=1)
    return count


def _fetch_job_status():
    """Descarga el último Cambio Proactivo por email desde el reporte 740."""
    base, auth = _bamboo()
    r = requests.get(
        f'{base}/reports/740', auth=auth,
        headers={'Accept': 'application/json'},
        params={'format': 'json'},
    )
    r.raise_for_status()
    lookup = {}
    for rec in r.json().get('employees', []):
        email = str(rec.get('employeeNumber') or '').strip().lower()
        if not email:
            continue
        d = _parse_date(rec.get('customEffectiveDate3'))
        if d and (email not in lookup or d > lookup[email]):
            lookup[email] = d
    print(f'  Job Status: {len(lookup)} empleados con Cambio Proactivo.')
    return lookup


def _read_variables(svc, month_dt):
    """Lee Variables para Rep del General sheet para el mes dado."""
    r = svc.spreadsheets().values().get(
        spreadsheetId=GENERAL_SID,
        range="'Variables para Rep'!A1:J50",
    ).execute()
    for row in r.get('values', [])[3:]:
        if not row:
            continue
        d = _parse_date(str(row[0]).strip())
        if d and d.year == month_dt.year and d.month == month_dt.month:
            def _f(i):
                try:
                    return float(str(row[i]).replace(',', '.'))
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
            }
    return None


def _read_bandas(svc):
    """Lee Bandas Div del General sheet. Retorna dict new_code → (min, med, max)."""
    r = svc.spreadsheets().values().get(
        spreadsheetId=GENERAL_SID,
        range="'Bandas Div'!A4:K500",
    ).execute()
    lookup = {}
    for row in r.get('values', []):
        if len(row) < 10:
            continue
        new_code = str(row[1]).strip()
        if not new_code:
            continue
        def _fv(i):
            try:
                return float(str(row[i]).replace(',', '.'))
            except (IndexError, ValueError):
                return None
        h, m, j = _fv(7), _fv(8), _fv(9)
        if h is not None and m is not None and j is not None:
            lookup[new_code] = (h, m, j)
    print(f'  Bandas Div: {len(lookup)} entradas.')
    return lookup


def update_compensaciones(svc, data, bamboo_lookup, snapshot_month, dry_run):
    """
    Calcula y escribe los campos de compensación en Histórico para el mes actual.
    Se ejecuta a partir del 3er día hábil del mes.
    Columnas: X(23) Costo, Y(24) New Code, Z-AC(25-28) Bandas+GAP,
              AD-AI(29-34) Antigüedad, Cambio Proactivo, Revisar, %GAP, Ajuste, Dispersión.
    """
    try:
        from generar_compensaciones import (
            to_float, calc_new_code, calc_costo_usd_h,
            calc_gap_banda, calc_antiguedad, calc_revisar_seniority,
            calc_pct_gap_banda, calc_ajuste_ultimo_ano,
            calc_dispersion_en_banda, MONTH_ABBR_ES,
        )
    except Exception as e:
        print(f'Error importando generar_compensaciones: {e}')
        return

    # Cargar emails con costo protegido (no se sobreescriben en la actualización automática)
    import json as _json, os as _os
    _protected_file = _os.path.join(_os.path.dirname(__file__), 'costos_protegidos.json')
    try:
        with open(_protected_file, encoding='utf-8') as _f:
            _emails_protegidos = set(e.lower().strip() for e in _json.load(_f).get('emails', []))
        print(f'Costos protegidos: {len(_emails_protegidos)} emails ({", ".join(sorted(_emails_protegidos))})')
    except FileNotFoundError:
        _emails_protegidos = set()
    except Exception as _e:
        print(f'Advertencia: no se pudo leer costos_protegidos.json: {_e}')
        _emails_protegidos = set()

    month_dt = datetime.datetime(
        int(snapshot_month[:4]), int(snapshot_month[5:7]), 1)
    today = datetime.date.today()

    print('\nCargando variables y bandas para compensaciones...')
    vars_dict = _read_variables(svc, month_dt)
    if not vars_dict:
        print(f'  Sin variables para {snapshot_month} en General sheet. '
              f'Verificá la solapa "Variables para Rep".')
        return
    bandas_lookup = _read_bandas(svc)
    print('Descargando Job Status...')
    job_lookup = _fetch_job_status()

    current_rows = [
        (i, row) for i, row in enumerate(data)
        if len(row) > MONTH_COL and str(row[MONTH_COL]).strip() == snapshot_month
    ]
    print(f'Calculando compensaciones para {len(current_rows)} filas de {snapshot_month}...')

    for i, row in current_rows:
        while len(row) < TOTAL_COLS:
            row.append('')

        email     = str(row[EMAIL_COL]).strip().lower()
        emp       = bamboo_lookup.get(email)
        code      = str(row[8]).strip()
        agreement = str(row[10]).strip()
        location  = str(row[5]).strip()
        bill      = to_float(row[11])
        payroll   = to_float(row[12])
        per       = str(row[13]).strip()

        # Usar fechas desde BambooHR (ISO, sin ambigüedad)
        hire_dt = _parse_date(emp.get('hireDate')) if emp else _parse_date(row[3])
        eff_dt1 = _parse_date(emp.get('customEffectiveDate')) if emp else _parse_date(row[7])

        # New Code
        new_code = calc_new_code(code, agreement, location) if code else ''
        row[24] = new_code  # Y

        # Costo: protegidos se saltean; Doppler → mensual; Resto → USD/H
        division   = str(row[46]).strip() if len(row) > 46 else ''
        is_doppler = 'doppler' in division.lower()
        if email in _emails_protegidos:
            costo = row[23]  # mantener valor existente, no tocar
        elif is_doppler:
            if agreement == 'Contractor' and bill:
                # Contractor Doppler: solo bill (mensual si per=Month, ×168 si Hour)
                costo = round(bill * 168, 4) if per == 'Hour' else bill
            elif bill or payroll:
                # Plus Fijo / Empleado / Temporal / Mixto: calc USD/H y convertir a mensual
                costo_h = calc_costo_usd_h(agreement, location, bill, payroll, per, vars_dict)
                costo = round(float(costo_h) * 168, 4) if costo_h != '' else ''
            else:
                costo = ''
        elif bill or payroll:
            costo = calc_costo_usd_h(agreement, location, bill, payroll, per, vars_dict)
        else:
            costo = ''
        row[23] = costo  # X

        # Bandas y derivados
        banda = bandas_lookup.get(new_code) if new_code else None
        if banda:
            b_min, b_med, b_max = banda
            row[25] = b_min  # Z
            row[26] = b_med  # AA
            row[27] = b_max  # AB
            gap = calc_gap_banda(costo, b_min, b_max) if costo != '' else ''
            row[28] = gap                                                         # AC
            row[32] = calc_pct_gap_banda(gap) if gap != '' else ''               # AG
            row[34] = calc_dispersion_en_banda(costo, b_min, b_max) \
                      if costo != '' else ''                                      # AI

        # Antigüedad
        if hire_dt:
            row[29] = calc_antiguedad(hire_dt, today)  # AD

        # Último Cambio Proactivo
        last_change = job_lookup.get(email)
        if last_change:
            row[30] = f"{MONTH_ABBR_ES[last_change.month]}-{last_change.year}"  # AE

        # Revisar seniority
        if code and eff_dt1:
            row[31] = calc_revisar_seniority(code, eff_dt1, today)  # AF

        # Ajuste último año
        if hire_dt:
            row[33] = calc_ajuste_ultimo_ano(month_dt, hire_dt, last_change)  # AH

    print(f'  {len(current_rows)} filas calculadas.')

    if dry_run:
        print('  [DRY-RUN] Sin escritura.')
        return

    BATCH = 100
    written = 0
    for start in range(0, len(current_rows), BATCH):
        chunk = current_rows[start:start + BATCH]
        vr = []
        for i, row in chunk:
            sr = i + 2
            vr.extend([
                {'range': f"'{TAB_HIST}'!X{sr}",          'values': [[row[23]]]},
                {'range': f"'{TAB_HIST}'!Y{sr}",          'values': [[row[24]]]},
                {'range': f"'{TAB_HIST}'!Z{sr}:AC{sr}",  'values': [row[25:29]]},
                {'range': f"'{TAB_HIST}'!AD{sr}:AI{sr}", 'values': [row[29:35]]},
            ])
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_BDD,
            body={'valueInputOption': 'USER_ENTERED', 'data': vr},
        ).execute()
        written += len(chunk)
    print(f'  {written} filas de compensación escritas.')


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    dry_run        = '--dry-run' in sys.argv
    snapshot_month = _today_snapshot()
    today          = datetime.date.today()
    biz_day        = _business_day_of_month()
    print(f'\n{"[DRY-RUN] " if dry_run else ""}actualizar_demografia.py')
    print(f'  Fecha: {today}  |  Mes: {snapshot_month}  |  Día hábil: {biz_day}\n')

    svc = _service()

    _load_doppler_seniority_map(svc)
    _load_td_project_map(svc)
    delete_ak_if_exists(svc, dry_run)

    demo_lookup  = load_demo_lookup(svc)
    header, data = load_historico(svc)
    all_emps     = fetch_all_bamboo()

    bamboo_lookup = {
        str(e.get('workEmail') or '').strip().lower(): e
        for e in all_emps
        if e.get('workEmail')
    }

    created_new = update_historico(
        svc, header, data, bamboo_lookup, demo_lookup, snapshot_month, dry_run)

    # Solo agregar faltantes si NO se creó el lote del mes nuevo
    # (en ese caso, ya se incluyeron todos los activos)
    if not created_new:
        add_missing_employees(
            svc, data, demo_lookup, bamboo_lookup, all_emps, snapshot_month, dry_run)

    rebuild_actualizado(svc, bamboo_lookup, demo_lookup, dry_run)

    if not dry_run:
        print('Aplicando formatos al Histórico...')
        meta    = svc.spreadsheets().get(spreadsheetId=SHEET_BDD).execute()
        tab_ids = {s['properties']['title']: s['properties']['sheetId']
                   for s in meta['sheets']}
        if TAB_HIST in tab_ids:
            _apply_sheet_formats(svc, tab_ids[TAB_HIST])

    print('\nListo.')
    print(f'  https://docs.google.com/spreadsheets/d/{SHEET_BDD}/edit')


if __name__ == '__main__':
    from alertas import run_with_alert
    run_with_alert(main, 'actualizar_demografia.py (BDD Bamboo diaria)')
