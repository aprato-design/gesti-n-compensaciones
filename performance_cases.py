import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from pathlib import Path
import re
from datetime import datetime
import uuid

# ── Config ────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Gestión Compensaciones", page_icon="🎯", layout="wide")

SPREADSHEET_ID = '15mTvxuwXGa8B8Cuc86qqDyb8eKhKrDlwRekx9I1dZ_0'
VARIABLES_SHEET_ID = '1vBmF6PdU9ZCeB-6pqX5hxQ-FM0G3yWgd7MqtDfFlt78'
SHEET_SUELDOS = 'Bamboo Compensaciones - Sueldos'
SHEET_BANDAS = 'Bandas Div'
SHEET_CASOS = 'Casos Performance'
HORAS = 168
COL_MIN_WAGE_10X = 14_235_000  # 10 × salario mínimo Colombia 2026 (COP)

TIPOS_CASO = ['Solo feedback', 'Ajuste compensación', 'Recat sin ajuste', 'Recat con ajuste']
STATUS_OPEN = 'Abierto'
STATUS_CLOSED = 'Cerrado'

try:
    READ_ONLY = bool(st.secrets.get('read_only', False))
except Exception:
    READ_ONLY = False

CASOS_COLS = [
    'id', 'created_at', 'created_by',
    'employee_email', 'employee_name', 'code', 'seniority',
    'hire_date', 'xm', 'agreement', 'per', 'currency',
    'current_bill', 'current_payroll', 'current_costo',
    'banda_min', 'banda_med', 'banda_max',
    'new_code_empleado',
    'target_code', 'target_banda_min', 'target_banda_med', 'target_banda_max',
    'proposed_bill', 'proposed_payroll', 'new_costo', 'differential',
    'tipo', 'notes', 'status', 'closed_at', 'closed_by',
]

NUM_COLS_CASOS = [
    'current_bill', 'current_payroll', 'current_costo',
    'banda_min', 'banda_med', 'banda_max',
    'target_banda_min', 'target_banda_med', 'target_banda_max',
    'proposed_bill', 'proposed_payroll', 'new_costo', 'differential',
]

# ── CSS ───────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
.caso-card {
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 8px;
    background: white;
}
.caso-card:hover { border-color: #1976d2; }
.badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 0.78rem;
    font-weight: 600;
}
.badge-feedback { background: #e3f2fd; color: #1565c0; }
.badge-ajuste   { background: #e8f5e9; color: #2e7d32; }
.badge-recat    { background: #fff3e0; color: #e65100; }
.section-header {
    font-size: 1rem;
    font-weight: 700;
    color: #424242;
    border-left: 3px solid #1976d2;
    padding-left: 8px;
    margin: 16px 0 8px 0;
}
.budget-box {
    background: #e8f5e9;
    border-radius: 8px;
    padding: 16px 24px;
    margin-bottom: 16px;
}
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {
    font-size: 0.72rem !important;
    line-height: 1.3 !important;
}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * {
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    line-height: 1.2 !important;
}
[data-testid="stMetricDelta"], [data-testid="stMetricDelta"] * {
    font-size: 0.75rem !important;
}
</style>
""", unsafe_allow_html=True)


# ── Auth ─────────────────────────────────────────────────────────────────────

@st.cache_resource
def get_gc():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes)
    except Exception:
        creds_file = Path(__file__).parent / 'talentserviceproject-1ce2ed91696b.json'
        creds = Credentials.from_service_account_file(str(creds_file), scopes=scopes)
    return gspread.authorize(creds)


# ── Data loading ──────────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def load_empleados() -> pd.DataFrame:
    gc = get_gc()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_SUELDOS)
    rows = ws.get_all_values()

    headers = rows[0]
    df = pd.DataFrame(rows[1:], columns=headers)

    df['Month'] = pd.to_datetime(df['Month'], errors='coerce')
    df = df[df['Month'].notna()]
    df = df.sort_values('Month')

    # Active only
    df = df[df['Movimiento'] == 'Activo']

    # Latest month per employee
    df = df.drop_duplicates(subset='Employee #', keep='last')

    df = df.rename(columns={
        'Employee #': 'email',
        'Last name, First name': 'name',
        'Code (Level)': 'code',
        'Seniority': 'seniority',
        'New Code': 'new_code',
        'xM': 'xm',
        'Agreement': 'agreement',
        'Per': 'per',
        'Bill': 'bill',
        'PayRoll': 'payroll',
        'Pay Rate - Currency Code': 'currency',
        'Costo USD/H': 'costo_usd_h',
        'Banda Min': 'banda_min',
        'Banda Med': 'banda_med',
        'Banda Max': 'banda_max',
        'Hire Date': 'hire_date',
    })

    for col in ['bill', 'payroll', 'costo_usd_h', 'banda_min', 'banda_med', 'banda_max']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    keep = ['email', 'name', 'code', 'seniority', 'new_code', 'xm', 'agreement',
            'per', 'currency', 'bill', 'payroll', 'costo_usd_h', 'banda_min',
            'banda_med', 'banda_max', 'hire_date']
    return df[[c for c in keep if c in df.columns]].reset_index(drop=True)


@st.cache_data(ttl=300, show_spinner=False)
def load_bandas() -> pd.DataFrame:
    gc = get_gc()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_BANDAS)
    rows = ws.get_all_values()

    def safe_float(v):
        try:
            return float(v.strip()) if v and v.strip() else None
        except ValueError:
            return None

    bandas = []
    for row in rows[3:]:
        if len(row) < 2:
            continue
        new_code = row[1].strip() if len(row) > 1 else ''
        if not new_code:
            continue
        # Prefer Costo/hs USD columns (7,8,9); fallback to Contractors (3,4,5)
        bmin = safe_float(row[7] if len(row) > 7 else '') or safe_float(row[3] if len(row) > 3 else '')
        bmed = safe_float(row[8] if len(row) > 8 else '') or safe_float(row[4] if len(row) > 4 else '')
        bmax = safe_float(row[9] if len(row) > 9 else '') or safe_float(row[5] if len(row) > 5 else '')
        bandas.append({
            'new_code': new_code,
            'position': row[0].strip() if row[0] else '',
            'banda_min': bmin,
            'banda_med': bmed,
            'banda_max': bmax,
            # separate sets for cross-modality display
            'contractor_min': safe_float(row[3] if len(row) > 3 else ''),
            'contractor_med': safe_float(row[4] if len(row) > 4 else ''),
            'contractor_max': safe_float(row[5] if len(row) > 5 else ''),
            'costo_usd_min': safe_float(row[7] if len(row) > 7 else ''),
            'costo_usd_med': safe_float(row[8] if len(row) > 8 else ''),
            'costo_usd_max': safe_float(row[9] if len(row) > 9 else ''),
        })

    return pd.DataFrame(bandas)


@st.cache_data(ttl=300, show_spinner=False)
def load_variables() -> dict:
    """Load most recent month variables from the variables sheet."""
    gc = get_gc()
    sh = gc.open_by_key(VARIABLES_SHEET_ID)
    ws = sh.worksheet('Variables para Rep')
    rows = ws.get_all_values()

    def clean(v):
        return v.strip().replace('$', '').replace(',', '') if v else ''

    # Header map: col index → month string (row index 2 has variable names, rows 4+ have data)
    # Find most recent data row (last row with a non-empty FX value)
    latest = {}
    for row in rows[4:]:
        if not row or not clean(row[1]):
            continue
        # Check if this row has meaningful data (FX ARS OF must be present)
        try:
            float(clean(row[1]))
        except (ValueError, IndexError):
            continue
        # Read all variables from this row
        var_names = ['', 'FX ARS OF', 'FX ARG Blue', 'FX COPs', 'Costo ARG',
                     'Costo COL mayor', 'Costo COL menor', 'Horas', 'Costo ARG USD',
                     'Deducciones Empleado ARG %', 'Deducciones Empleado COL %',
                     'Deducciones contractor %', 'Deducciones usd plus fijo %',
                     'Prepaga contractor (usd/h)']
        row_vars = {}
        for i, name in enumerate(var_names[1:], start=1):
            val = clean(row[i]) if i < len(row) else ''
            if val and val not in ('#DIV/0!', '#N/A', '#REF!', ''):
                try:
                    row_vars[name] = float(val)
                except ValueError:
                    pass
        if row_vars:
            latest = row_vars

    return latest


@st.cache_data(ttl=120, show_spinner=False)
def load_casos() -> pd.DataFrame:
    gc = get_gc()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = _ensure_casos_ws(sh)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=CASOS_COLS)
    df = pd.DataFrame(records)
    for col in NUM_COLS_CASOS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


# ── Cases sheet helpers ───────────────────────────────────────────────────────

def _ensure_casos_ws(sh):
    try:
        return sh.worksheet(SHEET_CASOS)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_CASOS, rows=1000, cols=len(CASOS_COLS))
        ws.update([CASOS_COLS])
        return ws


def _get_casos_ws():
    gc = get_gc()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return _ensure_casos_ws(sh)


def create_caso(data: dict) -> str:
    ws = _get_casos_ws()
    caso_id = str(uuid.uuid4())[:8].upper()
    data['id'] = caso_id
    data['created_at'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    data['status'] = STATUS_OPEN
    data['closed_at'] = ''
    data['closed_by'] = ''
    row = [str(data.get(col, '')) for col in CASOS_COLS]
    ws.append_row(row, value_input_option='USER_ENTERED')
    st.cache_data.clear()
    return caso_id


def update_caso(caso_id: str, data: dict):
    ws = _get_casos_ws()
    all_vals = ws.get_all_values()
    for i, row in enumerate(all_vals[1:], start=2):
        if row[0] == caso_id:
            new_row = []
            for j, col in enumerate(CASOS_COLS):
                new_row.append(str(data.get(col, row[j] if j < len(row) else '')))
            end_col = chr(ord('A') + len(CASOS_COLS) - 1)
            ws.update(f'A{i}:{end_col}{i}', [new_row], value_input_option='USER_ENTERED')
            st.cache_data.clear()
            return


def close_caso_sheet(caso_id: str, closed_by: str):
    ws = _get_casos_ws()
    all_vals = ws.get_all_values()
    si = CASOS_COLS.index('status') + 1
    ci = CASOS_COLS.index('closed_at') + 1
    bi = CASOS_COLS.index('closed_by') + 1
    for i, row in enumerate(all_vals[1:], start=2):
        if row[0] == caso_id:
            ws.update_cell(i, si, STATUS_CLOSED)
            ws.update_cell(i, ci, datetime.now().strftime('%Y-%m-%d %H:%M'))
            ws.update_cell(i, bi, closed_by)
            st.cache_data.clear()
            return


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_next_code(new_code: str, bandas_df: pd.DataFrame):
    if not new_code:
        return None
    m = re.match(r'^([A-Z]+)\s+0*(\d+)-C$', new_code.strip())
    if not m:
        return None
    prefix, num = m.group(1), int(m.group(2))
    for fmt in [f"{prefix} {num+1:02d}-C", f"{prefix} {num+1}-C"]:
        if fmt in bandas_df['new_code'].values:
            return fmt
    return None


def fmt_usd(val, zero_dash=True) -> str:
    try:
        v = float(val)
    except (TypeError, ValueError):
        return '—'
    if v == 0 and zero_dash:
        return '—'
    return f'${v:,.2f}'


def calc_costo_empresa(bill: float, payroll: float, per: str,
                       agreement: str, currency: str, vars: dict) -> float | None:
    """Calculate Costo USD/H using the same formulas as the main compensation sheet.

    Formulas (from 'Proceso' tab):
      Contractor Hour:    bill
      Contractor Month:   bill / 168
      Empleado ARS:       payroll / FX_OF * Costo_ARG / 168
      Plus Fijo:          (bill * Costo_ARG_USD + payroll * Costo_ARG / FX_OF) / 168
      Empleado COL:       payroll / FX_COPs * Costo_COL / 168
      Biweekly:           payroll / FX_OF * Costo_ARG / 168  (same as ARS)
    """
    if not vars:
        return None

    fx_of = vars.get('FX ARS OF', 0)
    costo_arg = vars.get('Costo ARG', 0)
    costo_arg_usd = vars.get('Costo ARG USD', 1.085)
    fx_cop = vars.get('FX COPs', 0)
    costo_col_mayor = vars.get('Costo COL mayor', 0)
    costo_col_menor = vars.get('Costo COL menor', 0)

    if agreement == 'Contractor':
        if per == 'Hour' and bill:
            return bill
        if per == 'Month' and bill:
            return bill / HORAS

    elif agreement == 'Empleado':
        if currency == 'ARS' and payroll and fx_of and costo_arg:
            return payroll / fx_of * costo_arg / HORAS
        if currency == 'COP' and payroll and fx_cop:
            costo_col = costo_col_mayor if payroll > COL_MIN_WAGE_10X else costo_col_menor
            return payroll / fx_cop * costo_col / HORAS
        if currency == 'USD' and per == 'Biweekly' and payroll and fx_of and costo_arg:
            return payroll / fx_of * costo_arg / HORAS
        if currency == 'USD' and per == 'Month' and bill:
            return bill / HORAS

    elif agreement == 'Plus Fijo' and fx_of and costo_arg:
        return (bill * costo_arg_usd + payroll * costo_arg / fx_of) / HORAS

    return None


def _banda_position_html(label: str, costo: float, bmin: float, bmed: float, bmax: float) -> str:
    """Return HTML showing band position as a visual bar + ok/% indicator (column AC format)."""
    if not bmin or not bmax:
        return f'<div style="padding:8px 0"><b>{label}:</b> —</div>'

    if bmin <= costo <= bmax:
        status_html = '<span style="color:#2e7d32;font-weight:700">ok</span>'
    elif costo < bmin:
        pct = (costo - bmin) / bmin * 100
        status_html = f'<span style="color:#c62828;font-weight:700">{pct:.1f}%</span>'
    else:
        pct = (costo - bmax) / bmax * 100
        status_html = f'<span style="color:#e65100;font-weight:700">+{pct:.1f}%</span>'

    band_width = bmax - bmin
    dot_pct = max(-12, min(112, (costo - bmin) / band_width * 100)) if band_width > 0 else 50
    med_pct = ((bmed - bmin) / band_width * 100) if (bmed and band_width > 0) else 50
    dot_color = '#2e7d32' if bmin <= costo <= bmax else ('#c62828' if costo < bmin else '#e65100')

    return f'''<div style="padding:6px 0 10px 0">
      <div style="display:flex;justify-content:space-between;align-items:center;font-size:0.82rem;margin-bottom:4px">
        <span style="font-weight:600;color:#424242">{label}</span>
        <span>{status_html}&nbsp;&nbsp;<span style="color:#757575;font-size:0.78rem">{fmt_usd(costo)}</span></span>
      </div>
      <div style="position:relative;height:8px;background:#e0e0e0;border-radius:4px;margin:4px 0 2px 0">
        <div style="position:absolute;left:{med_pct:.1f}%;top:-3px;width:2px;height:14px;background:#bdbdbd;transform:translateX(-50%)"></div>
        <div style="position:absolute;left:{dot_pct:.1f}%;top:50%;width:14px;height:14px;border-radius:50%;background:{dot_color};transform:translate(-50%,-50%);border:2px solid white;box-shadow:0 1px 3px rgba(0,0,0,0.35)"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:0.70rem;color:#9e9e9e;margin-top:3px">
        <span>{fmt_usd(bmin)}</span>
        <span>{fmt_usd(bmed) if bmed else ""}</span>
        <span>{fmt_usd(bmax)}</span>
      </div>
    </div>'''


def badge_html(tipo: str) -> str:
    if tipo == 'Solo feedback':
        cls = 'badge-feedback'
    elif 'Recat' in tipo:
        cls = 'badge-recat'
    else:
        cls = 'badge-ajuste'
    return f'<span class="badge {cls}">{tipo}</span>'


def get_banda_for_code(new_code: str, bandas_df: pd.DataFrame) -> dict:
    rows = bandas_df[bandas_df['new_code'] == new_code]
    if rows.empty:
        return {'banda_min': None, 'banda_med': None, 'banda_max': None}
    r = rows.iloc[0]
    return {'banda_min': r['banda_min'], 'banda_med': r['banda_med'], 'banda_max': r['banda_max']}


def get_cross_banda_for_code(new_code: str, agreement: str, bandas_df: pd.DataFrame) -> dict:
    """Return the equivalent band for the opposite modality by swapping the code suffix.
    Contractor (SEN 04-C)   → looks up SEN 04-ARG.
    Empleado / Plus Fijo (SEN 04-ARG) → looks up SEN 04-C.
    Returns dict with 'cross_code' plus band values, or all None if not found.
    """
    if not new_code:
        return {'cross_code': None, 'banda_min': None, 'banda_med': None, 'banda_max': None}
    if agreement == 'Contractor':
        if not new_code.endswith('-C'):
            return {'cross_code': None, 'banda_min': None, 'banda_med': None, 'banda_max': None}
        cross_code = new_code[:-2] + '-ARG'
    else:
        if not new_code.endswith('-ARG'):
            return {'cross_code': None, 'banda_min': None, 'banda_med': None, 'banda_max': None}
        cross_code = new_code[:-4] + '-C'
    rows = bandas_df[bandas_df['new_code'] == cross_code]
    if rows.empty:
        return {'cross_code': cross_code, 'banda_min': None, 'banda_med': None, 'banda_max': None}
    r = rows.iloc[0]
    return {'cross_code': cross_code, 'banda_min': r['banda_min'], 'banda_med': r['banda_med'], 'banda_max': r['banda_max']}


# ── Caso form ─────────────────────────────────────────────────────────────────

def show_caso_form(empleados_df: pd.DataFrame, bandas_df: pd.DataFrame,
                   caso: dict = None, force_readonly: bool = False):
    is_edit = caso is not None
    actually_closed = is_edit and caso.get('status') == STATUS_CLOSED
    is_closed = actually_closed or force_readonly
    back_view = 'closed_list' if actually_closed else 'open_list'
    heading = '📋 Ver caso' if is_closed else ('✏️ Editar caso' if is_edit else '➕ Nuevo caso')

    st.markdown(f'### {heading}')

    vars = load_variables()

    # ── Employee selector ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Colaborador</div>', unsafe_allow_html=True)

    if is_edit:
        emp_email = caso['employee_email']
        emp_row = empleados_df[empleados_df['email'] == emp_email]
        emp = emp_row.iloc[0] if not emp_row.empty else None
        st.markdown(f"**{caso['employee_name']}** — `{caso['employee_email']}`")
    else:
        emp_options = [''] + sorted(empleados_df['name'].tolist())
        sel = st.selectbox('Seleccionar colaborador', emp_options, key='form_emp_select')
        if not sel:
            st.info('Seleccioná un colaborador para continuar.')
            return
        emp_rows = empleados_df[empleados_df['name'] == sel]
        emp = emp_rows.iloc[0] if not emp_rows.empty else None
        emp_email = emp['email'] if emp is not None else ''

    # Resolve current values: from caso (if editing) or from sheet
    if is_edit:
        code = caso['code']
        seniority = caso['seniority']
        hire_date = caso.get('hire_date', '')
        xm = caso.get('xm', '')
        agreement = caso.get('agreement', '')
        per = caso.get('per', 'Month')
        currency = caso.get('currency', 'USD')
        cur_bill = float(caso.get('current_bill', 0) or 0)
        cur_payroll = float(caso.get('current_payroll', 0) or 0)
        cur_costo = float(caso.get('current_costo', 0) or 0)
        banda_min = float(caso.get('banda_min', 0) or 0)
        banda_med = float(caso.get('banda_med', 0) or 0)
        banda_max = float(caso.get('banda_max', 0) or 0)
        new_code_empleado = caso.get('new_code_empleado', '')
        emp_name = caso['employee_name']
    else:
        if emp is None:
            st.error('Empleado no encontrado.')
            return
        code = emp['code']
        seniority = emp['seniority']
        hire_date = emp.get('hire_date', '')
        xm = emp.get('xm', '')
        agreement = emp.get('agreement', '')
        per = emp.get('per', 'Month')
        currency = emp.get('currency', 'USD')
        cur_bill = float(emp.get('bill', 0) or 0)
        cur_payroll = float(emp.get('payroll', 0) or 0)
        cur_costo = float(emp.get('costo_usd_h', 0) or 0)
        banda_min = float(emp.get('banda_min', 0) or 0)
        banda_med = float(emp.get('banda_med', 0) or 0)
        banda_max = float(emp.get('banda_max', 0) or 0)
        new_code_empleado = emp.get('new_code', '')
        emp_name = emp['name']

    # ── Info grid ─────────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric('Code / Level', code or '—')
    c2.metric('Seniority', seniority or '—')
    c3.metric('Agreement', agreement or '—')
    c4.metric('Moneda / Per', f'{currency} / {per}')
    c5.metric('xM', xm or '—')
    c6.metric('Hire Date', hire_date or '—')

    c1, c2, c3 = st.columns(3)
    c1.metric('Bill actual', fmt_usd(cur_bill))
    c2.metric('PayRoll actual', fmt_usd(cur_payroll))
    c3.metric('Costo USD/H actual', fmt_usd(cur_costo))

    # ── Banda actual ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Banda actual</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.write(f'**Code:** `{new_code_empleado or "—"}`')
    c2.metric('Mínimo', fmt_usd(banda_min))
    c3.metric('Medio', fmt_usd(banda_med))
    c4.metric('Máximo', fmt_usd(banda_max))

    # ── Pares ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Pares — mismo code y siguiente nivel</div>',
                unsafe_allow_html=True)

    next_code = get_next_code(new_code_empleado, bandas_df)

    mask = empleados_df['new_code'] == new_code_empleado
    if next_code:
        mask = mask | (empleados_df['new_code'] == next_code)

    peers = empleados_df[mask].copy()

    if peers.empty:
        st.info('No se encontraron pares para este code.')
    else:
        peers['Nivel'] = peers['new_code'].apply(
            lambda c: '➡ Siguiente' if c == next_code else 'Actual'
        )
        peers_cols = ['name', 'Nivel', 'code', 'seniority', 'bill', 'payroll',
                      'costo_usd_h', 'banda_min', 'banda_med', 'banda_max']
        peers_col_names = ['Nombre', 'Nivel', 'Code', 'Seniority', 'Bill', 'PayRoll',
                           'Costo USD/H', 'Banda Mín', 'Banda Med', 'Banda Máx']
        fmt_map = {
            'Bill': fmt_usd, 'PayRoll': fmt_usd, 'Costo USD/H': fmt_usd,
            'Banda Mín': fmt_usd, 'Banda Med': fmt_usd, 'Banda Máx': fmt_usd,
        }

        if agreement == 'Plus Fijo':
            fx_blue = vars.get('FX ARG Blue', 0)
            if fx_blue:
                peers = peers.copy()
                peers['bruto_total'] = peers['payroll'] + peers['bill'] * fx_blue
                peers_cols.append('bruto_total')
                peers_col_names.append('Bruto Total ARS')
                fmt_map['Bruto Total ARS'] = lambda v: f'${v:,.0f}' if v else '—'

        display_peers = peers[peers_cols].copy()
        display_peers.columns = peers_col_names

        def _highlight_self(row):
            color = '#fff9c4' if row['Nombre'] == emp_name else (
                '#f3e5f5' if row['Nivel'] == '➡ Siguiente' else ''
            )
            return [f'background-color: {color}'] * len(row)

        styled = (
            display_peers.style
            .apply(_highlight_self, axis=1)
            .format(fmt_map)
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # ── Bandas otra modalidad ──────────────────────────────────────────────────
    cross_label = 'Costo USD/H (Empleado)' if agreement == 'Contractor' else 'Contractor'
    # codes to compare: always current, plus target if recat
    cross_codes = [c for c in [new_code_empleado, next_code] if c]
    cross_rows = []
    for c in cross_codes:
        cb = get_cross_banda_for_code(c, agreement, bandas_df)
        if cb.get('cross_code') and any(cb.get(k) for k in ('banda_min', 'banda_med', 'banda_max')):
            cross_rows.append({
                'Code': cb['cross_code'],
                'Mín': cb['banda_min'],
                'Medio': cb['banda_med'],
                'Máx': cb['banda_max'],
            })
    if cross_rows:
        st.markdown(
            f'<div class="section-header">Bandas {cross_label} — mismos levels</div>',
            unsafe_allow_html=True,
        )
        cross_df = pd.DataFrame(cross_rows)
        st.dataframe(
            cross_df.style.format({'Mín': fmt_usd, 'Medio': fmt_usd, 'Máx': fmt_usd}),
            use_container_width=True, hide_index=True,
        )

    # ── Propuesta ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Propuesta</div>', unsafe_allow_html=True)

    tipo_default = TIPOS_CASO.index(caso.get('tipo', TIPOS_CASO[0])) if is_edit and caso.get('tipo') in TIPOS_CASO else 0
    tipo = st.selectbox('Tipo de resultado', TIPOS_CASO, index=tipo_default,
                        disabled=is_closed, key='form_tipo')

    is_recat = 'Recat' in tipo
    has_ajuste = tipo not in ('Solo feedback', 'Recat sin ajuste')

    # Target code for recat — siempre renderizar el selectbox para evitar
    # que Streamlit pierda las claves de posición de los widgets que siguen
    all_codes = [''] + bandas_df['new_code'].tolist()
    stored_tc = caso.get('target_code', '') if is_edit else ''
    default_tc = stored_tc if stored_tc in all_codes else (next_code if next_code else '')
    default_idx = all_codes.index(default_tc) if default_tc in all_codes else 0

    target_code = ''
    target_banda = {'banda_min': None, 'banda_med': None, 'banda_max': None}

    if is_recat:
        target_code = st.selectbox(
            'Nuevo code (recat target)', all_codes, index=default_idx,
            disabled=is_closed, key='form_target_code',
            help='Seleccioná el code al que se recategoriza al colaborador',
        )
        if target_code:
            target_banda = get_banda_for_code(target_code, bandas_df)
            bt1, bt2, bt3 = st.columns(3)
            bt1.metric('Banda Mín (target)', fmt_usd(target_banda['banda_min']))
            bt2.metric('Banda Med (target)', fmt_usd(target_banda['banda_med']))
            bt3.metric('Banda Máx (target)', fmt_usd(target_banda['banda_max']))

    # Bill / PayRoll / new costo inputs
    if not is_closed and has_ajuste:
        col_b, col_p = st.columns(2)
        with col_b:
            prop_bill = st.number_input(
                f'Nuevo Bill ({currency} / {per})',
                min_value=0.0,
                value=float(caso.get('proposed_bill', cur_bill) or cur_bill) if is_edit else cur_bill,
                step=50.0, format='%.2f',
                key='form_prop_bill',
            )
        with col_p:
            prop_payroll = st.number_input(
                'Nuevo PayRoll',
                min_value=0.0,
                value=float(caso.get('proposed_payroll', cur_payroll) or cur_payroll) if is_edit else cur_payroll,
                step=50.0, format='%.2f',
                key='form_prop_payroll',
            )

        # % variación propuesta vs actual
        pct_items = [(l, c, p) for l, c, p in [
            ('Bill', cur_bill, prop_bill), ('PayRoll', cur_payroll, prop_payroll)
        ] if c]
        if pct_items:
            pct_cols = st.columns(len(pct_items))
            for col_w, (label, cur_v, prop_v) in zip(pct_cols, pct_items):
                pct = (prop_v - cur_v) / cur_v * 100
                col_w.metric(f'Var. {label}', f'{pct:+.1f}%')

        new_costo = calc_costo_empresa(prop_bill, prop_payroll, per, agreement, currency, vars) or 0.0

    elif is_closed and has_ajuste:
        prop_bill = float(caso.get('proposed_bill', 0) or 0)
        prop_payroll = float(caso.get('proposed_payroll', 0) or 0)
        new_costo = float(caso.get('new_costo', cur_costo) or cur_costo)
        col_b, col_p, col_c = st.columns(3)
        col_b.metric('Nuevo Bill', fmt_usd(prop_bill))
        col_p.metric('Nuevo PayRoll', fmt_usd(prop_payroll))
        col_c.metric('Nuevo Costo USD/H', fmt_usd(new_costo))
        # % variación (cerrado, solo lectura)
        pct_items = [(l, c, p) for l, c, p in [
            ('Bill', cur_bill, prop_bill), ('PayRoll', cur_payroll, prop_payroll)
        ] if c]
        if pct_items:
            pct_cols = st.columns(len(pct_items))
            for col_w, (label, cur_v, prop_v) in zip(pct_cols, pct_items):
                pct = (prop_v - cur_v) / cur_v * 100
                col_w.metric(f'Var. {label}', f'{pct:+.1f}%')
    else:
        prop_bill = cur_bill
        prop_payroll = cur_payroll
        new_costo = cur_costo

    differential = new_costo - cur_costo

    st.markdown('<div class="section-header">Resultado calculado</div>', unsafe_allow_html=True)
    rc1, rc2, rc3 = st.columns(3)
    rc1.metric('Nuevo Costo USD/H', fmt_usd(new_costo))
    delta_sign = '+' if differential >= 0 else ''
    rc2.metric('Diferencial USD/H', f'{delta_sign}{differential:.4f}',
               delta_color='normal' if differential >= 0 else 'inverse')
    rc3.metric('Impacto mensual (×168)', fmt_usd(differential * HORAS, zero_dash=False))

    # ── Posición en banda ──────────────────────────────────────────────────────
    if banda_min or banda_med or banda_max:
        # Band for the proposal: target if recat, else current
        if is_recat and target_code and (target_banda.get('banda_med') or target_banda.get('banda_min')):
            prop_bmin = float(target_banda.get('banda_min') or 0)
            prop_bmed = float(target_banda.get('banda_med') or 0)
            prop_bmax = float(target_banda.get('banda_max') or 0)
            prop_band_suffix = f' ({target_code})'
        else:
            prop_bmin, prop_bmed, prop_bmax = banda_min, banda_med, banda_max
            prop_band_suffix = ''

        st.markdown('<div class="section-header">Posición en banda</div>', unsafe_allow_html=True)
        bars_html = _banda_position_html('Actual', cur_costo, banda_min, banda_med, banda_max)
        if has_ajuste:
            bars_html += _banda_position_html(
                f'Propuesta{prop_band_suffix}', new_costo, prop_bmin, prop_bmed, prop_bmax
            )
        st.markdown(
            f'<div style="padding:8px 16px;background:#fafafa;border-radius:8px;border:1px solid #e0e0e0">{bars_html}</div>',
            unsafe_allow_html=True,
        )

    if agreement == 'Plus Fijo' and has_ajuste:
        fx_blue = vars.get('FX ARG Blue', 0)
        if fx_blue:
            bruto_actual = cur_payroll + cur_bill * fx_blue
            bruto_prop = prop_payroll + prop_bill * fx_blue
            pct_bruto = (bruto_prop - bruto_actual) / bruto_actual * 100 if bruto_actual else 0
            st.markdown('<div class="section-header">Bruto total (Plus Fijo)</div>',
                        unsafe_allow_html=True)
            rb1, rb2, rb3 = st.columns(3)
            rb1.metric('Bruto actual', f'${bruto_actual:,.0f} ARS')
            rb2.metric('Bruto propuesto', f'${bruto_prop:,.0f} ARS')
            rb3.metric('Var. Bruto Total', f'{pct_bruto:+.1f}%')

    # Notes
    notes_val = caso.get('notes', '') if is_edit else ''
    notes = st.text_area('Notas / Comentarios', value=notes_val,
                         disabled=is_closed, height=80, key='form_notes')

    # ── Actions ───────────────────────────────────────────────────────────────
    if is_closed:
        if st.button('← Volver'):
            st.session_state.view = back_view
            st.rerun()
        return

    # Creator email persisted in session
    if 'user_email' not in st.session_state:
        st.session_state.user_email = ''
    st.session_state.user_email = st.text_input(
        'Tu email (para registro)', value=st.session_state.user_email,
        placeholder='tu@makingsense.com', key='form_user_email',
    )

    col_save, col_close, col_cancel = st.columns([1, 1, 3])

    with col_save:
        if st.button('💾 Guardar', type='primary', use_container_width=True):
            if not st.session_state.user_email:
                st.error('Ingresá tu email para guardar.')
            else:
                _save_caso(
                    is_edit, caso, emp_email, emp_name, code, seniority,
                    hire_date, xm, agreement, per, currency,
                    cur_bill, cur_payroll, cur_costo,
                    banda_min, banda_med, banda_max,
                    new_code_empleado, target_code, target_banda,
                    prop_bill, prop_payroll, new_costo, differential,
                    tipo, notes, st.session_state.user_email
                )

    with col_close:
        if is_edit:
            if st.button('🔒 Cerrar caso', use_container_width=True):
                if not st.session_state.user_email:
                    st.error('Ingresá tu email para cerrar.')
                else:
                    with st.spinner('Cerrando caso...'):
                        close_caso_sheet(caso['id'], st.session_state.user_email)
                    st.success('✅ Caso cerrado.')
                    st.session_state.view = 'open_list'
                    st.rerun()

    with col_cancel:
        if st.button('← Volver', use_container_width=False):
            st.session_state.view = 'open_list'
            st.rerun()


def _save_caso(is_edit, caso, emp_email, emp_name, code, seniority,
               hire_date, xm, agreement, per, currency,
               cur_bill, cur_payroll, cur_costo,
               banda_min, banda_med, banda_max,
               new_code_empleado, target_code, target_banda,
               prop_bill, prop_payroll, new_costo, differential,
               tipo, notes, creator):
    data = {
        'created_by': creator,
        'employee_email': emp_email,
        'employee_name': emp_name,
        'code': code,
        'seniority': seniority,
        'hire_date': hire_date,
        'xm': xm,
        'agreement': agreement,
        'per': per,
        'currency': currency,
        'current_bill': cur_bill,
        'current_payroll': cur_payroll,
        'current_costo': cur_costo,
        'banda_min': banda_min,
        'banda_med': banda_med,
        'banda_max': banda_max,
        'new_code_empleado': new_code_empleado,
        'target_code': target_code,
        'target_banda_min': target_banda.get('banda_min', '') or '',
        'target_banda_med': target_banda.get('banda_med', '') or '',
        'target_banda_max': target_banda.get('banda_max', '') or '',
        'proposed_bill': prop_bill,
        'proposed_payroll': prop_payroll,
        'new_costo': new_costo,
        'differential': differential,
        'tipo': tipo,
        'notes': notes,
    }
    with st.spinner('Guardando...'):
        if is_edit:
            data['id'] = caso['id']
            data['created_at'] = caso.get('created_at', '')
            data['status'] = STATUS_OPEN
            data['closed_at'] = ''
            data['closed_by'] = ''
            update_caso(caso['id'], data)
        else:
            create_caso(data)
    st.success('✅ Caso guardado.')
    st.session_state.view = 'open_list'
    st.rerun()


# ── Cases list components ──────────────────────────────────────────────────────

def show_open_list(casos_df: pd.DataFrame):
    open_df = casos_df[casos_df['status'] == STATUS_OPEN] if not casos_df.empty else pd.DataFrame()

    col_h, col_btn = st.columns([4, 1])
    col_h.markdown(f'#### Casos abiertos ({len(open_df)})')
    if not READ_ONLY:
        with col_btn:
            if st.button('➕ Nuevo caso', type='primary', use_container_width=True):
                st.session_state.view = 'new'
                st.rerun()

    if open_df.empty:
        st.info('No hay casos abiertos.')
        return

    for _, row in open_df.iterrows():
        diff = float(row.get('differential', 0) or 0)
        diff_str = f'{diff:+.2f} USD/H' if diff != 0 else 'Sin ajuste'
        c1, c2, c3, c4, c5 = st.columns([3, 2, 2, 2, 1])
        c1.markdown(f"**{row.get('employee_name', '—')}**")
        c2.write(row.get('code', '—'))
        c3.markdown(badge_html(row.get('tipo', '—')), unsafe_allow_html=True)
        c4.write(diff_str)
        with c5:
            if st.button('Ver →', key=f"open_{row['id']}"):
                st.session_state.editing_id = row['id']
                st.session_state.view = 'edit'
                st.rerun()
        st.divider()


def show_closed_list(casos_df: pd.DataFrame):
    closed_df = casos_df[casos_df['status'] == STATUS_CLOSED] if not casos_df.empty else pd.DataFrame()

    # Budget summary
    if not closed_df.empty:
        total_diff_h = closed_df['differential'].sum()
        total_budget_mensual = total_diff_h * HORAS
        st.markdown(
            f'<div class="budget-box">'
            f'<b>💰 Presupuesto ejecutado en cambios salariales</b><br>'
            f'Diferencial total: <b>{total_diff_h:+.4f} USD/H</b> &nbsp;·&nbsp; '
            f'Impacto mensual (×168): <b>${total_budget_mensual:,.0f} USD</b> — '
            f'{len(closed_df)} caso{"s" if len(closed_df) != 1 else ""} cerrado{"s" if len(closed_df) != 1 else ""}'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown(f'#### Casos cerrados ({len(closed_df)})')

    if closed_df.empty:
        st.info('No hay casos cerrados.')
        return

    for _, row in closed_df.iterrows():
        diff = float(row.get('differential', 0) or 0)
        diff_str = f'{diff:+.2f} USD/H' if diff != 0 else 'Sin ajuste'
        c1, c2, c3, c4, c5, c6 = st.columns([3, 2, 2, 2, 2, 1])
        c1.markdown(f"**{row.get('employee_name', '—')}**")
        c2.write(row.get('code', '—'))
        c3.markdown(badge_html(row.get('tipo', '—')), unsafe_allow_html=True)
        c4.write(diff_str)
        c5.write(row.get('closed_at', '—'))
        with c6:
            if st.button('Ver', key=f"closed_{row['id']}"):
                st.session_state.editing_id = row['id']
                st.session_state.view = 'view_closed'
                st.rerun()
        st.divider()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    st.title('🎯 Gestión Compensaciones')

    if 'view' not in st.session_state:
        st.session_state.view = 'open_list'
    if 'editing_id' not in st.session_state:
        st.session_state.editing_id = None

    with st.spinner('Cargando datos...'):
        try:
            empleados_df = load_empleados()
            bandas_df = load_bandas()
            casos_df = load_casos()
        except Exception as e:
            st.error(f'Error al cargar datos: {e}')
            if st.button('🔄 Reintentar'):
                st.cache_data.clear()
                st.cache_resource.clear()
                st.rerun()
            return

    view = st.session_state.view

    # ── Form views (full screen, no tabs) ─────────────────────────────────────
    if view in ('new', 'edit', 'view_closed'):
        if view == 'new' and not READ_ONLY:
            show_caso_form(empleados_df, bandas_df)
        elif view == 'new':
            st.session_state.view = 'open_list'
            st.rerun()
        else:
            caso_row = casos_df[casos_df['id'] == st.session_state.editing_id]
            if caso_row.empty:
                st.error('Caso no encontrado.')
                if st.button('← Volver'):
                    st.session_state.view = 'open_list'
                    st.rerun()
            else:
                show_caso_form(empleados_df, bandas_df, caso_row.iloc[0].to_dict(),
                               force_readonly=READ_ONLY)
        return

    # ── List views with tabs ───────────────────────────────────────────────────
    open_count = (casos_df['status'] == STATUS_OPEN).sum() if not casos_df.empty else 0
    closed_count = (casos_df['status'] == STATUS_CLOSED).sum() if not casos_df.empty else 0

    tab_open, tab_closed = st.tabs([
        f'📂 Abiertos ({open_count})',
        f'✅ Cerrados ({closed_count})',
    ])

    with tab_open:
        show_open_list(casos_df)

    with tab_closed:
        show_closed_list(casos_df)


if __name__ == '__main__':
    main()
