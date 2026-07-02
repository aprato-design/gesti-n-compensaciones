# -*- coding: utf-8 -*-
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from pathlib import Path

# ─── Config ────────────────────────────────────────────────────────────────────
SHEET_ID      = '1vBmF6PdU9ZCeB-6pqX5hxQ-FM0G3yWgd7MqtDfFlt78'
VARIABLES_TAB = 'Variables para Rep'
CREDS_FILE    = Path(__file__).parent.parent / 'talentserviceproject-1ce2ed91696b.json'

FACTOR_MEXICO      = 1.3500
FACTOR_PASANTE     = 1.1500
FACTOR_CONTRACTOR  = 1.0000
MULT_PLUS_FIJO_USD = 1.085

AGREEMENTS = ['Empleado', 'Empleado SEC', 'Pasante', 'Plus Fijo', 'Contractor']
LOCATIONS  = ['Argentina', 'Colombia', 'Mexico']

MONTH_NAMES = {
    1: 'Enero',   2: 'Febrero',    3: 'Marzo',      4: 'Abril',
    5: 'Mayo',    6: 'Junio',      7: 'Julio',       8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
}


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _to_float(val):
    if not val:
        return None
    try:
        return float(str(val).replace(',', '').replace('$', '').strip())
    except (ValueError, TypeError):
        return None


def _parse_month_num(date_str):
    """'5/1' → 5  (first token before '/')."""
    try:
        return int(str(date_str).split('/')[0])
    except Exception:
        return None


# ─── Data loading ──────────────────────────────────────────────────────────────

def _get_service():
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    try:
        info = dict(st.secrets['gcp_service_account'])
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    except Exception:
        creds = Credentials.from_service_account_file(str(CREDS_FILE), scopes=scopes)
    return build('sheets', 'v4', credentials=creds)


@st.cache_data(ttl=3600)
def load_variables():
    svc = _get_service()
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"'{VARIABLES_TAB}'!A1:AZ30",
        valueRenderOption='FORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])
    if len(rows) < 5:
        return None

    # Row 3 (index 2) contains the variable names
    header = rows[2]

    def col_of(name):
        for j, h in enumerate(header):
            if str(h).strip().lower() == name.lower():
                return j
        return None

    idx = {
        'fx_ars':        col_of('FX ARS OF'),
        'fx_cop':        col_of('FX COPs'),
        'fx_mx':         col_of('FX MX'),
        'costo_arg':     col_of('Costo ARG'),
        'costo_col_men': col_of('Costo COL menor'),
    }

    # Find latest data row: last row where FX ARS OF has a value (data from row 5, index 4)
    latest = None
    for row in rows[4:]:
        ars_idx = idx['fx_ars']
        if ars_idx is not None and len(row) > ars_idx and row[ars_idx]:
            latest = row

    if not latest:
        return None

    def v(key):
        i = idx[key]
        if i is None:
            return None
        try:
            return _to_float(latest[i])
        except IndexError:
            return None

    month_num  = _parse_month_num(latest[0]) if latest else None
    month_name = MONTH_NAMES.get(month_num, latest[0] if latest else '?')

    return {
        'FX_ARS_OF':     v('fx_ars'),
        'FX_COP':        v('fx_cop'),
        'FX_MX':         v('fx_mx'),
        'COSTO_ARG':     v('costo_arg'),
        'COSTO_COL_MEN': v('costo_col_men'),
        'mes':           month_name,
    }


# ─── Calculations ──────────────────────────────────────────────────────────────

def calcular_costo(agreement, location, salario, bill, fx_mx_manual, v):
    fx_ars     = v['FX_ARS_OF']
    fx_cop     = v['FX_COP']
    fx_mx      = v['FX_MX'] or fx_mx_manual
    factor_arg = v['COSTO_ARG']
    factor_col = v['COSTO_COL_MEN']

    if agreement == 'Empleado SEC':
        if not (fx_ars and factor_arg):
            return None
        return (salario * factor_arg) / fx_ars

    if agreement == 'Empleado':
        if location == 'Argentina':
            if not (fx_ars and factor_arg):
                return None
            return (salario * factor_arg) / fx_ars
        if location == 'Colombia':
            if not (fx_cop and factor_col):
                return None
            return (salario * factor_col) / fx_cop
        if location == 'Mexico':
            if not fx_mx:
                return None
            return (salario * 2 * FACTOR_MEXICO) / fx_mx

    if agreement == 'Pasante':
        if not fx_ars:
            return None
        return (salario * FACTOR_PASANTE) / fx_ars

    if agreement == 'Contractor':
        return bill * FACTOR_CONTRACTOR

    if agreement == 'Plus Fijo':
        if not (fx_ars and factor_arg):
            return None
        return (salario * factor_arg / fx_ars) + bill * MULT_PLUS_FIJO_USD

    return None


# ─── Detail breakdown ──────────────────────────────────────────────────────────

def _show_detail(agreement, location, salario, bill, fx_mx_manual, v, costo):
    loc = location if agreement == 'Empleado' else (
          'Argentina' if agreement == 'Empleado SEC' else None)

    if agreement in ('Empleado', 'Empleado SEC'):
        if loc == 'Argentina':
            comp = salario * v['COSTO_ARG'] / v['FX_ARS_OF']
            rows = [
                ('Salario bruto', f'ARS {salario:,.2f}'),
                (f'Factor de costo ({v["mes"]})', f'{v["COSTO_ARG"]:.4f}'),
                ('FX ARS OF', f'${v["FX_ARS_OF"]:,.0f}'),
                ('Fórmula', f'({salario:,.2f} × {v["COSTO_ARG"]:.4f}) / {v["FX_ARS_OF"]:,.0f} = **USD {comp:,.2f}**'),
            ]
        elif loc == 'Colombia':
            comp = salario * v['COSTO_COL_MEN'] / v['FX_COP']
            rows = [
                ('Salario bruto', f'COP {salario:,.0f}'),
                (f'Factor de costo ({v["mes"]})', f'{v["COSTO_COL_MEN"]:.4f}'),
                ('FX COPs', f'${v["FX_COP"]:,.2f}'),
                ('Fórmula', f'({salario:,.0f} × {v["COSTO_COL_MEN"]:.4f}) / {v["FX_COP"]:,.2f} = **USD {comp:,.2f}**'),
            ]
        elif loc == 'Mexico':
            fx_mx = v['FX_MX'] or fx_mx_manual
            comp = salario * 2 * FACTOR_MEXICO / fx_mx
            rows = [
                ('Salario bruto', f'MXN {salario:,.2f}'),
                ('Factor de costo (fijo)', f'{FACTOR_MEXICO:.4f}'),
                ('FX MXN', f'${fx_mx:,.2f}'),
                ('Fórmula', f'({salario:,.2f} × 2 × {FACTOR_MEXICO:.4f}) / {fx_mx:,.2f} = **USD {comp:,.2f}**'),
            ]
        else:
            return

    elif agreement == 'Pasante':
        rows = [
            ('Salario bruto', f'ARS {salario:,.2f}'),
            ('Factor de costo (fijo)', f'{FACTOR_PASANTE:.4f}'),
            ('FX ARS OF', f'${v["FX_ARS_OF"]:,.0f}'),
            ('Fórmula', f'({salario:,.2f} × {FACTOR_PASANTE:.4f}) / {v["FX_ARS_OF"]:,.0f} = **USD {costo:,.2f}**'),
        ]

    elif agreement == 'Contractor':
        rows = [
            ('Bill mensual', f'USD {bill:,.2f}'),
            ('Factor de costo (fijo)', f'{FACTOR_CONTRACTOR:.4f}'),
            ('Fórmula', f'{bill:,.2f} × {FACTOR_CONTRACTOR:.4f} = **USD {costo:,.2f}**'),
        ]

    elif agreement == 'Plus Fijo':
        comp_ars = salario * v['COSTO_ARG'] / v['FX_ARS_OF']
        comp_usd = bill * MULT_PLUS_FIJO_USD
        rows = [
            ('Salario bruto (ARS)', f'ARS {salario:,.2f}'),
            ('Bill (USD)', f'USD {bill:,.2f}'),
            (f'Factor de costo ({v["mes"]})', f'{v["COSTO_ARG"]:.4f}'),
            ('FX ARS OF', f'${v["FX_ARS_OF"]:,.0f}'),
            ('Componente ARS', f'{salario:,.2f} × {v["COSTO_ARG"]:.4f} / {v["FX_ARS_OF"]:,.0f} = USD {comp_ars:,.2f}'),
            ('Componente USD', f'{bill:,.2f} × {MULT_PLUS_FIJO_USD} = USD {comp_usd:,.2f}'),
            ('Total', f'USD {comp_ars:,.2f} + USD {comp_usd:,.2f} = **USD {costo:,.2f}**'),
        ]
    else:
        return

    for label, val in rows:
        st.write(f'- **{label}:** {val}')


# ─── Brand CSS ─────────────────────────────────────────────────────────────────

def _inject_css():
    st.markdown("""
    <style>
    :root {
        --dp-yellow:       #FFDD00;
        --dp-yellow-dark:  #C8A500;
        --dp-yellow-light: #FFFBDE;
        --dp-text:         #525845;
        --dp-text-light:   #7A8070;
        --dp-border:       #F0E070;
        --dp-text-on-yel:  #2C2500;
    }
    html, body, [class*="css"] { color: var(--dp-text); }
    h1, h2, h3 { color: var(--dp-text) !important; font-weight: 700 !important; }

    .dp-header {
        background: linear-gradient(135deg, #FFDD00 0%, #E5C000 100%);
        padding: 1.1rem 2rem;
        margin-bottom: 1.5rem;
        border-radius: 0 0 12px 12px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .dp-logo {
        font-size: 1.4rem; font-weight: 900; color: var(--dp-text-on-yel);
        letter-spacing: -0.03em; text-transform: uppercase;
    }
    .dp-logo span {
        opacity: 0.6; font-weight: 400; font-size: 0.7rem;
        letter-spacing: 0.12em; margin-left: 0.6rem; vertical-align: middle;
    }
    .dp-header-right { color: rgba(44,37,0,0.7); font-size: 0.85rem; }

    [data-testid="stButton"] > button {
        background-color: var(--dp-yellow) !important;
        color: var(--dp-text-on-yel) !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
    }
    [data-testid="stButton"] > button:hover {
        background-color: var(--dp-yellow-dark) !important;
        color: white !important;
    }

    hr { border-color: var(--dp-border) !important; margin: 1rem 0 !important; }
    #MainMenu, footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)


def _header():
    st.markdown("""
    <div class="dp-header">
        <div class="dp-logo">Doppler<span>Talent Care</span></div>
        <div class="dp-header-right">Calculadora · Costo Mensual</div>
    </div>
    """, unsafe_allow_html=True)


# ─── Main app ──────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title='Costo Mensual – Doppler',
        page_icon='📊',
        layout='centered',
    )
    _inject_css()
    _header()

    with st.spinner('Cargando variables...'):
        vars_ = load_variables()

    if not vars_:
        st.error('No se pudieron cargar las variables del sheet. Verificar credenciales.')
        st.stop()

    # ── Period info banner ────────────────────────────────────────────────────
    info_parts = [f"Datos de **{vars_['mes']}**"]
    if vars_['COSTO_ARG']:
        info_parts.append(f"Factor AR: **{vars_['COSTO_ARG']:.4f}**")
    if vars_['COSTO_COL_MEN']:
        info_parts.append(f"Factor CO: **{vars_['COSTO_COL_MEN']:.4f}**")
    if vars_['FX_ARS_OF']:
        info_parts.append(f"FX ARS: **\\${vars_['FX_ARS_OF']:,.0f}**")
    if vars_['FX_COP']:
        info_parts.append(f"FX COP: **\\${vars_['FX_COP']:,.2f}**")
    if vars_['FX_MX']:
        info_parts.append(f"FX MX: **\\${vars_['FX_MX']:,.2f}**")
    st.info('  ·  '.join(info_parts))

    st.divider()

    # ── Agreement & Location ──────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    agreement = c1.selectbox('Agreement', AGREEMENTS)

    location = None
    if agreement == 'Empleado':
        location = c2.selectbox('Location', LOCATIONS)

    # FX MX fallback when not yet in the sheet
    fx_mx_manual = None
    needs_fx_mx = agreement == 'Empleado' and location == 'Mexico'
    if needs_fx_mx and not vars_.get('FX_MX'):
        st.warning('FX MXN todavía no está en el sheet de variables. Ingresá el valor manualmente.')
        fx_mx_manual = st.number_input(
            'FX MXN / USD', min_value=1.0, value=17.50, step=0.01, format='%.2f')

    # ── Salary / Bill inputs ──────────────────────────────────────────────────
    salario = 0.0
    bill    = 0.0

    if agreement == 'Contractor':
        bill = st.number_input(
            'Bill mensual (USD)', min_value=0.0, value=0.0,
            step=100.0, format='%.2f')

    elif agreement == 'Plus Fijo':
        ca, cb = st.columns(2)
        salario = ca.number_input(
            'Salario bruto (ARS)', min_value=0.0, value=0.0,
            step=10_000.0, format='%.2f')
        bill = cb.number_input(
            'Bill (USD)', min_value=0.0, value=0.0,
            step=100.0, format='%.2f')

    elif agreement == 'Pasante':
        salario = st.number_input(
            'Salario bruto mensual (ARS)', min_value=0.0, value=0.0,
            step=10_000.0, format='%.2f')

    elif agreement == 'Empleado' and location == 'Colombia':
        salario = st.number_input(
            'Salario bruto mensual (COP)', min_value=0.0, value=0.0,
            step=100_000.0, format='%.0f')

    elif agreement == 'Empleado' and location == 'Mexico':
        salario = st.number_input(
            'Salario bruto mensual (MXN)', min_value=0.0, value=0.0,
            step=1_000.0, format='%.2f')

    else:  # Empleado Argentina, Empleado SEC
        salario = st.number_input(
            'Salario bruto mensual (ARS)', min_value=0.0, value=0.0,
            step=10_000.0, format='%.2f')

    # ── Result ────────────────────────────────────────────────────────────────
    has_input = salario > 0 or bill > 0
    if has_input:
        costo = calcular_costo(agreement, location, salario, bill, fx_mx_manual, vars_)

        st.divider()

        if costo is None:
            st.error('No se pudo calcular. Verificar que FX MX esté disponible.')
        else:
            st.metric(label='Costo mensual', value=f'USD {costo:,.2f}')

            with st.expander('Ver detalle del cálculo'):
                _show_detail(agreement, location, salario, bill, fx_mx_manual, vars_, costo)


if __name__ == '__main__':
    main()
