# -*- coding: utf-8 -*-
"""
App Streamlit: Evolución de Costo Mensual.

A diferencia de generar_evolucion_costo.py (que escribe un HTML estático con
los datos embebidos), esta app lee los datos en vivo desde Google Sheets en
cada carga y los renderiza en memoria — nunca quedan guardados en el repo.
"""

import json
import streamlit as st

from generar_evolucion_costo import (
    _service, read_old_sheet, read_comp2021, read_comp2022,
    read_comp2023, read_comp2024, read_comp2025, read_bdd, combinar,
    HTML_TEMPLATE,
)

st.set_page_config(page_title="Evolución de Costo Mensual", layout="wide")


@st.cache_data(ttl=3600, show_spinner="Cargando datos de compensaciones...")
def cargar_personas():
    svc = _service()
    old_data    = read_old_sheet(svc)
    comp21_data = read_comp2021(svc)
    comp22_data = read_comp2022(svc)
    comp23_data = read_comp2023(svc)
    comp24_data = read_comp2024(svc)
    comp25_data = read_comp2025(svc)
    (bdd_data, nombres, doppler_emails, divisiones,
     countries_map, seniorities_map, departments_map) = read_bdd(svc)
    return combinar(
        old_data, comp21_data, comp22_data, comp23_data, comp24_data,
        comp25_data, bdd_data, nombres, doppler_emails, divisiones,
        countries_map, seniorities_map, departments_map,
    )


def construir_html(personas):
    codes = sorted({d['code'] for p in personas for d in p['datos'] if d.get('code')})
    agreements = sorted({d['agreement'] for p in personas for d in p['datos'] if d.get('agreement')})
    divisions = sorted({p['division'] for p in personas if p.get('division')})
    countries = sorted({p['country'] for p in personas if p.get('country')})
    levels = sorted({p['level'] for p in personas if p.get('level')})
    departments = sorted({p['department'] for p in personas if p.get('department')})
    payload = {
        'personas': personas,
        'codes': codes,
        'agreements': agreements,
        'divisions': divisions,
        'countries': countries,
        'levels': levels,
        'departments': departments,
    }
    data_json = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    return HTML_TEMPLATE.replace('__DATA__', data_json)


personas = cargar_personas()
html = construir_html(personas)
st.iframe(html, height=1600)
