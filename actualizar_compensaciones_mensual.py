# -*- coding: utf-8 -*-
"""
Calcula y escribe los campos de compensación en la Base de Datos Bamboo
para el mes anterior (mes vencido).

Se ejecuta automáticamente a partir del 3er día hábil de cada mes
usando el calendario de días hábiles de Argentina.

Uso:
    python actualizar_compensaciones_mensual.py [--mes YYYY-MM] [--force] [--dry-run]

    --mes YYYY-MM  Mes a procesar (default: mes anterior)
    --force        Omite la validación del día hábil
    --dry-run      Muestra qué haría sin escribir nada
"""

import sys, os, datetime
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

# ─── DÍAS HÁBILES ARGENTINA ──────────────────────────────────────────────────
# Actualizar la lista al inicio de cada año.
# Incluye feriados nacionales inamovibles, feriados puente y días no laborables.
# Los sábados y domingos ya se excluyen por lógica de weekday.

ARGENTINA_HOLIDAYS = {
    2026: {
        datetime.date(2026,  1,  1),  # Año Nuevo
        datetime.date(2026,  2, 16),  # Carnaval lunes
        datetime.date(2026,  2, 17),  # Carnaval martes
        datetime.date(2026,  3, 24),  # Día de la Memoria
        datetime.date(2026,  4,  2),  # Día del Veterano
        datetime.date(2026,  4,  3),  # Viernes Santo
        datetime.date(2026,  5,  1),  # Día del Trabajador
        datetime.date(2026,  5, 25),  # Revolución de Mayo
        datetime.date(2026,  7,  9),  # Día de la Independencia
        datetime.date(2026,  8, 17),  # Paso a la Inmortalidad del Gral. San Martín
        datetime.date(2026, 10, 12),  # Día del Respeto a la Diversidad Cultural
        datetime.date(2026, 11, 20),  # Día de la Soberanía Nacional
        datetime.date(2026, 12,  8),  # Inmaculada Concepción
        datetime.date(2026, 12, 25),  # Navidad
    },
    2027: {
        datetime.date(2027,  1,  1),  # Año Nuevo
        datetime.date(2027,  2,  8),  # Carnaval lunes
        datetime.date(2027,  2,  9),  # Carnaval martes
        datetime.date(2027,  3, 24),  # Día de la Memoria
        datetime.date(2027,  3, 26),  # Viernes Santo
        datetime.date(2027,  4,  2),  # Día del Veterano
        datetime.date(2027,  5, 25),  # Revolución de Mayo
        datetime.date(2027,  7,  9),  # Día de la Independencia
        datetime.date(2027,  8, 16),  # San Martín (3er lunes de agosto)
        datetime.date(2027, 10, 11),  # Diversidad Cultural (2do lunes de octubre)
        datetime.date(2027, 12,  8),  # Inmaculada Concepción
        datetime.date(2027, 12, 25),  # Navidad
    },
}


def _is_arg_business_day(d: datetime.date) -> bool:
    if d.weekday() >= 5:
        return False
    return d not in ARGENTINA_HOLIDAYS.get(d.year, set())


def _arg_business_day_of_month(today: datetime.date = None) -> int:
    """Retorna el número de día hábil (Argentina) del mes hasta hoy inclusive."""
    today = today or datetime.date.today()
    count = 0
    d = today.replace(day=1)
    while d <= today:
        if _is_arg_business_day(d):
            count += 1
        d += datetime.timedelta(days=1)
    return count


def _prev_month_str() -> str:
    """Retorna el primer día del mes anterior como 'YYYY-MM-01'."""
    first_of_this = datetime.date.today().replace(day=1)
    last_of_prev  = first_of_this - datetime.timedelta(days=1)
    return last_of_prev.strftime('%Y-%m-01')


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    dry_run = '--dry-run' in sys.argv
    force   = '--force'   in sys.argv

    # Determinar mes objetivo
    target_month = None
    for arg in sys.argv[1:]:
        if arg.startswith('--mes'):
            parts = arg.split('=', 1)
            raw = parts[1].strip() if len(parts) == 2 else (
                sys.argv[sys.argv.index(arg) + 1]
                if sys.argv.index(arg) + 1 < len(sys.argv) else None)
            if raw:
                target_month = raw[:7] + '-01'   # YYYY-MM → YYYY-MM-01
    if not target_month:
        target_month = _prev_month_str()

    today   = datetime.date.today()
    biz_day = _arg_business_day_of_month(today)

    print(f'\n{"[DRY-RUN] " if dry_run else ""}actualizar_compensaciones_mensual.py')
    print(f'  Fecha: {today}  |  Día hábil AR: {biz_day}  |  Mes a procesar: {target_month}\n')

    if not force and biz_day < 3:
        print(f'Hoy es el día hábil {biz_day} del mes (AR). '
              f'Las compensaciones se calculan a partir del 3er día hábil.')
        print('Usá --force para ejecutar igualmente.')
        return

    # Importar funciones del script diario
    from actualizar_demografia import (
        _service, fetch_all_bamboo, load_historico,
        load_demo_lookup, rebuild_actualizado, update_compensaciones,
        _load_doppler_seniority_map, _load_td_project_map,
    )

    svc = _service()

    # Mapas que usa rebuild_actualizado (vía _apply_daily_cols); sin cargarlos
    # Actualizado pierde el Project de Talent Discussion y el Seniority Doppler.
    _load_doppler_seniority_map(svc)
    _load_td_project_map(svc)

    print('Descargando directorio de BambooHR...')
    all_emps = fetch_all_bamboo()
    bamboo_lookup = {
        str(e.get('workEmail') or '').strip().lower(): e
        for e in all_emps if e.get('workEmail')
    }

    print('Leyendo Histórico...')
    _, data = load_historico(svc)

    # Calcular y escribir compensaciones para el mes objetivo
    update_compensaciones(svc, data, bamboo_lookup, target_month, dry_run)

    # Reconstruir Actualizado para reflejar los cambios
    print('\nReconstruyendo Actualizado...')
    demo_lookup = load_demo_lookup(svc)
    rebuild_actualizado(svc, bamboo_lookup, demo_lookup, dry_run)

    print('\nListo.')


if __name__ == '__main__':
    from alertas import run_with_alert
    run_with_alert(main, 'actualizar_compensaciones_mensual.py (BDD Bamboo mensual)')
