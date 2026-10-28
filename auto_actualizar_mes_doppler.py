# -*- coding: utf-8 -*-
"""
Wrapper de ejecución automática mensual para Doppler.

Windows Task Scheduler lo llama cada día a las 09:30.
El script verifica si hoy es el 3er día hábil del mes (feriados Argentina).
Si lo es — y no corrió ya este mes — ejecuta actualizar_mes_sheets_doppler.py
con el mes anterior como argumento.

Ventana de gracia: hasta el 6to día hábil inclusive (por si el equipo estuvo apagado).

Log: auto_actualizacion_doppler.log
"""

import sys
import os
import datetime
import subprocess
import logging

try:
    import holidays as holidays_lib
    def _ar_holidays(year):
        return holidays_lib.Argentina(years=year)
except ImportError:
    def _ar_holidays(year):
        return {}

WORK_DIR   = os.path.dirname(os.path.abspath(__file__))
SCRIPT     = os.path.join(WORK_DIR, 'actualizar_mes_sheets_doppler.py')
LOG_FILE   = os.path.join(WORK_DIR, 'auto_actualizacion_doppler.log')
STATE_FILE = os.path.join(WORK_DIR, 'auto_actualizacion_doppler_estado.txt')

MONTH_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo',    4: 'Abril',
    5: 'Mayo',  6: 'Junio',   7: 'Julio',    8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
}

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s  %(levelname)-8s  %(message)s',
    encoding='utf-8',
)


def is_business_day(d, ar_hol):
    return d.weekday() < 5 and d not in ar_hol


def nth_business_day(year, month, n, ar_hol):
    import calendar
    count = 0
    day = datetime.date(year, month, 1)
    last = calendar.monthrange(year, month)[1]
    while day.day <= last:
        if is_business_day(day, ar_hol):
            count += 1
            if count == n:
                return day
        day += datetime.timedelta(days=1)
    return None


def already_ran_this_month(year, month):
    if not os.path.exists(STATE_FILE):
        return False
    with open(STATE_FILE, encoding='utf-8') as f:
        content = f.read().strip()
    return content == f'{year}-{month:02d}'


def mark_ran(year, month):
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        f.write(f'{year}-{month:02d}')


def prev_month(year, month):
    if month == 1:
        return year - 1, 12
    return year, month - 1


def main():
    today = datetime.date.today()
    year, month = today.year, today.month

    logging.info(f'Verificando Doppler: hoy={today}')

    ar_hol   = _ar_holidays(year)
    third_bd = nth_business_day(year, month, 3, ar_hol)
    sixth_bd = nth_business_day(year, month, 6, ar_hol)

    if third_bd is None:
        logging.warning('No se pudo calcular el 3er día hábil. Saliendo.')
        return

    logging.info(f'3er día hábil: {third_bd}  |  6to: {sixth_bd}')

    if not (third_bd <= today <= (sixth_bd or today)):
        logging.info('Fuera de ventana de ejecución Doppler. Sin acción.')
        return

    if already_ran_this_month(year, month):
        logging.info(f'Ya se ejecutó Doppler para {year}-{month:02d}. Sin acción.')
        return

    proc_year, proc_month = prev_month(year, month)
    mes_arg = MONTH_ES[proc_month]
    ano_arg = str(proc_year)

    logging.info(f'Iniciando actualización Doppler para {mes_arg} {ano_arg}...')
    print(f'[auto_actualizar_mes_doppler] Ejecutando: {mes_arg} {ano_arg}')

    result = subprocess.run(
        [sys.executable, SCRIPT, mes_arg, ano_arg],
        capture_output=True,
        text=True,
        encoding='utf-8',
        cwd=WORK_DIR,
    )

    if result.stdout:
        logging.info('STDOUT:\n' + result.stdout)
    if result.stderr:
        logging.warning('STDERR:\n' + result.stderr)

    if result.returncode == 0:
        mark_ran(year, month)
        logging.info(f'Actualización Doppler completada. Return code: {result.returncode}')
        print(f'[auto_actualizar_mes_doppler] Completado OK.')
    else:
        logging.error(f'Error en script Doppler. Return code: {result.returncode}')
        print(f'[auto_actualizar_mes_doppler] ERROR — revisar {LOG_FILE}')
        # El error del script interno se re-lanza para que run_with_alert mande
        # el correo y la tarea programada termine con código 1.
        raise RuntimeError(
            f'{SCRIPT} {mes_arg} {ano_arg} terminó con código {result.returncode}. '
            f'Log: {LOG_FILE}' + chr(10) * 2 + 'Últimas líneas de STDERR:' + chr(10)
            + chr(10).join((result.stderr or '').splitlines()[-40:]))


if __name__ == '__main__':
    from alertas import run_with_alert
    run_with_alert(main, 'Doppler Compensaciones Mensual (auto_actualizar_mes_doppler.py)')
