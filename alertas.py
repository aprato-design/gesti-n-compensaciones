# -*- coding: utf-8 -*-
"""
Alertas por correo cuando falla un script programado.

El correo lo manda un Apps Script publicado como app web (alertas_mail.gs),
desde la cuenta de quien lo publicó. La URL y el token se leen de
alertas_config.json (no versionado):

    {"webhook_url": "https://script.google.com/macros/s/.../exec",
     "token": "...", "to": "aprato@makingsense.com"}

Uso:
    from alertas import run_with_alert
    run_with_alert(main, 'nombre del proceso')
"""

import os, sys, json, socket, datetime, traceback

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'alertas_config.json')


def send_alert(subject, body):
    """Envía el correo de alerta. Nunca levanta excepción (devuelve True/False)."""
    try:
        import requests
        with open(CONFIG_FILE, encoding='utf-8') as f:
            cfg = json.load(f)
        r = requests.post(cfg['webhook_url'], timeout=60, json={
            'token': cfg['token'], 'to': cfg.get('to', ''),
            'subject': subject, 'body': body,
        })
        ok = r.ok and 'OK' in r.text
        if not ok:
            print(f'Advertencia: la alerta por correo no se envió ({r.status_code}): {r.text[:200]}')
        return ok
    except Exception as e:
        print(f'Advertencia: no se pudo enviar la alerta por correo: {e}')
        return False


def run_with_alert(main, proceso):
    """Corre main(); si falla, manda el correo con el error y sale con código 1."""
    try:
        main()
    except Exception:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)
        ahora = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
        send_alert(
            f'[BDD Bamboo] Falló {proceso} – {ahora}',
            f'El proceso "{proceso}" falló el {ahora} en {socket.gethostname()}.\n'
            f'No se escribieron (o se escribieron parcialmente) los datos de esta corrida.\n\n'
            f'Error:\n{tb}',
        )
        sys.exit(1)
