"""Lenkovich, Nelson: New Code fijo OM 05-ARG (Bamboo tiene FinSsrAdv) + banda/GAP may-oct 2026.
Mar-Oct: banda vigente OM 05-ARG de Bandas Div (250 / 2.000.000 / 700), que segun el usuario rige
desde marzo 2026, calculada con las variables de cada mes. (Una primera pasada estimaba May-Jul con
bandas viejas: corregido 19/11 por indicacion del usuario.)
Uso: python _fix_lenkovich_om05.py [--apply]"""
import sys, json, datetime, warnings; warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2 import service_account
from googleapiclient.discovery import build
from generar_compensaciones import calc_gap_banda, calc_pct_gap_banda, calc_dispersion_en_banda
SID = '15mTvxuwXGa8B8Cuc86qqDyb8eKhKrDlwRekx9I1dZ_0'; TAB = 'Bamboo Compensaciones - Sueldos'
creds = service_account.Credentials.from_service_account_file('talentserviceproject-1ce2ed91696b.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
svc = build('sheets', 'v4', credentials=creds)
rows = svc.spreadsheets().values().get(spreadsheetId=SID, range=f"'{TAB}'!A1:AL20000", valueRenderOption='UNFORMATTED_VALUE').execute()['values']
# Variables por mes (FX ARS, Costo ARG, Costo ARG USD) de 'Variables para Rep'
V = {46082: (1373.0, 1.5105), 46113: (1382.0, 1.5123), 46143: (1399.0, 1.5072), 46174: (1482.0, 1.5093), 46204: (1476.0, 1.5040),
     46235: (1499.5, 1.5023), 46266: (1508.0, 1.5058), 46296: (1525.0, 1.5093)}
CAU = 1.085
def banda(m):
    fx, ca = V[m]; base = 2000000*ca/fx
    lo, hi = (base + 250*CAU)/168, (base + 700*CAU)/168
    return round(lo, 4), round((lo+hi)/2, 4), round(hi, 4)
plan = []
for i, row in enumerate(rows[1:], start=2):
    if len(row) > 24 and row[1] == 'Lenkovich, Nelson' and row[0] in V:
        assert row[8] == 'FinSsrAdv' and row[10] == 'Plus Fijo', row
        bmin, bmed, bmax = banda(row[0]); costo = row[23]
        gap = calc_gap_banda(costo, bmin, bmax)
        new = {'Y': 'OM 05-ARG', 'Z': bmin, 'AA': bmed, 'AB': bmax, 'AC': gap,
               'AG': calc_pct_gap_banda(gap), 'AI': round(calc_dispersion_en_banda(costo, bmin, bmax), 6)}
        old = {'Y': row[24], 'Z': row[25], 'AA': row[26], 'AB': row[27], 'AC': row[28],
               'AG': row[32] if len(row) > 32 else '', 'AI': row[34] if len(row) > 34 else ''}
        plan.append((i, row[0], costo, old, new))
assert len(plan) == 8, len(plan)
for i, m, c, o, n in sorted(plan, key=lambda x: x[1]):
    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=m)
    g = n['AC'] if n['AC'] == 'ok' else f"{n['AC']*100:+.1f}%"
    print(f"{d:%b-%Y} fila {i}: costo {c:.2f} | banda {n['Z']:.2f}-{n['AB']:.2f} | GAP {g} | {n['AG']}")
if '--apply' in sys.argv:
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    json.dump([(i, m, o) for i, m, c, o, n in plan], open(f'_backup_lenkovich_om05_{ts}.json', 'w'), ensure_ascii=False, indent=1)
    data = [{'range': f"'{TAB}'!{col}{i}", 'values': [[val]]} for i, m, c, o, n in plan for col, val in n.items()]
    res = svc.spreadsheets().values().batchUpdate(spreadsheetId=SID, body={'valueInputOption': 'RAW', 'data': data}).execute()
    print('Escritas', res.get('totalUpdatedCells'), 'celdas. Backup', ts)
