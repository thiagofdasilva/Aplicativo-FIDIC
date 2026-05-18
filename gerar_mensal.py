"""
Gera dados mensais por FIDC com status de pagamento do CSV ACS Faturas.
Chave de cruzamento: str(NOTA_FISCAL).lstrip('0') + str(PARCELA).zfill(3)
"""
import openpyxl, json, csv, io
from datetime import datetime
from collections import defaultdict

# ─── CSV ACS FATURAS ──────────────────────────────────────────────────────────
with open('ACS Faturas.csv', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

reader = csv.DictReader(io.StringIO(content))
fn = reader.fieldnames
col_nf, col_par, col_val, col_rst, col_sts = fn[1], fn[3], fn[6], fn[7], fn[17]

csv_lookup = {}   # key -> {'status', 'valor', 'valor_restante'}
for r in reader:
    nf  = str(r[col_nf]).lstrip('0')
    par = str(r[col_par]).zfill(3)
    key = nf + par
    try:    valor = float(r[col_val])
    except: valor = 0.0
    try:    restante = float(r[col_rst])
    except: restante = 0.0
    csv_lookup[key] = {
        'status':   r[col_sts],
        'valor':    valor,
        'restante': restante,
    }

print(f'CSV carregado: {len(csv_lookup):,} registros')
print(f'Status únicos: {set(v["status"] for v in csv_lookup.values())}')

# ─── EXCEL ────────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(
    'CONTROLE BOLETOS OPERAÇÃO FIDC - BANCOS - 2025.xlsx', data_only=True
)

SHEET_NAMES = {
    'DAYCOVAL':'DAYCOVAL','SAFRA':'SAFRA','MULTIPLIKE':'MULTIPLIKE','C6':'C6',
    'CREDVALE':'CREDVALE','ALL SEC':'ALL SEC','AUDAX':'AUDAX','KANASTRA':'KANASTRA',
    'YAALEH':'YAALEH','ONE 7':'ONE 7','JUMP':'JUMP','OPHIR':'OPHIR',
    'KREDITON':'KREDITON','ASIA':'ASIA','SICOOB':'SICOOB','RED ASSET':'RED ASSET',
    'MAIN3':'MAIN3','MANCHESTER':'MANCHESTER','ASA':'ASA',
}
ARTICO = next((s for s in wb.sheetnames if 'RTICO' in s), None)

def process(ws, display_name):
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return {}
    header = rows[0]
    hc = [str(h).strip().upper() if h else '' for h in header]

    val_idx  = next((i for i,h in enumerate(hc) if 'VALOR' in h and 'R$' in h), None)
    nf_idx   = next((i for i,h in enumerate(hc) if 'NOTA' in h), None)
    par_idx  = next((i for i,h in enumerate(hc) if 'PARCELA' in h), None)
    data_idx = next((i for i,h in enumerate(hc) if h == 'DATA'), None)
    venc_idx = next((i for i,h in enumerate(hc) if 'VENCIMENTO' in h), None)
    desagio_idx = next((i for i,h in enumerate(hc) if 'DESAGIO' in h), None)

    if val_idx is None or data_idx is None:
        print(f'  AVISO: colunas não encontradas em {display_name}')
        return {}

    monthly = defaultdict(lambda: {
        'valor': 0.0, 'desagio': 0.0, 'volumetria': 0, 'prazos': [],
        'v_pago': 0.0, 'v_aberto': 0.0, 'v_sem_info': 0.0,
        'n_pago': 0,   'n_aberto': 0,  'n_sem_info': 0,
    })
    total_check = 0.0
    matched = no_match = 0

    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        d = row[data_idx]
        if not isinstance(d, datetime):
            continue
        try:
            v = float(row[val_idx]) if row[val_idx] is not None else 0.0
        except (TypeError, ValueError):
            v = 0.0
        if v <= 0:
            continue

        # Chave de cruzamento: NF + PARCELA.zfill(3)
        try:
            nf_raw = str(row[nf_idx]).lstrip('0') if nf_idx is not None and row[nf_idx] is not None else ''
            par_raw = str(int(float(str(row[par_idx])))).zfill(3) if par_idx is not None and row[par_idx] is not None else '001'
            key = nf_raw + par_raw
        except:
            key = ''

        info = csv_lookup.get(key)
        status = info['status'] if info else 'Sem registro'

        ym = d.strftime('%Y-%m')
        m  = monthly[ym]
        m['valor']      += v
        m['volumetria'] += 1
        total_check     += v

        # Deságio
        if desagio_idx is not None and row[desagio_idx] is not None:
            try:
                des = float(row[desagio_idx])
                if 0 < des < v:
                    m['desagio'] += des
            except (TypeError, ValueError):
                pass

        # Prazo
        if venc_idx is not None and row[venc_idx] is not None:
            vd = row[venc_idx]
            if isinstance(vd, datetime):
                p = (vd - d).days
                if 0 < p < 1000:
                    m['prazos'].append(p)

        # Status
        if status == 'Pago':
            m['v_pago']  += v
            m['n_pago']  += 1
            matched += 1
        elif status in ('Não pago', 'N�о pago', 'Nao pago') or (info and 'pago' in status.lower() and ('n' in status[:2].lower())):
            m['v_aberto']  += v
            m['n_aberto']  += 1
            matched += 1
        elif status == 'Parcialmente pago':
            m['v_aberto']  += v   # tratar como aberto
            m['n_aberto']  += 1
            matched += 1
        else:
            m['v_sem_info'] += v
            m['n_sem_info'] += 1
            no_match += 1

    # Finalizar meses
    result = {}
    for ym, data in sorted(monthly.items()):
        ps = data.pop('prazos')
        result[ym] = {
            'valor':      round(data['valor'], 2),
            'desagio':    round(data['desagio'], 2),
            'volumetria': data['volumetria'],
            'prazo_medio': round(sum(ps) / len(ps)) if ps else 0,
            'v_pago':     round(data['v_pago'], 2),
            'v_aberto':   round(data['v_aberto'], 2),
            'v_sem_info': round(data['v_sem_info'], 2),
            'n_pago':     data['n_pago'],
            'n_aberto':   data['n_aberto'],
            'n_sem_info': data['n_sem_info'],
        }

    total_emb = sum(m['valor'] for m in result.values())
    pct = round(matched/(matched+no_match)*100,1) if (matched+no_match) else 0
    print(f'  {display_name:<15} total={total_emb:>15,.2f}  match={pct}%  '
          f'pago={sum(m["n_pago"] for m in result.values())}  '
          f'aberto={sum(m["n_aberto"] for m in result.values())}  '
          f'sem_info={sum(m["n_sem_info"] for m in result.values())}')
    return result


print('\nProcessando FIDICs...')
all_monthly = {}
for sk, dn in SHEET_NAMES.items():
    if sk in wb.sheetnames:
        m = process(wb[sk], dn)
        if m:
            all_monthly[dn] = m

if ARTICO:
    m = process(wb[ARTICO], 'ARTICO')
    if m:
        all_monthly['ARTICO'] = m

all_months = sorted({ym for fund in all_monthly.values() for ym in fund})

out = {'monthly': all_monthly, 'all_months': all_months}
with open('fidic_monthly_v2.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

sz = len(json.dumps(out, ensure_ascii=False, separators=(',', ':')))
print(f'\nSalvo: fidic_monthly_v2.json  ({sz/1024:.1f} KB)')
