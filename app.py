from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import openpyxl
import math
import io

app = Flask(__name__)
CORS(app)

# ── Helpers ──────────────────────────────────────────────────────────────────

def clean_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace('$', '').replace('%', '').replace(',', '.').strip()
    try:
        return float(s)
    except:
        return None

def ns_to_z(ns):
    if ns >= 0.99: return 2.326
    if ns >= 0.98: return 2.054
    if ns >= 0.97: return 1.881
    if ns >= 0.95: return 1.645
    return 1.282

def classify_abc(products):
    total_va = sum(p['D'] * p['pv'] for p in products)
    if total_va == 0:
        for p in products:
            p['clase'] = 'C'
        return
    sorted_p = sorted(products, key=lambda x: x['D'] * x['pv'], reverse=True)
    cum = 0
    for p in sorted_p:
        cum += p['D'] * p['pv'] / total_va * 100
        p['clase'] = 'A' if cum <= 80 else ('B' if cum <= 95 else 'C')
    # Write clase back to original list by code
    clase_map = {p['code']: p['clase'] for p in sorted_p}
    for p in products:
        p['clase'] = clase_map.get(p['code'], 'C')

def calc_eoq(p, params):
    kpct  = params['kpct'] / 100
    c_loc = params['cloc']
    alt   = params['alt'] or 2.0
    z_map = {'A': params['za'], 'B': params['zb'], 'C': params['zc']}
    z     = z_map.get(p['clase'], 1.645)

    ms    = p['months']
    n     = len(ms)
    D     = p['D']
    d_avg = D / 12

    variance  = sum((x - d_avg) ** 2 for x in ms) / (n - 1) if n > 1 else 0
    sigma_m   = math.sqrt(variance)
    cv        = sigma_m / d_avg if d_avg > 0 else 0
    sigma_lt  = sigma_m * math.sqrt(p['lt'] / 30)
    SS        = z * sigma_lt

    m2u = p['vol'] / alt
    h   = p['cu'] * 0.25 + m2u * c_loc * 12
    K   = kpct * d_avg * p['cu']

    EOQ   = math.sqrt(2 * D * K / h) if h > 0 else 0
    n_ped = D / EOQ if EOQ > 0 else 0
    ciclo = 365 / n_ped if n_ped > 0 else 0
    PR    = (D / 365) * p['lt'] + SS
    inv_p = EOQ / 2 + SS
    CIT   = (D / EOQ) * K + inv_p * h if EOQ > 0 else 0
    CIT_b = math.sqrt(2 * D * K * h) if EOQ > 0 else 0
    cSS   = SS * h
    rent  = ((p['pv'] - p['cu']) / p['pv']) * 100 if p['pv'] > 0 else 0

    return {
        'z': round(z, 3),
        'sigma_m': round(sigma_m, 1),
        'cv': round(cv, 3),
        'SS': round(SS, 1),
        'h': round(h, 2),
        'K': round(K, 2),
        'EOQ': round(EOQ),
        'n_ped': round(n_ped, 2),
        'ciclo': round(ciclo, 1),
        'PR': round(PR, 1),
        'inv_p': round(inv_p, 1),
        'CIT': round(CIT, 2),
        'CIT_b': round(CIT_b, 2),
        'cSS': round(cSS, 2),
        'rent': round(rent, 2),
        'val_inv': round(inv_p * p['cu'], 2),
    }

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/parse-excel', methods=['POST'])
def parse_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo'}), 400

    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'El archivo debe ser .xlsx o .xls'}), 400

    try:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        ws = wb[wb.sheetnames[0]]

        # Parámetros globales (filas 5-9, col B)
        ns_a  = clean_num(ws.cell(5, 2).value) or 0.99
        ns_b  = clean_num(ws.cell(6, 2).value) or 0.97
        ns_c  = clean_num(ws.cell(7, 2).value) or 0.95
        kpct  = clean_num(ws.cell(8, 2).value) or 0.04
        cloc_raw = ws.cell(9, 2).value
        cloc  = clean_num(cloc_raw) or 9500

        params = {
            'za':   ns_to_z(ns_a),
            'zb':   ns_to_z(ns_b),
            'zc':   ns_to_z(ns_c),
            'kpct': round(kpct * 100, 2),
            'cloc': cloc,
            'alt':  2.0,
        }

        # Productos desde fila 14
        products = []
        for row in range(14, ws.max_row + 1):
            code = ws.cell(row, 1).value
            if not code or str(code).strip() == '':
                break
            name  = str(ws.cell(row, 2).value or '').strip()
            vol   = clean_num(ws.cell(row, 3).value) or 0.05
            lt    = clean_num(ws.cell(row, 5).value) or 10
            # Columnas F-Q = 6-17 → 12 meses
            months_raw = [clean_num(ws.cell(row, c).value) for c in range(6, 18)]
            D_total    = clean_num(ws.cell(row, 18).value)
            pv         = clean_num(ws.cell(row, 19).value) or 0
            cu         = clean_num(ws.cell(row, 20).value) or 0

            # Imputar meses nulos con promedio de los válidos
            valid  = [m for m in months_raw if m is not None and m > 0]
            avg    = sum(valid) / len(valid) if valid else 0
            months = [m if (m is not None and m > 0) else avg for m in months_raw]
            D      = D_total or sum(months)

            products.append({
                'code':   str(code).strip(),
                'name':   name,
                'vol':    vol,
                'lt':     lt,
                'months': months,
                'D':      D,
                'pv':     pv,
                'cu':     cu,
                'clase':  'A',  # se sobreescribe en classify_abc
            })

        if not products:
            return jsonify({'error': 'No se encontraron productos en el archivo'}), 400

        classify_abc(products)
        results = []
        for p in products:
            r = calc_eoq(p, params)
            results.append({**p, **r})

        return jsonify({'params': params, 'products': results})

    except Exception as e:
        return jsonify({'error': f'Error procesando el archivo: {str(e)}'}), 500


@app.route('/api/recalc', methods=['POST'])
def recalc():
    body = request.get_json()
    products = body.get('products', [])
    params   = body.get('params', {})

    if not products:
        return jsonify({'error': 'Sin productos'}), 400

    classify_abc(products)
    results = []
    for p in products:
        r = calc_eoq(p, params)
        results.append({**p, **r})

    return jsonify({'products': results})


if __name__ == '__main__':
    app.run(debug=True)
