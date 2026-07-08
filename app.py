from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import openpyxl
import math
import io

app = Flask(__name__)
CORS(app)

# ── Helpers ──────────────────────────────────────────────────────────────────

def clean_num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace('$','').replace('%','').replace(',','.').strip()
    try: return float(s)
    except: return None

def ns_to_z(ns):
    if ns >= 0.99: return 2.326
    if ns >= 0.98: return 2.054
    if ns >= 0.97: return 1.881
    if ns >= 0.95: return 1.645
    return 1.282

def classify_abc(products):
    total_va = sum(p['D'] * p['pv'] for p in products)
    if total_va == 0:
        for p in products: p['clase'] = 'C'
        return
    sorted_p = sorted(products, key=lambda x: x['D'] * x['pv'], reverse=True)
    cum = 0
    for p in sorted_p:
        cum += p['D'] * p['pv'] / total_va * 100
        p['clase'] = 'A' if cum <= 80 else ('B' if cum <= 95 else 'C')
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
        'z': round(z, 3), 'sigma_m': round(sigma_m, 1), 'cv': round(cv, 3),
        'SS': round(SS, 1), 'h': round(h, 2), 'K': round(K, 2),
        'EOQ': round(EOQ), 'n_ped': round(n_ped, 2), 'ciclo': round(ciclo, 1),
        'PR': round(PR, 1), 'inv_p': round(inv_p, 1),
        'CIT': round(CIT, 2), 'CIT_b': round(CIT_b, 2),
        'cSS': round(cSS, 2), 'rent': round(rent, 2),
        'val_inv': round(inv_p * p['cu'], 2),
    }

def calc_rotation(p):
    """Rotación de inventario, sell-through, backorders y días de inventario."""
    D     = p['D']
    ms    = p['months']
    inv_p = p.get('inv_p', D / 24)  # inventario promedio estimado si no hay EOQ

    # Rotación = ventas anuales / inventario promedio
    rotacion = D / inv_p if inv_p > 0 else 0

    # Días de inventario (DSI)
    dsi = 365 / rotacion if rotacion > 0 else 0

    # Sell-through rate mensual promedio
    # (ventas / (ventas + stock promedio)) * 100
    stock_prom_mensual = inv_p
    ventas_mes = D / 12
    sell_through = (ventas_mes / (ventas_mes + stock_prom_mensual)) * 100 if (ventas_mes + stock_prom_mensual) > 0 else 0

    # Variabilidad mensual para detectar backorders potenciales
    d_avg = D / 12
    variance = sum((x - d_avg) ** 2 for x in ms) / (len(ms) - 1) if len(ms) > 1 else 0
    sigma_m = math.sqrt(variance)

    # Meses con demanda > inventario promedio (riesgo de backorder)
    meses_riesgo = sum(1 for m in ms if m > inv_p)
    pct_riesgo = meses_riesgo / len(ms) * 100

    # Backorder estimado = suma de excesos sobre el inventario promedio
    backorder_est = sum(max(0, m - inv_p) for m in ms)

    return {
        'rotacion': round(rotacion, 2),
        'dsi': round(dsi, 1),
        'sell_through': round(sell_through, 1),
        'meses_riesgo': meses_riesgo,
        'pct_riesgo': round(pct_riesgo, 1),
        'backorder_est': round(backorder_est, 1),
    }

def calc_forecast(months, periods=6, alpha=None):
    """
    Suavizado exponencial simple (Holt-Winters simple).
    Si alpha=None, lo optimiza minimizando MSE.
    Retorna forecast para los próximos `periods` meses y métricas.
    """
    if len(months) < 3:
        avg = sum(months) / len(months)
        return {
            'forecast': [round(avg, 1)] * periods,
            'alpha': 0.3,
            'mae': 0,
            'mape': 0,
            'fitted': months,
            'trend': 0,
        }

    def ses(data, a):
        fitted = [data[0]]
        for i in range(1, len(data)):
            fitted.append(a * data[i-1] + (1 - a) * fitted[i-1])
        return fitted

    def mse(data, a):
        f = ses(data, a)
        return sum((data[i] - f[i])**2 for i in range(1, len(data))) / (len(data) - 1)

    # Optimizar alpha si no se provee
    if alpha is None:
        best_a, best_mse = 0.1, float('inf')
        for a in [i/10 for i in range(1, 10)]:
            m = mse(months, a)
            if m < best_mse:
                best_mse, best_a = m, a
        alpha = best_a

    fitted = ses(months, alpha)
    last = fitted[-1]

    # Detectar tendencia lineal simple
    n = len(months)
    x_avg = (n - 1) / 2
    y_avg = sum(months) / n
    trend = sum((i - x_avg) * (months[i] - y_avg) for i in range(n)) / \
            sum((i - x_avg)**2 for i in range(n)) if n > 1 else 0

    # Forecast: suavizado + tendencia ajustada
    forecast = []
    val = last
    for i in range(periods):
        val = alpha * months[-1] + (1 - alpha) * val + trend * 0.3
        forecast.append(round(max(0, val), 1))

    # MAE y MAPE
    errors = [abs(months[i] - fitted[i]) for i in range(1, len(months))]
    mae = sum(errors) / len(errors) if errors else 0
    pct_errors = [abs(months[i] - fitted[i]) / months[i] * 100
                  for i in range(1, len(months)) if months[i] > 0]
    mape = sum(pct_errors) / len(pct_errors) if pct_errors else 0

    return {
        'forecast': forecast,
        'alpha': round(alpha, 2),
        'mae': round(mae, 1),
        'mape': round(mape, 1),
        'fitted': [round(f, 1) for f in fitted],
        'trend': round(trend, 2),
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
        ns_a  = clean_num(ws.cell(5, 2).value) or 0.99
        ns_b  = clean_num(ws.cell(6, 2).value) or 0.97
        ns_c  = clean_num(ws.cell(7, 2).value) or 0.95
        kpct  = clean_num(ws.cell(8, 2).value) or 0.04
        cloc  = clean_num(ws.cell(9, 2).value) or 9500
        params = {
            'za': ns_to_z(ns_a), 'zb': ns_to_z(ns_b), 'zc': ns_to_z(ns_c),
            'kpct': round(kpct * 100, 2), 'cloc': cloc, 'alt': 2.0,
        }
        products = []
        for row in range(14, ws.max_row + 1):
            code = ws.cell(row, 1).value
            if not code or str(code).strip() == '': break
            months_raw = [clean_num(ws.cell(row, c).value) for c in range(6, 18)]
            valid  = [m for m in months_raw if m is not None and m > 0]
            avg    = sum(valid) / len(valid) if valid else 0
            months = [m if (m is not None and m > 0) else avg for m in months_raw]
            D_total = clean_num(ws.cell(row, 18).value)
            D = D_total or sum(months)
            products.append({
                'code':   str(code).strip(),
                'name':   str(ws.cell(row, 2).value or '').strip(),
                'vol':    clean_num(ws.cell(row, 3).value) or 0.05,
                'lt':     clean_num(ws.cell(row, 5).value) or 10,
                'months': months,
                'D':      D,
                'pv':     clean_num(ws.cell(row, 19).value) or 0,
                'cu':     clean_num(ws.cell(row, 20).value) or 0,
                'clase':  'A',
            })
        if not products:
            return jsonify({'error': 'No se encontraron productos'}), 400
        classify_abc(products)
        results = []
        for p in products:
            r = calc_eoq(p, params)
            p.update(r)
            rot = calc_rotation(p)
            fc  = calc_forecast(p['months'])
            results.append({**p, **rot, 'forecast': fc})
        return jsonify({'params': params, 'products': results})
    except Exception as e:
        return jsonify({'error': f'Error: {str(e)}'}), 500

@app.route('/api/recalc', methods=['POST'])
def recalc():
    body     = request.get_json()
    products = body.get('products', [])
    params   = body.get('params', {})
    if not products: return jsonify({'error': 'Sin productos'}), 400
    classify_abc(products)
    results = []
    for p in products:
        r = calc_eoq(p, params)
        p.update(r)
        rot = calc_rotation(p)
        fc  = calc_forecast(p['months'])
        results.append({**p, **rot, 'forecast': fc})
    return jsonify({'products': results})

if __name__ == '__main__':
    app.run(debug=True)
