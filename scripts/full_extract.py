import openpyxl, datetime, json, sys
from collections import defaultdict, Counter

def dstr(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.date): return v.strftime('%Y-%m-%d')
    return v

def strip_or_none(v):
    return v.strip() if isinstance(v, str) else v

def sheet_rows(ws, header_row, id_col=1, filter_fn=None):
    headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(r, id_col).value is None:
            continue
        row = {headers[c - 1]: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
        if filter_fn and not filter_fn(row):
            continue
        rows.append(row)
    return rows


def extract_data(wb):
    data = {}

    # ---------------- Ventas Finn ----------------
    ws = wb['Ventas Finn']
    # El filtro por id_col (columna 1) no alcanza: puede haber filas de
    # plantilla o con fórmulas que dejan un valor en la columna 1 sin que la
    # fila tenga datos reales de venta (vta_id vacío). La tabla pedidos_venta
    # exige vta_id NOT NULL en Supabase, así que acá se descartan
    # explícitamente las filas sin vta_id, en vez de confiar en la columna 1
    # como proxy.
    rows = sheet_rows(ws, 3, filter_fn=lambda row: row.get('vta_id') is not None)
    ventas = []
    for row in rows:
        ventas.append({
            'vta_id': row['vta_id'],
            'fecha': dstr(row['Fecha']),
            'cliente': row['Cliente'],
            'categoria': row['Categoria'],
            'semillero': row['Semillero'],
            'variedad': row['Variedad'],
            'cantidad': row['Cantidad'],
            'vendedor': row['Vendedor'],
            'tratamiento': row['Tratamiento'],
            'pack': strip_or_none(row['Unidad']),
            'fecha_entrega': dstr(row['Fecha Entrega']),
            'kg_total': row['Kg Total'],
            'control_pendiente': row['Control'],
        })
    data['ventas'] = ventas
    data['kpi_ventas'] = {
        'total_kg': sum(v['kg_total'] or 0 for v in ventas),
        'pedidos': len(set(v['vta_id'] for v in ventas)),
        'clientes': len(set(v['cliente'] for v in ventas)),
    }
    por_variedad, por_vendedor, por_fecha = defaultdict(float), defaultdict(float), defaultdict(float)
    for v in ventas:
        por_variedad[v['variedad']] += v['kg_total'] or 0
        por_vendedor[v['vendedor']] += v['kg_total'] or 0
        por_fecha[v['fecha']] += v['kg_total'] or 0
    data['por_variedad'] = dict(por_variedad)
    data['por_vendedor'] = dict(por_vendedor)
    data['por_fecha'] = dict(por_fecha)

    # ---------------- Balanza ----------------
    ws = wb['Balanza']
    rows = sheet_rows(ws, 1)
    balanza = []
    for row in rows:
        balanza.append({
            'fecha': dstr(row['Fecha']),
            'chofer': row['Chofer'],
            'maquina': row['Maq. Limpieza'],
            'campo': row['Campo'],
            'variedad': row['Variedad'],
            'categoria': row['Categoria'],
            'neto': row['Neto'],
            'silo': row['SILO'],
            'merma': row['Merma '] or 0,
            'merma_viento': row['Merma viento '] or 0,
        })
    data['balanza'] = balanza
    total_neto = sum(b['neto'] or 0 for b in balanza)
    total_merma = sum(b['merma'] for b in balanza)
    total_mv = sum(b['merma_viento'] for b in balanza)
    data['kpi_balanza'] = {
        'total_neto': total_neto,
        'camiones': len(balanza),
        'campos': len(set(b['campo'] for b in balanza)),
        'merma_prom_pct': (total_merma + total_mv) / total_neto * 100 if total_neto else 0,
    }
    bal_por_variedad, bal_por_campo, bal_por_fecha = defaultdict(float), defaultdict(float), defaultdict(float)
    for b in balanza:
        bal_por_variedad[b['variedad']] += b['neto'] or 0
        bal_por_campo[b['campo']] += b['neto'] or 0
        bal_por_fecha[b['fecha']] += b['neto'] or 0
    data['bal_por_variedad'] = dict(bal_por_variedad)
    data['bal_por_campo'] = dict(bal_por_campo)
    data['bal_por_fecha'] = dict(bal_por_fecha)

    # ---------------- Intermedio -> lotes / kpi_planta / calidad ----------------
    ws = wb['Intermedio']
    rows = sheet_rows(ws, 3, id_col=2, filter_fn=lambda row: row['Fin'] is not None)
    lotes = []
    for row in rows:
        lotes.append({
            'lote': row['Lote Final'],
            'procedencia': row['Procedencia'],
            'variedad': row['Variedad'],
            'area': row['Area'],
            'kg': row['Kg Intermedio'],
            'kg_hr': row['Kg/hr'],
            'paro': row['Paro'],
            'detalle_paro': row['Detalle paro'],
            'fin': dstr(row['Fin']),
            'pureza': row['Pureza'],
            'ok_calidad': row['Ok Calidad'],
            'pg': row['PG'],
        })
    data['lotes'] = lotes
    kghr_vals = [l['kg_hr'] for l in lotes if l['kg_hr'] is not None]
    data['kpi_planta'] = {
        'kg_procesado': sum(l['kg'] or 0 for l in lotes),
        'rendimiento_prom': (sum(kghr_vals) / len(kghr_vals)) if kghr_vals else 0,
        'paradas': sum(1 for l in lotes if l['paro']),
        'lotes': len(lotes),
    }
    int_kg, int_kghr_list = defaultdict(float), defaultdict(list)
    for l in lotes:
        int_kg[l['variedad']] += l['kg'] or 0
        if l['kg_hr'] is not None:
            int_kghr_list[l['variedad']].append(l['kg_hr'])
    data['int_por_variedad_kg'] = dict(int_kg)
    data['int_rendimiento_var'] = {k: sum(v) / len(v) for k, v in int_kghr_list.items()}
    data['calidad_counts'] = dict(Counter((l['pureza'] or 'Sin dato') for l in lotes))

    # ---------------- Prod. Final -> prodfinal / estados_pedidos ----------------
    ws = wb['Prod. Final']
    rows = sheet_rows(ws, 3)
    prodfinal = []
    for row in rows:
        prodfinal.append({
            'vta_id': row['vta_id'],
            'venta': row['VENTA'],
            'estado': row['Estado'],
            'cliente': row['Firma'],
            'variedad': row['Variedad'],
            'presentacion': strip_or_none(row['Presentacion']),
            'kg': row['Cantidad KG'],
            'lote_final': row['Lote Final'],
            'fecha_entrega': dstr(row['fecha_entrega']),
        })
    data['prodfinal'] = prodfinal
    estados = defaultdict(lambda: {'count': 0, 'kg': 0.0})
    for p in prodfinal:
        estados[p['estado']]['count'] += 1
        estados[p['estado']]['kg'] += p['kg'] or 0
    data['estados_pedidos'] = {k: dict(v) for k, v in estados.items()}

    # ---------------- Stock 2026 -> resumen_marca / bolsas_by_semilla ----------------
    ws = wb['Stock 2026']
    resumen_marca = []
    r = 12
    while ws.cell(r, 6).value is not None:
        resumen_marca.append({
            'marca': ws.cell(r, 6).value,
            'pintados': ws.cell(r, 7).value,
            'sin_pintar': ws.cell(r, 8).value,
            'nuevos': ws.cell(r, 9).value,
        })
        r += 1
    data['resumen_marca'] = resumen_marca

    bolsas = defaultdict(float)
    r = 27
    while True:
        marca, semilla = ws.cell(r, 1).value, ws.cell(r, 3).value
        if marca is None and semilla is None:
            break
        if semilla is not None:
            bolsas[semilla] += ws.cell(r, 6).value or 0
        r += 1
        if r > ws.max_row + 5:
            break
    data['bolsas_by_semilla'] = dict(bolsas)

    # ---------------- Plan Curado ----------------
    ws = wb['Plan Curado']
    plan_curado = []
    r = 9
    while ws.cell(r, 1).value is not None:
        plan_curado.append({
            'variedad': ws.cell(r, 1).value,
            'tratamiento': ws.cell(r, 2).value,
            'pendiente_total': ws.cell(r, 4).value,
            'curado_disp': ws.cell(r, 5).value,
            'neto_a_curar': ws.cell(r, 6).value,
        })
        r += 1
    data['plan_curado'] = plan_curado
    weekly = []
    for c in range(7, 18):
        label = ws.cell(8, c).value
        if isinstance(label, datetime.datetime):
            label = label.strftime('%Y-%m-%d')
        weekly.append({'semana': label, 'kg': ws.cell(4, c).value or 0})
    data['plan_curado_weekly'] = weekly
    data['total_a_curar'] = ws.cell(5, 5).value

    # ---------------- Intermedio curado ----------------
    ws = wb['Intermedio curado']
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    lotes_curado = []
    for r in range(4, ws.max_row + 1):
        row = {headers[c - 1]: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
        if not row.get('Lote Intermedio') and not row.get('Lote Final'):
            continue
        lotes_curado.append({
            'id': row.get('ID'),
            'lote_intermedio': row.get('Lote Intermedio'),
            'lote_final': row.get('Lote Final'),
            'variedad': row.get('Variedad'),
            'categoria': row.get('Categoría'),
            'semillero': row.get('Semillero'),
            'kg_disp_inter': row.get('Kg. Disp. Inter.'),
            'kg_consumidos': row.get('Kg. Consumidos'),
            'area_int': row.get('Area Int.'),
            'kg_producidos': row.get('Kg. Producidos'),
            'kg_disponibles': row.get('Kg. Disponibles'),
            'area_final': row.get('Area Final'),
            'tratamiento': row.get('Tratamiento'),
            'complemento': row.get('Complemento'),
            'hora_inicio': dstr(row.get('Hora Inicio')),
            'hora_fin': dstr(row.get('Hora fin')),
            'horas_prod': row.get('Horas Prod.'),
            'kg_hs': row.get('Kg/Hs'),
        })
    data['lotes_curado'] = lotes_curado
    kg_prod_curado = sum(l['kg_producidos'] or 0 for l in lotes_curado)
    kghs_vals = [l['kg_hs'] for l in lotes_curado if l['kg_hs']]
    data['kpi_curado'] = {
        'lotes': len(lotes_curado),
        'kg_producidos': kg_prod_curado,
        'rendimiento_prom': (sum(kghs_vals) / len(kghs_vals)) if kghs_vals else 0,
    }

    # ---------------- Despacho ----------------
    ws = wb['Despacho']
    headers = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
    despachos = []
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            continue
        row = {headers[c - 1]: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
        despachos.append({
            'pedido': row['Pedido'],
            'cliente': row['Cliente'],
            'dia_carga': dstr(row['Dia de carga']),
            'variedad': row['Producto'],
            'presentacion': strip_or_none(row['Presentación']),
            'cantidad': row['Cantidad'],
            'kg': row['Cant. kg'],
            'lote_final': row['Lote Final'],
            'estado': row['Estado'],
        })
    data['despachos'] = despachos
    data['despacho_filas_totales'] = ws.max_row

    # ---------------- Calendario Produccion ----------------
    ws = wb['Calendario Produccion']
    total_gessi = total_blomar = 0
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            continue
        total_gessi += ws.cell(r, 5).value or 0
        total_blomar += ws.cell(r, 8).value or 0
    data['calendario_total_gessi'] = total_gessi
    data['calendario_total_blomar'] = total_blomar

    # ---------------- Chequeo ----------------
    ws = wb['Chequeo']
    data['chequeo'] = {'filas': ws.cell(4, 2).value}

    return data


# ---------------- Hojas del Excel (mirror) ----------------
def format_cell(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'Sí' if v else 'No'
    if isinstance(v, datetime.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime('%Y-%m-%d')
        return v.strftime('%Y-%m-%d %H:%M')
    if isinstance(v, datetime.date):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.time):
        return str(v)
    if isinstance(v, (int, float)):
        if abs(v - round(v)) < 1e-6:
            return f'{round(v):,}'.replace(',', '.')
        rv = round(v, 2)
        s = f'{rv:,.2f}'
        intpart, decpart = s.split('.')
        intpart = intpart.replace(',', '.')
        return f'{intpart},{decpart}'
    return str(v)


def used_bounds(ws):
    max_r = max_c = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                if cell.row > max_r: max_r = cell.row
                if cell.column > max_c: max_c = cell.column
    return max_r, max_c


def extract_sheets(wb):
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows, cols = used_bounds(ws)
        merges, covered = [], set()
        for mrange in ws.merged_cells.ranges:
            r0, c0, r1, c1 = mrange.min_row, mrange.min_col, mrange.max_row, mrange.max_col
            if r0 > rows or c0 > cols:
                continue
            merges.append({'r': r0, 'c': c0, 'rowspan': r1 - r0 + 1, 'colspan': c1 - c0 + 1})
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    if (rr, cc) != (r0, c0):
                        covered.add((rr, cc))
        grid = []
        for r in range(1, rows + 1):
            row_out = []
            for c in range(1, cols + 1):
                row_out.append(None if (r, c) in covered else format_cell(ws.cell(r, c).value))
            grid.append(row_out)
        out[name] = {'rows': rows, 'cols': cols, 'grid': grid, 'merges': merges}
    return out


if __name__ == '__main__':
    src = sys.argv[1]
    out_data = sys.argv[2]
    out_sheets = sys.argv[3]
    wb = openpyxl.load_workbook(src, data_only=True, keep_vba=True)
    data = extract_data(wb)
    sheets = extract_sheets(wb)
    json.dump(data, open(out_data, 'w', encoding='utf-8'), ensure_ascii=False)
    json.dump(sheets, open(out_sheets, 'w', encoding='utf-8'), ensure_ascii=False)
    print('wrote', out_data, out_sheets)
