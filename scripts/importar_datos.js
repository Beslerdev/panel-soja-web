/*
 * Importador Excel -> Supabase, en un solo paso (Node.js).
 *
 * Hace lo mismo que full_extract.py + generate_import_sql.py + pegar el SQL
 * a mano en Supabase, pero de una sola corrida: lee el Excel, arma el SQL de
 * reemplazo (DELETE + INSERT por tabla), y lo manda directo a Supabase por
 * HTTP (a la función `exec_import_sql`, que solo puede ejecutar la clave
 * secreta del servidor). Nada pasa por el chat ni por copiar y pegar — por
 * eso tarda segundos y no minutos.
 *
 * Uso:
 *   node importar_datos.js "SOJA CAMPAÑA 2526.xlsm"
 *
 * Necesita, en la carpeta de arriba (panel-soja-web/config.env), las mismas
 * variables que usa el servidor: SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY.
 * Necesita también el paquete "xlsx" instalado (npm install xlsx).
 */
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ---------------- utilidades de fecha/hora/texto ----------------

function dstr(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) {
    const y = v.getUTCFullYear();
    const m = String(v.getUTCMonth() + 1).padStart(2, '0');
    const d = String(v.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // Una celda con FORMATO DE NUMERO (en lugar de formato de fecha) no llega
  // como Date sino como el "numero de serie" de Excel: 46281 es 16/09/2026.
  // Pasa cuando alguien tipea una fecha en una celda que quedo sin formatear
  // — apareció en Intermedio!V80 ("Fin" del lote S264-2STOO) y rompia la
  // importacion entera con "column fin is of type date but expression is of
  // type integer". Se convierte a fecha en vez de cortar la corrida.
  // El rango 20000-60000 (1954 a 2064) evita tomar por fecha un numero que
  // en realidad es otra cosa; cualquier otro valor se deja como estaba.
  if (typeof v === 'number' && Number.isFinite(v) && v >= 20000 && v <= 60000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${dd}`;
  }
  return v;
}

// Los encabezados del Excel a veces traen espacios al borde ('PG ', 'Merma ').
// Se recortan al armar el mapa de columnas, asi el mapeo no depende de esos
// espacios: si alguien los agrega o los saca, la columna se sigue encontrando.
// (La columna 'PG ' de la hoja Intermedio se venia importando vacia por esto.)
function hdrName(v) {
  return typeof v === 'string' ? v.trim() : v;
}

function stripOrNone(v) {
  return typeof v === 'string' ? v.trim() : v;
}

function fmtTimeCell(v) {
  // Para columnas 'text' que en el Excel guardan horarios (Calendario Produccion).
  if (v === null || v === undefined) return null;
  if (v instanceof Date) {
    const h = String(v.getUTCHours()).padStart(2, '0');
    const mi = String(v.getUTCMinutes()).padStart(2, '0');
    const s = String(v.getUTCSeconds()).padStart(2, '0');
    return `${h}:${mi}:${s}`;
  }
  if (typeof v === 'string' && v.trim() === '') return null;
  return v;
}

function cleanHours(v) {
  // Las horas de Calendario Produccion vienen de restas de horarios en Excel
  // y a veces arrastran errores de punto flotante mínusculos
  // (9.000000000000002 en vez de 9.0). Si está a menos de 1e-6 de un entero,
  // se redondea.
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') {
    const rv = Math.round(v);
    return Math.abs(v - rv) < 1e-6 ? rv : Math.round(v * 1e6) / 1e6;
  }
  return v;
}

// ---------------- acceso a celdas (1-indexado, como openpyxl) ----------------

function cellAt(ws, r, c) {
  const addr = xlsx.utils.encode_cell({ r: r - 1, c: c - 1 });
  const cell = ws[addr];
  if (!cell) return null;
  const v = cell.v === undefined ? null : cell.v;
  // openpyxl (data_only=True) devuelve None para celdas de fórmula cuyo
  // resultado cacheado es texto vacío (p.ej. IF(...,"") sin cumplirse) — la
  // librería xlsx en cambio preserva la cadena vacía. Igualamos el criterio
  // acá para que los mismos filtros (is not None / !== null) den el mismo
  // resultado en los dos lenguajes.
  if (v === '') return null;
  return v;
}

function sheetBounds(ws) {
  const range = xlsx.utils.decode_range(ws['!ref']);
  return { maxRow: range.e.r + 1, maxCol: range.e.c + 1 };
}

function sheetRows(ws, headerRow, idCol, filterFn) {
  idCol = idCol || 1;
  const { maxRow, maxCol } = sheetBounds(ws);
  const headers = [];
  for (let c = 1; c <= maxCol; c++) headers.push(hdrName(cellAt(ws, headerRow, c)));
  const rows = [];
  for (let r = headerRow + 1; r <= maxRow; r++) {
    if (cellAt(ws, r, idCol) === null) continue;
    const row = {};
    for (let c = 1; c <= maxCol; c++) row[headers[c - 1]] = cellAt(ws, r, c);
    if (filterFn && !filterFn(row)) continue;
    row.__fila = r; // numero de fila del Excel, para poder señalarla en los avisos
    rows.push(row);
  }
  return rows;
}

// ---------------- extracción por hoja (mismas hojas que full_extract.py) ----------------

function extractVentas(wb) {
  const ws = wb.Sheets['Ventas Finn'];
  // El filtro por idCol (columna 1) no alcanza: puede haber filas de plantilla
  // o con fórmulas que dejan un valor en la columna 1 sin que la fila tenga
  // datos reales de venta (vta_id vacío). La tabla pedidos_venta exige
  // vta_id NOT NULL en Supabase, así que acá se descartan explícitamente las
  // filas sin vta_id, en vez de confiar en la columna 1 como proxy.
  const rows = sheetRows(ws, 3, 1, (row) => row['vta_id'] !== null && row['vta_id'] !== undefined);
  // Ventas canceladas: cantidad en 0 (ver el comentario de esVentaCancelada).
  const vivas = rows.filter((row) => {
    if (esVentaCancelada(row)) {
      OMITIDAS.ventasCanceladas.push({ fila: row.__fila, vta_id: row['vta_id'], cliente: row['Cliente'] });
      return false;
    }
    return true;
  });
  return vivas.map((row) => ({
    vta_id: row['vta_id'],
    fecha: dstr(row['Fecha']),
    cliente: row['Cliente'],
    categoria: row['Categoria'],
    semillero: row['Semillero'],
    variedad: row['Variedad'],
    cantidad: row['Cantidad'],
    vendedor: row['Vendedor'],
    tratamiento: row['Tratamiento'],
    presentacion: stripOrNone(row['Unidad']),
    fecha_entrega: dstr(row['Fecha Entrega']),
    kg_total: row['Kg Total'],
    control_pendiente: row['Control'],
    // Clave de la línea (columna "id"): pedido + variedad + cliente + envase + cantidad.
    // Es lo que "Prod. Final" usa en "Concatenado" para apuntar a esta venta;
    // con eso el panel calcula vencidos / por vencer por línea.
    linea_id: row['id'],
    __fila: row.__fila,
  }));
}

function extractBalanza(wb) {
  const ws = wb.Sheets['Balanza'];
  const rows = sheetRows(ws, 1);
  return rows.map((row) => ({
    fecha: dstr(row['Fecha']),
    chofer: row['Chofer'],
    maquina: row['Maq. Limpieza'],
    campo: row['Campo'],
    variedad: row['Variedad'],
    categoria: row['Categoria'],
    neto: row['Neto'],
    silo: row['SILO'],
    merma: row['Merma'] || 0,
    merma_viento: row['Merma viento'] || 0,
    __fila: row.__fila,
  }));
}

function extractLotes(wb) {
  const ws = wb.Sheets['Intermedio'];
  const rows = sheetRows(ws, 3, 2, (row) => row['Fin'] !== null && row['Fin'] !== undefined);
  return rows.map((row) => ({
    lote: row['Lote Final'],
    procedencia: row['Procedencia'],
    variedad: row['Variedad'],
    area: row['Area'],
    kg: row['Kg Intermedio'],
    kg_hr: row['Kg/hr'],
    paro: row['Paro'],
    detalle_paro: row['Detalle paro'],
    fin: dstr(row['Fin']),
    pureza: row['Pureza'],
    ok_calidad: row['Ok Calidad'],
    pg: row['PG'],
    pg_curado: row['PG Curado'],
    p1000: row['P1000'],
    __fila: row.__fila,
  }));
}

function extractProdFinal(wb) {
  const ws = wb.Sheets['Prod. Final'];
  const rows = sheetRows(ws, 3);
  // Filas cuyo BUSCARV no encontro la venta: la venta esta cancelada.
  const utiles = rows.filter((row) => {
    if (esFilaSinVenta(row)) {
      OMITIDAS.prodSinVenta.push({ fila: row.__fila, venta: row['VENTA'] });
      return false;
    }
    return true;
  });
  return utiles.map((row) => ({
    vta_id: row['vta_id'],
    venta: row['VENTA'],
    estado: row['Estado'],
    cliente: row['Firma'],
    variedad: row['Variedad'],
    presentacion: stripOrNone(row['Presentacion']),
    kg: row['Cantidad KG'],
    lote_final: row['Lote Final'],
    fecha_entrega: dstr(row['fecha_entrega']),
    // Para "Vencidos y por vencer": a qué línea de venta pertenece la fila
    // (Concatenado = id de Ventas Finn) y cuánto de ella ya salió de planta.
    linea_id: row['Concatenado'],
    cantidad: row['Cantidad'],
    despachado: row['Despachado'],
    pend_despacho: row['Pend. Despacho'],
    __fila: row.__fila,
  }));
}

function extractDespachos(wb) {
  const ws = wb.Sheets['Despacho'];
  const { maxRow, maxCol } = sheetBounds(ws);
  const headers = [];
  for (let c = 1; c <= maxCol; c++) headers.push(hdrName(cellAt(ws, 2, c)));
  const out = [];
  for (let r = 3; r <= maxRow; r++) {
    if (cellAt(ws, r, 1) === null) continue;
    const row = {};
    for (let c = 1; c <= maxCol; c++) row[headers[c - 1]] = cellAt(ws, r, c);
    let lote_final = row['Lote Final'];
    if (lote_final === 0) lote_final = null; // placeholder del Excel para "sin lote todavía"
    out.push({
      pedido: row['Pedido'],
      cliente: row['Cliente'],
      dia_carga: dstr(row['Dia de carga']),
      variedad: row['Producto'],
      presentacion: stripOrNone(row['Presentación']),
      cantidad: row['Cantidad'],
      kg: row['Cant. kg'],
      lote_final,
      estado: row['Estado'],
      // Columna nueva (2026-10-08): número de camión del día. Junto con
      // "Dia de carga" identifica el camión real que salió de planta —
      // varias líneas (pedidos) pueden compartir un mismo camión. Puede
      // venir vacía en despachos anteriores a que se agregara esta columna.
      camion: row['Camión'],
      __fila: r,
    });
  }
  return out;
}

function extractStock(wb) {
  const ws = wb.Sheets['Stock 2026'];
  const resumenMarca = [];
  let r = 12;
  while (cellAt(ws, r, 6) !== null) {
    resumenMarca.push({
      marca: cellAt(ws, r, 6),
      pintados: cellAt(ws, r, 7),
      sin_pintar: cellAt(ws, r, 8),
      nuevos: cellAt(ws, r, 9),
    });
    r += 1;
  }

  const bolsas = {};
  const { maxRow } = sheetBounds(ws);
  r = 27;
  while (true) {
    const marca = cellAt(ws, r, 1);
    const semilla = cellAt(ws, r, 3);
    if (marca === null && semilla === null) break;
    if (semilla !== null) {
      bolsas[semilla] = (bolsas[semilla] || 0) + (cellAt(ws, r, 6) || 0);
    }
    r += 1;
    if (r > maxRow + 5) break;
  }
  const bolsasList = Object.entries(bolsas).map(([semilla, bolsas_]) => ({ semilla, bolsas: bolsas_ }));
  return { resumenMarca, bolsasList };
}

function extractPlanCurado(wb) {
  const ws = wb.Sheets['Plan Curado'];
  const { maxCol } = sheetBounds(ws);
  const headerRow = 8;
  // Columnas por nombre de encabezado, no por posicion fija: si se inserta
  // o reordena una columna en esta hoja (pasó el 2026-10-19, al agregar
  // "Categoria" -Primu/Original- antes de "Pendiente total", corriendo
  // todo un lugar a la derecha) esto sigue encontrando cada dato por su
  // nombre en vez de leer lo que haya quedado en la posicion vieja.
  const col = {};
  for (let c = 1; c <= maxCol; c++) {
    const h = hdrName(cellAt(ws, headerRow, c));
    if (h) col[h] = c;
  }
  const cVariedad = col['Variedad'] || 1;
  const cTratamiento = col['Tratamiento'] || 2;
  const cPendiente = col['Pendiente total'];
  const cCuradoDisp = col['Curado disp.'];
  const cNeto = col['Neto a curar'];
  const cSinFecha = col['Sin fecha'];

  const planCurado = [];
  let r = headerRow + 1;
  while (cellAt(ws, r, cVariedad) !== null) {
    planCurado.push({
      variedad: cellAt(ws, r, cVariedad),
      tratamiento: cellAt(ws, r, cTratamiento),
      pendiente_total: cPendiente ? cellAt(ws, r, cPendiente) : null,
      curado_disp: cCuradoDisp ? cellAt(ws, r, cCuradoDisp) : null,
      neto_a_curar: cNeto ? cellAt(ws, r, cNeto) : null,
    });
    r += 1;
  }

  // Igual que arriba: arranca donde esté de verdad la columna "Sin fecha"
  // (antes era siempre la 7), y llega hasta la ultima columna real de la
  // hoja en vez de cortar en la 17 — asi tampoco se pierde la ultima
  // semana del cronograma cuando se agregan mas columnas de fecha.
  const weekly = [];
  const startCol = cSinFecha || 7;
  for (let c = startCol; c <= maxCol; c++) {
    let label = cellAt(ws, headerRow, c);
    if (label instanceof Date) label = dstr(label);
    weekly.push({ semana: label, kg: cellAt(ws, 4, c) || 0 });
  }
  return { planCurado, weekly };
}

function extractLotesCurado(wb) {
  const ws = wb.Sheets['Intermedio curado'];
  const { maxRow, maxCol } = sheetBounds(ws);
  const headers = [];
  for (let c = 1; c <= maxCol; c++) headers.push(hdrName(cellAt(ws, 3, c)));
  const out = [];
  for (let r = 4; r <= maxRow; r++) {
    const row = {};
    for (let c = 1; c <= maxCol; c++) row[headers[c - 1]] = cellAt(ws, r, c);
    if (!row['Lote Intermedio'] && !row['Lote Final']) continue;
    out.push({
      lote_intermedio: row['Lote Intermedio'],
      lote_final: row['Lote Final'],
      variedad: row['Variedad'],
      categoria: row['Categoría'],
      semillero: row['Semillero'],
      kg_disp_inter: row['Kg. Disp. Inter.'],
      kg_consumidos: row['Kg. Consumidos'],
      area_int: row['Area Int.'],
      kg_producidos: row['Kg. Producidos'],
      kg_disponibles: row['Kg. Disponibles'],
      area_final: row['Area Final'],
      tratamiento: row['Tratamiento'],
      complemento: row['Complemento'],
      hora_inicio: dstr(row['Hora Inicio']),
      hora_fin: dstr(row['Hora fin']),
      horas_prod: row['Horas Prod.'],
      kg_hs: row['Kg/Hs'],
      __fila: r,
    });
  }
  return out;
}

function extractVariedades(wb) {
  const ws = wb.Sheets['MASTER'];
  const out = [];
  for (let r = 3; r <= 13; r++) {
    const variedad = cellAt(ws, r, 3);
    if (variedad === null || String(variedad).trim() === '') continue;
    out.push({
      variedad: String(variedad).trim(),
      ciclo: cellAt(ws, r, 4),
      semilla: cellAt(ws, r, 5),
      semillero: cellAt(ws, r, 6),
      tecnologia: cellAt(ws, r, 7),
    });
  }
  return out;
}

function extractCalendario(wb) {
  const ws = wb.Sheets['Calendario Produccion'];
  const { maxRow } = sheetBounds(ws);
  const out = [];
  for (let r = 5; r <= maxRow; r++) {
    const fecha = cellAt(ws, r, 1);
    if (fecha === null) continue;
    out.push({
      fecha: dstr(fecha),
      dia: cellAt(ws, r, 2),
      gessi_inicio: fmtTimeCell(cellAt(ws, r, 3)),
      gessi_fin: fmtTimeCell(cellAt(ws, r, 4)),
      horas_gessi: cleanHours(cellAt(ws, r, 5)),
      blomar_inicio: fmtTimeCell(cellAt(ws, r, 6)),
      blomar_fin: fmtTimeCell(cellAt(ws, r, 7)),
      horas_blomar: cleanHours(cellAt(ws, r, 8)),
      observacion: cellAt(ws, r, 9),
      __fila: r,
    });
  }
  return out;
}

// ---------------- ventas canceladas ----------------

/*
 * Una venta cancelada, en el Excel, es una fila con vta_id cargado y Cantidad
 * en 0. No es una convención inventada acá: es la misma condición del formato
 * condicional de la hoja "Ventas Finn", el que las muestra tachadas en rojo —
 *
 *     rango A4:Y434   ->   AND($B4<>"", $J4=0)   ->   tachado
 *     (columna B = vta_id, columna J = Cantidad)
 *
 * Esas ventas no se importan. Y las filas de "Prod. Final" que apuntan a una
 * de ellas tampoco: su BUSCARV no encuentra la venta y devuelve "SIN VENTA"
 * en toda la fila, así que no hay dato que cargar.
 *
 * Nada de esto se descarta en silencio: lo omitido se cuenta y se informa por
 * pantalla al final de la corrida.
 */

const OMITIDAS = { ventasCanceladas: [], prodSinVenta: [] };

function esVentaCancelada(row) {
  const cant = row['Cantidad'];
  return cant === 0 || cant === null || cant === undefined || cant === '';
}

function esFilaSinVenta(row) {
  const marca = (v) => typeof v === 'string' && v.trim().toUpperCase() === 'SIN VENTA';
  return marca(row['vta_id']) || marca(row['Cantidad KG']) || marca(row['Firma']);
}

// ---------------- control previo de columnas numéricas ----------------

/*
 * Antes de mandar nada a Supabase, revisa que las columnas que en la base son
 * numéricas traigan números. Si una celda del Excel tiene texto —el caso típico
 * es un BUSCARV que no encuentra la venta y devuelve "SIN VENTA"— Supabase
 * rechaza el bloque entero con un error como:
 *
 *   invalid input syntax for type numeric: "SIN VENTA"
 *
 * que no dice en qué hoja ni en qué fila está el problema. Este control lo
 * encuentra antes y lo informa con hoja, fila y columna.
 *
 * No corrige ni convierte nada: si todo está bien, la importación sigue
 * exactamente igual que antes.
 */

// Por tabla: la hoja del Excel, las columnas numéricas (con el nombre que
// tienen en el Excel, para que el aviso sea reconocible), y las columnas que
// sirven para identificar la fila.
// Ojo: en 'lotes', las columnas pg / pg_curado / p1000 son TEXT en la base
// (los valores de laboratorio a veces vienen con texto), asi que no se controlan aca.
const COLUMNAS_NUMERICAS = {
  pedidos_venta:         { hoja: 'Ventas Finn',           num: { cantidad: 'Cantidad', kg_total: 'Kg Total', control_pendiente: 'Control' },                          id: ['vta_id', 'cliente'] },
  balanza:               { hoja: 'Balanza',               num: { neto: 'Neto', merma: 'Merma', merma_viento: 'Merma viento' },                                          id: ['fecha', 'chofer'] },
  lotes:                 { hoja: 'Intermedio',            num: { kg: 'Kg Intermedio', kg_hr: 'Kg/hr' },                                                                             id: ['lote', 'variedad'] },
  produccion_final:      { hoja: 'Prod. Final',           num: { kg: 'Cantidad KG', cantidad: 'Cantidad', despachado: 'Despachado', pend_despacho: 'Pend. Despacho' },                                                                                   id: ['venta', 'vta_id', 'cliente'] },
  despachos:             { hoja: 'Despacho',              num: { cantidad: 'Cantidad', kg: 'Cant. kg' },                                                                id: ['pedido', 'cliente'] },
  stock_resumen_marca:   { hoja: 'Stock 2026',            num: { pintados: 'Pintados', sin_pintar: 'Sin pintar', nuevos: 'Nuevos' },                                    id: ['marca'] },
  stock_bolsas_semilla:  { hoja: 'Stock 2026',            num: { bolsas: 'Bolsas' },                                                                                    id: ['semilla'] },
  plan_curado:           { hoja: 'Plan Curado',           num: { pendiente_total: 'Pendiente total', curado_disp: 'Curado disp.', neto_a_curar: 'Neto a curar' },       id: ['variedad'] },
  plan_curado_semanal:   { hoja: 'Plan Curado',           num: { kg: 'Kg' },                                                                                            id: ['semana'] },
  lotes_curado:          { hoja: 'Intermedio curado',     num: { kg_disp_inter: 'Kg. Disp. Inter.', kg_consumidos: 'Kg. Consumidos', kg_producidos: 'Kg. Producidos', kg_disponibles: 'Kg. Disponibles', horas_prod: 'Horas Prod.', kg_hs: 'Kg/Hs' }, id: ['lote_final', 'lote_intermedio'] },
  calendario_produccion: { hoja: 'Calendario Produccion', num: { horas_gessi: 'Horas Gessi', horas_blomar: 'Horas Blomar' },                                            id: ['fecha'] },
};

// Vacío o número está bien. Solo molesta el texto que no sea un número.
function esNumeroAceptable(v) {
  if (v === null || v === undefined || v === '') return true;
  if (typeof v === 'number') return Number.isFinite(v);
  if (typeof v === 'boolean') return false;
  if (typeof v === 'string') {
    const t = v.trim();
    if (t === '') return true;
    return /^-?\d+([.,]\d+)?$/.test(t.replace(/\s/g, ''));
  }
  return false;
}

function revisarNumericos(table, rows) {
  const def = COLUMNAS_NUMERICAS[table];
  if (!def) return [];
  const problemas = [];
  rows.forEach((row, i) => {
    const malas = [];
    for (const [col, enExcel] of Object.entries(def.num)) {
      if (!esNumeroAceptable(row[col])) malas.push({ col: enExcel, valor: row[col] });
    }
    if (!malas.length) return;
    const ident = def.id.map((k) => row[k]).find((v) => v !== null && v !== undefined && v !== '' && v !== 'SIN VENTA');
    problemas.push({
      tabla: table,
      hoja: def.hoja,
      fila: row.__fila || null,
      posicion: i + 1,
      ident: ident !== undefined ? String(ident) : null,
      malas,
    });
  });
  return problemas;
}

function mensajeDeProblemas(problemas) {
  const lineas = [];
  lineas.push('');
  lineas.push('='.repeat(64));
  lineas.push('  NO SE IMPORTO NADA');
  lineas.push('='.repeat(64));
  lineas.push('');
  lineas.push('  Hay celdas con texto donde la base espera un numero.');
  lineas.push('  Supabase rechazaria la carga completa, asi que se corta antes.');
  lineas.push('');
  const porHoja = {};
  for (const p of problemas) (porHoja[p.hoja] = porHoja[p.hoja] || []).push(p);
  for (const [hoja, lista] of Object.entries(porHoja)) {
    lineas.push(`  Hoja "${hoja}":`);
    for (const p of lista.slice(0, 15)) {
      const donde = p.fila ? `fila ${p.fila}` : `linea ${p.posicion} de la tabla`;
      const quien = p.ident ? ` (${p.ident})` : '';
      const qué = p.malas.map((m) => `${m.col} = ${JSON.stringify(m.valor)}`).join(', ');
      lineas.push(`    ${donde}${quien}  ->  ${qué}`);
    }
    if (lista.length > 15) lineas.push(`    ... y ${lista.length - 15} filas mas`);
    lineas.push('');
  }
  lineas.push('  El caso mas comun es "SIN VENTA": el BUSCARV de "Prod. Final" no');
  lineas.push('  encuentra la venta en "Ventas Finn" porque le cambiaron la');
  lineas.push('  presentacion o la cantidad, y la clave concatenada dejo de');
  lineas.push('  coincidir. Revisa esas ventas y volve a correr.');
  lineas.push('');
  lineas.push('  La base quedo intacta: no se envio nada.');
  return lineas.join('\n');
}

// ---------------- armado del SQL ----------------

function esc(v) {
  if (v === null || v === undefined) return 'NULL';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return 'NULL';
    return String(v);
  }
  let sv = String(v).trim();
  if (sv === '') return 'NULL';
  sv = sv.replace(/'/g, "''");
  return `'${sv}'`;
}

function insertStmt(table, columns, rows) {
  if (!rows.length) return null;
  const colsSql = columns.join(', ');
  const valuesSql = rows.map((row) => `(${columns.map((c) => esc(row[c])).join(', ')})`);
  return `INSERT INTO ${table} (${colsSql}) VALUES\n${valuesSql.join(',\n')};`;
}

function buildSql(xlsxPath) {
  const wb = xlsx.readFile(xlsxPath, { cellDates: true });
  OMITIDAS.ventasCanceladas = [];
  OMITIDAS.prodSinVenta = [];

  const ventas = extractVentas(wb);
  const balanza = extractBalanza(wb);
  const lotes = extractLotes(wb);
  const prodfinal = extractProdFinal(wb);
  const despachos = extractDespachos(wb);
  const { resumenMarca, bolsasList } = extractStock(wb);
  const { planCurado, weekly } = extractPlanCurado(wb);
  const lotesCurado = extractLotesCurado(wb);
  const variedades = extractVariedades(wb);
  const calendario = extractCalendario(wb);

  const blocks = [];
  const counts = {};
  const problemas = [];

  function addTable(table, columns, rows) {
    counts[table] = rows.length;
    problemas.push(...revisarNumericos(table, rows));
    // "WHERE true" es necesario: Supabase tiene activada la protección
    // "safeupdate", que bloquea cualquier DELETE/UPDATE sin WHERE (para
    // evitar borrados accidentales de toda la tabla). Con WHERE true se
    // sigue borrando todo, pero ya no dispara esa protección.
    blocks.push(`DELETE FROM ${table} WHERE true;`);
    const stmt = insertStmt(table, columns, rows);
    if (stmt) blocks.push(stmt);
  }

  addTable('variedades', ['variedad', 'ciclo', 'semilla', 'semillero', 'tecnologia'], variedades);

  addTable(
    'pedidos_venta',
    ['vta_id', 'fecha', 'cliente', 'categoria', 'semillero', 'variedad', 'cantidad', 'vendedor',
      'tratamiento', 'presentacion', 'fecha_entrega', 'kg_total', 'control_pendiente', 'linea_id'],
    ventas
  );

  addTable(
    'balanza',
    ['fecha', 'chofer', 'maquina', 'campo', 'variedad', 'categoria', 'neto', 'silo', 'merma', 'merma_viento'],
    balanza
  );

  addTable(
    'lotes',
    ['lote', 'procedencia', 'variedad', 'area', 'kg', 'kg_hr', 'paro', 'detalle_paro', 'fin', 'pureza', 'ok_calidad', 'pg', 'pg_curado', 'p1000'],
    lotes
  );

  addTable(
    'produccion_final',
    ['vta_id', 'venta', 'estado', 'cliente', 'variedad', 'presentacion', 'kg', 'lote_final', 'fecha_entrega',
      'linea_id', 'cantidad', 'despachado', 'pend_despacho'],
    prodfinal
  );

  addTable(
    'despachos',
    ['pedido', 'cliente', 'dia_carga', 'variedad', 'presentacion', 'cantidad', 'kg', 'lote_final', 'estado', 'camion'],
    despachos
  );

  // Cupos y Programacion de Camiones: cada vez que se actualiza el reporte de
  // Despachos, sincroniza las cargas automaticamente (crea las nuevas, marca
  // como despachadas las que ya salieron) dentro de la misma transaccion.
  blocks.push('SELECT cupos_sincronizar_desde_despachos();');

  addTable('stock_resumen_marca', ['marca', 'pintados', 'sin_pintar', 'nuevos'], resumenMarca);

  addTable('stock_bolsas_semilla', ['semilla', 'bolsas'], bolsasList);

  addTable(
    'plan_curado',
    ['variedad', 'tratamiento', 'pendiente_total', 'curado_disp', 'neto_a_curar'],
    planCurado
  );

  const pcw = weekly.filter((w) => w.semana && w.semana !== 'Sin fecha');
  addTable('plan_curado_semanal', ['semana', 'kg'], pcw);

  addTable(
    'lotes_curado',
    ['lote_intermedio', 'lote_final', 'variedad', 'categoria', 'semillero', 'kg_disp_inter',
      'kg_consumidos', 'area_int', 'kg_producidos', 'kg_disponibles', 'area_final', 'tratamiento',
      'complemento', 'hora_inicio', 'hora_fin', 'horas_prod', 'kg_hs'],
    lotesCurado
  );

  addTable(
    'calendario_produccion',
    ['fecha', 'dia', 'gessi_inicio', 'gessi_fin', 'horas_gessi', 'blomar_inicio', 'blomar_fin', 'horas_blomar', 'observacion'],
    calendario
  );

  if (problemas.length) {
    const e = new Error(mensajeDeProblemas(problemas));
    e.esProblemaDeDatos = true;
    throw e;
  }

  const sql = `BEGIN;\n\n${blocks.join('\n\n')}\n\nCOMMIT;\n`;
  return { sql, counts, omitidas: { ventasCanceladas: [...OMITIDAS.ventasCanceladas], prodSinVenta: [...OMITIDAS.prodSinVenta] } };
}

function stripTransactionWrapper(sql) {
  // build_sql() envuelve todo en BEGIN;...COMMIT; para cuando se pega a mano
  // en el editor SQL. Acá no hace falta: exec_import_sql ya corre dentro de
  // una sola transacción (la del llamado HTTP), así que si algo falla a
  // mitad de camino, Postgres deshace todo solo.
  let s = sql.trim();
  if (s.startsWith('BEGIN;')) s = s.slice('BEGIN;'.length);
  if (s.trimEnd().endsWith('COMMIT;')) s = s.trimEnd().slice(0, -'COMMIT;'.length);
  return s.trim();
}

function cargarConfigEnv() {
  // Lee panel-soja-web/config.env a mano (mismo archivo que usa el
  // servidor) sin depender de ningún paquete nuevo — así este script anda
  // apenas se copian los .js a la carpeta scripts/, sin instalar nada.
  if (process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY) return;
  const configPath = path.join(__dirname, '..', 'config.env');
  if (!fs.existsSync(configPath)) return;
  const contenido = fs.readFileSync(configPath, 'utf8');
  for (const linea of contenido.split('\n')) {
    const l = linea.trim();
    if (!l || l.startsWith('#') || !l.includes('=')) continue;
    const idx = l.indexOf('=');
    const clave = l.slice(0, idx).trim();
    const valor = l.slice(idx + 1).trim();
    if (process.env[clave] === undefined) process.env[clave] = valor;
  }
}

async function enviarASupabase(sql) {
  cargarConfigEnv();
  const supabaseUrl = process.env.SUPABASE_URL;
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!supabaseUrl || !serviceKey) {
    throw new Error('Faltan SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY en config.env');
  }
  const endpoint = `${supabaseUrl.replace(/\/$/, '')}/rest/v1/rpc/exec_import_sql`;
  const res = await fetch(endpoint, {
    method: 'POST',
    headers: {
      apikey: serviceKey,
      Authorization: `Bearer ${serviceKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ sql: stripTransactionWrapper(sql) }),
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Supabase respondió ${res.status}: ${body}`);
  }
}

module.exports = { buildSql, esc, dstr, fmtTimeCell, cleanHours, stripOrNone, stripTransactionWrapper,
  esNumeroAceptable, revisarNumericos, mensajeDeProblemas, COLUMNAS_NUMERICAS,
  esVentaCancelada, esFilaSinVenta };

// ---------------- CLI ----------------

async function main() {
  const xlsxPath = process.argv[2];
  if (!xlsxPath) {
    console.log('Uso: node importar_datos.js "SOJA CAMPAÑA 2526.xlsm"');
    process.exit(1);
  }
  if (!fs.existsSync(xlsxPath)) {
    console.log(`No encuentro el archivo: ${xlsxPath}`);
    process.exit(1);
  }

  const t0 = Date.now();
  console.log(`Leyendo ${path.basename(xlsxPath)} y armando el SQL...`);
  let sql, counts, omitidas;
  try {
    ({ sql, counts, omitidas } = buildSql(xlsxPath));
  } catch (err) {
    if (err.esProblemaDeDatos) {
      console.log(err.message);
      process.exit(1);
    }
    throw err;
  }
  const t1 = Date.now();
  console.log(`  listo en ${((t1 - t0) / 1000).toFixed(1)}s. Filas por tabla:`);
  for (const [t, n] of Object.entries(counts)) console.log(`    ${t}: ${n}`);

  const vc = omitidas.ventasCanceladas, pv = omitidas.prodSinVenta;
  if (vc.length || pv.length) {
    console.log('');
    console.log('  Ventas canceladas (cantidad 0, las que se ven tachadas en el Excel):');
    if (vc.length) {
      console.log(`    ${vc.length} en "Ventas Finn" -> no van a pedidos_venta`);
      for (const o of vc) console.log(`       fila ${o.fila}  ${o.vta_id}  (${o.cliente || 'sin cliente'})`);
    }
    if (pv.length) {
      console.log(`    ${pv.length} en "Prod. Final" -> no van a produccion_final (su venta esta cancelada)`);
      for (const o of pv) console.log(`       fila ${o.fila}  venta ${o.venta}`);
    }
  }

  console.log('Enviando a Supabase...');
  try {
    await enviarASupabase(sql);
  } catch (err) {
    console.log(`ERROR: ${err.message}`);
    process.exit(1);
  }
  const t2 = Date.now();
  console.log(`Listo. Supabase actualizado en ${((t2 - t1) / 1000).toFixed(1)}s (total ${((t2 - t0) / 1000).toFixed(1)}s).`);
  console.log('El panel ya está al día — no hace falta hacer nada más.');
}

if (require.main === module) {
  main();
}
