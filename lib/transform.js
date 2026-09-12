// Convierte las filas crudas de las tablas de Supabase en el mismo formato
// (DATA) que ya espera el front-end del panel — agregados y KPIs incluidos.
// Es la versión en JS de lo que full_extract.py calculaba a partir del Excel.

function sum(arr, pick) {
  return arr.reduce((acc, x) => acc + (pick(x) || 0), 0);
}

function groupSum(arr, keyFn, valFn) {
  const out = {};
  for (const x of arr) {
    const k = keyFn(x);
    out[k] = (out[k] || 0) + (valFn(x) || 0);
  }
  return out;
}

function avg(values) {
  return values.length ? sum(values, (v) => v) / values.length : 0;
}

// -------- Despacho: normalización de "día de carga" --------
// En el Excel, la columna "Dia de carga" a veces se carga como fecha real
// (Excel/el importador la guarda como 'YYYY-MM-DD') y a veces como texto
// libre tecleado a mano ('martes 15/09', 'miercoles 16-09' — con '/' o '-',
// con o sin día de la semana bien escrito). Eso hace que la tabla de
// Despachos del panel se vea con dos formatos distintos mezclados. Acá se
// normalizan los dos casos a un único formato de exhibición: "Dia dd-mm"
// (día de la semana capitalizado, sin tilde, y fecha con guion) — sin tocar
// el valor crudo que vive en Supabase, solo cómo se muestra en el panel.
const DIAS_ES = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];

function stripAccents(s) {
  return s.normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function capitalize(s) {
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

function formatDiaCarga(v) {
  if (v === null || v === undefined) return v;
  const s = String(v).trim();
  if (!s) return v;

  // Caso 1: fecha real (guardada como 'YYYY-MM-DD...') — se calcula el día
  // de la semana a partir de la fecha.
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    const [, y, m, d] = isoMatch;
    const date = new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
    const dia = stripAccents(DIAS_ES[date.getUTCDay()]);
    return `${capitalize(dia)} ${d}-${m}`;
  }

  // Caso 2: texto libre con día de la semana + fecha ('martes 15/09',
  // 'miercoles 16-09') — se pareja el separador y se capitaliza el día.
  const freeMatch = s.match(/^([a-zA-ZÀ-ÿ]+)\s+(\d{1,2})[\/\-](\d{1,2})/);
  if (freeMatch) {
    const [, dia, d, m] = freeMatch;
    const dd = d.padStart(2, '0');
    const mm = m.padStart(2, '0');
    return `${capitalize(stripAccents(dia))} ${dd}-${mm}`;
  }

  // Formato inesperado: se deja tal cual en vez de ocultar un dato raro.
  return s;
}

// results: { pedidos_venta, balanza, lotes, produccion_final, despachos,
//            stock_resumen_marca, stock_bolsas_semilla, plan_curado,
//            plan_curado_semanal, lotes_curado, calendario_produccion }
// cada valor es el array de filas tal cual las devuelve Supabase (select('*')).
function transformRows(results) {
  const data = {};

  // -------- Ventas Finn --------
  const ventas = results.pedidos_venta.map((r) => ({
    vta_id: r.vta_id,
    fecha: r.fecha,
    cliente: r.cliente,
    categoria: r.categoria,
    semillero: r.semillero,
    variedad: r.variedad,
    cantidad: r.cantidad,
    vendedor: r.vendedor,
    tratamiento: r.tratamiento,
    pack: r.presentacion,
    fecha_entrega: r.fecha_entrega,
    kg_total: r.kg_total,
    control_pendiente: r.control_pendiente,
  }));
  data.ventas = ventas;
  data.kpi_ventas = {
    total_kg: sum(ventas, (v) => v.kg_total),
    pedidos: new Set(ventas.map((v) => v.vta_id)).size,
    clientes: new Set(ventas.map((v) => v.cliente)).size,
  };
  data.por_variedad = groupSum(ventas, (v) => v.variedad, (v) => v.kg_total);
  data.por_vendedor = groupSum(ventas, (v) => v.vendedor, (v) => v.kg_total);
  data.por_fecha = groupSum(ventas, (v) => v.fecha, (v) => v.kg_total);

  // -------- Balanza --------
  const balanza = results.balanza.map((r) => ({
    fecha: r.fecha,
    chofer: r.chofer,
    maquina: r.maquina,
    campo: r.campo,
    variedad: r.variedad,
    categoria: r.categoria,
    neto: r.neto,
    silo: r.silo,
    merma: r.merma || 0,
    merma_viento: r.merma_viento || 0,
  }));
  data.balanza = balanza;
  const totalNeto = sum(balanza, (b) => b.neto);
  const totalMerma = sum(balanza, (b) => b.merma);
  const totalMv = sum(balanza, (b) => b.merma_viento);
  data.kpi_balanza = {
    total_neto: totalNeto,
    camiones: balanza.length,
    campos: new Set(balanza.map((b) => b.campo)).size,
    merma_prom_pct: totalNeto ? ((totalMerma + totalMv) / totalNeto) * 100 : 0,
  };
  data.bal_por_variedad = groupSum(balanza, (b) => b.variedad, (b) => b.neto);
  data.bal_por_campo = groupSum(balanza, (b) => b.campo, (b) => b.neto);
  data.bal_por_fecha = groupSum(balanza, (b) => b.fecha, (b) => b.neto);

  // -------- Intermedio (lotes) --------
  const lotes = results.lotes.map((r) => ({
    lote: r.lote,
    procedencia: r.procedencia,
    variedad: r.variedad,
    area: r.area,
    kg: r.kg,
    kg_hr: r.kg_hr,
    paro: r.paro,
    detalle_paro: r.detalle_paro,
    fin: r.fin,
    pureza: r.pureza,
    ok_calidad: r.ok_calidad,
    pg: r.pg,
  }));
  data.lotes = lotes;
  const kghrVals = lotes.map((l) => l.kg_hr).filter((v) => v !== null && v !== undefined);
  data.kpi_planta = {
    kg_procesado: sum(lotes, (l) => l.kg),
    rendimiento_prom: avg(kghrVals),
    paradas: lotes.filter((l) => l.paro).length,
    lotes: lotes.length,
  };
  data.int_por_variedad_kg = groupSum(lotes, (l) => l.variedad, (l) => l.kg);
  const kghrByVar = {};
  for (const l of lotes) {
    if (l.kg_hr !== null && l.kg_hr !== undefined) {
      (kghrByVar[l.variedad] = kghrByVar[l.variedad] || []).push(l.kg_hr);
    }
  }
  data.int_rendimiento_var = Object.fromEntries(
    Object.entries(kghrByVar).map(([k, v]) => [k, avg(v)])
  );
  const calidadCounts = {};
  for (const l of lotes) {
    const k = l.pureza || 'Sin dato';
    calidadCounts[k] = (calidadCounts[k] || 0) + 1;
  }
  data.calidad_counts = calidadCounts;

  // -------- Prod. Final --------
  const prodfinal = results.produccion_final.map((r) => ({
    vta_id: r.vta_id,
    venta: r.venta,
    estado: r.estado,
    cliente: r.cliente,
    variedad: r.variedad,
    presentacion: r.presentacion,
    kg: r.kg,
    lote_final: r.lote_final,
    fecha_entrega: r.fecha_entrega,
  }));
  data.prodfinal = prodfinal;
  const estados = {};
  for (const p of prodfinal) {
    const k = p.estado;
    if (!estados[k]) estados[k] = { count: 0, kg: 0 };
    estados[k].count += 1;
    estados[k].kg += p.kg || 0;
  }
  data.estados_pedidos = estados;

  // -------- Despacho --------
  data.despachos = results.despachos.map((r) => ({
    pedido: r.pedido,
    cliente: r.cliente,
    dia_carga: formatDiaCarga(r.dia_carga),
    variedad: r.variedad,
    presentacion: r.presentacion,
    cantidad: r.cantidad,
    kg: r.kg,
    lote_final: r.lote_final,
    estado: r.estado,
  }));

  // -------- Stock --------
  data.resumen_marca = results.stock_resumen_marca.map((r) => ({
    marca: r.marca,
    pintados: r.pintados,
    sin_pintar: r.sin_pintar,
    nuevos: r.nuevos,
  }));
  data.bolsas_by_semilla = Object.fromEntries(
    results.stock_bolsas_semilla.map((r) => [r.semilla, r.bolsas])
  );

  // -------- Plan Curado --------
  data.plan_curado = results.plan_curado.map((r) => ({
    variedad: r.variedad,
    tratamiento: r.tratamiento,
    pendiente_total: r.pendiente_total,
    curado_disp: r.curado_disp,
    neto_a_curar: r.neto_a_curar,
  }));
  data.plan_curado_weekly = results.plan_curado_semanal
    .slice()
    .sort((a, b) => (a.semana || '').localeCompare(b.semana || ''))
    .map((r) => ({ semana: r.semana, kg: r.kg || 0 }));

  // -------- Intermedio curado --------
  const lotesCurado = results.lotes_curado.map((r) => ({
    lote_intermedio: r.lote_intermedio,
    lote_final: r.lote_final,
    variedad: r.variedad,
    categoria: r.categoria,
    semillero: r.semillero,
    kg_disp_inter: r.kg_disp_inter,
    kg_consumidos: r.kg_consumidos,
    area_int: r.area_int,
    kg_producidos: r.kg_producidos,
    kg_disponibles: r.kg_disponibles,
    area_final: r.area_final,
    tratamiento: r.tratamiento,
    complemento: r.complemento,
    hora_inicio: r.hora_inicio,
    hora_fin: r.hora_fin,
    horas_prod: r.horas_prod,
    kg_hs: r.kg_hs,
  }));
  data.lotes_curado = lotesCurado;
  const kghsVals = lotesCurado.map((l) => l.kg_hs).filter(Boolean);
  data.kpi_curado = {
    lotes: lotesCurado.length,
    kg_producidos: sum(lotesCurado, (l) => l.kg_producidos),
    rendimiento_prom: avg(kghsVals),
  };

  // -------- Calendario Producción --------
  data.calendario_total_gessi = sum(results.calendario_produccion, (r) => r.horas_gessi);
  data.calendario_total_blomar = sum(results.calendario_produccion, (r) => r.horas_blomar);

  return data;
}

module.exports = { transformRows, sum, groupSum, avg, formatDiaCarga };
