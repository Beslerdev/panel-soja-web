// -------- Vencidos y por vencer --------
// Cruza Ventas Finn (pedidos_venta) con Prod. Final (produccion_final) y
// Despacho (despachos) para saber, por cada línea de venta, cuánto falta
// despachar, cuánto de eso ya está terminado en planta y cuánto falta
// producir, y si la fecha de entrega ya pasó o vence en los próximos días.
//
// Es la misma lógica de los reportes "Vencidos Soja" (PDF/Excel):
//   - kg despachados  = Prod. Final "Despachado" (unidades) × kg por unidad
//   - kg listos       = filas "2- Terminado": "Pend. Despacho" × kg por unidad
//   - kg en proceso   = filas "3- En Proceso": "Pend. Despacho" × kg por unidad
//   - kg sin producir = pendiente − listos − en proceso
// Nunca se cuenta más que lo pedido: si Prod. Final tiene filas duplicadas o
// con kilos de más, se topea en el kg total de la venta.
//
// La unión es por la clave de línea: columna "id" de Ventas Finn =
// columna "Concatenado" de Prod. Final (linea_id en las dos tablas).

const DIAS_POR_VENCER = 7;

const ESTADO = {
  VENCIDO: 'Vencido',
  POR_VENCER: 'Por vencer',
  A_TIEMPO: 'A tiempo',
  ENTREGADO: 'Entregado',
};

// Fecha de hoy en Argentina, como 'YYYY-MM-DD'. Render corre en UTC: sin
// esto, entre las 21 y las 24 hs el panel ya contaría el día siguiente.
function hoyArgentina(now = new Date()) {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Argentina/Buenos_Aires',
    year: 'numeric', month: '2-digit', day: '2-digit',
  }).format(now);
}

function diasEntre(desde, hasta) {
  const a = Date.UTC(+desde.slice(0, 4), +desde.slice(5, 7) - 1, +desde.slice(8, 10));
  const b = Date.UTC(+hasta.slice(0, 4), +hasta.slice(5, 7) - 1, +hasta.slice(8, 10));
  return Math.round((b - a) / 86400000);
}

const num = (v) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};

const clave = (v) => String(v ?? '').trim().toUpperCase();

const ddmm = (iso) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : '');

const nombrePedido = (vtaId) =>
  String(vtaId || '')
    .replace('Pedido de Venta - Semillas - ', 'PV ')
    .replace('PEDVTA - ', 'PEDVTA ');

// Programación de carga cargada en la hoja Despacho, para lo que todavía no
// salió: fecha del día de carga o el texto ("A confirmar por el cliente").
function programacionPorVenta(despachos) {
  const out = new Map();
  for (const d of despachos) {
    if (!d.pedido || d.estado === 'Despachado') continue;
    const raw = d.dia_carga == null ? '' : String(d.dia_carga).trim();
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    let txt = iso ? `${iso[3]}/${iso[2]}` : raw;
    if (d.estado === 'En Planta') txt = txt ? `${txt} · en planta` : 'en planta';
    if (!txt) continue;
    const k = clave(d.pedido);
    const lista = out.get(k) || [];
    if (!lista.includes(txt)) lista.push(txt);
    out.set(k, lista);
  }
  return out;
}

function calcularLineas(results, hoy) {
  // Prod. Final agrupado por línea de venta
  const prod = new Map();
  for (const p of results.produccion_final) {
    if (!p.linea_id) continue;
    const cant = num(p.cantidad);
    const ku = cant ? num(p.kg) / cant : 0;
    const pend = num(p.pend_despacho) * ku;
    const g = prod.get(p.linea_id) || { desp: 0, listo: 0, proc: 0, ventas: [] };
    g.desp += num(p.despachado) * ku;
    if (p.estado === '2- Terminado') g.listo += pend;
    if (p.estado === '3- En Proceso') g.proc += pend;
    if (p.venta) g.ventas.push(p.venta);
    prod.set(p.linea_id, g);
  }
  const prog = programacionPorVenta(results.despachos);

  const lineas = [];
  for (const v of results.pedidos_venta) {
    const kgPedido = num(v.kg_total);
    if (kgPedido <= 0 || !v.fecha_entrega) continue;
    const g = prod.get(v.linea_id) || { desp: 0, listo: 0, proc: 0, ventas: [] };
    const desp = Math.min(g.desp, kgPedido);
    const pend = Math.max(kgPedido - desp, 0);
    const listo = Math.min(g.listo, pend);
    const proc = Math.min(g.proc, pend - listo);
    const sinProd = Math.max(pend - listo - proc, 0);
    const fecha = String(v.fecha_entrega).slice(0, 10);
    const dias = diasEntre(fecha, hoy); // > 0: días de atraso
    let estado = ESTADO.A_TIEMPO;
    if (pend <= 0) estado = ESTADO.ENTREGADO;
    else if (dias > 0) estado = ESTADO.VENCIDO;
    else if (dias >= -DIAS_POR_VENCER) estado = ESTADO.POR_VENCER;

    const programacion = [];
    for (const vta of g.ventas) {
      for (const t of prog.get(clave(vta)) || []) if (!programacion.includes(t)) programacion.push(t);
    }

    lineas.push({
      pedido: nombrePedido(v.vta_id),
      cliente: v.cliente,
      variedad: v.variedad,
      envase: v.presentacion,
      vendedor: v.vendedor,
      fecha_entrega: fecha,
      dias_atraso: dias,
      estado,
      kg_pedido: Math.round(kgPedido),
      kg_despachado: Math.round(desp),
      kg_pendiente: Math.round(pend),
      kg_listo: Math.round(listo),
      kg_en_proceso: Math.round(proc),
      kg_sin_producir: Math.round(sinProd),
      programacion: programacion.join(' | '),
    });
  }
  return lineas;
}

// Lista de llamados: un renglón por cliente con algo vencido o por vencer.
// Orden: 1) vencido con carga lista (más kg listos primero), 2) vencido sin
// carga lista, 3) por vencer con carga lista, 4) por vencer sin carga.
function porCliente(lineas) {
  const g = new Map();
  for (const r of lineas) {
    if (r.estado !== ESTADO.VENCIDO && r.estado !== ESTADO.POR_VENCER) continue;
    const c = g.get(r.cliente) || {
      cliente: r.cliente, kg_vencido: 0, kg_por_vencer: 0, kg_listo: 0, kg_listo_vencido: 0,
      kg_sin_producir: 0, dias_atraso: 0, prox_entrega: '', pedidos: new Set(), variedades: new Set(), prog: new Set(),
    };
    if (r.estado === ESTADO.VENCIDO) {
      c.kg_vencido += r.kg_pendiente;
      c.kg_listo_vencido += r.kg_listo;
      c.dias_atraso = Math.max(c.dias_atraso, r.dias_atraso);
    } else {
      c.kg_por_vencer += r.kg_pendiente;
      if (!c.prox_entrega || r.fecha_entrega < c.prox_entrega) c.prox_entrega = r.fecha_entrega;
    }
    c.kg_listo += r.kg_listo;
    c.kg_sin_producir += r.kg_sin_producir + r.kg_en_proceso;
    c.pedidos.add(r.pedido);
    c.variedades.add(r.variedad);
    if (r.programacion) r.programacion.split(' | ').forEach((p) => c.prog.add(p));
    g.set(r.cliente, c);
  }

  const filas = [...g.values()].map((c) => {
    const aConfirmar = [...c.prog].some((p) => p.startsWith('A confirmar'));
    const grupo = c.kg_vencido > 0 ? (c.kg_listo_vencido > 0 ? 1 : 2) : (c.kg_listo > 0 ? 3 : 4);
    let accion;
    if (grupo === 1 && c.kg_listo_vencido < c.kg_vencido * 0.25) accion = 'Llamar · liberar lo poco que hay listo y acordar fecha del resto';
    else if (grupo === 1) accion = aConfirmar ? 'Llamar ya · carga lista esperando confirmación' : 'Llamar ya · confirmar recepción de carga lista';
    else if (grupo === 2) accion = c.kg_listo > 0 ? 'Vencido · acordar fecha (hay carga lista de lo que vence esta semana)' : 'Vencido · avisar atraso y acordar nueva fecha';
    else if (grupo === 3) accion = 'Por vencer · confirmar día de carga';
    else accion = 'Por vencer · avisar si no llegamos';
    return {
      cliente: c.cliente,
      grupo,
      accion,
      kg_vencido: c.kg_vencido,
      kg_por_vencer: c.kg_por_vencer,
      kg_listo: c.kg_listo,
      kg_sin_producir: c.kg_sin_producir,
      dias_atraso: c.dias_atraso,
      prox_entrega: c.prox_entrega,
      pedidos: [...c.pedidos].join(', '),
      variedades: [...c.variedades].join(', '),
      programacion: [...c.prog].join(' | ') || 'Sin programar',
    };
  });
  filas.sort((a, b) =>
    a.grupo - b.grupo
    || ((a.grupo === 1 || a.grupo === 3) ? b.kg_listo - a.kg_listo : 0)
    || (b.kg_vencido + b.kg_por_vencer) - (a.kg_vencido + a.kg_por_vencer));
  filas.forEach((f, i) => { f.prioridad = i + 1; });
  return filas;
}

function porVariedad(lineas) {
  const g = new Map();
  for (const r of lineas) {
    if (r.estado !== ESTADO.VENCIDO && r.estado !== ESTADO.POR_VENCER) continue;
    const v = g.get(r.variedad) || { variedad: r.variedad, kg_vencido: 0, kg_listo: 0, kg_resto: 0, kg_por_vencer: 0, clientes: new Set() };
    if (r.estado === ESTADO.VENCIDO) {
      v.kg_vencido += r.kg_pendiente;
      v.kg_listo += r.kg_listo;
      v.kg_resto += r.kg_pendiente - r.kg_listo;
      v.clientes.add(r.cliente);
    } else {
      v.kg_por_vencer += r.kg_pendiente;
    }
    g.set(r.variedad, v);
  }
  return [...g.values()]
    .map((v) => ({ ...v, clientes: v.clientes.size }))
    .sort((a, b) => b.kg_vencido - a.kg_vencido || b.kg_por_vencer - a.kg_por_vencer);
}

function porFecha(lineas) {
  const g = new Map();
  for (const r of lineas) {
    if (r.kg_pendiente <= 0) continue;
    const f = g.get(r.fecha_entrega) || { fecha: r.fecha_entrega, kg_listo: 0, kg_resto: 0 };
    f.kg_listo += r.kg_listo;
    f.kg_resto += r.kg_pendiente - r.kg_listo;
    g.set(r.fecha_entrega, f);
  }
  return [...g.values()].sort((a, b) => a.fecha.localeCompare(b.fecha));
}

function calcularVencidos(results, hoy = hoyArgentina()) {
  const lineas = calcularLineas(results, hoy);
  const venc = lineas.filter((r) => r.estado === ESTADO.VENCIDO);
  const prox = lineas.filter((r) => r.estado === ESTADO.POR_VENCER);
  const s = (rs, k) => rs.reduce((a, r) => a + r[k], 0);
  const hasta = new Date(Date.UTC(+hoy.slice(0, 4), +hoy.slice(5, 7) - 1, +hoy.slice(8, 10) + DIAS_POR_VENCER))
    .toISOString().slice(0, 10);
  return {
    hoy,
    hasta,
    dias_por_vencer: DIAS_POR_VENCER,
    resumen: {
      kg_vencido: s(venc, 'kg_pendiente'),
      kg_listo_vencido: s(venc, 'kg_listo'),
      kg_sin_producir_vencido: s(venc, 'kg_sin_producir') + s(venc, 'kg_en_proceso'),
      clientes_vencidos: new Set(venc.map((r) => r.cliente)).size,
      pedidos_vencidos: new Set(venc.map((r) => r.pedido)).size,
      kg_por_vencer: s(prox, 'kg_pendiente'),
      kg_listo_por_vencer: s(prox, 'kg_listo'),
      // Si la columna linea_id está vacía (importador viejo), no hay cruce
      // posible con Prod. Final: el panel lo avisa en vez de mostrar todo
      // como "sin producir".
      sin_cruce: !results.produccion_final.some((p) => p.linea_id),
    },
    clientes: porCliente(lineas),
    variedades: porVariedad(lineas),
    fechas: porFecha(lineas),
    lineas: lineas.filter((r) => r.estado !== ESTADO.ENTREGADO),
  };
}

module.exports = { calcularVencidos, hoyArgentina, ESTADO, DIAS_POR_VENCER, ddmm };
