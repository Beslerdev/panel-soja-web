// lib/cupos.js
//
// Motor de reglas de negocio para el Sistema de Cupos y Programación de
// Camiones/Carros (DS Hnos). Esta capa NO valida capacidad "a mano": delega
// toda la parte crítica (capacidad, atomicidad, liberar el cupo viejo al
// reprogramar, bloquear cambios de modalidad incompatibles) a las funciones
// `cupos_*` de Postgres (SECURITY DEFINER) creadas en la migración
// "cupos_programacion_camiones". Acá solo viven: el armado de las llamadas,
// el algoritmo de "buscar el próximo día disponible" (que si no, haría un
// viaje HTTP por cada día candidato) y la propuesta de fecha previa a
// confirmar.
//
// Uso desde server.js:
//   const { crearServicioCupos } = require('./lib/cupos');
//   const cupos = crearServicioCupos(supabaseAdmin);
//   await cupos.resumenSemana('2026-10-26');

// ---------------- Utilidades de fecha (todo en UTC, formato 'YYYY-MM-DD') ----------------
// Se evita el objeto Date salvo para aritmética en UTC: así no hay sorpresas
// de huso horario entre el servidor y Postgres, que también trabaja en UTC
// para el tipo `date`.

function parseISODate(s) {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

function formatISODate(d) {
  return d.toISOString().slice(0, 10);
}

function hoyISO() {
  return formatISODate(new Date());
}

function addDays(fechaISO, n) {
  const d = parseISODate(fechaISO);
  d.setUTCDate(d.getUTCDate() + n);
  return formatISODate(d);
}

// 0 = domingo, 1 = lunes, ... 6 = sábado (igual que extract(dow) de Postgres).
function diaSemana(fechaISO) {
  return parseISODate(fechaISO).getUTCDay();
}

function esDomingo(fechaISO) {
  return diaSemana(fechaISO) === 0;
}

// Lunes de la semana que contiene fechaISO.
function lunesDeSemana(fechaISO) {
  const dow = diaSemana(fechaISO);
  const offset = (dow + 6) % 7; // lunes=0 ... domingo=6
  return addDays(fechaISO, -offset);
}

const NOMBRES_DIA = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function nombreDia(fechaISO) {
  return NOMBRES_DIA[diaSemana(fechaISO)];
}

// ---------------- Helper genérico para llamar a los RPC ----------------

async function llamarRpc(supabaseAdmin, nombreFn, params) {
  const { data, error } = await supabaseAdmin.rpc(nombreFn, params);
  if (error) {
    // Los `raise exception` de Postgres llegan en error.message, ya en
    // español y listos para mostrar al usuario (son los mensajes que
    // escribimos a mano en la migración).
    throw new Error(error.message || `Error ejecutando ${nombreFn}`);
  }
  return data;
}

// ---------------- Lecturas ----------------

async function resumenSemana(supabaseAdmin, lunes) {
  const lunesFinal = lunes || lunesDeSemana(hoyISO());
  return llamarRpc(supabaseAdmin, 'cupos_resumen_semana', { p_lunes: lunesFinal });
}

async function capacidadRango(supabaseAdmin, desde, hasta) {
  const { data, error } = await supabaseAdmin.rpc('cupos_capacidad_rango', { p_desde: desde, p_hasta: hasta });
  if (error) throw new Error(`Error consultando el rango de capacidad: ${error.message}`);
  return data;
}

async function capacidadDia(supabaseAdmin, fecha, tipoEquipo) {
  return llamarRpc(supabaseAdmin, 'cupos_capacidad', { p_fecha: fecha, p_tipo: tipoEquipo });
}

async function obtenerDia(supabaseAdmin, fecha) {
  const [capCamion, capCarro] = await Promise.all([
    capacidadDia(supabaseAdmin, fecha, 'camion'),
    capacidadDia(supabaseAdmin, fecha, 'carro'),
  ]);
  const { data: cargas, error } = await supabaseAdmin
    .from('cupos_cargas')
    .select('*')
    .eq('fecha_asignada', fecha)
    .in('estado', ['programado', 'despachado'])
    .order('tipo_equipo', { ascending: true })
    .order('created_at', { ascending: true });
  if (error) throw new Error(`Error leyendo las cargas del día: ${error.message}`);
  return { fecha, dia_semana: nombreDia(fecha), camion: capCamion, carro: capCarro, cargas };
}

async function obtenerCarga(supabaseAdmin, cargaId) {
  const { data, error } = await supabaseAdmin.from('cupos_cargas').select('*').eq('id', cargaId).single();
  if (error || !data) throw new Error('La carga no existe.');
  return data;
}

async function listarCargas(supabaseAdmin, filtros = {}) {
  let q = supabaseAdmin.from('cupos_cargas').select('*').order('created_at', { ascending: false });
  if (filtros.estado) q = q.eq('estado', filtros.estado);
  if (filtros.fechaDesde) q = q.gte('fecha_asignada', filtros.fechaDesde);
  if (filtros.fechaHasta) q = q.lte('fecha_asignada', filtros.fechaHasta);
  if (filtros.pedido) q = q.ilike('pedido', `%${filtros.pedido}%`);
  const { data, error } = await q;
  if (error) throw new Error(`Error listando las cargas: ${error.message}`);
  return data;
}

async function historialCarga(supabaseAdmin, cargaId) {
  const { data, error } = await supabaseAdmin
    .from('cupos_historial')
    .select('*')
    .eq('carga_id', cargaId)
    .order('created_at', { ascending: false });
  if (error) throw new Error(`Error leyendo el historial: ${error.message}`);
  return data;
}

async function historialReciente(supabaseAdmin, limite = 100) {
  const { data, error } = await supabaseAdmin
    .from('cupos_historial')
    .select('*')
    .order('created_at', { ascending: false })
    .limit(limite);
  if (error) throw new Error(`Error leyendo el historial: ${error.message}`);
  return data;
}

// ---------------- Búsqueda del próximo día disponible ----------------
//
// Recorre fechas en orden cronológico desde `fechaInicio`, en tandas de 6
// semanas por llamada (un solo viaje HTTP cubre 6 semanas en vez de 1 por
// día), hasta `horizonteSemanas` semanas. Los domingos ya vienen excluidos
// por cupos_capacidad_rango. Un día sin modalidad configurada (`modalidad ===
// null`) se considera no operativo y se salta. Devuelve el primer día con al
// menos un cupo libre del tipo pedido, o null si no hay nada en el horizonte.

async function buscarProximoDisponible(supabaseAdmin, fechaInicio, tipoEquipo, opts = {}) {
  const horizonteSemanas = opts.horizonteSemanas || 12;
  const diasPorTanda = 42; // 6 semanas
  let desde = fechaInicio;
  let semanasRecorridas = 0;

  while (semanasRecorridas < horizonteSemanas) {
    const hasta = addDays(desde, diasPorTanda - 1);
    const fila_rango = await capacidadRango(supabaseAdmin, desde, hasta);
    for (const fila of fila_rango) {
      if (fila.fecha < fechaInicio) continue;
      if (!fila.modalidad) continue; // día sin modalidad = no operativo
      const maximo = tipoEquipo === 'carro' ? fila.carro_maximo : fila.camion_maximo;
      const ocupados = tipoEquipo === 'carro' ? fila.carro_ocupados : fila.camion_ocupados;
      const disponibles = maximo - ocupados;
      if (disponibles > 0) {
        return { fecha: fila.fecha, disponibles, maximo, modalidad: fila.modalidad };
      }
    }
    desde = addDays(hasta, 1);
    semanasRecorridas += diasPorTanda / 7;
  }
  return null;
}

// ---------------- Propuesta / confirmación de asignación ----------------
//
// Flujo en dos pasos pedido por el enunciado: "proponer" es de solo lectura
// (no escribe nada) y "confirmar" vuelve a validar en el servidor (dentro de
// cupos_asignar, con lock) antes de escribir — así nunca se confirma sobre
// una disponibilidad que ya cambió entre que se propuso y que el usuario
// aceptó.

async function proponerAsignacion(supabaseAdmin, { cargaId, fechaDeseada }) {
  const carga = await obtenerCarga(supabaseAdmin, cargaId);
  if (!['pendiente', 'programado'].includes(carga.estado)) {
    throw new Error(`La carga está "${carga.estado}": no se puede programar sin reabrirla primero.`);
  }

  const hoy = hoyISO();
  let fechaBase = fechaDeseada || carga.fecha_deseada || hoy;
  if (fechaBase < hoy) fechaBase = hoy; // no se programa en el pasado
  if (esDomingo(fechaBase)) fechaBase = addDays(fechaBase, 1); // el domingo nunca es operativo

  const capExacta = await capacidadDia(supabaseAdmin, fechaBase, carga.tipo_equipo);
  if (capExacta.operativo && capExacta.disponibles > 0) {
    return {
      ok: true,
      exacta: true,
      carga_id: cargaId,
      fecha_propuesta: fechaBase,
      disponibles: capExacta.disponibles,
      maximo: capExacta.maximo,
      modalidad: capExacta.modalidad,
    };
  }

  const proxima = await buscarProximoDisponible(supabaseAdmin, fechaBase, carga.tipo_equipo);
  if (!proxima) {
    return {
      ok: false,
      carga_id: cargaId,
      motivo: capExacta.operativo
        ? `El ${fechaBase} está completo y no se encontró disponibilidad en las próximas semanas.`
        : `El ${fechaBase} no tiene una modalidad configurada y no se encontró disponibilidad en las próximas semanas.`,
    };
  }

  return {
    ok: true,
    exacta: false,
    carga_id: cargaId,
    fecha_solicitada: fechaBase,
    fecha_propuesta: proxima.fecha,
    disponibles: proxima.disponibles,
    maximo: proxima.maximo,
    modalidad: proxima.modalidad,
  };
}

async function confirmarAsignacion(supabaseAdmin, { cargaId, fecha, usuario, nombre }) {
  return llamarRpc(supabaseAdmin, 'cupos_asignar', {
    p_carga_id: cargaId,
    p_fecha: fecha,
    p_usuario: usuario,
    p_nombre: nombre,
  });
}

// ---------------- Cancelar / despachar / reabrir ----------------

async function cancelar(supabaseAdmin, { cargaId, usuario, nombre, motivo }) {
  return llamarRpc(supabaseAdmin, 'cupos_cancelar', {
    p_carga_id: cargaId,
    p_usuario: usuario,
    p_nombre: nombre,
    p_motivo: motivo || null,
  });
}

async function despachar(supabaseAdmin, { cargaId, usuario, nombre }) {
  return llamarRpc(supabaseAdmin, 'cupos_despachar', {
    p_carga_id: cargaId,
    p_usuario: usuario,
    p_nombre: nombre,
  });
}

async function reabrir(supabaseAdmin, { cargaId, usuario, nombre, motivo }) {
  return llamarRpc(supabaseAdmin, 'cupos_reabrir', {
    p_carga_id: cargaId,
    p_usuario: usuario,
    p_nombre: nombre,
    p_motivo: motivo || null,
  });
}

// ---------------- Cambio de modalidad ----------------
// Puede devolver { ok:false, bloqueos:[...] } sin lanzar error — eso es un
// resultado normal, no una falla, así que el llamador lo distingue mirando
// `ok`, no con try/catch.

async function cambiarModalidad(supabaseAdmin, { fecha, modalidad, usuario, nombre }) {
  return llamarRpc(supabaseAdmin, 'cupos_cambiar_modalidad', {
    p_fecha: fecha,
    p_nueva_modalidad: modalidad,
    p_usuario: usuario,
    p_nombre: nombre,
  });
}

// ---------------- Alta de una carga ----------------

async function crearCarga(supabaseAdmin, { pedido, cliente, producto, tipoEquipo, patente, fechaDeseada, usuario, nombre, observaciones }) {
  if (!pedido || !String(pedido).trim()) throw new Error('El número de pedido es obligatorio.');
  if (!['camion', 'carro'].includes(tipoEquipo)) throw new Error('El tipo de equipo debe ser "camion" o "carro".');

  const { data, error } = await supabaseAdmin
    .from('cupos_cargas')
    .insert({
      pedido: String(pedido).trim(),
      cliente: cliente || null,
      producto: producto || null,
      tipo_equipo: tipoEquipo,
      patente: patente || null,
      fecha_deseada: fechaDeseada || null,
      estado: 'pendiente',
      observaciones: observaciones || null,
      usuario_usuario: usuario,
      usuario_nombre: nombre,
    })
    .select()
    .single();

  if (error) {
    if (error.code === '23505') {
      throw new Error(
        `Ya existe una programación activa para el pedido "${pedido}". Cancelá o reprogramá la existente antes de crear una nueva.`
      );
    }
    throw new Error(`Error creando la carga: ${error.message}`);
  }

  await supabaseAdmin.from('cupos_historial').insert({
    carga_id: data.id,
    pedido: data.pedido,
    tipo_movimiento: 'creacion',
    fecha_anterior: null,
    fecha_nueva: data.fecha_deseada,
    estado_anterior: null,
    estado_nuevo: 'pendiente',
    detalle: null,
    usuario_usuario: usuario,
    usuario_nombre: nombre,
  });

  return data;
}

// ---------------- Fachada ----------------

function crearServicioCupos(supabaseAdmin) {
  return {
    resumenSemana: (lunes) => resumenSemana(supabaseAdmin, lunes),
    capacidadRango: (desde, hasta) => capacidadRango(supabaseAdmin, desde, hasta),
    capacidadDia: (fecha, tipoEquipo) => capacidadDia(supabaseAdmin, fecha, tipoEquipo),
    obtenerDia: (fecha) => obtenerDia(supabaseAdmin, fecha),
    obtenerCarga: (cargaId) => obtenerCarga(supabaseAdmin, cargaId),
    listarCargas: (filtros) => listarCargas(supabaseAdmin, filtros),
    historialCarga: (cargaId) => historialCarga(supabaseAdmin, cargaId),
    historialReciente: (limite) => historialReciente(supabaseAdmin, limite),
    buscarProximoDisponible: (fechaInicio, tipoEquipo, opts) =>
      buscarProximoDisponible(supabaseAdmin, fechaInicio, tipoEquipo, opts),
    proponerAsignacion: (args) => proponerAsignacion(supabaseAdmin, args),
    confirmarAsignacion: (args) => confirmarAsignacion(supabaseAdmin, args),
    cancelar: (args) => cancelar(supabaseAdmin, args),
    despachar: (args) => despachar(supabaseAdmin, args),
    reabrir: (args) => reabrir(supabaseAdmin, args),
    cambiarModalidad: (args) => cambiarModalidad(supabaseAdmin, args),
    crearCarga: (args) => crearCarga(supabaseAdmin, args),
  };
}

module.exports = {
  crearServicioCupos,
  // utilidades de fecha, exportadas para reutilizar en server.js (p. ej. para
  // calcular el lunes de "hoy" al armar la URL por defecto de /api/cupos/semana)
  hoyISO,
  lunesDeSemana,
  addDays,
  diaSemana,
  esDomingo,
  nombreDia,
};
