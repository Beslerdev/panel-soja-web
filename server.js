const path = require('path');
require('dotenv').config({ path: path.join(__dirname, 'config.env') });

const express = require('express');
const session = require('express-session');
const { createClient } = require('@supabase/supabase-js');
const { transformRows } = require('./lib/transform');
const { crearServicioCupos, hoyISO, lunesDeSemana } = require('./lib/cupos');

const {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  SUPABASE_SERVICE_ROLE_KEY,
  SESSION_SECRET,
  PORT,
} = process.env;

for (const [name, val] of Object.entries({ SUPABASE_URL, SUPABASE_ANON_KEY, SUPABASE_SERVICE_ROLE_KEY, SESSION_SECRET })) {
  if (!val) {
    console.error(`Falta la variable de entorno ${name}. Revisá el .env (o las env vars de Render).`);
    process.exit(1);
  }
}

// Cliente con la clave pública: solo se usa para validar usuario/contraseña contra Supabase Auth.
// Esta clave no puede leer ninguna tabla (RLS activo, sin políticas públicas).
const supabaseAuth = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Cliente con la clave service_role: puede leer todo, evita RLS. Solo se usa acá, del lado del servidor.
// NUNCA exponer SUPABASE_SERVICE_ROLE_KEY al navegador.
const supabaseAdmin = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false },
});

const cupos = crearServicioCupos(supabaseAdmin);

const EMAIL_DOMAIN = '@panel-soja.local';

const app = express();
app.set('trust proxy', 1); // Render (y otros PaaS) hacen proxy por HTTP interno; sin esto, la cookie "secure" nunca se guarda.
app.use(express.json());
app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      maxAge: 1000 * 60 * 60 * 12, // 12 horas
      sameSite: 'lax',
    },
  })
);

function requireAuthPage(req, res, next) {
  if (req.session && req.session.user) return next();
  return res.redirect('/login.html');
}

// Sincroniza Cupos desde Despachos automáticamente, sin depender de que
// alguien la corra a mano: se dispara sola en cada llamada autenticada a la
// API (cualquier uso del panel), con un margen de 60s entre corridas para no
// pegarle a Supabase en cada click. Si falla, no bloquea la respuesta real:
// solo queda logueado en la consola del servidor (Render).
let ultimoSyncCupos = 0;
async function sincronizarCuposSiCorresponde() {
  const ahora = Date.now();
  if (ahora - ultimoSyncCupos < 60_000) return;
  ultimoSyncCupos = ahora;
  try {
    const { error } = await supabaseAdmin.rpc('cupos_sincronizar_desde_despachos');
    if (error) console.error('[cupos-sync-auto]', error.message);
  } catch (err) {
    console.error('[cupos-sync-auto]', err && err.message ? err.message : err);
  }
}

function requireAuthApi(req, res, next) {
  if (req.session && req.session.user) {
    sincronizarCuposSiCorresponde().finally(next);
    return;
  }
  return res.status(401).json({ error: 'No autenticado' });
}

// ---------------- Auth ----------------

app.post('/api/login', async (req, res) => {
  const { usuario, password } = req.body || {};
  if (!usuario || !password) {
    return res.status(400).json({ error: 'Usuario y contraseña requeridos' });
  }
  const email = usuario.trim().toLowerCase() + EMAIL_DOMAIN;

  const { data, error } = await supabaseAuth.auth.signInWithPassword({ email, password });
  if (error || !data.user) {
    return res.status(401).json({ error: 'Usuario o contraseña incorrectos' });
  }

  req.session.user = {
    id: data.user.id,
    usuario: data.user.user_metadata?.usuario || usuario,
    nombre_completo: data.user.user_metadata?.nombre_completo || usuario,
  };
  res.json({ ok: true });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// ---------------- Datos del panel ----------------

async function buildPanelData() {
  const tables = [
    'pedidos_venta',
    'balanza',
    'lotes',
    'produccion_final',
    'despachos',
    'stock_resumen_marca',
    'stock_bolsas_semilla',
    'plan_curado',
    'plan_curado_semanal',
    'lotes_curado',
    'calendario_produccion',
    'oferta_variedades',
  ];

  const results = {};
  let ultimaCarga = null;
  for (const t of tables) {
    const { data, error } = await supabaseAdmin.from(t).select('*');
    if (error) throw new Error(`Error leyendo ${t}: ${error.message}`);
    results[t] = data || [];
    for (const fila of results[t]) {
      if (!fila.created_at) continue;
      const marca = new Date(fila.created_at).getTime();
      if (!Number.isNaN(marca) && (ultimaCarga === null || marca > ultimaCarga)) ultimaCarga = marca;
    }
  }

  return { data: transformRows(results), actualizado: ultimaCarga !== null ? new Date(ultimaCarga).toISOString() : null };
}

app.get('/api/data', requireAuthApi, async (req, res) => {
  try {
    const { data, actualizado } = await buildPanelData();
    res.json({
      data,
      usuario: req.session.user,
      // Momento real de la última carga de datos (el más nuevo entre los "created_at"
      // que deja Supabase al reemplazar las tablas), no el momento en que este usuario
      // abrió el panel.
      actualizado,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error al leer los datos de Supabase' });
  }
});

// ---------------- Cupos y programación de camiones/carros ----------------

const FECHA_RE = /^\d{4}-\d{2}-\d{2}$/;

// Envuelve un handler async: cualquier error se devuelve como
// { error: mensaje } con 400 (son validaciones de negocio, con mensajes en
// español ya listos para mostrar), salvo que sea un error realmente
// inesperado del servidor, que loguea y devuelve 500.
function rutaCupos(fn) {
  return (req, res) => {
    fn(req, res).catch((err) => {
      console.error('[cupos]', err && err.message ? err.message : err);
      res.status(400).json({ error: (err && err.message) || 'Error inesperado' });
    });
  };
}

function usuarioDeSesion(req) {
  const u = req.session.user || {};
  return { usuario: u.usuario, nombre: u.nombre_completo };
}

// Resumen semanal para el panel de 6 tarjetas (lunes a sábado).
app.get(
  '/api/cupos/semana',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const lunes = req.query.lunes && FECHA_RE.test(req.query.lunes) ? req.query.lunes : lunesDeSemana(hoyISO());
    const data = await cupos.resumenSemana(lunes);
    res.json(data);
  })
);

// Detalle de un día puntual (modalidad, capacidad de camión y carro, cargas).
app.get(
  '/api/cupos/dia/:fecha',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    if (!FECHA_RE.test(req.params.fecha)) return res.status(400).json({ error: 'Fecha inválida' });
    const data = await cupos.obtenerDia(req.params.fecha);
    res.json(data);
  })
);

// Capacidad de un rango de fechas (para una vista de calendario / disponibilidad rápida).
app.get(
  '/api/cupos/disponibilidad',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { desde, hasta } = req.query;
    if (!FECHA_RE.test(desde || '') || !FECHA_RE.test(hasta || '')) {
      return res.status(400).json({ error: 'Parámetros "desde" y "hasta" (YYYY-MM-DD) requeridos' });
    }
    const data = await cupos.capacidadRango(desde, hasta);
    res.json(data);
  })
);

// Listado de cargas (con filtros opcionales por estado, pedido y rango de fecha asignada).
app.get(
  '/api/cupos/cargas',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { estado, pedido, desde, hasta } = req.query;
    const data = await cupos.listarCargas({ estado, pedido, fechaDesde: desde, fechaHasta: hasta });
    res.json(data);
  })
);

// Alta de una nueva carga (queda en estado "pendiente de programación").
app.post(
  '/api/cupos/cargas',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const { pedido, cliente, producto, tipoEquipo, patente, fechaDeseada, observaciones } = req.body || {};
    const data = await cupos.crearCarga({ pedido, cliente, producto, tipoEquipo, patente, fechaDeseada, observaciones, usuario, nombre });
    res.json(data);
  })
);

// Historial reciente de movimientos (todas las cargas, el más nuevo primero).
app.get(
  '/api/cupos/historial',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const limite = Math.min(parseInt(req.query.limite, 10) || 100, 500);
    const data = await cupos.historialReciente(limite);
    res.json(data);
  })
);

// Cambia (o fija por primera vez) la modalidad de un día. Si hay cargas que
// no entrarían en la nueva capacidad, no cambia nada: devuelve { ok:false,
// bloqueos:[...] } con el detalle de qué cargas lo impiden.
app.post(
  '/api/cupos/modalidad',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const { fecha, modalidad } = req.body || {};
    if (!FECHA_RE.test(fecha || '')) return res.status(400).json({ error: 'Fecha inválida' });
    const data = await cupos.cambiarModalidad({ fecha, modalidad, usuario, nombre });
    res.json(data);
  })
);

// Detalle de una carga puntual.
app.get(
  '/api/cupos/cargas/:id',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const data = await cupos.obtenerCarga(req.params.id);
    res.json(data);
  })
);

// Historial de una carga puntual.
app.get(
  '/api/cupos/cargas/:id/historial',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const data = await cupos.historialCarga(req.params.id);
    res.json(data);
  })
);

// Paso 1 de la asignación: solo lectura. Busca si la fecha pedida tiene
// cupo; si no, propone la próxima fecha disponible (sin escribir nada
// todavía — hace falta "confirmar" para que se guarde de verdad).
app.post(
  '/api/cupos/cargas/:id/proponer',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { fechaDeseada } = req.body || {};
    const data = await cupos.proponerAsignacion({ cargaId: req.params.id, fechaDeseada });
    res.json(data);
  })
);

// Paso 2: confirma y escribe. El servidor vuelve a validar la disponibilidad
// en el momento de confirmar (dentro de cupos_asignar, con lock por fecha),
// así que si cambió algo entre proponer y confirmar, esto falla con un
// mensaje claro en vez de sobre-ocupar el día.
app.post(
  '/api/cupos/cargas/:id/confirmar',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const { fecha } = req.body || {};
    if (!FECHA_RE.test(fecha || '')) return res.status(400).json({ error: 'Fecha inválida' });
    const data = await cupos.confirmarAsignacion({ cargaId: req.params.id, fecha, usuario, nombre });
    res.json(data);
  })
);

// Cancela una carga programada o pendiente.
app.post(
  '/api/cupos/cargas/:id/cancelar',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const { motivo } = req.body || {};
    const data = await cupos.cancelar({ cargaId: req.params.id, usuario, nombre, motivo });
    res.json(data);
  })
);

// Marca una carga programada como despachada.
app.post(
  '/api/cupos/cargas/:id/despachar',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const data = await cupos.despachar({ cargaId: req.params.id, usuario, nombre });
    res.json(data);
  })
);

// Reapertura explícita de una carga cancelada o despachada (la única forma
// de volver a programarla, tal como pide el enunciado).
app.post(
  '/api/cupos/cargas/:id/reabrir',
  requireAuthApi,
  rutaCupos(async (req, res) => {
    const { usuario, nombre } = usuarioDeSesion(req);
    const { motivo } = req.body || {};
    const data = await cupos.reabrir({ cargaId: req.params.id, usuario, nombre, motivo });
    res.json(data);
  })
);

// ---------------- Páginas ----------------

app.use(express.static(path.join(__dirname, 'public'), { index: false }));

app.get('/', requireAuthPage, (req, res) => {
  res.sendFile(path.join(__dirname, 'private', 'panel_ds_soja.html'));
});

app.listen(PORT || 3000, () => {
  console.log(`Panel de Campaña de Soja escuchando en el puerto ${PORT || 3000}`);
});
