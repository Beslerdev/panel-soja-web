const path = require('path');
require('dotenv').config({ path: path.join(__dirname, 'config.env') });

const express = require('express');
const session = require('express-session');
const { createClient } = require('@supabase/supabase-js');
const { transformRows } = require('./lib/transform');

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

function requireAuthApi(req, res, next) {
  if (req.session && req.session.user) return next();
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
  ];

  const results = {};
  for (const t of tables) {
    const { data, error } = await supabaseAdmin.from(t).select('*');
    if (error) throw new Error(`Error leyendo ${t}: ${error.message}`);
    results[t] = data || [];
  }

  return transformRows(results);
}

app.get('/api/data', requireAuthApi, async (req, res) => {
  try {
    const data = await buildPanelData();
    res.json({
      data,
      usuario: req.session.user,
      actualizado: new Date().toISOString(),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error al leer los datos de Supabase' });
  }
});

// ---------------- Páginas ----------------

app.use(express.static(path.join(__dirname, 'public'), { index: false }));

app.get('/', requireAuthPage, (req, res) => {
  res.sendFile(path.join(__dirname, 'private', 'panel_ds_soja.html'));
});

app.listen(PORT || 3000, () => {
  console.log(`Panel de Campaña de Soja escuchando en el puerto ${PORT || 3000}`);
});
