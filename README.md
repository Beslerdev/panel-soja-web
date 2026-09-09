# Panel de Campaña de Soja — DS Hnos

Servidor Node/Express que sirve el panel con datos en vivo desde Supabase y login real (Supabase Auth). Reemplaza al Excel como fuente de datos y al patrón de login inseguro del panel de Ventas.

## Estructura

```
panel-soja-web/
├── server.js              # servidor Express: login, sesión, API de datos, servido de páginas
├── lib/transform.js        # convierte las filas de Supabase en el formato que espera el panel (KPIs, agrupaciones)
├── public/
│   └── login.html          # pantalla de login (pública)
├── private/
│   └── panel_ds_soja.html  # el panel — solo se sirve si hay sesión activa
├── package.json
├── config.env.example             # copiar a config.env y completar
└── .gitignore
```

## Cómo funciona

1. `login.html` manda usuario/contraseña a `POST /api/login`.
2. El servidor arma un email interno (`usuario.toLowerCase() + '@panel-soja.local'`) y se lo pasa a Supabase Auth con la clave pública (`SUPABASE_ANON_KEY`). Esa clave solo puede hablar con Auth — no puede leer ninguna tabla, porque las 12 tablas tienen RLS activo sin políticas públicas.
3. Si es válido, se guarda una sesión de servidor (cookie `httpOnly`, no accesible desde JS del navegador).
4. `GET /` solo sirve `panel_ds_soja.html` si hay sesión — si no, redirige a `/login.html`.
5. `GET /api/data` (protegido) lee las 12 tablas de Supabase con la clave `service_role` (evita RLS, nunca llega al navegador) y arma el mismo JSON que antes estaba embebido a mano en el HTML.
6. El panel (`panel_ds_soja.html`) ya no trae los datos incrustados: al cargar, pide `/api/data` y renderiza con lo que recibe.

## Correr en local

```bash
npm install
cp config.env.example config.env
# completar SUPABASE_ANON_KEY, SUPABASE_SERVICE_ROLE_KEY y SESSION_SECRET en config.env
npm start
```

Abrir `http://localhost:3000` — redirige a login. Usá cualquiera de los 5 usuarios de la Fase 3 (ver `diseno_base_datos_supabase.md`).

Las claves se consiguen en el panel de Supabase, proyecto **panel-campana-soja**: Project Settings → API Keys.
- `SUPABASE_ANON_KEY`: la clave "publishable" (o "anon" legacy).
- `SUPABASE_SERVICE_ROLE_KEY`: la clave "service_role" — **secreta**, nunca la subas a git ni la compartas.

## Nota sobre "Hojas del Excel"

La pestaña que mostraba un espejo crudo de las hojas del Excel se sacó del panel en vivo (decisión tomada en la Fase 4): ya no tiene sentido una vez que la fuente de datos es Supabase y no el Excel. Toda la información útil que tenía ya está organizada en las otras 6 pestañas.

## Deploy — Fase 5

Deploy en Render (conectado a GitHub, igual que el panel de Ventas): este repo se conecta como Web Service en Render, con `npm install` como build command y `npm start` como start command. Las variables de entorno (`SUPABASE_URL`, `SUPABASE_ANON_KEY`, `SUPABASE_SERVICE_ROLE_KEY`, `SESSION_SECRET`) se cargan en la configuración de Render — nunca en este repo.

Nota: este repo es público (para poder subir los archivos sin compartir credenciales de git). No contiene ninguna clave real — `config.env` está en `.gitignore` y nunca se subió.
