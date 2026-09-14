/*
 * Importador de la solapa "Oferta" (Excel -> Supabase), en un solo paso.
 *
 * Es la versión para panel-soja-web del mismo Excel que ya actualiza el
 * panel de Oferta viejo (proyecto "Proyecto consultas" / ventas-dashboard):
 * lee "ventas.xlsx" (Semillero, Varied, Variedad, Oferta, Venta, Disponible
 * — en bolsas), arma el SQL de reemplazo completo para la tabla
 * oferta_variedades, y lo manda directo a Supabase (misma función
 * exec_import_sql que ya usa importar_datos.js). Nada pasa por el chat.
 *
 * Uso:
 *   node importar_oferta.js "ventas.xlsx"
 *
 * Necesita, en la carpeta de arriba (panel-soja-web/config.env), las mismas
 * variables que usa el servidor: SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY.
 * Necesita también el paquete "xlsx" instalado (ya lo instala
 * actualizar_datos.bat/actualizar_oferta.bat la primera vez).
 */
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

function toNumber(value) {
  if (value === undefined || value === null) return 0;
  const limpio = String(value).replace(/[^\d.-]/g, '');
  const numero = Number(limpio);
  return Number.isNaN(numero) ? 0 : numero;
}

function esc(v) {
  if (v === null || v === undefined) return 'NULL';
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

function extractOferta(xlsxPath) {
  const wb = xlsx.readFile(xlsxPath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const allRows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  // Mismas 2 filas de encabezado que descarta convertExcel.js en el
  // proyecto original.
  const rows = allRows.slice(2);
  const result = rows.map((row) => ({
    semillero: row[0] ?? null,
    varied: row[1] ?? null,
    variedad: row[2] ?? null,
    oferta_primu: toNumber(row[3]),
    venta_primu: toNumber(row[4]),
    disponible_primu: toNumber(row[5]),
  }));
  return result.filter((r) => r.variedad && !String(r.variedad).toLowerCase().includes('total'));
}

function buildSql(xlsxPath) {
  const oferta = extractOferta(xlsxPath);
  const blocks = [
    // "WHERE true" es necesario: Supabase tiene activada la protección
    // "safeupdate", que bloquea cualquier DELETE/UPDATE sin WHERE.
    'DELETE FROM oferta_variedades WHERE true;',
  ];
  const stmt = insertStmt(
    'oferta_variedades',
    ['semillero', 'varied', 'variedad', 'oferta_primu', 'venta_primu', 'disponible_primu'],
    oferta
  );
  if (stmt) blocks.push(stmt);
  return { sql: blocks.join('\n\n'), counts: { oferta_variedades: oferta.length } };
}

function cargarConfigEnv() {
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
    body: JSON.stringify({ sql }),
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Supabase respondió ${res.status}: ${body}`);
  }
}

module.exports = { buildSql, extractOferta };

// ---------------- CLI ----------------

async function main() {
  const xlsxPath = process.argv[2];
  if (!xlsxPath) {
    console.log('Uso: node importar_oferta.js "ventas.xlsx"');
    process.exit(1);
  }
  if (!fs.existsSync(xlsxPath)) {
    console.log(`No encuentro el archivo: ${xlsxPath}`);
    process.exit(1);
  }

  const t0 = Date.now();
  console.log(`Leyendo ${path.basename(xlsxPath)} y armando el SQL...`);
  const { sql, counts } = buildSql(xlsxPath);
  const t1 = Date.now();
  console.log(`  listo en ${((t1 - t0) / 1000).toFixed(1)}s. Filas: ${counts.oferta_variedades}`);

  console.log('Enviando a Supabase...');
  try {
    await enviarASupabase(sql);
  } catch (err) {
    console.log(`ERROR: ${err.message}`);
    process.exit(1);
  }
  const t2 = Date.now();
  console.log(`Listo. Supabase actualizado en ${((t2 - t1) / 1000).toFixed(1)}s (total ${((t2 - t0) / 1000).toFixed(1)}s).`);
  console.log('La solapa Oferta del panel ya está al día.');
}

if (require.main === module) {
  main();
}
