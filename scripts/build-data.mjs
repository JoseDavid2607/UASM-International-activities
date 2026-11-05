// Genera data/events.json y data/evidence.json a partir de Google Drive.
// Se ejecuta en GitHub Actions (ver .github/workflows/update-data.yml); la API key vive en un secret
// y nunca llega al navegador.
//
// Variables de entorno:
//   DRIVE_API_KEY      (obligatoria) API key con Drive API habilitada y SIN restricción por referrer
//   DRIVE_FILE_ID      (opcional)    ID del .xlsx de actividades
//   EVIDENCE_ROOT_ID   (opcional)    ID de la carpeta "evidencias"
//   XLSX_FILE          (opcional)    ruta local a un .xlsx (omite la descarga; útil para pruebas)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const DRIVE_FILE_ID = process.env.DRIVE_FILE_ID || '1Hu7iJYf-FWxiGQaP2VSgF203C877Nj8M';
const EVIDENCE_ROOT_ID = process.env.EVIDENCE_ROOT_ID || '1NhEF_I9U13BmA9nszeI4Z9tRRnN8oTYz';
const DATA_SHEET = 'Actividades';
const FOLDER_MIME = 'application/vnd.google-apps.folder';
const CATS = { 'Convocatoria': 'convocatoria', 'Movilidad': 'intercambio', 'Sesión informativa': 'reunion', 'Evento': 'evento' };

export function normKey(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

/** Celda de fecha -> {y,m,d} sin depender de zona horaria. Acepta serial de Excel, Date o texto AAAA-MM-DD / DD/MM/AAAA. */
export function parseExcelDate(v) {
  let y, m, d;
  if (typeof v === 'number' && isFinite(v)) {
    const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
  } else if (v instanceof Date && !isNaN(v)) {
    const dt = new Date(v.getTime() + 12 * 3600000);
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
  } else if (typeof v === 'string') {
    const t = v.trim();
    let mt = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (mt) { y = +mt[1]; m = +mt[2]; d = +mt[3]; }
    else if ((mt = t.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/))) { d = +mt[1]; m = +mt[2]; y = +mt[3]; }
    else return null;
  } else return null;
  const chk = new Date(Date.UTC(y, m - 1, d));
  if (y < 2000 || y > 2100 || chk.getUTCFullYear() !== y || chk.getUTCMonth() !== m - 1 || chk.getUTCDate() !== d) return null;
  return { y, m, d };
}

/** Solo enlaces http(s). */
export function safeURL(v) {
  const t = String(v || '').trim();
  return /^https?:\/\//i.test(t) ? t : null;
}

const asDay = (x) => Date.UTC(x.y, x.m - 1, x.d);

/** Fila del Excel -> evento (o null con el motivo en `warn`). */
export function rowToEvent(row, warn = () => {}) {
  const estado = String(row['Estado'] || 'Confirmada').trim();
  if (estado === 'Cancelada') return null;
  const cat = CATS[String(row['Categoría'] || '').trim()];
  const program = String(row['Programa'] || '').trim();
  const title = String(row['Título'] || '').trim();
  const start = parseExcelDate(row['Fecha inicio']);
  if (!cat || !program || !title || !start) {
    warn(`Fila omitida (categoría, programa, título o fecha de inicio inválidos): ${JSON.stringify(row)}`);
    return null;
  }
  const ev = {
    year: start.y, m: start.m, d: start.d, cat, program, title,
    desc: String(row['Descripción corta'] || '').trim(),
    long: String(row['Descripción larga'] || row['Descripción corta'] || '').trim(),
    estado,
    link: safeURL(row['Enlace / formulario']),
    sessionLink: safeURL(row['Enlace de la sesión']),
    recording: safeURL(row['Enlace de la grabación']),
  };
  const end = parseExcelDate(row['Fecha fin (opcional)']);
  if (end) {
    if (asDay(end) > asDay(start)) { ev.y2 = end.y; ev.m2 = end.m; ev.d2 = end.d; }
    else if (asDay(end) < asDay(start)) warn(`Fecha fin anterior a la de inicio; se ignora la fecha fin: ${title}`);
  }
  return ev;
}

export function workbookToEvents(buf, warn) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const name = wb.SheetNames.includes(DATA_SHEET) ? DATA_SHEET : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null });
  return rows.map((r) => rowToEvent(r, warn)).filter(Boolean);
}

async function driveJSON(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Drive respondió HTTP ${res.status} (${(await res.text()).slice(0, 200)})`);
  return res.json();
}

/** Recorre la carpeta de evidencias; devuelve { títuloNormalizado: [archivos] }. */
export async function listEvidence(key, rootId, fetchJSON = driveJSON) {
  const byTitle = {};
  const queue = [{ id: rootId, name: '' }];
  while (queue.length) {
    const folder = queue.shift();
    let pageToken = '';
    do {
      const q = encodeURIComponent(`'${folder.id}' in parents and trashed=false`);
      const fields = encodeURIComponent('nextPageToken,files(id,name,mimeType,webViewLink)');
      const data = await fetchJSON(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=${fields}&pageSize=200&orderBy=name&key=${key}${pageToken ? '&pageToken=' + pageToken : ''}`);
      for (const f of data.files || []) {
        if (f.mimeType === FOLDER_MIME) queue.push({ id: f.id, name: f.name });
        else if (folder.name) {
          (byTitle[normKey(folder.name)] ||= []).push({
            id: f.id, name: f.name, mimeType: f.mimeType,
            url: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
            thumb: `https://drive.google.com/thumbnail?id=${f.id}&sz=w400`, // miniatura estable para archivos públicos
          });
        }
      }
      pageToken = data.nextPageToken || '';
    } while (pageToken);
  }
  return byTitle;
}

async function main() {
  const key = process.env.DRIVE_API_KEY;
  const outDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'data');
  fs.mkdirSync(outDir, { recursive: true });

  let buf;
  if (process.env.XLSX_FILE) buf = fs.readFileSync(process.env.XLSX_FILE);
  else {
    if (!key) throw new Error('Falta DRIVE_API_KEY');
    const res = await fetch(`https://www.googleapis.com/drive/v3/files/${DRIVE_FILE_ID}?alt=media&key=${key}`);
    if (!res.ok) throw new Error(`No se pudo descargar el Excel: HTTP ${res.status} (${(await res.text()).slice(0, 200)})`);
    buf = Buffer.from(await res.arrayBuffer());
  }
  const events = workbookToEvents(buf, (m) => console.warn('⚠', m));
  if (!events.length) throw new Error('El Excel no produjo ninguna actividad válida; no se sobrescriben los datos.');

  let evidence = {};
  if (key) {
    try { evidence = await listEvidence(key, EVIDENCE_ROOT_ID); }
    catch (e) { console.warn('⚠ No se pudieron listar las evidencias:', e.message); evidence = null; }
  }

  const write = (name, obj) => fs.writeFileSync(path.join(outDir, name), JSON.stringify(obj, null, 1) + '\n');
  write('events.json', { events });
  if (evidence) write('evidence.json', evidence); // si falló el listado se conserva el archivo anterior
  console.log(`✓ ${events.length} actividades; ${evidence ? Object.keys(evidence).length + ' carpetas de evidencias' : 'evidencias sin cambios'}`);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) main().catch((e) => { console.error('✗', e.message); process.exit(1); });
