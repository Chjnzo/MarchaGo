/**
 * MarchaGo — Backend (Google Apps Script) per cuestionario + admin ronde.
 *
 * ► SETUP INIZIALE (fai queste 3 cose una volta sola):
 *  A. Nel foglio Google: menu Extensiones → Apps Script → cancella tutto,
 *     incolla questo file, salva.
 *  B. In alto a sinistra clicca l'icona "⚙ Configuración del proyecto" →
 *     "Propiedades del script" → "Agregar propiedad":
 *       Nome:  ADMIN_PASSCODE
 *       Valore: (una password a tua scelta, es. "marchago-2026")
 *     Salva. Questa è la password che digiterete voi ragazzi nell'admin.
 *  C. Deploy: "Implementar" → "Nueva implementación" → tipo "Aplicación web".
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *     Copia la URL "/exec" e incollala in Cuestionario.html E in admin.html
 *     nella costante WEBHOOK_URL.
 *
 * ► DOPO OGNI MODIFICA DELLO SCRIPT: "Implementar" → "Gestionar implementaciones"
 *   → matita ✏️ → "Nueva versión" → Implementar. La URL resta uguale.
 *
 * ► TAB DEL FOGLIO (creati automaticamente):
 *   - "Respuestas"  → una riga per ogni invio del cuestionario.
 *   - "Rondas"      → registro delle ronde dinamiche pubblicate dall'admin.
 */

const SHEET_RESPONSES = 'Respuestas';
const SHEET_RONDAS = 'Rondas';
const ADMIN_PASSCODE_PROP = 'ADMIN_PASSCODE';

// Campi della Ronda 1 (hardcoded nel cuestionario) — ordine per il foglio.
const BASE_FIELDS = [
  'q-resp', 'q-founded', 'family', 'legal',
  'q-fulltime', 'q-parttime', 'q-billing', 'q-cnae',
  'q-correcciones',
  'q-origen', 'q-ejemplos', 'q-publico', 'q-necesidad',
  'q-proceso', 'q-evolucion', 'q-alcance',
  'q-mision', 'q-vision', 'q-objetivos', 'q-valores',
  'q-geo', 'q-eco', 'q-legal', 'q-social', 'q-cierre'
];

const BASE_LABELS = {
  'q-resp': 'Nombre y puesto',
  'q-founded': 'Fecha fundación',
  'family': '¿Empresa familiar?',
  'legal': 'Forma jurídica',
  'q-fulltime': 'Empleados tiempo completo',
  'q-parttime': 'Empleados tiempo parcial',
  'q-billing': 'Facturación 2024',
  'q-cnae': 'CNAE',
  'q-correcciones': 'Correcciones a lo que ya sabíamos',
  'q-origen': 'Origen de la idea',
  'q-ejemplos': 'Ejemplos de eventos',
  'q-publico': 'Público y clientes',
  'q-necesidad': 'Necesidad que cubren',
  'q-proceso': 'Proceso de trabajo',
  'q-evolucion': 'Evolución de la empresa',
  'q-alcance': 'Alcance geográfico',
  'q-mision': 'Qué significa el lema',
  'q-vision': 'Visión a 5 años',
  'q-objetivos': 'Objetivos actuales',
  'q-valores': 'Valores internos',
  'q-geo': 'Entorno geográfico',
  'q-eco': 'Entorno económico',
  'q-legal': 'Entorno político-legal',
  'q-social': 'Entorno sociocultural',
  'q-cierre': 'Factores que más afectan'
};

/* ============================== ROUTING ============================== */

function doGet(e) {
  return handleList();
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return jsonOut({ ok: false, error: 'No payload' });
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'respond';
    if (action === 'respond')       return handleRespond(data);
    if (action === 'list')          return handleList();
    if (action === 'publish')       return handlePublish(data);
    if (action === 'delete-ronda')  return handleDeleteRonda(data);
    if (action === 'check-pass')    return jsonOut({ ok: checkPasscode(data.passcode) });
    return jsonOut({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/* ============================== ACTIONS ============================== */

function handleList() {
  const sheet = getSheet(SHEET_RONDAS);
  const rondas = [];
  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    rows.forEach(function (r) {
      if (r[5] === 'published') {
        rondas.push({
          id: r[0],
          number: r[1],
          title: r[2],
          subtitle: r[3],
          publishedAt: r[4] ? (r[4] instanceof Date ? r[4].toISOString() : String(r[4])) : '',
          questions: safeJson(r[6], [])
        });
      }
    });
  }
  return jsonOut({ ok: true, rondas: rondas });
}

function handlePublish(data) {
  if (!checkPasscode(data.passcode)) return jsonOut({ ok: false, error: 'passcode' });
  const r = data.ronda || {};
  if (!r.title || !r.number || !Array.isArray(r.questions) || !r.questions.length) {
    return jsonOut({ ok: false, error: 'ronda incompleta' });
  }
  const sheet = ensureRondasSheet();
  const id = 'ronda-' + Date.now();
  sheet.appendRow([
    id,
    Number(r.number),
    String(r.title),
    String(r.subtitle || ''),
    new Date(),
    'published',
    JSON.stringify(r.questions)
  ]);
  return jsonOut({ ok: true, id: id });
}

function handleDeleteRonda(data) {
  if (!checkPasscode(data.passcode)) return jsonOut({ ok: false, error: 'passcode' });
  const sheet = getSheet(SHEET_RONDAS);
  if (!sheet) return jsonOut({ ok: false, error: 'sin tab Rondas' });
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return jsonOut({ ok: true });
    }
  }
  return jsonOut({ ok: false, error: 'no encontrada' });
}

function handleRespond(data) {
  const sheet = ensureResponsesSheet();
  const dyn = getDynamicFields();
  const allFields = BASE_FIELDS.concat(dyn.map(function (f) { return f.id; }));
  const labelOf = {};
  Object.keys(BASE_LABELS).forEach(function (k) { labelOf[k] = BASE_LABELS[k]; });
  dyn.forEach(function (f) { labelOf[f.id] = f.label; });

  const expected = ['Timestamp', 'Dispositivo'].concat(
    allFields.map(function (f) { return labelOf[f] || f; })
  );

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expected);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, expected.length).setFontWeight('bold');
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol < expected.length) {
      const missing = expected.slice(lastCol);
      sheet.getRange(1, lastCol + 1, 1, missing.length)
           .setValues([missing]).setFontWeight('bold');
    }
  }

  const row = [new Date(), data._ua || ''].concat(
    allFields.map(function (f) { return data[f] != null ? String(data[f]) : ''; })
  );
  sheet.appendRow(row);
  return jsonOut({ ok: true });
}

/* ============================== HELPERS ============================== */

function getDynamicFields() {
  const sheet = getSheet(SHEET_RONDAS);
  const out = [];
  if (!sheet || sheet.getLastRow() < 2) return out;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  // Ordina per numero di ronda per stabilità colonne.
  rows.sort(function (a, b) { return (a[1] || 0) - (b[1] || 0); });
  rows.forEach(function (r) {
    if (r[5] === 'published') {
      const qs = safeJson(r[6], []);
      qs.forEach(function (q) {
        if (q && q.id) out.push({ id: q.id, label: q.label || q.id });
      });
    }
  });
  return out;
}

function ensureRondasSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RONDAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RONDAS);
    sheet.appendRow(['id', 'number', 'title', 'subtitle', 'published_at', 'status', 'questions_json']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setColumnWidth(7, 400);
  }
  return sheet;
}

function ensureResponsesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_RESPONSES) || ss.insertSheet(SHEET_RESPONSES);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function checkPasscode(pc) {
  const stored = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSCODE_PROP);
  return !!stored && String(pc || '') === stored;
}

function safeJson(s, fallback) {
  try { return JSON.parse(s); } catch (_) { return fallback; }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
