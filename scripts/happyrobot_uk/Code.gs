/**
 * ============================================================
 * HAPPY ROBOT UK - PIPELINE DE APPOINTMENTS (Firstwork -> llamadas)
 * ============================================================
 *
 * Adaptación UK del pipeline de US (Code.gs). Solo cubre la PRIMERA
 * parte del proceso: limpiar el DF de Firstwork y preparar la cola de
 * personas a llamar. NO incluye retry_queue, Supabase ni dedupe de
 * resultados (eso lo gestiona el equipo de llamadas).
 *
 * Origen (DF que se actualiza con la cron de Firstwork):
 *   Spreadsheet 1kljAAP_hRSgAGLQE9PeuIZeBqceM-I2ObrPDUY0Yo1I
 *   Pestaña gid=1479377624 ("UK - HappyRobot Calls - Firstwork DF")
 *
 * Destino (esta spreadsheet, el script va vinculado a ella):
 *   Spreadsheet 1f_zAxe3PjcqVBz_pZM7cEj3NAJTT4K-liFvfSu2bEmM
 *   Hojas: clean, tomorrow_appointments, data_error_reports
 *
 * DIFERENCIAS CON US
 * ------------------
 * - Headers del DF: "(Onsite Interview)" en vez de "(Onboarding
 *   Appointment)"; workerId/prospectId en vez de WorkerID:/ProspectID:.
 * - No hay "State": en UK el nombre del evento no lleva ", XX".
 *   Se elimina la columna y toda la lógica de STATE_OVERRIDES.
 * - Las horas vienen en UTC ("2026-11-05 14:00:00+00"). Se convierten a
 *   hora local UK (Europe/London, respeta GMT/BST) y se guarda también la
 *   hora UTC original en ISO para que no haya ambigüedad.
 * - Los arrays (Event Name / Location / Status / Start Time / Booking ID)
 *   son JSON válido, así que el parseo es directo.
 * - Solo pasan citas con status SCHEDULED (en US solo se quitaba
 *   CANCELLED; en UK además existen EXPIRED y COMPLETED).
 * - Se añaden Phone Number, Email, Current Stage y Application ID a la
 *   salida, para que el equipo de llamadas tenga todo en una fila.
 * - Fechas calculadas en Europe/London (US usaba Europe/Madrid).
 * - ID: workerId > prospect_<prospectId> > application_<Application ID>
 *   (este último fallback no existía en US).
 *
 * USO
 * ---
 * 1. Extensiones > Apps Script en la spreadsheet destino, pega este
 *    archivo como Code.gs.
 * 2. Ejecuta una vez runDailyUpdate() a mano para autorizar permisos.
 * 3. Ejecuta una vez installHourlyTrigger() para programarlo cada hora.
 * ============================================================
 */


/* ============================================================
 * CONFIG
 * ============================================================ */

const RAW_SPREADSHEET_ID = '1kljAAP_hRSgAGLQE9PeuIZeBqceM-I2ObrPDUY0Yo1I';
const RAW_SHEET_GID = 1479377624;          // se busca por gid: el nombre de la pestaña puede cambiar
const RAW_SHEET_NAME_FALLBACK = 'UK - HappyRobot Calls - Firstwork DF';

// Spreadsheet destino. Si el script está vinculado a ella se usa la activa.
const TARGET_SPREADSHEET_ID = '1f_zAxe3PjcqVBz_pZM7cEj3NAJTT4K-liFvfSu2bEmM';

const UK_TZ = 'Europe/London';

// tomorrow_appointments acepta citas entre HOY y HOY + N días (igual que US).
const CAPTURE_WINDOW_DAYS = 2;

// 'clean' solo guarda citas en esta ventana (el DF trae histórico desde 2025).
const RAW_WINDOW_PAST_DAYS = 3;
const RAW_WINDOW_FUTURE_DAYS = 30;

// Solo estas citas se llaman. Resto (CANCELLED, EXPIRED, COMPLETED) se descartan.
const VALID_BOOKING_STATUSES = ['SCHEDULED'];

// Candidatos en estas etapas no se llaman aunque tengan cita SCHEDULED.
// (Ajustar con el equipo de operaciones.)
const EXCLUDED_STAGES = ['Rejected', 'Inactive Applicants'];

// Eventos de prueba que nunca deben llamarse (match parcial, sin mayúsculas).
const EXCLUDED_EVENT_PATTERNS = [/testing workplace/i, /\btest\b/i];

// Prefijos de teléfono permitidos. Números de otro país van a data_error_reports.
// Pon [] para no filtrar por país.
const ALLOWED_PHONE_PREFIXES = ['+44'];

const CLEAN_HEADER = [
  'First Name',
  'ID',
  'Event Name',
  'Event Location',
  'Start Time',        // hora local UK, formato MM/dd/yyyy hh:mm a (igual que US)
  'Time Zone',         // Europe/London
  'Booking ID',
  'Phone Number',
  'Email',
  'Current Stage',
  'Application ID',
  'Start Time UTC'     // ISO 8601, sin ambigüedad
];

const ERROR_HEADER = [
  'First Name',
  'ID',
  'Phone Number',
  'Event Name',
  'Start Time',
  'Reason',
  'Source Row'
];

// Headers del DF de Firstwork UK
const SRC = {
  applicationId: ['Application ID'],
  stage:         ['Current Stage'],
  eventName:     ['All Children: Event Name (Onsite Interview)'],
  location:      ['All Children: Event Location (Onsite Interview)'],
  status:        ['All Children: Booking Status (Onsite Interview)'],
  startTime:     ['All Children: Event Start Time (Onsite Interview)'],
  bookingId:     ['All Children: Booking ID (Onsite Interview)'],
  workerId:      ['workerId', 'WorkerID:'],
  prospectId:    ['prospectId', 'ProspectID:'],
  firstName:     ['First Name'],
  email:         ['Email'],
  phone:         ['Phone Number']
};


/* ============================================================
 * ENTRY POINTS
 * ============================================================ */

function doGet(e)  { return jsonOut_(runDailyUpdate()); }
function doPost(e) { return jsonOut_(runDailyUpdate()); }

function runDailyUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn('Pipeline ya en ejecución. Se omite esta corrida.');
    return { status: 'skipped' };
  }
  try {
    const stats = buildAppointmentsDF();
    const added = updateTomorrowAppointments();
    return Object.assign({ status: 'ok', new_appointments_added: added }, stats);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { status: 'error', message: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/** Ejecutar UNA vez para crear el trigger horario (borra duplicados previos). */
function installHourlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runDailyUpdate')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runDailyUpdate').timeBased().everyHours(1).create();
  console.log('Trigger horario creado para runDailyUpdate().');
}


/* ============================================================
 * 1. RAW -> clean + data_error_reports
 * ============================================================ */

function buildAppointmentsDF() {
  const ss = getTargetSpreadsheet_();
  const rawSheet = getRawSheet_();

  const data = rawSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('El DF de Firstwork está vacío; no se sobrescribe clean.');
  }

  const headers = data[0].map(h => String(h).trim());
  const col = {};
  Object.keys(SRC).forEach(k => { col[k] = findCol_(headers, SRC[k]); });

  const todayKey = dateKeyInTz_(new Date(), UK_TZ);
  const minKey = addDaysToKey_(todayKey, -RAW_WINDOW_PAST_DAYS);
  const maxKey = addDaysToKey_(todayKey, RAW_WINDOW_FUTURE_DAYS);
  const captureMaxKey = addDaysToKey_(todayKey, CAPTURE_WINDOW_DAYS);
  // Los errores solo se reportan para citas que entrarían en la cola de llamadas.
  const inCaptureWindow = k => k && k >= todayKey && k <= captureMaxKey;

  const cleanRows = [CLEAN_HEADER];
  const errorRows = [ERROR_HEADER];
  const seenBookings = new Set();
  const stats = { clean: 0, errors: 0, not_scheduled: 0, out_of_window: 0, duplicates: 0 };

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const sourceRow = r + 1;

    const firstName = str_(row[col.firstName]);
    const worker = str_(row[col.workerId]);
    const prospect = str_(row[col.prospectId]);
    const phone = normalizePhone_(row[col.phone]);
    const email = str_(row[col.email]);
    const stage = str_(row[col.stage]);
    const applicationId = str_(row[col.applicationId]);

    const names = parseCellArray_(row[col.eventName]);
    const locs = parseCellArray_(row[col.location]);
    const statuses = parseCellArray_(row[col.status]);
    const times = parseCellArray_(row[col.startTime]);
    const bookings = parseCellArray_(row[col.bookingId]);

    if (names.length === 0 && times.length === 0) continue;

    // Mismo criterio que US: workerId preferente, si no prospect_<id>.
    // En UK ~10% de candidatos con cita no tienen ninguno de los dos todavía:
    // se usa application_<Application ID> para no perder esas llamadas.
    const id = worker ? worker
      : prospect ? 'prospect_' + prospect
      : applicationId ? 'application_' + applicationId
      : '';

    const pushError = (eventName, startLocal, dateKey, reason) => {
      if (!inCaptureWindow(dateKey)) return;
      errorRows.push([firstName, id, phone, eventName, startLocal, reason, sourceRow]);
    };

    const n = names.length;
    if (locs.length !== n || times.length !== n || statuses.length !== n || bookings.length !== n) {
      // No sabemos qué fecha tiene cada cita, así que se reporta si alguna cae en la ventana.
      const anyKey = times.map(t => (parseUtcTimestamp_(t) ? dateKeyInTz_(parseUtcTimestamp_(t), UK_TZ) : ''))
        .find(k => inCaptureWindow(k)) || '';
      pushError(names.join(' | '), times.join(' | '), anyKey,
        `Count mismatch: ${n} names / ${locs.length} locations / ${statuses.length} statuses / ` +
        `${times.length} times / ${bookings.length} booking IDs`);
      continue;
    }

    for (let j = 0; j < n; j++) {
      const eventName = names[j];
      const location = locs[j];
      const status = String(statuses[j] || '').toUpperCase();
      const bookingId = bookings[j];

      const startDate = parseUtcTimestamp_(times[j]);
      if (!startDate) {
        // Sin fecha no podemos saber si está en ventana: solo se reporta si estaba SCHEDULED.
        if (VALID_BOOKING_STATUSES.indexOf(status) !== -1) {
          errorRows.push([firstName, id, phone, eventName, times[j], 'Unparseable start time', sourceRow]);
        }
        continue;
      }

      const dateKey = dateKeyInTz_(startDate, UK_TZ);
      if (dateKey < minKey || dateKey > maxKey) { stats.out_of_window++; continue; }

      if (VALID_BOOKING_STATUSES.indexOf(status) === -1) { stats.not_scheduled++; continue; }

      const startLocal = Utilities.formatDate(startDate, UK_TZ, 'MM/dd/yyyy hh:mm a');
      const startUtcIso = Utilities.formatDate(startDate, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

      // ---- Validaciones (van a data_error_reports, no a clean) ----
      let reason = '';
      if (!id) reason = 'Missing workerId, prospectId and Application ID';
      else if (!bookingId) reason = 'Missing Booking ID';
      else if (!location) reason = 'Empty event location';
      else if (/PENDING/i.test(location)) reason = 'PENDING location';
      else if (!phone) reason = 'Missing phone number';
      else if (!isAllowedPhone_(phone)) reason = 'Non-UK phone number';
      else if (EXCLUDED_STAGES.indexOf(stage) !== -1) reason = `Excluded stage: ${stage}`;
      else if (EXCLUDED_EVENT_PATTERNS.some(p => p.test(eventName))) reason = 'Test event';

      if (reason) {
        pushError(eventName, startLocal, dateKey, reason);
        continue;
      }

      // El DF puede repetir una misma reserva si la persona tiene varias aplicaciones.
      if (seenBookings.has(bookingId)) { stats.duplicates++; continue; }
      seenBookings.add(bookingId);

      cleanRows.push([
        firstName,
        id,
        eventName,
        location,
        startLocal,
        UK_TZ,
        bookingId,
        phone,
        email,
        stage,
        applicationId,
        startUtcIso
      ]);
    }
  }

  // Orden cronológico (más fácil de revisar a mano)
  const body = cleanRows.slice(1).sort((a, b) => a[11] < b[11] ? -1 : a[11] > b[11] ? 1 : 0);

  writeSheet_(ss, 'clean', [CLEAN_HEADER].concat(body));
  writeSheet_(ss, 'data_error_reports', errorRows);

  stats.clean = body.length;
  stats.errors = errorRows.length - 1;
  console.log(
    `✅ clean: ${stats.clean} | ⚠️ errores (ventana de captura): ${stats.errors} | ` +
    `🚫 no SCHEDULED: ${stats.not_scheduled} | 📅 fuera de ventana: ${stats.out_of_window} | ` +
    `♻️ duplicados: ${stats.duplicates}`
  );
  return stats;
}


/* ============================================================
 * 2. clean -> tomorrow_appointments (append + dedupe)
 * ============================================================ */

function updateTomorrowAppointments() {
  const ss = getTargetSpreadsheet_();
  const clean = ss.getSheetByName('clean');
  if (!clean) throw new Error("No existe la hoja 'clean'");

  const data = clean.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim());
  const idx = name => {
    const i = header.indexOf(name);
    if (i === -1) throw new Error(`Falta la columna '${name}' en 'clean'`);
    return i;
  };
  const iId = idx('ID'), iEvent = idx('Event Name'), iTime = idx('Start Time'), iBooking = idx('Booking ID');
  const iUtc = idx('Start Time UTC');

  const expectedHeader = ['ID_ROW'].concat(header);

  let tom = ss.getSheetByName('tomorrow_appointments');
  if (!tom) tom = ss.insertSheet('tomorrow_appointments');

  const existingKeys = new Set();
  let maxIdRow = 0;

  if (tom.getLastRow() === 0) {
    tom.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]);
  } else {
    const tomHeader = tom.getRange(1, 1, 1, Math.max(tom.getLastColumn(), 1)).getValues()[0]
      .map(h => String(h).trim());
    // Nunca reescribimos la cabecera si ya tiene datos: Zapier mapea por columnas.
    if (expectedHeader.some((h, k) => tomHeader[k] !== h)) {
      throw new Error(
        "La cabecera de 'tomorrow_appointments' no coincide con 'clean'. " +
        'Esperada: ' + expectedHeader.join(', ')
      );
    }
    if (tom.getLastRow() > 1) {
      const rows = tom.getRange(2, 1, tom.getLastRow() - 1, expectedHeader.length).getValues();
      rows.forEach(r => {
        const n = parseInt(r[0], 10);
        if (!isNaN(n) && n > maxIdRow) maxIdRow = n;
        existingKeys.add(appointmentKey_(r[iBooking + 1], r[iId + 1], r[iEvent + 1], r[iTime + 1]));
      });
    }
  }

  const todayKey = dateKeyInTz_(new Date(), UK_TZ);
  const maxKey = addDaysToKey_(todayKey, CAPTURE_WINDOW_DAYS);
  const now = Date.now();
  console.log(`🔍 Ventana de captura: ${todayKey} -> ${maxKey} (${UK_TZ})`);

  let idRow = maxIdRow + 1;
  const out = [];
  let skippedPast = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const start = new Date(String(row[iUtc]));
    if (isNaN(start.getTime())) continue;

    const dateKey = dateKeyInTz_(start, UK_TZ);
    if (dateKey < todayKey || dateKey > maxKey) continue;

    // Cita de hoy que ya ha empezado: no tiene sentido llamar.
    if (start.getTime() <= now) { skippedPast++; continue; }

    const key = appointmentKey_(row[iBooking], row[iId], row[iEvent], row[iTime]);
    if (existingKeys.has(key)) continue;
    existingKeys.add(key);

    out.push([idRow++].concat(row.map(v => cellToString_(v))));
  }

  if (out.length > 0) {
    const range = tom.getRange(tom.getLastRow() + 1, 1, out.length, out[0].length);
    range.setNumberFormat('@');   // texto: conserva +44 y fechas tal cual
    range.setValues(out);
  }

  console.log(`✅ Nuevos appointments añadidos: ${out.length} | Ya pasados (hoy): ${skippedPast}`);
  return out.length;
}


/* ============================================================
 * HELPERS
 * ============================================================ */

function getTargetSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
}

function getRawSheet_() {
  const ext = SpreadsheetApp.openById(RAW_SPREADSHEET_ID);
  const sheet = ext.getSheets().find(s => s.getSheetId() === RAW_SHEET_GID) ||
                ext.getSheetByName(RAW_SHEET_NAME_FALLBACK);
  if (!sheet) {
    throw new Error(`No se encuentra la pestaña gid=${RAW_SHEET_GID} en la spreadsheet del DF de Firstwork.`);
  }
  return sheet;
}

function findCol_(headers, names) {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i !== -1) return i;
  }
  throw new Error(`Columna no encontrada en el DF (probé: ${names.join(' | ')})`);
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

function cellToString_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, UK_TZ, 'MM/dd/yyyy hh:mm a');
  return str_(v);
}

/** '["a","b"]' -> ['a','b'];  'a' -> ['a'];  '' -> [] */
function parseCellArray_(raw) {
  const s = str_(raw);
  if (!s) return [];
  try {
    const parsed = JSON.parse(s);
    if (Array.isArray(parsed)) return parsed.map(v => str_(v));
    return [str_(parsed)];
  } catch (e) {
    return [s];
  }
}

/**
 * '2026-11-05 14:00:00+00' | '2026-11-05T14:00:00+01:00' | '...Z' -> Date (instante real).
 * Sin offset se asume UTC (es lo que exporta Firstwork UK).
 */
function parseUtcTimestamp_(raw) {
  const s = str_(raw);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(Z|[+-]\d{2}(?::?\d{2})?)?$/i);
  if (!m) return null;
  let offsetMin = 0;
  if (m[7] && m[7].toUpperCase() !== 'Z') {
    const sign = m[7][0] === '-' ? -1 : 1;
    const digits = m[7].slice(1).replace(':', '');
    offsetMin = sign * (parseInt(digits.slice(0, 2), 10) * 60 + parseInt(digits.slice(2) || '0', 10));
  }
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) - offsetMin * 60000;
  return new Date(ms);
}

/** Fecha 'yyyy-MM-dd' de un instante en una zona horaria (comparable como string). */
function dateKeyInTz_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function addDaysToKey_(key, days) {
  const [y, m, d] = key.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  return Utilities.formatDate(dt, 'UTC', 'yyyy-MM-dd');
}

/** Teléfono a E.164 simple: quita espacios/guiones; 07xxx -> +447xxx. */
function normalizePhone_(raw) {
  let s = str_(raw).replace(/[\s\-().]/g, '');
  if (!s) return '';
  if (s.indexOf('00') === 0) s = '+' + s.slice(2);
  if (/^0\d{9,10}$/.test(s)) s = '+44' + s.slice(1);
  if (/^44\d{9,10}$/.test(s)) s = '+' + s;
  return s;
}

function isAllowedPhone_(phone) {
  if (!ALLOWED_PHONE_PREFIXES.length) return true;
  return ALLOWED_PHONE_PREFIXES.some(p => phone.indexOf(p) === 0);
}

function appointmentKey_(bookingId, personId, eventName, startTime) {
  const b = str_(bookingId);
  return b ? `BOOKING|${b}` : `LEGACY|${str_(personId)}|${str_(eventName)}|${cellToString_(startTime)}`;
}

function writeSheet_(ss, name, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  if (rows.length > 0) {
    const range = sheet.getRange(1, 1, rows.length, rows[0].length);
    range.setNumberFormat('@');
    range.setValues(rows);
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
