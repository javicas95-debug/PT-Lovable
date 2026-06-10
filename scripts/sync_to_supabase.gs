// ============================================================
// PT-Lovable — Sincronización Google Sheets → Supabase
// ============================================================
// Pegar en: Google Sheets → Extensiones → Apps Script
// Ejecutar: syncAll()
// ============================================================

const SUPABASE_URL = "https://kanzirszawsgwpltdsyr.supabase.co";
const SUPABASE_KEY = "YOUR_ANON_KEY_HERE"; // reemplazar con la clave anon de Supabase

function excelDateToISO(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return value.toISOString();
  }
  const num = parseFloat(value);
  if (isNaN(num) || num < 1) return null;
  try {
    const date = new Date((num - 25569) * 86400 * 1000);
    if (isNaN(date.getTime())) return null;
    return date.toISOString();
  } catch (e) {
    return null;
  }
}

function excelDateToDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return value.toISOString().split("T")[0];
  }
  const num = parseFloat(value);
  if (isNaN(num) || num < 1) return null;
  try {
    const date = new Date((num - 25569) * 86400 * 1000);
    if (isNaN(date.getTime())) return null;
    return date.toISOString().split("T")[0];
  } catch (e) {
    return null;
  }
}

function cleanPhone(phone) {
  if (!phone) return null;
  return String(phone).replace(/\.0$/, "").replace("E11", "").replace("E+11", "");
}

function toBoolean(val) {
  return val === 1 || val === "1" || val === true;
}

function upsertToSupabase(table, rows) {
  if (rows.length === 0) return;
  const url = `${SUPABASE_URL}/rest/v1/${table}`;
  const BATCH_SIZE = 500;
  let synced = 0;

  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batch = rows.slice(i, i + BATCH_SIZE);
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SUPABASE_KEY,
        "Authorization": `Bearer ${SUPABASE_KEY}`,
        "Prefer": "resolution=merge-duplicates"
      },
      payload: JSON.stringify(batch),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200 && code !== 201) {
      Logger.log(`Error en tabla ${table} (lote ${i}-${i + BATCH_SIZE}): ${response.getContentText()}`);
    } else {
      synced += batch.length;
    }
  }
  Logger.log(`✅ ${table}: ${synced} filas sincronizadas`);
}

function syncFwApplications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("firstwork_data");
  if (!sheet) { Logger.log("Hoja 'firstwork_data' no encontrada"); return; }
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({
      application_id: r[0],
      created_on: excelDateToISO(r[1]),
      current_stage: r[2] || null,
      last_modified: excelDateToISO(r[3]),
      prospect_id: r[4] || null,
      workplace_id: r[5] || null,
      workplace_name: String(r[6] || ""),
      external_vacancy_id: r[7] || null,
      worker_id: r[8] || null,
      email: r[9] || null,
      full_name: r[10] || null,
      phone_number: cleanPhone(r[11]),
      application_url: r[12] || null,
      cv_uploaded: toBoolean(r[13]),
      cv_url: r[14] || null
    });
  }
  upsertToSupabase("fw_applications", rows);
}

function syncCacheShortlisted() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cache_shortlisted");
  if (!sheet) { Logger.log("Hoja 'cache_shortlisted' no encontrada"); return; }
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[3]) continue;
    rows.push({
      lead_uid: r[3],
      month: excelDateToDate(r[0]),
      week: excelDateToDate(r[1]),
      day: excelDateToDate(r[2]),
      sourcing_mode: r[4] || null,
      vacancy_request_uid: r[5] || null,
      client_name: r[6] || null,
      position_name: r[7] || null,
      municipality: r[8] || null,
      client_agency: r[9] || null,
      lead_status: r[10] || null,
      lang_code: r[11] || "pt_PT",
      lead_created_at: excelDateToISO(r[12]),
      updated_at: excelDateToISO(r[13]),
      external_source_type: r[14] || null,
      worker_profile_id: r[15] || null,
      ats_id: String(r[16] || ""),
      candidate_id: r[17] || null,
      prospect_uid: r[18] || null,
      first_name: r[19] || null,
      last_name: r[20] || null,
      phone: cleanPhone(r[21]),
      email: r[22] || null,
      applied_on: excelDateToISO(r[23]),
      called_on: excelDateToISO(r[24]),
      contacted_on: excelDateToISO(r[25]),
      is_hired_same_client: toBoolean(r[26]),
      is_hired_different_client: toBoolean(r[27]),
      has_firstwork: toBoolean(r[28])
    });
  }
  upsertToSupabase("cache_shortlisted", rows);
}

function syncStatusVacancies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("status_vacancies");
  if (!sheet) { Logger.log("Hoja 'status_vacancies' no encontrada"); return; }
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[2]) continue;
    rows.push({
      client_name: r[0] || null,
      workplace_name: r[1] || null,
      id: r[2],
      position_name: r[3] || null,
      job_starts_at: excelDateToDate(r[4]),
      workers_requested: parseInt(r[5]) || 0,
      status: r[6] || null,
      created_at: excelDateToISO(r[7]),
      shift_pattern: r[8] || null,
      workplace_id: r[9] || null,
      flow_version: parseFloat(r[10]) || null,
      applicants_count: parseInt(r[11]) || 0
    });
  }
  upsertToSupabase("status_vacancies", rows);
}

function syncPlacements() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("placements");
  if (!sheet) { Logger.log("Hoja 'placements' no encontrada"); return; }
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[3]) continue;
    rows.push({
      country_code: r[0] || "PT",
      client_id: r[1] || null,
      workplace_id: r[2] || null,
      placement_id: r[3],
      worker_id: r[4] || null,
      starts_at: excelDateToDate(r[5]),
      status: r[6] || null,
      email: r[7] || null,
      phone: cleanPhone(r[8]),
      worker_shift_planned: parseInt(r[9]) || 0,
      worker_shift_and_clocked: parseInt(r[10]) || 0
    });
  }
  upsertToSupabase("placements", rows);
}

function syncAll() {
  Logger.log("🚀 Iniciando sincronización completa...");
  syncStatusVacancies();
  syncFwApplications();
  syncCacheShortlisted();
  syncPlacements();
  Logger.log("✅ Sincronización completada");
}
