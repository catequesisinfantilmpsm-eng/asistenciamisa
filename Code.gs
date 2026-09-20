const REVIEW_CONFIG = Object.freeze({
  SPREADSHEET_ID: '1gVFUsGjPbwR9DMIK41a5XTAcBulXGlG9FjjmtqzXVH0',
  REGISTRY_SPREADSHEET_ID: '1YMQPGC8pWdukQ1sQJ7XDHRPQA3fa2GeSC57ue-tImuU',
  REGISTRY_SHEETS: ['CATEQUIZANDOS1', 'CATEQUIZANDOS'],
  HIDDEN_SHEETS: ['CONEXIÓN REGISTRO'],
  INDEX_CACHE_SECONDS: 600,
  GROUP_CACHE_SECONDS: 0
});

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  const safeCallback = /^[A-Za-z_$][\w$]*$/.test(callback) ? callback : '';
  let result;
  try {
    const action = String(p.action || 'index');
    if (action === 'group') result = getGroup_(String(p.sheet || ''), false);
    else if (action === 'alerts') result = getAlerts_();
    else result = getIndex_();
  } catch (error) {
    result = { ok: false, error: 'No fue posible leer las asistencias: ' + String(error && error.message || error) };
  }
  const json = JSON.stringify(result);
  return ContentService.createTextOutput(safeCallback ? safeCallback + '(' + json + ');' : json)
    .setMimeType(safeCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getIndex_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('review-index-v5');
  if (cached) return JSON.parse(cached);
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.SPREADSHEET_ID);
  const groups = ss.getSheets().filter(isPublicGroup_).map(sheet => {
    const labels = sheet.getRange('A1:A2').getDisplayValues();
    const title = String(labels[0][0] || '');
    const subtitle = String(labels[1][0] || '');
    const stageMatch = title.match(/·\s*([^·]+)\s*·/);
    const groupMatch = subtitle.match(/GRUPO:\s*([^·]+)/i);
    const cateMatch = subtitle.match(/CATEQUISTA:\s*([^·]+)/i);
    return {
      sheet: sheet.getName(),
      stage: stageMatch ? stageMatch[1].trim() : 'CATEKIDS',
      group: groupMatch ? groupMatch[1].trim() : sheet.getName(),
      catechist: cateMatch ? cateMatch[1].trim() : '',
      childrenCount: Math.max(0, sheet.getLastRow() - 5)
    };
  });
  const result = { ok: true, updated: formatNow_(ss), groups: groups };
  try { cache.put('review-index-v5', JSON.stringify(result), REVIEW_CONFIG.INDEX_CACHE_SECONDS); } catch (_) {}
  return result;
}

function getGroup_(sheetName, includeFamily) {
  if (!sheetName) throw new Error('Falta seleccionar el grupo.');
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isPublicGroup_(sheet)) throw new Error('No se encontró el grupo solicitado.');
  const families = includeFamily ? safeFamilies_() : { byCode: {}, byName: {} };
  const group = readGroup_(sheet, ss.getSpreadsheetTimeZone() || 'America/Mexico_City', families);
  return { ok: true, updated: formatNow_(ss), group: group };
}

function getAlerts_() {
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
  const families = safeFamilies_();
  const alerts = [];
  ss.getSheets().filter(isPublicGroup_).forEach(sheet => {
    const group = readGroup_(sheet, tz, families);
    group.children.filter(c => c.absences >= 3).forEach(c => alerts.push({
      sheet: group.sheet, stage: group.stage, group: group.group,
      catechist: group.catechist, child: c
    }));
  });
  alerts.sort((a, b) => b.child.absences - a.child.absences || a.child.name.localeCompare(b.child.name, 'es'));
  return { ok: true, updated: formatNow_(ss), alerts: alerts };
}

function readGroup_(sheet, tz, families) {
  const labels = sheet.getRange('A1:A2').getDisplayValues();
  const title = String(labels[0][0] || '');
  const subtitle = String(labels[1][0] || '');
  const stageMatch = title.match(/·\s*([^·]+)\s*·/);
  const groupMatch = subtitle.match(/GRUPO:\s*([^·]+)/i);
  const cateMatch = subtitle.match(/CATEQUISTA:\s*([^·]+)/i);
  const rows = sheet.getLastRow() > 5 ? sheet.getRange(6, 1, sheet.getLastRow() - 5, Math.max(sheet.getLastColumn(), 6)).getValues() : [];
  const raw = rows.filter(r => r[2]).map(r => {
    const code = String(r[1] || ''), name = String(r[2] || '');
    const family = families.byCode[normalizeCode_(code)] || families.byName[normalizeText_(name)] || {};
    const obligations = {};
    r.slice(5).filter(v => v instanceof Date && !isNaN(v)).forEach(v => {
      const item = obligationFromDate_(v, tz);
      if (item) obligations[item.key] = item;
    });
    return { code: code, name: name, tutor: family.tutor || '', whatsapp: family.whatsapp || '', obligations: obligations };
  });
  const scheduleMap = {};
  raw.forEach(c => Object.keys(c.obligations).forEach(k => scheduleMap[k] = c.obligations[k]));
  const schedule = Object.keys(scheduleMap).sort().map(k => scheduleMap[k]);
  const children = raw.map(c => {
    const attendance = schedule.map(item => ({ key: item.key, label: item.label, attended: !!c.obligations[item.key] }));
    const attended = attendance.filter(x => x.attended).length;
    return { code: c.code, name: c.name, tutor: c.tutor, whatsapp: c.whatsapp, attendance: attendance, attended: attended, absences: schedule.length - attended };
  });
  const possible = children.length * schedule.length;
  const completed = children.reduce((sum, child) => sum + child.attended, 0);
  return {
    sheet: sheet.getName(),
    stage: stageMatch ? stageMatch[1].trim() : 'CATEKIDS',
    group: groupMatch ? groupMatch[1].trim() : sheet.getName(),
    catechist: cateMatch ? cateMatch[1].trim() : '',
    schedule: schedule,
    attendancePercent: possible ? Math.round(completed * 1000 / possible) / 10 : 0,
    children: children
  };
}

// El sábado y el domingo se convierten en la misma obligación dominical.
function obligationFromDate_(date, tz) {
  const day = Number(Utilities.formatDate(date, tz, 'u'));
  const time = Utilities.formatDate(date, tz, 'HH:mm');
  if (time === '09:36' || day === 4) return null; // Actividad CATEKIDS y Hora Santa no generan falta de misa.
  const d = new Date(date.getTime());
  if (day === 6) d.setDate(d.getDate() + 1);
  if (day !== 6 && day !== 7 && time !== '07:12') return null;
  const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const sunday = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
  const sat = new Date(d.getTime()); sat.setDate(sat.getDate() - 1);
  const saturday = Utilities.formatDate(sat, tz, 'dd/MM');
  return { key: key, label: time === '07:12' ? sunday + ' · Solemne' : saturday + '–' + sunday.substring(0, 5) };
}

function isPublicGroup_(sheet) { return !sheet.isSheetHidden() && REVIEW_CONFIG.HIDDEN_SHEETS.indexOf(sheet.getName()) < 0; }
function formatNow_(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'America/Mexico_City', 'dd/MM/yyyy HH:mm'); }
function safeFamilies_() { try { return readFamilies_(); } catch (_) { return { byCode: {}, byName: {} }; } }

function readFamilies_() {
  const result = { byCode: {}, byName: {} };
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.REGISTRY_SPREADSHEET_ID);
  let sheet = null;
  for (let i = 0; i < REVIEW_CONFIG.REGISTRY_SHEETS.length && !sheet; i++) sheet = ss.getSheetByName(REVIEW_CONFIG.REGISTRY_SHEETS[i]);
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getDataRange().getDisplayValues();
  let hr = 0;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(normalizeText_);
    if (row.some(v => v.indexOf('NOMBRE') >= 0) && row.some(v => v.indexOf('CODIGO') >= 0)) { hr = r; break; }
  }
  const headers = values[hr].map(normalizeText_);
  const col = names => {
    for (let n = 0; n < names.length; n++) {
      const target = normalizeText_(names[n]);
      let i = headers.indexOf(target); if (i >= 0) return i;
      i = headers.findIndex(h => h.indexOf(target) >= 0); if (i >= 0) return i;
    } return -1;
  };
  const ci=col(['CODIGO','CODIGO QR','FIRMA']), ni=col(['NOMBRE DEL ALUMNO','NOMBRES','NOMBRE']), ti=col(['MADRE DE FAMILIA','NOMBRE DEL TUTOR','TUTOR']), pi=col(['WHATSAPP DEL TUTOR','WHATSAPP','TELEFONO']);
  values.slice(hr + 1).forEach(row => {
    const name=ni>=0?String(row[ni]||'').trim():'', code=ci>=0?String(row[ci]||'').trim():'';
    if (!name && !code) return;
    const family={tutor:ti>=0?String(row[ti]||'').trim():'',whatsapp:pi>=0?cleanPhone_(row[pi]):''};
    if(code)result.byCode[normalizeCode_(code)]=family;if(name)result.byName[normalizeText_(name)]=family;
  });
  return result;
}
function normalizeText_(v){return String(v||'').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');}
function normalizeCode_(v){return normalizeText_(v).replace(/[^A-Z0-9]/g,'');}
function cleanPhone_(v){let p=String(v||'').replace(/\D/g,'');if(p.length===10)p='52'+p;return p;}
