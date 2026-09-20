const REVIEW_CONFIG = Object.freeze({
  SPREADSHEET_ID: '1gVFUsGjPbwR9DMIK41a5XTAcBulXGlG9FjjmtqzXVH0',
  REGISTRY_SPREADSHEET_ID: '1YMQPGC8pWdukQ1sQJ7XDHRPQA3fa2GeSC57ue-tImuU',
  CATECHIST_SPREADSHEET_ID: '1hUpHGQAYPA0IBRPqxDp-GvcL_Xq_oQfF25unUGDC26A',
  PERSEVERANTE_SPREADSHEET_ID: '1hPTqr2zn4FpC8szHPBORdjx35p1572Yb5y8cPgLo7as',
  REGISTRY_SHEETS: ['CATEQUIZANDOS', 'CATEQUIZANDOS1'],
  ATTENDANCE_SHEETS: ['QR ASISTENCIA', 'ASISTENCIA A MISA', 'ASISTENCIA OTRA MISA'],
  HIDDEN_SHEETS: ['CONEXIÓN REGISTRO', 'RANKING DE ASISTENCIA', 'ÍNDICE POR GRUPO'],
  INDEX_CACHE_SECONDS: 600,
  GROUP_CACHE_SECONDS: 20
});

// Índice ligero: evita abrir y recorrer más de 35 hojas al cargar la portada.
const REVIEW_GROUPS = Object.freeze([
  ['SILVIA','INICIACIÓN','Ini1','SILVIA Y XIOMARA',22],['MARLEN','INICIACIÓN','Ini2','MARLEN Y ANA LUCIA',20],
  ['SAMANTHA','INICIACIÓN','Ini3','SAMANTHA Y LUZ',20],['AURORA','INICIACIÓN','Ini4','AURORA Y MARIA GUADALUPE',19],
  ['TADEO','INICIACIÓN','Ini5','TADEO Y HECTOR',22],['ANA','INICIACIÓN','Ini6','ANA Y DEVANY',20],
  ['EDIL','INICIACIÓN','Ini7','EDIL Y CONCHITA',17],['JOHANA','INICIACIÓN','Ini8','JOHANA Y GUSTAVO',20],
  ['MILDRED INI','INICIACIÓN','Ini9','MILDRED Y FAUSTINO',20],['LILIANA','INICIACIÓN','Ini10','LILIANA Y ALDO',21],
  ['VICKY','CONFIRMACIÓN','Conf1','VICTORIA',15],['HILIAMOR','CONFIRMACIÓN','Conf2','HILIAMOR Y FELIPE DE JESUS',21],
  ['RAQUEL','CONFIRMACIÓN','Conf3','RAQUEL Y IRENE',15],['CECILIA','CONFIRMACIÓN','Conf4','CECILIA',20],
  ['ESTHER','CONFIRMACIÓN','Conf5','ESTHER Y NANDIYELI',15],['FIDE DE JESUS','CONFIRMACIÓN','Conf6','FIDENCIO DE JESUS Y GRECIA',14],
  ['DON CHUY','CONFIRMACIÓN','Conf7','JESUS',19],['CONCHITA CONF','CONFIRMACIÓN','Conf8','CONCHITA Y ALDE',19],
  ['ARACELI','CONFIRMACIÓN','Conf9','ARACELY, XIMENA Y TANIA',13],['ESTRELLA','CONFIRMACIÓN','Conf10','ESTRELLA',4],
  ['MANDY','COMUNIÓN','Com1','MANDY',18],['TINA','COMUNIÓN','Com2','TINA Y JUVENCIA',19],
  ['ROX','COMUNIÓN','Com3','ROX, JOSELYN Y DEVANY',16],['MILDRED COM','COMUNIÓN','Com4','MILDRED Y FAUSTINO',14],
  ['PATY','COMUNIÓN','Com5','PATY Y MARISOL',18],['JORGE','COMUNIÓN','Com6','JORGE',13],
  ['MONSE','COMUNIÓN','Com7','MONSE Y YOSELIN',22],['CONCHITA COM','COMUNIÓN','Com8','CONCHITA',15],
  ['ALIZA','COMUNIÓN','Com9','ALIZA',20],['RICARDA','COMUNIÓN','Com10','RICARDA Y SEBASTIAN',13],
  ['CAYETANA','COMUNIÓN','Com11','CAYETANA Y KARLA',10],['LIDIA','COMUNIÓN','Com12','MA. LIDIA',7],
  ['HILIAMOR PERS','PERSEVERANTES','Pers1','CATEQUISTA HILIAMOR',6],['FIDENCIO','PERSEVERANTES','Pers2','CATEQUISTA FIDENCIO',7],
  ['EDIL PERS','PERSEVERANTES','Pers3','CATEQUISTA EDIL',1]
]);

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  const safeCallback = /^[A-Za-z_$][\w$]*$/.test(callback) ? callback : '';
  let result;
  try {
    const action = String(p.action || 'index');
    if (action === 'group') result = getGroup_(String(p.sheet || ''), false);
    else if (action === 'alerts') result = getAlerts_(String(p.sheets || p.sheet || ''));
    else if (action === 'catechists') result = getCatechists_();
    else result = getIndex_();
  } catch (error) {
    result = { ok: false, error: 'No fue posible leer las asistencias: ' + String(error && error.message || error) };
  }
  const json = JSON.stringify(result);
  return ContentService.createTextOutput(safeCallback ? safeCallback + '(' + json + ');' : json)
    .setMimeType(safeCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getIndex_() {
  const groups = REVIEW_GROUPS.map(row => ({
    sheet: row[0], stage: row[1], group: row[2], catechist: row[3], childrenCount: row[4]
  }));
  return {
    ok: true,
    updated: Utilities.formatDate(new Date(), 'America/Mexico_City', 'dd/MM/yyyy HH:mm'),
    groups: groups
  };
}

function getGroup_(sheetName, includeFamily) {
  if (!sheetName) throw new Error('Falta seleccionar el grupo.');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'review-group-v7-' + normalizeCode_(sheetName);
  if (!includeFamily && REVIEW_CONFIG.GROUP_CACHE_SECONDS > 0) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isPublicGroup_(sheet)) throw new Error('No se encontró el grupo solicitado.');
  const families = includeFamily ? safeFamilies_() : { byCode: {}, byName: {} };
  const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
  const attendanceIndex = safeAttendanceIndex_(tz);
  const registryChildren = safeRegistryChildren_();
  const group = readGroup_(sheet, tz, families, attendanceIndex, registryChildren);
  const result = { ok: true, updated: formatNow_(ss), group: group };
  if (!includeFamily && REVIEW_CONFIG.GROUP_CACHE_SECONDS > 0) {
    try { cache.put(cacheKey, JSON.stringify(result), REVIEW_CONFIG.GROUP_CACHE_SECONDS); } catch (_) {}
  }
  return result;
}

function getAlerts_(sheetNames) {
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
  const families = safeFamilies_();
  const attendanceIndex = safeAttendanceIndex_(tz);
  const registryChildren = safeRegistryChildren_();
  const alerts = [];
  const requested = String(sheetNames || '').split('|').map(v => v.trim()).filter(Boolean);
  const sheets = requested.length
    ? requested.map(name => ss.getSheetByName(name)).filter(sheet => sheet && isPublicGroup_(sheet))
    : [];
  sheets.forEach(sheet => {
    const group = readGroup_(sheet, tz, families, attendanceIndex, registryChildren);
    group.children.filter(c => c.absences >= 3).forEach(c => alerts.push({
      sheet: group.sheet, stage: group.stage, group: group.group,
      catechist: group.catechist, child: c
    }));
  });
  alerts.sort((a, b) => b.child.absences - a.child.absences || a.child.name.localeCompare(b.child.name, 'es'));
  return { ok: true, updated: formatNow_(ss), alerts: alerts };
}

function getCatechists_() {
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.REGISTRY_SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
  const attendanceIndex = safeAttendanceIndex_(tz);
  const roster = readCatechistRoster_();
  const massSchedule = Object.keys(attendanceIndex.schedule).sort().map(k => Object.assign({ type: 'misa' }, attendanceIndex.schedule[k]));
  const holySchedule = Object.keys(attendanceIndex.holyHourSchedule).map(k => Object.assign({ type: 'hora_santa' }, attendanceIndex.holyHourSchedule[k]));
  const schedule = massSchedule.concat(holySchedule).sort((a, b) => a.key.localeCompare(b.key));
  const catechists = roster.map(person => {
    const registered = attendanceIndex.byCode[normalizeCode_(person.code)]
      || attendanceIndex.byName[normalizeText_(person.name)] || {};
    const holyHour = attendanceIndex.holyHourByCode[normalizeCode_(person.code)]
      || attendanceIndex.holyHourByName[normalizeText_(person.name)] || {};
    const attendance = schedule.map(item => ({
      key: item.key,
      label: item.label,
      type: item.type,
      attended: item.type === 'hora_santa' ? !!holyHour[item.key] : !!registered[item.key]
    }));
    const attended = attendance.filter(item => item.type === 'misa' && item.attended).length;
    return {
      code: person.code,
      name: person.name,
      attendance: attendance,
      attended: attended,
      absences: massSchedule.length - attended,
      holyHours: attendance.filter(item => item.type === 'hora_santa' && item.attended).length
    };
  });
  return { ok: true, updated: formatNow_(ss), schedule: schedule, catechists: catechists };
}

function readCatechistRoster_() {
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.CATECHIST_SPREADSHEET_ID);
  const unique = {};
  ss.getSheets().filter(s => !s.isSheetHidden() && s.getLastRow() > 0).forEach(sheet => {
    const values = sheet.getDataRange().getDisplayValues();
    let start = 0;
    if (values.length && values[0].some(v => normalizeText_(v).indexOf('NOMBRE') >= 0)) start = 1;
    values.slice(start).forEach(row => {
      const name = String(row[0] || '').trim();
      const code = String(row[2] || '').trim();
      if (!name || (!code && normalizeText_(name) === 'NOMBRE')) return;
      const key = normalizeCode_(code) || normalizeText_(name);
      unique[key] = { name: name, code: code };
    });
  });
  return Object.keys(unique).map(k => unique[k]).sort((a, b) => a.name.localeCompare(b.name, 'es'));
}

function readGroup_(sheet, tz, families, attendanceIndex, registryChildren) {
  const labels = sheet.getRange('A1:A2').getDisplayValues();
  const title = String(labels[0][0] || '');
  const subtitle = String(labels[1][0] || '');
  const stageMatch = title.match(/·\s*([^·]+)\s*·/);
  const groupMatch = subtitle.match(/GRUPO:\s*([^·]+)/i);
  const cateMatch = subtitle.match(/CATEQUISTA:\s*([^·]+)/i);
  const rows = sheet.getLastRow() > 5
    ? sheet.getRange(6, 1, sheet.getLastRow() - 5, Math.max(sheet.getLastColumn(), 6)).getValues()
    : [];

  let raw = rows.filter(r => r[2]).map(r => {
    const code = String(r[1] || '').trim();
    const name = String(r[2] || '').trim();
    const codeKey = normalizeCode_(code);
    const nameKey = normalizeText_(name);
    const family = families.byCode[codeKey] || families.byName[nameKey] || {};
    const planned = {};
    r.slice(5).filter(v => v instanceof Date && !isNaN(v)).forEach(v => {
      const item = obligationFromDate_(v, tz);
      if (item) planned[item.key] = item;
    });
    const registered = attendanceIndex.byCode[codeKey] || attendanceIndex.byName[nameKey] || {};
    const holyHour = attendanceIndex.holyHourByCode[codeKey] || attendanceIndex.holyHourByName[nameKey] || {};
    return {
      code: code,
      name: name,
      tutor: family.tutor || '',
      whatsapp: family.whatsapp || '',
      planned: planned,
      registered: registered,
      holyHour: holyHour
    };
  });

  // Incorpora automáticamente altas y cambios realizados en CATEQUIZANDOS.
  // La hoja particular del grupo sigue aportando el calendario; el padrón vivo
  // aporta nombres y códigos para las listas, búsquedas y alertas.
  const groupName = groupMatch ? groupMatch[1].trim() : sheet.getName();
  let registryGroup = registryChildren.byGroup[normalizeCode_(sheet.getName())]
    || registryChildren.byGroup[normalizeCode_(groupName)]
    || [];
  if (!registryGroup.length && raw.length) {
    const possibleKeys = Object.keys(registryChildren.byGroup).sort((a, b) => b.length - a.length);
    const matchedKey = possibleKeys.find(key => raw.some(child => normalizeCode_(child.code).indexOf(key) === 0));
    if (matchedKey) registryGroup = registryChildren.byGroup[matchedKey];
  }
  const merged = {};
  raw.forEach(child => {
    const key = normalizeCode_(child.code) || normalizeText_(child.name);
    if (key) merged[key] = child;
  });
  registryGroup.forEach(item => {
    const codeKey = normalizeCode_(item.code);
    const nameKey = normalizeText_(item.name);
    const key = codeKey || nameKey;
    if (!key) return;
    const existingPair = Object.keys(merged).map(k => [k, merged[k]])
      .find(pair => normalizeText_(pair[1].name) === nameKey);
    const existing = merged[key] || (existingPair && existingPair[1]);
    if (existingPair && existingPair[0] !== key) delete merged[existingPair[0]];
    const registered = attendanceIndex.byCode[codeKey] || attendanceIndex.byName[nameKey] || {};
    const holyHour = attendanceIndex.holyHourByCode[codeKey] || attendanceIndex.holyHourByName[nameKey] || {};
    const child = existing || {
      code: item.code,
      name: item.name,
      tutor: '',
      whatsapp: '',
      planned: {},
      registered: registered,
      holyHour: holyHour
    };
    child.code = item.code || child.code;
    child.name = item.name || child.name;
    child.tutor = item.tutor || child.tutor || '';
    child.whatsapp = item.whatsapp || child.whatsapp || '';
    child.registered = registered;
    child.holyHour = holyHour;
    merged[key] = child;
  });
  raw = Object.keys(merged).map(k => merged[k]).sort((a, b) => a.name.localeCompare(b.name, 'es'));

  // Conserva el calendario que ya existe en las hojas de grupo.
  const scheduleMap = {};
  raw.forEach(c => Object.keys(c.planned).forEach(k => scheduleMap[k] = c.planned[k]));

  // También incluye fines de semana reales encontrados en REGISTRO GENERAL.
  // Así un registro manual, QR, AppSheet o formulario aparece aunque todavía
  // no se haya propagado como fecha dentro de la hoja particular del grupo.
  raw.forEach(c => Object.keys(c.registered).forEach(k => {
    scheduleMap[k] = scheduleMap[k] || attendanceIndex.schedule[k];
  }));

  const massSchedule = Object.keys(scheduleMap).sort().map(k => scheduleMap[k]).filter(Boolean);
  const displayScheduleMap = {};
  massSchedule.forEach(item => displayScheduleMap[item.key] = Object.assign({ type: 'misa' }, item));
  Object.keys(attendanceIndex.holyHourSchedule).forEach(k => {
    displayScheduleMap[k] = Object.assign({ type: 'hora_santa' }, attendanceIndex.holyHourSchedule[k]);
  });
  const schedule = Object.keys(displayScheduleMap).sort().map(k => displayScheduleMap[k]);
  const children = raw.map(c => {
    const attendance = schedule.map(item => ({
      key: item.key,
      label: item.label,
      type: item.type,
      attended: item.type === 'hora_santa'
        ? !!c.holyHour[item.key]
        : (!!c.planned[item.key] || !!c.registered[item.key])
    }));
    const attended = attendance.filter(x => x.type === 'misa' && x.attended).length;
    return {
      code: c.code,
      name: c.name,
      tutor: c.tutor,
      whatsapp: c.whatsapp,
      attendance: attendance,
      attended: attended,
      absences: massSchedule.length - attended,
      holyHours: attendance.filter(x => x.type === 'hora_santa' && x.attended).length
    };
  });

  const possible = children.length * massSchedule.length;
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

function safeAttendanceIndex_(tz) {
  try {
    const cache = CacheService.getScriptCache();
    const key = 'attendance-index-v3';
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
    const result = readAttendanceIndex_(tz);
    try { cache.put(key, JSON.stringify(result), 45); } catch (_) {}
    return result;
  }
  catch (_) { return { byCode: {}, byName: {}, schedule: {}, holyHourByCode: {}, holyHourByName: {}, holyHourSchedule: {} }; }
}

function readAttendanceIndex_(tz) {
  const result = { byCode: {}, byName: {}, schedule: {}, holyHourByCode: {}, holyHourByName: {}, holyHourSchedule: {} };
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.REGISTRY_SPREADSHEET_ID);
  let foundSheets = 0;

  REVIEW_CONFIG.ATTENDANCE_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    foundSheets++;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const values = range.getValues();
    const displayed = range.getDisplayValues();

    let hr = 0;
    for (let r = 0; r < Math.min(displayed.length, 10); r++) {
      const headers = displayed[r].map(normalizeText_);
      if (headers.some(h => h.indexOf('FECHA') >= 0) && headers.some(h => h.indexOf('CODIGO') >= 0)) {
        hr = r;
        break;
      }
    }

    const headers = displayed[hr].map(normalizeText_);
    const col = names => {
      for (let n = 0; n < names.length; n++) {
        const target = normalizeText_(names[n]);
        let i = headers.indexOf(target);
        if (i >= 0) return i;
        i = headers.findIndex(h => h.indexOf(target) >= 0);
        if (i >= 0) return i;
      }
      return -1;
    };

    const fi = col(['FECHA', 'FECHA DE MISA']);
    const ci = col(['CODIGO', 'CÓDIGO', 'CODIGO QR ENCONTRADO']);
    const ni = col(['NOMBRE REGISTRADO', 'NOMBRE', 'NOMBRE RECIBIDO POR WHATSAPP']);
    const ii = col(['INCIDENCIA']);
    const mi = col(['MOVIMIENTO']);
    const si = col(['ESTADO DE VINCULACION', 'ESTADO DE VINCULACIÓN']);
    if (fi < 0 || (ci < 0 && ni < 0)) return;

    for (let r = hr + 1; r < values.length; r++) {
      const row = values[r];
      const displayRow = displayed[r];
      const incidence = ii >= 0 ? normalizeText_(displayRow[ii]) : '';
      const movement = mi >= 0 ? normalizeText_(displayRow[mi]) : '';
      const linkStatus = si >= 0 ? normalizeText_(displayRow[si]) : '';
      const isHolyHour = incidence.indexOf('HORA SANTA') >= 0;

      // QR ASISTENCIA puede contener Hora Santa u otras incidencias.
      if (incidence && incidence.indexOf('MISA DOMINICAL') < 0 && incidence.indexOf('MISA EN OTRA PARROQUIA') < 0 && !isHolyHour) continue;
      if (movement && movement !== 'ENTRADA' && movement !== 'SALIDA') continue;

      // En ASISTENCIA OTRA MISA sólo se toman registros correctamente vinculados.
      if (sheetName === 'ASISTENCIA OTRA MISA' && linkStatus !== 'VINCULADA') continue;

      const date = coerceDate_(row[fi], displayRow[fi]);
      if (!date) continue;
      if (isHolyHour && Number(Utilities.formatDate(date, tz, 'u')) !== 4) continue;

      const code = ci >= 0 ? String(displayRow[ci] || '').trim() : '';
      const name = ni >= 0 ? String(displayRow[ni] || '').trim() : '';
      const codeKey = code && normalizeText_(code) !== 'NO ENCONTRADO' ? normalizeCode_(code) : '';
      const nameKey = normalizeText_(name);
      if (!codeKey && !nameKey) continue;

      if (isHolyHour) {
        const holyItem = holyHourFromDate_(date, tz);
        result.holyHourSchedule[holyItem.key] = holyItem;
        if (codeKey) {
          result.holyHourByCode[codeKey] = result.holyHourByCode[codeKey] || {};
          result.holyHourByCode[codeKey][holyItem.key] = true;
        }
        if (nameKey) {
          result.holyHourByName[nameKey] = result.holyHourByName[nameKey] || {};
          result.holyHourByName[nameKey][holyItem.key] = true;
        }
        continue;
      }

      const item = obligationFromDate_(date, tz);
      if (!item) continue;

      result.schedule[item.key] = item;
      if (codeKey) {
        result.byCode[codeKey] = result.byCode[codeKey] || {};
        result.byCode[codeKey][item.key] = true;
      }
      if (nameKey) {
        result.byName[nameKey] = result.byName[nameKey] || {};
        result.byName[nameKey][item.key] = true;
      }
    }
  });

  if (!foundSheets) throw new Error('No se encontraron las tres hojas configuradas para consultar asistencias.');
  return result;
}

function holyHourFromDate_(date, tz) {
  return {
    key: Utilities.formatDate(date, tz, 'yyyy-MM-dd') + '-HS',
    label: Utilities.formatDate(date, tz, 'dd/MM') + ' · Hora Santa'
  };
}

function coerceDate_(raw, displayed) {
  const text = String(displayed || raw || '').trim();
  // La fecha visible se interpreta al mediodía para impedir que UTC la mueva
  // al día anterior. Esto es esencial para reconocer correctamente los jueves.
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+.*)?$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  if (raw instanceof Date && !isNaN(raw)) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate(), 12, 0, 0);
  }
  return null;
}

// El sábado y el domingo se convierten en la misma obligación dominical.
function obligationFromDate_(date, tz) {
  const day = Number(Utilities.formatDate(date, tz, 'u'));
  const time = Utilities.formatDate(date, tz, 'HH:mm');
  if (time === '09:36' || day === 4) return null;
  const d = new Date(date.getTime());
  if (day === 6) d.setDate(d.getDate() + 1);
  if (day !== 6 && day !== 7 && time !== '07:12') return null;
  const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const sunday = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
  const sat = new Date(d.getTime());
  sat.setDate(sat.getDate() - 1);
  const saturday = Utilities.formatDate(sat, tz, 'dd/MM');
  return { key: key, label: time === '07:12' ? sunday + ' · Solemne' : saturday + '–' + sunday.substring(0, 5) };
}

function isPublicGroup_(sheet) {
  return !sheet.isSheetHidden() && REVIEW_CONFIG.HIDDEN_SHEETS.indexOf(sheet.getName()) < 0;
}

function formatNow_(ss) {
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'America/Mexico_City', 'dd/MM/yyyy HH:mm');
}

function safeRegistryChildren_() {
  try { return readRegistryChildren_(); }
  catch (_) { return { byGroup: {} }; }
}

function readRegistryChildren_() {
  const result = { byGroup: {} };
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.REGISTRY_SPREADSHEET_ID);
  let sheet = null;
  for (let i = 0; i < REVIEW_CONFIG.REGISTRY_SHEETS.length && !sheet; i++) {
    sheet = ss.getSheetByName(REVIEW_CONFIG.REGISTRY_SHEETS[i]);
  }
  if (!sheet || sheet.getLastRow() < 2) return result;

  const values = sheet.getDataRange().getDisplayValues();
  let hr = 0;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(normalizeText_);
    if (row.some(v => v.indexOf('NOMBRE') >= 0) && row.some(v => v.indexOf('CODIGO') >= 0)) {
      hr = r;
      break;
    }
  }

  const headers = values[hr].map(normalizeText_);
  const col = names => {
    for (let n = 0; n < names.length; n++) {
      const target = normalizeText_(names[n]);
      let i = headers.indexOf(target);
      if (i >= 0) return i;
      i = headers.findIndex(h => h.indexOf(target) >= 0);
      if (i >= 0) return i;
    }
    return -1;
  };

  const gi = col(['CODIGO DE GRUPO', 'GRUPO']);
  const ni = col(['NOMBRE DEL ALUMNO', 'NOMBRES', 'NOMBRE']);
  const ci = col(['CODIGO', 'CODIGO QR', 'FIRMA']);
  const si = col(['CURSO', 'ETAPA']);
  const ai = col(['CATEQUISTA']);
  const ti = col(['MADRE DE FAMILIA', 'NOMBRE DEL TUTOR', 'TUTOR']);
  const pi = col(['WHATSAPP DEL TUTOR', 'WHATSAPP', 'TELEFONO']);
  if (gi < 0 || ni < 0) throw new Error('CATEQUIZANDOS no contiene las columnas de grupo y nombre.');

  values.slice(hr + 1).forEach(row => {
    const groupCode = String(row[gi] || '').trim();
    const name = String(row[ni] || '').trim();
    const code = ci >= 0 ? String(row[ci] || '').trim() : '';
    if (!groupCode || (!name && !code)) return;
    const key = normalizeCode_(groupCode);
    result.byGroup[key] = result.byGroup[key] || [];
    result.byGroup[key].push({
      groupCode: groupCode,
      name: name,
      code: code,
      stage: si >= 0 ? String(row[si] || '').trim() : '',
      catechist: ai >= 0 ? String(row[ai] || '').trim() : '',
      tutor: ti >= 0 ? String(row[ti] || '').trim() : '',
      whatsapp: pi >= 0 ? cleanPhone_(row[pi]) : ''
    });
  });

  // Complementa el padrón con la hoja GENERAL de REGISTRO PERSEVERANTE.
  // E=grupo, G=nombre, H=código QR, I=catequista.
  try {
    const persSS = SpreadsheetApp.openById(REVIEW_CONFIG.PERSEVERANTE_SPREADSHEET_ID);
    const persSheet = persSS.getSheetByName('GENERAL');
    if (persSheet && persSheet.getLastRow() > 1) {
      const persRows = persSheet.getRange(2, 1, persSheet.getLastRow() - 1, Math.max(9, persSheet.getLastColumn())).getDisplayValues();
      persRows.forEach(row => {
        const groupCode = String(row[4] || '').trim();
        const name = String(row[6] || '').trim();
        const code = String(row[7] || '').trim();
        if (!groupCode || (!name && !code)) return;
        const key = normalizeCode_(groupCode);
        result.byGroup[key] = result.byGroup[key] || [];
        result.byGroup[key].push({
          groupCode: groupCode,
          name: name,
          code: code,
          stage: 'PERSEVERANTES',
          catechist: String(row[8] || '').replace(/^CATEQUISTA\s*/i, '').trim(),
          tutor: '',
          whatsapp: ''
        });
      });
    }
  } catch (_) {}

  Object.keys(result.byGroup).forEach(key => {
    const unique = {};
    result.byGroup[key].forEach(item => {
      const itemKey = normalizeCode_(item.code) || normalizeText_(item.name);
      if (itemKey) unique[itemKey] = item;
    });
    result.byGroup[key] = Object.keys(unique).map(k => unique[k]).sort((a, b) => a.name.localeCompare(b.name, 'es'));
  });
  return result;
}

function safeFamilies_() {
  try { return readFamilies_(); }
  catch (_) { return { byCode: {}, byName: {} }; }
}

function readFamilies_() {
  const result = { byCode: {}, byName: {} };
  const ss = SpreadsheetApp.openById(REVIEW_CONFIG.REGISTRY_SPREADSHEET_ID);
  let sheet = null;
  for (let i = 0; i < REVIEW_CONFIG.REGISTRY_SHEETS.length && !sheet; i++) {
    sheet = ss.getSheetByName(REVIEW_CONFIG.REGISTRY_SHEETS[i]);
  }
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getDataRange().getDisplayValues();
  let hr = 0;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    const row = values[r].map(normalizeText_);
    if (row.some(v => v.indexOf('NOMBRE') >= 0) && row.some(v => v.indexOf('CODIGO') >= 0)) {
      hr = r;
      break;
    }
  }
  const headers = values[hr].map(normalizeText_);
  const col = names => {
    for (let n = 0; n < names.length; n++) {
      const target = normalizeText_(names[n]);
      let i = headers.indexOf(target);
      if (i >= 0) return i;
      i = headers.findIndex(h => h.indexOf(target) >= 0);
      if (i >= 0) return i;
    }
    return -1;
  };
  const ci = col(['CODIGO', 'CODIGO QR', 'FIRMA']);
  const ni = col(['NOMBRE DEL ALUMNO', 'NOMBRES', 'NOMBRE']);
  const ti = col(['MADRE DE FAMILIA', 'NOMBRE DEL TUTOR', 'TUTOR']);
  const pi = col(['WHATSAPP DEL TUTOR', 'WHATSAPP', 'TELEFONO']);
  values.slice(hr + 1).forEach(row => {
    const name = ni >= 0 ? String(row[ni] || '').trim() : '';
    const code = ci >= 0 ? String(row[ci] || '').trim() : '';
    if (!name && !code) return;
    const family = {
      tutor: ti >= 0 ? String(row[ti] || '').trim() : '',
      whatsapp: pi >= 0 ? cleanPhone_(row[pi]) : ''
    };
    if (code) result.byCode[normalizeCode_(code)] = family;
    if (name) result.byName[normalizeText_(name)] = family;
  });
  return result;
}

function normalizeText_(v) {
  return String(v || '').trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizeCode_(v) {
  return normalizeText_(v).replace(/[^A-Z0-9]/g, '');
}

function cleanPhone_(v) {
  let p = String(v || '').replace(/\D/g, '');
  if (p.length === 10) p = '52' + p;
  return p;
}
