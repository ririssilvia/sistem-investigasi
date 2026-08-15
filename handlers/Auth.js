/**
 * Auth.gs — Sesi & Autentikasi Login Per-Site
 */

const SESSION_TAB = 'Sessions';
const SESSION_DURATION_NORMAL_MS = 8 * 60 * 60 * 1000;
const SESSION_DURATION_REMEMBER_MS = 30 * 24 * 60 * 60 * 1000;
const SESSION_BRIDGE_TTL_SECONDS = 8 * 60 * 60;

function attemptLogin(username, password, rememberMe) {
  const users = sheetToObjects(getMasterSS(), CONFIG.TAB_USERS);
  const user = users.find(u => String(u.Username || '').toLowerCase().trim() === String(username || '').toLowerCase().trim());
  if (!user) return { success: false, message: 'Username tidak ditemukan.' };
  if (String(user.Status || '').toLowerCase() !== 'aktif') return { success: false, message: 'Akun Anda tidak aktif. Silakan hubungi admin.' };
  if (String(user.Password || '').trim() !== String(password).trim()) return { success: false, message: 'Password salah.' };
  return { success: true, token: createSession(user, !!rememberMe) };
}

function createSession(user, rememberMe) {
  const token = Utilities.getUuid();
  const durationMs = rememberMe ? SESSION_DURATION_REMEMBER_MS : SESSION_DURATION_NORMAL_MS;
  const expiresAt = new Date(Date.now() + durationMs);
  let sheet = getMasterSS().getSheetByName(SESSION_TAB);
  if (!sheet) {
    sheet = getMasterSS().insertSheet(SESSION_TAB);
    sheet.appendRow(['Token', 'Username', 'Role', 'Site', 'Nama', 'ExpiresAt']);
  }
  sheet.appendRow([token, user.Username, user.Role, user.Site, user.Nama, expiresAt]);
  registerSessionBridge_(token);
  return token;
}

function normalizeSessionToken_(token) {
  return String(token == null ? '' : token).replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
}

function getSessionBridgeKey_() {
  return 'HSE_SESSION_' + Session.getTemporaryActiveUserKey();
}

function registerSessionBridge_(token) {
  const normalizedToken = normalizeSessionToken_(token);
  if (!normalizedToken) return;
  try {
    CacheService.getScriptCache().put(getSessionBridgeKey_(), normalizedToken, SESSION_BRIDGE_TTL_SECONDS);
  } catch (e) {
    // Cache is only a fallback; normal token authentication remains authoritative.
  }
}

function getBridgedSessionToken_() {
  try {
    return normalizeSessionToken_(CacheService.getScriptCache().get(getSessionBridgeKey_()));
  } catch (e) {
    return '';
  }
}

function resolveSessionToken_(token) {
  const supplied = normalizeSessionToken_(token);
  if (supplied && getValidSession(supplied)) {
    registerSessionBridge_(supplied);
    return supplied;
  }
  const bridged = getBridgedSessionToken_();
  if (bridged && getValidSession(bridged)) return bridged;
  return supplied || bridged || '';
}

function getValidSession(token) {
  try {
    const normalizedToken = normalizeSessionToken_(token);
    if (!normalizedToken) return null;
    const sessions = sheetToObjects(getMasterSS(), SESSION_TAB);
    const session = sessions.find(s => normalizeSessionToken_(s.Token) === normalizedToken);
    if (!session) return null;

    const expiresAt = session.ExpiresAt instanceof Date ? session.ExpiresAt : new Date(session.ExpiresAt);
    if (isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) return null;

    session.Token = normalizedToken;
    return session;
  } catch (e) {
    return null;
  }
}

function logoutSession(token) {
  const normalizedToken = normalizeSessionToken_(token);
  const sheet = getMasterSS().getSheetByName(SESSION_TAB);
  if (!sheet) return { success: true };
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalizeSessionToken_(values[i][0]) === normalizedToken) { sheet.deleteRow(i + 1); break; }
  }
  try { CacheService.getScriptCache().remove(getSessionBridgeKey_()); } catch (e) {}
  return { success: true };
}

function requireSession(token) {
  const resolvedToken = resolveSessionToken_(token);
  const session = getValidSession(resolvedToken);
  if (!session) throw new Error('Sesi telah berakhir atau tidak valid. Silakan login kembali.');
  registerSessionBridge_(session.Token);
  return session;
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/[\s_\-\/]+/g, ' ');
}

function findHeaderIndex_(headers, candidates) {
  const normalized = headers.map(normalizeHeader_);
  const wanted = candidates.map(normalizeHeader_);
  for (const candidate of wanted) {
    const exact = normalized.indexOf(candidate);
    if (exact !== -1) return exact;
  }
  for (let i = 0; i < normalized.length; i++) {
    if (wanted.some(candidate => candidate && (normalized[i].includes(candidate) || candidate.includes(normalized[i])))) return i;
  }
  return -1;
}

function serializeCellForClient_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value;
  return String(value);
}

/**
 * Membuat mapping ID_Perusahaan -> nama Perusahaan dari DB_Master_System.
 * Data transaksi lama boleh berisi ID/kode, tetapi yang ditampilkan di laporan
 * harus tetap nama perusahaan seperti master.
 */
function getCompanyMap_() {
  const map = {};
  const sheet = getMasterSS().getSheetByName(CONFIG.TAB_PERUSAHAAN);
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return map;

  const headers = values[0].map(v => String(v || ''));
  const idIdx = findHeaderIndex_(headers, ['ID_Perusahaan', 'ID Perusahaan', 'Kode Perusahaan', 'Company ID']);
  const nameIdx = findHeaderIndex_(headers, ['Perusahaan', 'Nama Perusahaan', 'Company', 'Company Name']);
  if (idIdx === -1 || nameIdx === -1) return map;

  values.slice(1).forEach(row => {
    const id = String(row[idIdx] == null ? '' : row[idIdx]).trim();
    const name = String(row[nameIdx] == null ? '' : row[nameIdx]).trim();
    if (id && name) map[id.toLowerCase()] = name;
  });

  return map;
}

function getIncidentSheetData_() {
  const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan di DB Transaksi.');
  const values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) return { sheet: sheet, headers: [], rows: [] };
  return { sheet: sheet, headers: values[0].map(v => String(v || '')), rows: values.slice(1).filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined)) };
}

function getFilteredIncidentRows_(token, filters) {
  const session = requireSession(token);
  const data = getIncidentSheetData_();
  const headers = data.headers;
  const rawRows = data.rows;
  filters = filters || {};

  const siteIdx = findHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site', 'Business Unit', 'Business Unit (BU)']);
  const dateIdx = findHeaderIndex_(headers, ['Tanggal Kejadian', 'Tanggal', 'Date of Incident', 'Incident Date']);
  const statusIdx = findHeaderIndex_(headers, ['Status Laporan', 'Status', 'Status Incident', 'Status Insiden']);
  const idIdx = findHeaderIndex_(headers, ['Nomor Insiden', 'No Insiden', 'No. Insiden', 'ID Insiden', 'Incident Number']);
  const companyIdx = findHeaderIndex_(headers, ['Perusahaan', 'Nama Perusahaan', 'ID_Perusahaan', 'ID Perusahaan', 'Company']);
  const chronologyIdx = findHeaderIndex_(headers, ['Kronologis', 'Uraian Kronologis Kejadian', 'Kronologi']);
  const role = String(session.Role || '').trim().toLowerCase();
  const isAdmin = role === 'admin';
  const sessionSite = String(session.Site || '').trim().toLowerCase();
  const companyMap = getCompanyMap_();

  const sites = [];
  const statuses = [];
  rawRows.forEach(row => {
    if (siteIdx !== -1) { const v = String(row[siteIdx] || '').trim(); if (v && !sites.includes(v)) sites.push(v); }
    if (statusIdx !== -1) { const v = String(row[statusIdx] || '').trim(); if (v && !statuses.includes(v)) statuses.push(v); }
  });
  sites.sort((a,b) => a.localeCompare(b));
  statuses.sort((a,b) => a.localeCompare(b));

  let startDate = filters.startDate ? new Date(String(filters.startDate) + 'T00:00:00') : null;
  let endDate = filters.endDate ? new Date(String(filters.endDate) + 'T23:59:59') : null;
  if (startDate && isNaN(startDate.getTime())) startDate = null;
  if (endDate && isNaN(endDate.getTime())) endDate = null;
  const search = String(filters.search || '').trim().toLowerCase();
  const selectedSite = String(filters.site || '').trim().toLowerCase();
  const selectedStatus = String(filters.status || '').trim().toLowerCase();

  const filteredRows = rawRows.filter(row => {
    if (!isAdmin && siteIdx !== -1 && String(row[siteIdx] || '').trim().toLowerCase() !== sessionSite) return false;
    if (selectedSite && siteIdx !== -1 && String(row[siteIdx] || '').trim().toLowerCase() !== selectedSite) return false;
    if (selectedStatus && statusIdx !== -1 && String(row[statusIdx] || '').trim().toLowerCase() !== selectedStatus) return false;

    if (startDate || endDate) {
      let rowDate = dateIdx !== -1 ? row[dateIdx] : null;
      if (!(Object.prototype.toString.call(rowDate) === '[object Date]')) rowDate = new Date(rowDate);
      if (isNaN(rowDate.getTime())) return false;
      if (startDate && rowDate < startDate) return false;
      if (endDate && rowDate > endDate) return false;
    }

    if (search) {
      const companyRaw = companyIdx !== -1 ? String(row[companyIdx] || '') : '';
      const companyName = companyMap[companyRaw.trim().toLowerCase()] || companyRaw;
      const haystack = [idIdx, chronologyIdx, siteIdx, statusIdx].filter(i => i !== -1).map(i => String(row[i] || '').toLowerCase()).join(' ') + ' ' + companyName.toLowerCase();
      if (!haystack.includes(search)) return false;
    }
    return true;
  });

  // Jangan mengubah struktur DB_Transaksi. Hanya nilai perusahaan pada response
  // yang dikirim ke halaman laporan/export yang di-resolve menjadi nama master.
  const displayRows = filteredRows.map(row => {
    const copy = row.slice();
    if (companyIdx !== -1) {
      const raw = String(copy[companyIdx] == null ? '' : copy[companyIdx]).trim();
      if (raw && companyMap[raw.toLowerCase()]) copy[companyIdx] = companyMap[raw.toLowerCase()];
    }
    return copy;
  });

  return {
    session: session,
    headers: headers,
    rows: displayRows,
    indices: { id: idIdx, date: dateIdx, site: siteIdx, company: companyIdx, status: statusIdx },
    filterOptions: { sites: isAdmin ? sites : (session.Site ? [String(session.Site)] : []), statuses: statuses },
    isAdmin: isAdmin
  };
}

function getFilteredIncidents(token, filters) {
  try {
    const result = getFilteredIncidentRows_(token, filters || {});
    return {
      success: true,
      session: {
        Username: String(result.session.Username || ''),
        Role: String(result.session.Role || ''),
        Site: String(result.session.Site || ''),
        Nama: String(result.session.Nama || '')
      },
      headers: result.headers,
      rows: result.rows.map(row => row.map(serializeCellForClient_)),
      indices: result.indices,
      filterOptions: result.filterOptions,
      isAdmin: result.isAdmin,
      total: result.rows.length
    };
  } catch (err) {
    return { success: false, message: err && err.message ? err.message : String(err), headers: [], rows: [], total: 0 };
  }
}

function exportIncidentExcel(token, filters) {
  const result = getFilteredIncidentRows_(token, filters || {});
  if (!result.rows.length) return { success: false, message: 'Tidak ada data yang sesuai dengan filter untuk diexport.' };

  const tempName = 'EXPORT_INCIDENT_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const tempSS = SpreadsheetApp.create(tempName);
  const tempSheet = tempSS.getSheets()[0];
  tempSheet.setName('Data_Input');
  tempSheet.getRange(1, 1, 1, result.headers.length).setValues([result.headers]);
  tempSheet.getRange(2, 1, result.rows.length, result.headers.length).setValues(result.rows);
  tempSheet.setFrozenRows(1);
  tempSheet.getRange(1, 1, 1, result.headers.length).setFontWeight('bold');
  tempSheet.autoResizeColumns(1, result.headers.length);
  SpreadsheetApp.flush();

  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
    throw new Error('Gagal membuat file Excel. HTTP ' + code);
  }

  const excelBlob = response.getBlob().setName(tempName + '.xlsx');
  const outputFile = DriveApp.createFile(excelBlob);
  outputFile.setName(tempName + '.xlsx');
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  return {
    success: true,
    fileName: outputFile.getName(),
    fileId: outputFile.getId(),
    downloadUrl: outputFile.getDownloadUrl(),
    total: result.rows.length
  };
}
