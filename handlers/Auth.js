/**
 * Auth.gs — Sesi & Autentikasi Login Per-Site
 */

const SESSION_TAB = 'Sessions';
const SESSION_DURATION_NORMAL_MS = 8 * 60 * 60 * 1000;      // 8 Jam
const SESSION_DURATION_REMEMBER_MS = 30 * 24 * 60 * 60 * 1000; // 30 Hari

/**
 * Pintu masuk utama Web App
 */
function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  const session = token ? getValidSession(token) : null;

  if (session) {
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.userName = session.Nama || session.Username;
    template.userRole = session.Role;
    template.userSite = session.Site;
    template.token = token;
    return template.evaluate()
      .setTitle('Monitoring Investigasi HSE')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const template = HtmlService.createTemplateFromFile('Login');
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('Monitoring Investigasi HSE — Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Dipanggil dari Login.html
 */
function attemptLogin(username, password, rememberMe) {
  const users = sheetToObjects(getMasterSS(), CONFIG.TAB_USERS);
  
  const user = users.find(u =>
    String(u.Username || '').toLowerCase().trim() === String(username || '').toLowerCase().trim()
  );

  if (!user) {
    return { success: false, message: 'Username tidak ditemukan.' };
  }
  
  if (String(user.Status || '').toLowerCase() !== 'aktif') {
    return { success: false, message: 'Akun Anda tidak aktif. Silakan hubungi admin.' };
  }

  const sheetPassword = String(user.Password || '').trim();
  if (sheetPassword !== String(password).trim()) {
    return { success: false, message: 'Password salah.' };
  }

  const token = createSession(user, !!rememberMe);
  return { success: true, token: token };
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
  return token;
}

function getValidSession(token) {
  try {
    const sessions = sheetToObjects(getMasterSS(), SESSION_TAB);
    const session = sessions.find(s => s.Token === token);
    if (!session) return null;
    if (new Date(session.ExpiresAt) < new Date()) return null;
    return session;
  } catch (e) {
    return null;
  }
}

function logoutSession(token) {
  const sheet = getMasterSS().getSheetByName(SESSION_TAB);
  if (!sheet) return { success: true };
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === token) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}

function requireSession(token) {
  const session = getValidSession(token);
  if (!session) {
    throw new Error('Sesi telah berakhir atau tidak valid. Silakan login kembali.');
  }
  return session;
}

/**
 * Mengambil data terfilter otomatis berdasarkan Site user yang sedang login
 */
function getFilteredIncidents(token) {
  const session = requireSession(token);
  const ssTransaksi = getTransaksiSS();
  const sheet = ssTransaksi.getSheetByName(CONFIG.TAB_INSIDEN);
  
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { success: true, session: session, headers: [], rows: [] };

  const headers = values[0];
  let siteColIdx = headers.indexOf('Site/Business Unit (BU)');
  if (siteColIdx === -1) siteColIdx = headers.indexOf('Site');

  const filteredRows = values.slice(1).filter(row => {
    if (!row.some(cell => cell !== '' && cell !== null)) return false;

    // Admin bisa lihat semua data
    if (String(session.Role).toLowerCase() === 'admin') return true;

    // User biasa hanya melihat data dari Site miliknya
    if (siteColIdx !== -1) {
      return String(row[siteColIdx]).trim().toLowerCase() === String(session.Site).trim().toLowerCase();
    }
    return true;
  });

  return {
    success: true,
    session: session,
    headers: headers,
    rows: filteredRows
  };
}