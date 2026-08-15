/** Backend Data Laporan Insiden */
function normalizeIncidentHeader_(v) {
  return String(v || '').trim().toLowerCase().replace(/[\s_\-\/]+/g, ' ');
}

function incidentHeaderIndex_(headers, candidates) {
  const h = headers.map(normalizeIncidentHeader_);
  const c = candidates.map(normalizeIncidentHeader_);
  for (const x of c) {
    const i = h.indexOf(x);
    if (i !== -1) return i;
  }
  for (let i = 0; i < h.length; i++) {
    if (c.some(x => x && (h[i].includes(x) || x.includes(h[i])))) return i;
  }
  return -1;
}

function getIncidentListData_(token, filters) {
  const session = requireSession(token);
  const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [], filterOptions: {sites: [], statuses: []} };

  const headers = values[0].map(String);
  const rows = values.slice(1).filter(r => r.some(v => v !== '' && v !== null && v !== undefined));

  const idIdx = incidentHeaderIndex_(headers, ['Nomor Insiden','No Insiden','No. Insiden','ID Insiden','Incident Number']);
  const dateIdx = incidentHeaderIndex_(headers, ['Tanggal Kejadian','Tanggal','Date of Incident','Incident Date']);
  const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)','Site','Business Unit','Business Unit (BU)']);
  const statusIdx = incidentHeaderIndex_(headers, ['Status Laporan','Status','Status Incident','Status Insiden']);
  const companyIdx = incidentHeaderIndex_(headers, ['Perusahaan','Nama Perusahaan','Perusahaan Karyawan','Perusahaan Terlibat','Nama Perusahaan Terlibat','Company','Company Name']);
  const employeeIdx = incidentHeaderIndex_(headers, ['Karyawan','Nama Karyawan','Nama Karyawan Terlibat','Employee','Employee Name']);
  const classIdx = incidentHeaderIndex_(headers, ['Klasifikasi','Klasifikasi Insiden','Klasifikasi Kecelakaan','Classification']);
  const chronologyIdx = incidentHeaderIndex_(headers, ['Kronologis','Uraian Kronologis Kejadian','Kronologi']);

  const companyMap = {};
  const master = getMasterSS().getSheetByName(CONFIG.TAB_PERUSAHAAN);
  if (master) {
    const mv = master.getDataRange().getValues();
    if (mv.length > 1) {
      const mh = mv[0].map(String);
      const mid = incidentHeaderIndex_(mh, ['ID_Perusahaan','ID Perusahaan','Kode Perusahaan','Company ID']);
      const mname = incidentHeaderIndex_(mh, ['Perusahaan','Nama Perusahaan','Company','Company Name']);
      if (mid !== -1 && mname !== -1) mv.slice(1).forEach(r => {
        const id = String(r[mid] == null ? '' : r[mid]).trim();
        const name = String(r[mname] == null ? '' : r[mname]).trim();
        if (id && name) companyMap[id.toLowerCase()] = name;
      });
    }
  }

  filters = filters || {};
  const search = String(filters.search || '').trim().toLowerCase();
  const selectedSite = String(filters.site || '').trim().toLowerCase();
  const selectedStatus = String(filters.status || '').trim().toLowerCase();
  let startDate = filters.startDate ? new Date(String(filters.startDate) + 'T00:00:00') : null;
  let endDate = filters.endDate ? new Date(String(filters.endDate) + 'T23:59:59') : null;
  if (startDate && isNaN(startDate.getTime())) startDate = null;
  if (endDate && isNaN(endDate.getTime())) endDate = null;

  const role = String(session.Role || '').trim().toLowerCase();
  const isAdmin = role === 'admin';
  const sessionSite = String(session.Site || '').trim().toLowerCase();

  const filtered = rows.filter(row => {
    const site = siteIdx === -1 ? '' : String(row[siteIdx] || '').trim();
    const status = statusIdx === -1 ? '' : String(row[statusIdx] || '').trim();
    if (!isAdmin && siteIdx !== -1 && site.toLowerCase() !== sessionSite) return false;
    if (selectedSite && site.toLowerCase() !== selectedSite) return false;
    if (selectedStatus && status.toLowerCase() !== selectedStatus) return false;
    if (startDate || endDate) {
      let d = dateIdx === -1 ? null : row[dateIdx];
      if (!(Object.prototype.toString.call(d) === '[object Date]')) d = new Date(d);
      if (!d || isNaN(d.getTime())) return false;
      if (startDate && d < startDate) return false;
      if (endDate && d > endDate) return false;
    }
    if (search) {
      const text = row.map(v => String(v == null ? '' : v).toLowerCase()).join(' ');
      if (!text.includes(search)) return false;
    }
    return true;
  });

  // Penting: kelompokkan berdasarkan nomor insiden. Semua karyawan dan perusahaan
  // pada baris dengan nomor sama akan dikirim sebagai satu laporan.
  const grouped = {};
  filtered.forEach(row => {
    const id = idIdx === -1 ? '' : String(row[idIdx] || '').trim();
    const key = id || Utilities.getUuid();
    if (!grouped[key]) grouped[key] = { row: row.slice(), rows: [] };
    grouped[key].rows.push(row);
  });

  const outRows = [];
  Object.keys(grouped).forEach(key => {
    const g = grouped[key];
    const base = g.row.slice();
    const employees = [];
    const companies = [];
    g.rows.forEach(row => {
      if (employeeIdx !== -1) {
        const e = String(row[employeeIdx] == null ? '' : row[employeeIdx]).trim();
        if (e && !employees.some(x => x.toLowerCase() === e.toLowerCase())) employees.push(e);
      }
      if (companyIdx !== -1) {
        let c = String(row[companyIdx] == null ? '' : row[companyIdx]).trim();
        if (c && companyMap[c.toLowerCase()]) c = companyMap[c.toLowerCase()];
        if (c && !companies.some(x => x.toLowerCase() === c.toLowerCase())) companies.push(c);
      }
    });
    // Jika perusahaan belum ada di kolom transaksi, jangan mengarang nilai.
    // Nilai '-' hanya muncul jika memang tidak ada nilai perusahaan pada transaksi.
    if (companyIdx !== -1) base[companyIdx] = companies.join(', ');
    if (employeeIdx !== -1) base[employeeIdx] = employees.join(', ');
    outRows.push(base);
  });

  const sites = [];
  const statuses = [];
  rows.forEach(r => {
    if (siteIdx !== -1) { const v = String(r[siteIdx] || '').trim(); if (v && !sites.includes(v)) sites.push(v); }
    if (statusIdx !== -1) { const v = String(r[statusIdx] || '').trim(); if (v && !statuses.includes(v)) statuses.push(v); }
  });

  return {
    headers: headers,
    rows: outRows,
    filterOptions: { sites: isAdmin ? sites.sort() : (session.Site ? [String(session.Site)] : []), statuses: statuses.sort() },
    indices: { id: idIdx, date: dateIdx, site: siteIdx, company: companyIdx, status: statusIdx, classification: classIdx, employee: employeeIdx, chronology: chronologyIdx },
    total: outRows.length
  };
}

function getFilteredIncidents(token, filters) {
  try {
    const result = getIncidentListData_(token, filters || {});
    return {
      success: true,
      session: requireSession(token),
      headers: result.headers,
      rows: result.rows.map(r => r.map(v => Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : (v == null ? '' : v))),
      indices: result.indices,
      filterOptions: result.filterOptions,
      total: result.total
    };
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e), headers: [], rows: [], total: 0 };
  }
}
