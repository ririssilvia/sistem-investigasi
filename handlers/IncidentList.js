function getIncidentListData_(token, filters) {
  const session = requireSession(token);
  const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { headers: [], rows: [], filterOptions: {sites: [], statuses: []} };

  const headers = values[0].map(v => String(v || '').trim());
  const rows = values.slice(1).filter(r => r.some(v => v !== '' && v !== null && v !== undefined));

  // DETEKSI HEADER PERSIS SESUAI SPREADSHEET KAMU
  const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden']);
  const dateIdx = incidentHeaderIndex_(headers, ['Tanggal Kejadian', 'Tanggal']);
  const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
  const companyIdx = incidentHeaderIndex_(headers, ['Perusahan', 'Perusahaan', 'Nama Perusahaan']);
  const classIdx = incidentHeaderIndex_(headers, ['Klasifikasi Kecelakaan', 'Klasifikasi']);
  const employeeIdx = incidentHeaderIndex_(headers, ['Karyawan Terlibat / Nama', 'Karyawan Terlibat', 'Karyawan']);
  const statusIdx = incidentHeaderIndex_(headers, ['Status Laporan', 'Status']);
  const chronologyIdx = incidentHeaderIndex_(headers, ['Kronologis']);

  filters = filters || {};
  const search = String(filters.search || '').trim().toLowerCase();
  const selectedSite = String(filters.site || '').trim().toLowerCase();
  const selectedStatus = String(filters.status || '').trim().toLowerCase();
  
  let startDate = filters.startDate ? new Date(String(filters.startDate) + 'T00:00:00') : null;
  let endDate = filters.endDate ? new Date(String(filters.endDate) + 'T23:59:59') : null;

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
      const rawDateStr = dateIdx === -1 ? '' : String(row[dateIdx] || '').trim();
      let d = new Date(rawDateStr);
      if (isNaN(d.getTime())) return false;
      if (startDate && d < startDate) return false;
      if (endDate && d > endDate) return false;
    }

    if (search) {
      const text = row.map(v => String(v == null ? '' : v).toLowerCase()).join(' ');
      if (!text.includes(search)) return false;
    }
    return true;
  });

  // Grouping 1 No Insiden -> Banyak Baris
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
      // Ambil Karyawan (Kolom O)
      if (employeeIdx !== -1) {
        const e = String(row[employeeIdx] == null ? '' : row[employeeIdx]).trim();
        if (e && e !== '-' && !employees.some(x => x.toLowerCase() === e.toLowerCase())) {
          employees.push(e);
        }
      }

      // Ambil Perusahaan/Perusahan (Kolom L)
      if (companyIdx !== -1) {
        const c = String(row[companyIdx] == null ? '' : row[companyIdx]).trim();
        if (c && c !== '-' && !companies.some(x => x.toLowerCase() === c.toLowerCase())) {
          companies.push(c);
        }
      }
    });

    if (companyIdx !== -1) base[companyIdx] = companies.length > 0 ? companies.join(', ') : '-';
    if (employeeIdx !== -1) base[employeeIdx] = employees.length > 0 ? employees.join(', ') : '-';
    
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
    indices: { 
      id: idIdx, 
      date: dateIdx, 
      site: siteIdx, 
      company: companyIdx, 
      status: statusIdx, 
      classification: classIdx, 
      employee: employeeIdx, 
      chronology: chronologyIdx 
    },
    total: outRows.length,
    session: session
  };
}