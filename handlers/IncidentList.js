function getIncidentListData_(token, filters) {
  const session = requireSession(token);
  const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { headers: [], rows: [], filterOptions: {sites: [], statuses: []} };

  const headers = values[0].map(v => String(v || '').trim());
  const rows = values.slice(1).filter(r => r.some(v => v !== '' && v !== null && v !== undefined));
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
  const startDate = filters.startDate ? new Date(String(filters.startDate) + 'T00:00:00') : null;
  const endDate = filters.endDate ? new Date(String(filters.endDate) + 'T23:59:59') : null;
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
      const d = new Date(rawDateStr);
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
        if (e && e !== '-' && !employees.some(x => x.toLowerCase() === e.toLowerCase())) employees.push(e);
      }
      if (companyIdx !== -1) {
        const c = String(row[companyIdx] == null ? '' : row[companyIdx]).trim();
        if (c && c !== '-' && !companies.some(x => x.toLowerCase() === c.toLowerCase())) companies.push(c);
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
    headers,
    rows: outRows,
    filterOptions: { sites: isAdmin ? sites.sort() : (session.Site ? [String(session.Site)] : []), statuses: statuses.sort() },
    indices: { id: idIdx, date: dateIdx, site: siteIdx, company: companyIdx, status: statusIdx, classification: classIdx, employee: employeeIdx, chronology: chronologyIdx },
    total: outRows.length,
    session
  };
}

/** Hapus seluruh row yang memiliki No Insiden sama. User hanya boleh menghapus OPEN pada site-nya; Admin boleh semuanya. */
function deleteIncident(token, idInsiden) {
  const session = requireSession(token);
  const id = String(idInsiden || '').trim();
  if (!id) return { success: false, message: 'No. Insiden tidak valid.' };

  const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return { success: false, message: 'Data laporan tidak ditemukan.' };

  const headers = values[0].map(v => String(v || '').trim());
  const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
  const statusIdx = headers.length >= 38 ? 37 : incidentHeaderIndex_(headers, ['Status', 'Status Laporan']);
  const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
  if (idIdx < 0 || statusIdx < 0) throw new Error('Kolom No Insiden atau Status tidak ditemukan.');
  if (headers.length >= 38 && statusIdx !== 37) throw new Error('Struktur status tidak sesuai: AL harus Status.');

  const matches = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx] || '').trim() === id) matches.push({ row: i + 1, status: String(values[i][statusIdx] || '').trim(), site: siteIdx >= 0 ? String(values[i][siteIdx] || '').trim() : '' });
  }
  if (!matches.length) return { success: false, message: 'Laporan tidak ditemukan.' };

  const isAdmin = String(session.Role || '').trim().toLowerCase() === 'admin';
  if (!isAdmin && siteIdx >= 0 && matches.some(m => m.site.toLowerCase() !== String(session.Site || '').trim().toLowerCase())) {
    throw new Error('Anda tidak memiliki akses untuk menghapus laporan ini.');
  }
  if (!isAdmin && matches.some(m => m.status.toLowerCase() === 'close')) {
    throw new Error('Laporan yang sudah Close tidak dapat dihapus.');
  }

  // Hapus dari bawah agar nomor row tidak bergeser.
  matches.map(m => m.row).sort((a, b) => b - a).forEach(rowNumber => sheet.deleteRow(rowNumber));
  return { success: true, message: 'Laporan berhasil dihapus.', deletedRows: matches.length, idInsiden: id };
}
