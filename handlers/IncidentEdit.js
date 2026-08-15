/** Ambil data detail insiden berdasarkan Nomor Insiden */
function getIncidentDetailForEdit(token, incidentId) {
  try {
    const session = requireSession(token);
    const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) throw new Error('Data insiden kosong.');

    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
    const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
    if (idIdx === -1) throw new Error('Kolom No Insiden tidak ditemukan.');

    const rows = values.slice(1).filter(r => String(r[idIdx] || '').trim() === String(incidentId).trim());
    if (!rows.length) throw new Error('Data insiden tidak ditemukan: ' + incidentId);

    if (String(session.Role || '').toLowerCase() !== 'admin' && siteIdx !== -1 &&
        String(rows[0][siteIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) {
      throw new Error('Anda tidak memiliki akses ke laporan pada site tersebut.');
    }

    return JSON.stringify({ success: true, headers: headers, rows: rows, session: session });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}

/** Simpan Update Data Insiden ke Spreadsheet */
function updateIncidentData(token, incidentId, updatedRowsData) {
  try {
    const session = requireSession(token);
    const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

    const values = sheet.getDataRange().getDisplayValues();
    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
    const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
    if (idIdx === -1) throw new Error('Kolom No Insiden tidak ditemukan.');

    const oldRows = values.slice(1).filter(r => String(r[idIdx] || '').trim() === String(incidentId).trim());
    if (!oldRows.length) throw new Error('Data insiden tidak ditemukan.');
    if (String(session.Role || '').toLowerCase() !== 'admin' && siteIdx !== -1 &&
        String(oldRows[0][siteIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) {
      throw new Error('Anda tidak memiliki akses untuk memperbarui laporan ini.');
    }
    if (!Array.isArray(updatedRowsData) || !updatedRowsData.length) throw new Error('Minimal satu data karyawan harus tersedia.');
    if (updatedRowsData.some(r => !Array.isArray(r) || r.length !== headers.length)) {
      throw new Error('Struktur data update tidak sesuai dengan kolom Data_Input.');
    }

    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][idIdx] || '').trim() === String(incidentId).trim()) sheet.deleteRow(i + 1);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, updatedRowsData.length, headers.length).setValues(updatedRowsData);

    const riwayatSheet = getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT);
    if (riwayatSheet) {
      riwayatSheet.appendRow([new Date(), incidentId, 'Updated', session.Nama + ' (' + session.Site + ')', 'Pembaruan data laporan insiden']);
    }
    return JSON.stringify({ success: true, message: 'Data insiden berhasil diperbarui!' });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}

/** Hapus seluruh transaksi satu nomor insiden. */
function deleteIncident(token, incidentId) {
  try {
    const session = requireSession(token);
    const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) throw new Error('Data insiden kosong.');
    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
    const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
    const rows = values.slice(1);
    const matches = rows.filter(r => String(r[idIdx] || '').trim() === String(incidentId).trim());
    if (!matches.length) throw new Error('Data insiden tidak ditemukan.');
    if (String(session.Role || '').toLowerCase() !== 'admin' && siteIdx !== -1 &&
        String(matches[0][siteIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) {
      throw new Error('Anda tidak memiliki akses untuk menghapus laporan ini.');
    }
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][idIdx] || '').trim() === String(incidentId).trim()) sheet.deleteRow(i + 1);
    }
    const riwayatSheet = getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT);
    if (riwayatSheet) riwayatSheet.appendRow([new Date(), incidentId, 'Deleted', 'Deleted', 'Laporan dihapus oleh ' + (session.Nama || session.Username)]);
    return { success: true, message: 'Laporan ' + incidentId + ' berhasil dihapus.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/** Data detail read-only untuk halaman Detail. */
function getIncidentDetail(token, incidentId) {
  const raw = getIncidentDetailForEdit(token, incidentId);
  try {
    const parsed = JSON.parse(raw);
    if (!parsed.success) return raw;
    const ss = getTransaksiSS();
    const attach = ss.getSheetByName(CONFIG.TAB_LAMPIRAN);
    const history = ss.getSheetByName(CONFIG.TAB_RIWAYAT);
    const attachments = attach ? sheetToObjects(ss, CONFIG.TAB_LAMPIRAN).filter(o => Object.values(o).some(v => String(v || '').trim() === String(incidentId).trim())) : [];
    const histories = history ? sheetToObjects(ss, CONFIG.TAB_RIWAYAT).filter(o => Object.values(o).some(v => String(v || '').trim() === String(incidentId).trim())) : [];
    parsed.attachments = attachments;
    parsed.history = histories;
    return JSON.stringify(parsed);
  } catch (e) {
    return JSON.stringify({ success: false, message: e.message });
  }
}

function incidentHeaderIndex_(headers, candidates) {
  const normalized = headers.map(h => String(h || '').trim().toLowerCase().replace(/[\s_\-\/]+/g, ' '));
  for (const candidate of candidates) {
    const c = String(candidate).trim().toLowerCase().replace(/[\s_\-\/]+/g, ' ');
    const exact = normalized.indexOf(c);
    if (exact !== -1) return exact;
  }
  for (let i = 0; i < normalized.length; i++) {
    if (candidates.some(candidate => {
      const c = String(candidate).trim().toLowerCase().replace(/[\s_\-\/]+/g, ' ');
      return c && (normalized[i].includes(c) || c.includes(normalized[i]));
    })) return i;
  }
  return -1;
}
