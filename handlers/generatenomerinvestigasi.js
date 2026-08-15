/**
 * HANDLER GENERATE NOMOR INVESTIGASI & TRANSAKSI LAPORAN INSIDEN
 * File: handlers/generatenomerinvestigasi.js
 */

function generateIncidentID(siteName, tglKejadian) {
  const ss = getTransaksiSS();
  const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
  if (!sheet) throw new Error('Tab Data_Input tidak ditemukan di DB Transaksi.');

  const dateObj = tglKejadian ? new Date(tglKejadian) : new Date();
  const dd = String(dateObj.getDate()).padStart(2, '0');
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const yyyy = dateObj.getFullYear();
  const dateStr = `${dd}${mm}${yyyy}`;
  const cleanSite = String(siteName || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const prefixTarget = `${dateStr}-INV-${cleanSite}-`;

  // Satu laporan dapat mempunyai beberapa row karena beberapa karyawan.
  // Counter harus menghitung NOMOR INSIDEN UNIK, bukan jumlah row transaksi.
  let nextCounter = 1;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const noColumn = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
    const used = new Set();
    noColumn.forEach(row => {
      const noInc = String(row[0] || '').trim();
      if (noInc.indexOf(prefixTarget) === 0) used.add(noInc);
    });
    used.forEach(noInc => {
      const match = noInc.match(/-(\d+)$/);
      if (match) nextCounter = Math.max(nextCounter, Number(match[1]) + 1);
    });
  }

  const counterStr = String(nextCounter).padStart(4, '0');
  return `${dateStr}-INV-${cleanSite}-${counterStr}`;
}

function saveIncidentReport(token, formData) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan di DB Transaksi.');

    const noInsiden = generateIncidentID(formData.site, formData.tglKejadian);
    const timestamp = new Date();
    const empList = (formData.karyawanList && formData.karyawanList.length > 0) ? formData.karyawanList : [{}];

    // Lock mencegah dua submit bersamaan mendapatkan nomor insiden yang sama.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const lockedNoInsiden = generateIncidentID(formData.site, formData.tglKejadian);
      empList.forEach(emp => {
        const rowData = [
          timestamp,
          lockedNoInsiden,
          session.Nama || session.Username,
          formData.hari || '',
          formData.tglKejadian || '',
          formData.bulan || '',
          formData.jamKejadian || '',
          formData.shift || '',
          formData.lokasi || '',
          formData.kronologis || '',
          formData.site || '',
          emp.perusahaan || '',
          emp.departemen || '',
          emp.klasifikasi || '',
          emp.karyawan || '',
          emp.jabatan || '',
          emp.umur || '',
          emp.masaKerja || '',
          formData.alatTerlibat || '',
          formData.jenisAlat || '',
          formData.lossCost || '',
          formData.jenisKontak || '',
          formData.sumberKecelakaan || '',
          formData.tta || '',
          formData.ketTta || '',
          formData.kta || '',
          formData.ketKta || '',
          formData.faktorManusia || '',
          formData.ketManusia || '',
          formData.faktorPekerjaan || '',
          formData.ketPekerjaan || '',
          formData.kurangKendali || '',
          formData.ketKendali || '',
          formData.tindakanPerbaikan || '',
          formData.tindakanPencegahan || '',
          formData.dueDate || '',
          '',
          'Open'
        ];
        sheet.appendRow(rowData);
      });
      return { success: true, message: 'Data insiden berhasil disimpan!', noInsiden: lockedNoInsiden };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { success: false, message: err.message };
  }
}