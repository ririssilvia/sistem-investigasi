/**
 * HANDLER GENERATE NOMOR INVESTIGASI & TRANSAKSI LAPORAN INSIDEN
 * File: handlers/generatenomerinvestigasi.js
 */

function generateIncidentID(siteName, tglKejadian) {
  const ss = getTransaksiSS();
  const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
  
  let dateObj = tglKejadian ? new Date(tglKejadian) : new Date();
  let dd = String(dateObj.getDate()).padStart(2, '0');
  let mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  let yyyy = dateObj.getFullYear();
  let dateStr = `${dd}${mm}${yyyy}`;

  let cleanSite = String(siteName).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

  let lastRow = sheet.getLastRow();
  let dailyCounter = 1;

  if (lastRow > 1) {
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const prefixTarget = `${dateStr}-INV-${cleanSite}`;
    
    data.forEach(row => {
      let noInc = String(row[0]);
      if (noInc.indexOf(prefixTarget) === 0) {
        dailyCounter++;
      }
    });
  }

  let counterStr = String(dailyCounter).padStart(4, '0');
  return `${dateStr}-INV-${cleanSite}-${counterStr}`;
}

function saveIncidentReport(token, formData) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);

    const noInsiden = generateIncidentID(formData.site, formData.tglKejadian);
    const timestamp = new Date();

    const empList = (formData.karyawanList && formData.karyawanList.length > 0) 
      ? formData.karyawanList 
      : [{}];

    empList.forEach(emp => {
      const rowData = [
        timestamp,                          
        noInsiden,                          
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

    return { 
      success: true, 
      message: 'Data insiden berhasil disimpan!', 
      noInsiden: noInsiden 
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}