/** Handler edit/detail/delete laporan insiden. */
function incidentIsAdmin_(session) { return String(session && session.Role || '').trim().toLowerCase() === 'admin'; }
function incidentStatusIndex_(headers) { if (headers.length >= 38) return 37; return incidentHeaderIndex_(headers, ['Status','Status Laporan','Status Incident','Status Insiden']); }
function incidentCompletionDateIndex_(headers) { if (headers.length >= 37) return 36; return incidentHeaderIndex_(headers, ['Completion Date','Tanggal Completion','Completion']); }

function getIncidentDetailForEdit(token, incidentId, allowClosedRead) {
  try {
    const session=requireSession(token),sheet=getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN); if(!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values=sheet.getDataRange().getDisplayValues(); if(values.length<=1) throw new Error('Data insiden kosong.');
    const headers=values[0].map(v=>String(v||'').trim()),idIdx=incidentHeaderIndex_(headers,['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']),siteIdx=incidentHeaderIndex_(headers,['Site/Business Unit (BU)','Site']),statusIdx=incidentStatusIndex_(headers);
    if(idIdx<0) throw new Error('Kolom No Insiden tidak ditemukan.');
    const rows=values.slice(1).filter(r=>String(r[idIdx]||'').trim()===String(incidentId).trim()); if(!rows.length) throw new Error('Data insiden tidak ditemukan: '+incidentId);
    if(!incidentIsAdmin_(session)&&siteIdx>=0&&String(rows[0][siteIdx]||'').trim().toLowerCase()!==String(session.Site||'').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses ke laporan pada site tersebut.');
    const status=statusIdx>=0?String(rows[0][statusIdx]||'').trim():'';
    if(!allowClosedRead&&!incidentIsAdmin_(session)&&status.toLowerCase()==='close') throw new Error('Laporan yang sudah Close hanya dapat dilihat oleh user biasa.');
    return JSON.stringify({success:true,headers:headers,rows:rows,session:session,status:status});
  } catch(e){return JSON.stringify({success:false,message:e.message});}
}

function updateIncidentData(token,incidentId,updatedRowsData){
  try{
    const session=requireSession(token),sheet=getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN); if(!sheet)throw new Error('Tab Data_Input tidak ditemukan.');
    const values=sheet.getDataRange().getDisplayValues(),headers=values[0].map(v=>String(v||'').trim());
    const idIdx=incidentHeaderIndex_(headers,['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']),siteIdx=incidentHeaderIndex_(headers,['Site/Business Unit (BU)','Site']),statusIdx=incidentStatusIndex_(headers),completionIdx=incidentCompletionDateIndex_(headers);
    if(idIdx<0)throw new Error('Kolom No Insiden tidak ditemukan.');
    if(headers.length>=38&&(statusIdx!==37||completionIdx!==36))throw new Error('Struktur Data_Input tidak sesuai: AK harus Completion Date dan AL harus Status.');
    const oldRows=values.slice(1).filter(r=>String(r[idIdx]||'').trim()===String(incidentId).trim()); if(!oldRows.length)throw new Error('Data insiden tidak ditemukan.');
    const isAdmin=incidentIsAdmin_(session);
    if(!isAdmin&&siteIdx>=0&&String(oldRows[0][siteIdx]||'').trim().toLowerCase()!==String(session.Site||'').trim().toLowerCase())throw new Error('Anda tidak memiliki akses untuk memperbarui laporan ini.');
    const currentStatus=statusIdx>=0?String(oldRows[0][statusIdx]||'').trim():'';
    if(!isAdmin&&currentStatus.toLowerCase()==='close')throw new Error('Laporan yang sudah Close tidak dapat diedit kembali.');
    if(!Array.isArray(updatedRowsData)||!updatedRowsData.length||updatedRowsData.some(r=>!Array.isArray(r)||r.length!==headers.length))throw new Error('Struktur data update tidak sesuai dengan Data_Input.');

    // Field yang dikelola sistem tidak boleh diubah oleh browser.
    // Untuk user biasa, Site juga dikunci mengikuti session agar edit tidak dapat memindahkan laporan ke site lain.
    const originalSite=siteIdx>=0?String(oldRows[0][siteIdx]||'').trim():'';
    updatedRowsData.forEach(r=>{
      r[idIdx]=incidentId;
      if(statusIdx>=0)r[statusIdx]=oldRows[0][statusIdx];
      if(completionIdx>=0)r[completionIdx]=oldRows[0][completionIdx];
      if(!isAdmin&&siteIdx>=0)r[siteIdx]=originalSite;
    });

    // Replace hanya row dengan No. Insiden ini. Jumlah karyawan boleh berubah,
    // tetapi laporan lain tidak disentuh dan identifier/status tetap dikelola backend.
    const lastRow=sheet.getLastRow();
    for(let i=lastRow;i>=2;i--){
      const rowId=String(sheet.getRange(i,idIdx+1).getDisplayValue()||'').trim();
      if(rowId===String(incidentId).trim())sheet.deleteRow(i);
    }
    const insertAt=sheet.getLastRow()+1;
    sheet.getRange(insertAt,1,updatedRowsData.length,headers.length).setValues(updatedRowsData);

    const h=getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT);
    if(h)h.appendRow(['HIS-'+Utilities.getUuid().substring(0,8).toUpperCase(),incidentId,currentStatus,currentStatus,'Data laporan diperbarui',new Date(),session.Nama||session.Username]);
    return JSON.stringify({success:true,message:'Data insiden berhasil diperbarui!',updatedRows:updatedRowsData.length,previousRows:oldRows.length});
  }catch(e){return JSON.stringify({success:false,message:e.message});}
}

function deleteIncident(token,incidentId){
  try{
    const session=requireSession(token),sheet=getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);if(!sheet)throw new Error('Tab Data_Input tidak ditemukan.');
    const values=sheet.getDataRange().getDisplayValues(),headers=values[0].map(v=>String(v||'').trim()),idIdx=incidentHeaderIndex_(headers,['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']),siteIdx=incidentHeaderIndex_(headers,['Site/Business Unit (BU)','Site']),statusIdx=incidentStatusIndex_(headers);
    const matches=values.slice(1).filter(r=>String(r[idIdx]||'').trim()===String(incidentId).trim());if(!matches.length)throw new Error('Data insiden tidak ditemukan.');
    const isAdmin=incidentIsAdmin_(session);
    if(!isAdmin&&siteIdx>=0&&String(matches[0][siteIdx]||'').trim().toLowerCase()!==String(session.Site||'').trim().toLowerCase())throw new Error('Anda tidak memiliki akses untuk menghapus laporan ini.');
    const status=statusIdx>=0?String(matches[0][statusIdx]||'').trim():'';
    if(status.toLowerCase()==='close'&&!isAdmin)throw new Error('Laporan yang sudah Close tidak dapat dihapus.');
    for(let i=sheet.getLastRow();i>=2;i--){if(String(sheet.getRange(i,idIdx+1).getDisplayValue()||'').trim()===String(incidentId).trim())sheet.deleteRow(i);}
    const h=getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT);if(h)h.appendRow(['HIS-'+Utilities.getUuid().substring(0,8).toUpperCase(),incidentId,status,status,'Laporan dihapus',new Date(),session.Nama||session.Username]);
    return{success:true,message:'Laporan '+incidentId+' berhasil dihapus.'};
  }catch(e){return{success:false,message:e.message};}
}

function getIncidentDetail(token,incidentId){const raw=getIncidentDetailForEdit(token,incidentId,true);try{const parsed=JSON.parse(raw);if(!parsed.success)return raw;const ss=getTransaksiSS(),a=ss.getSheetByName(CONFIG.TAB_LAMPIRAN),h=ss.getSheetByName(CONFIG.TAB_RIWAYAT);parsed.attachments=a?sheetToObjects(ss,CONFIG.TAB_LAMPIRAN).filter(o=>Object.values(o).some(v=>String(v||'').trim()===String(incidentId).trim())):[];parsed.history=h?sheetToObjects(ss,CONFIG.TAB_RIWAYAT).filter(o=>Object.values(o).some(v=>String(v||'').trim()===String(incidentId).trim())):[];return JSON.stringify(parsed);}catch(e){return JSON.stringify({success:false,message:e.message});}}
function incidentHeaderIndex_(headers,candidates){const h=headers.map(x=>String(x||'').trim().toLowerCase().replace(/[\s_\-\/]+/g,' '));for(const c0 of candidates){const c=String(c0).trim().toLowerCase().replace(/[\s_\-\/]+/g,' '),i=h.indexOf(c);if(i>=0)return i}for(let i=0;i<h.length;i++)if(candidates.some(c0=>{const c=String(c0).trim().toLowerCase().replace(/[\s_\-\/]+/g,' ');return c&&(h[i].includes(c)||c.includes(h[i]))}))return i;return -1;}