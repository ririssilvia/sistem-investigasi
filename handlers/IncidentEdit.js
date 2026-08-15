/** Handler edit/detail/delete laporan insiden. */
function getIncidentDetailForEdit(token, incidentId) {
  try {
    const session = requireSession(token), sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values = sheet.getDataRange().getDisplayValues(); if (values.length <= 1) throw new Error('Data insiden kosong.');
    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']);
    const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)','Site']); if (idIdx < 0) throw new Error('Kolom No Insiden tidak ditemukan.');
    const rows = values.slice(1).filter(r => String(r[idIdx] || '').trim() === String(incidentId).trim()); if (!rows.length) throw new Error('Data insiden tidak ditemukan: ' + incidentId);
    if (String(session.Role || '').toLowerCase() !== 'admin' && siteIdx >= 0 && String(rows[0][siteIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses ke laporan pada site tersebut.');
    return JSON.stringify({success:true,headers:headers,rows:rows,session:session});
  } catch(e){return JSON.stringify({success:false,message:e.message});}
}
function updateIncidentData(token, incidentId, updatedRowsData) {
  try {
    const session=requireSession(token), sheet=getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN); if(!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values=sheet.getDataRange().getDisplayValues(),headers=values[0].map(v=>String(v||'').trim());
    const idIdx=incidentHeaderIndex_(headers,['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']),siteIdx=incidentHeaderIndex_(headers,['Site/Business Unit (BU)','Site']);
    const oldRows=values.slice(1).filter(r=>String(r[idIdx]||'').trim()===String(incidentId).trim()); if(!oldRows.length) throw new Error('Data insiden tidak ditemukan.');
    if(String(session.Role||'').toLowerCase()!=='admin'&&siteIdx>=0&&String(oldRows[0][siteIdx]||'').trim().toLowerCase()!==String(session.Site||'').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses untuk memperbarui laporan ini.');
    if(!Array.isArray(updatedRowsData)||!updatedRowsData.length||updatedRowsData.some(r=>!Array.isArray(r)||r.length!==headers.length)) throw new Error('Struktur data update tidak sesuai dengan Data_Input.');
    for(let i=values.length-1;i>=1;i--) if(String(values[i][idIdx]||'').trim()===String(incidentId).trim()) sheet.deleteRow(i+1);
    sheet.getRange(sheet.getLastRow()+1,1,updatedRowsData.length,headers.length).setValues(updatedRowsData);
    const h=getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT); if(h) h.appendRow(['HIS-'+Utilities.getUuid().substring(0,8).toUpperCase(),incidentId,'Updated','Updated','Pembaruan data laporan insiden',new Date(),session.Nama||session.Username]);
    return JSON.stringify({success:true,message:'Data insiden berhasil diperbarui!'});
  } catch(e){return JSON.stringify({success:false,message:e.message});}
}
function deleteIncident(token,incidentId){
  try{
    const session=requireSession(token),sheet=getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN); if(!sheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values=sheet.getDataRange().getDisplayValues(),headers=values[0].map(v=>String(v||'').trim()),idIdx=incidentHeaderIndex_(headers,['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']),siteIdx=incidentHeaderIndex_(headers,['Site/Business Unit (BU)','Site']);
    const rows=values.slice(1),matches=rows.filter(r=>String(r[idIdx]||'').trim()===String(incidentId).trim()); if(!matches.length) throw new Error('Data insiden tidak ditemukan.');
    if(String(session.Role||'').toLowerCase()!=='admin'&&siteIdx>=0&&String(matches[0][siteIdx]||'').trim().toLowerCase()!==String(session.Site||'').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses untuk menghapus laporan ini.');
    for(let i=values.length-1;i>=1;i--) if(String(values[i][idIdx]||'').trim()===String(incidentId).trim()) sheet.deleteRow(i+1);
    const h=getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT); if(h) h.appendRow(['HIS-'+Utilities.getUuid().substring(0,8).toUpperCase(),incidentId,'Deleted','Deleted','Laporan dihapus',new Date(),session.Nama||session.Username]);
    return {success:true,message:'Laporan '+incidentId+' berhasil dihapus.'};
  }catch(e){return {success:false,message:e.message};}
}
function getIncidentDetail(token,incidentId){
  const raw=getIncidentDetailForEdit(token,incidentId); try{const parsed=JSON.parse(raw); if(!parsed.success)return raw; const ss=getTransaksiSS(),a=ss.getSheetByName(CONFIG.TAB_LAMPIRAN),h=ss.getSheetByName(CONFIG.TAB_RIWAYAT); parsed.attachments=a?sheetToObjects(ss,CONFIG.TAB_LAMPIRAN).filter(o=>Object.values(o).some(v=>String(v||'').trim()===String(incidentId).trim())):[]; parsed.history=h?sheetToObjects(ss,CONFIG.TAB_RIWAYAT).filter(o=>Object.values(o).some(v=>String(v||'').trim()===String(incidentId).trim())):[]; return JSON.stringify(parsed);}catch(e){return JSON.stringify({success:false,message:e.message});}}
function incidentHeaderIndex_(headers,candidates){const h=headers.map(x=>String(x||'').trim().toLowerCase().replace(/[\s_\-\/]+/g,' '));for(const c0 of candidates){const c=String(c0).trim().toLowerCase().replace(/[\s_\-\/]+/g,' '),i=h.indexOf(c);if(i>=0)return i}for(let i=0;i<h.length;i++)if(candidates.some(c0=>{const c=String(c0).trim().toLowerCase().replace(/[\s_\-\/]+/g,' ');return c&&(h[i].includes(c)||c.includes(h[i]))}))return i;return -1;}