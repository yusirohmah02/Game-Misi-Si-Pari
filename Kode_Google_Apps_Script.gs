/**
 * Rekap Misi Si Pari. Tempel kode ini di Extensions > Apps Script
 * pada Google Sheets milik guru. Deploy sebagai Web app.
 * Akses: Anyone; jalankan sebagai pemilik (guru).
 * Jangan bagikan tautan Google Sheets kepada siswa.
 */
const SHEET_NAME = 'Rekap Si Pari';
const HEADERS = ['Waktu diterima','Token','Nama siswa','Misi 1','Misi 2','Misi 3','Misi 4','Total / 100','Refleksi'];
function getSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(SHEET_NAME);
  if(!sh){sh=ss.insertSheet(SHEET_NAME);sh.appendRow(HEADERS);sh.setFrozenRows(1)}
  return sh;
}
function doPost(e){
  const p=(e&&e.parameter)||{};
  const token=String(p.token||'').trim().slice(0,120);
  const name=String(p.nama||'').trim().slice(0,45);
  const scoreKeys=['misi1','misi2','misi3','misi4'];
  const vals=scoreKeys.map(k=>Number(p[k]));
  if(!token||!name||!vals.every(n=>Number.isInteger(n)&&n>=0&&n<=25&&n%5===0)){
    return ContentService.createTextOutput('INVALID');
  }
  const total=vals.reduce((a,b)=>a+b,0);
  if(total!==Number(p.total))return ContentService.createTextOutput('INVALID_TOTAL');
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const sh=getSheet_();
    const finder=sh.getRange(1,2,sh.getLastRow(),1).createTextFinder(token).matchEntireCell(true).findNext();
    if(!finder){
      // Leading apostrophe prevents spreadsheet formula injection in user text.
      const safe=v=>"'"+String(v).replace(/[\r\n]/g,' ').slice(0,500);
      sh.appendRow([new Date(),token,safe(name),...vals,total,safe(p.refleksi||'')]);
    }
    return ContentService.createTextOutput('OK');
  }finally{lock.releaseLock()}
}
function doGet(e){
  const p=(e&&e.parameter)||{};
  if(p.action==='public_scores'){
    const callback=String(p.callback||'');
    if(!/^[A-Za-z_$][A-Za-z0-9_$]{0,60}$/.test(callback))return ContentService.createTextOutput('INVALID_CALLBACK');
    const data=getPublicScores_();
    return jsonp_(callback,{rows:data});
  }
  if(p.action!=='status')return ContentService.createTextOutput('Misi Si Pari Rekap Aktif');
  const callback=String(p.callback||'');
  if(!/^[A-Za-z_$][A-Za-z0-9_$]{0,60}$/.test(callback))return ContentService.createTextOutput('INVALID_CALLBACK');
  const token=String(p.token||'').slice(0,120);
  let found=false;
  if(token){const sh=getSheet_();const cell=sh.getRange(1,2,sh.getLastRow(),1).createTextFinder(token).matchEntireCell(true).findNext();found=!!cell}
  const payload=JSON.stringify({token:token,found:found});
  return ContentService.createTextOutput(callback+'('+payload+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function jsonp_(callback,data){return ContentService.createTextOutput(callback+"("+JSON.stringify(data)+");").setMimeType(ContentService.MimeType.JAVASCRIPT)}

// Dashboard hanya dibuka dari dalam Google Sheets milik guru.
// Jangan membagikan akses spreadsheet kepada siswa.
function onOpen(){
  SpreadsheetApp.getUi().createMenu('🌾 Misi Si Pari')
    .addItem('📊 Dashboard Guru', 'showTeacherDashboard')
    .addToUi();
}
function showTeacherDashboard(){
  const html=HtmlService.createHtmlOutputFromFile('DashboardGuru')
    .setTitle('Dashboard Guru — Misi Si Pari').setWidth(1100).setHeight(680);
  SpreadsheetApp.getUi().showModelessDialog(html,'📊 Dashboard Guru — Misi Si Pari');
}
function getTeacherScores(){
  const values=getSheet_().getDataRange().getValues();
  return values.slice(1).reverse().slice(0,2000).map(r=>({
    waktu:r[0] instanceof Date?Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'):String(r[0]||''),
    nama:String(r[2]||'').replace(/^'/,''),misi1:r[3],misi2:r[4],misi3:r[5],misi4:r[6],total:r[7],refleksi:String(r[8]||'').replace(/^'/,'')
  }));
}

// Publik hanya dapat membaca nama panggilan, waktu dan skor.
// Token dan refleksi tidak dipublikasikan. Tidak ada endpoint edit/hapus.
function getPublicScores_(){
 const sh=getSheet_();const n=sh.getLastRow();if(n<=1)return [];
 const values=sh.getRange(2,1,Math.min(n-1,2000),9).getValues();
 return values.reverse().map(r=>({
  waktu:r[0] instanceof Date?Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm'):String(r[0]||''),
  nama:String(r[2]||'').replace(/^'/,''),
  misi1:Number(r[3])||0,misi2:Number(r[4])||0,misi3:Number(r[5])||0,misi4:Number(r[6])||0,total:Number(r[7])||0
 }));
}
