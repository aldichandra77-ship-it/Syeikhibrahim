// ===== Silsilah Keturunan Syaikh Ibrahim =====
const KODE_ADMIN="Aldi@Syaikh2026!#";
const KODE_KELUARGA="";
const FOTO_FOLDER_ID="GANTI_DENGAN_FOLDER_ID";

const H=["id","json","nama","orang_tua_id","status","domisili","pasangan","lahir","wafat","catatan","menunggu_persetujuan"];

function sh_(n,h){const ss=SpreadsheetApp.getActiveSpreadsheet();let s=ss.getSheetByName(n);if(!s){s=ss.insertSheet(n);s.appendRow(h)}return s}

function setup(){
  sh_("Anggota",H);
  sh_("Riwayat",["t","json"]);
  const s=sh_("Pengaturan",["json"]);
  if(!s.getRange(2,1).getValue()){
    s.getRange(2,1).setValue(JSON.stringify({
      title:"Silsilah Keturunan Syaikh Ibrahim",
      sub:"",
      contact:true,
      approve:false
    }));
  }
}

function doGet(){
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Silsilah Keturunan Syaikh Ibrahim")
    .addMetaTag("viewport","width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}

function doPost(e){
  try{
    const q=JSON.parse(e.postData.contents);
    const ok={getAll:getAll,check:check,save:save,addLog:addLog,uploadFoto:uploadFoto};
    if(!ok[q.fn])throw new Error("Fungsi tidak dikenal");
    return json_({result:ok[q.fn].apply(null,q.args||[])});
  }catch(err){
    return json_({error:String(err.message||err)});
  }
}

function cfg_(){const v=sh_("Pengaturan",["json"]).getRange(2,1).getValue();return v?JSON.parse(v):null}

function check(c){
  c=String(c||"");
  if(KODE_ADMIN&&KODE_ADMIN.indexOf("GANTI")!==0&&c===KODE_ADMIN)return "admin";
  if(KODE_KELUARGA===""||c===KODE_KELUARGA)return "keluarga";
  return "";
}

function getAll(code){
  const a=sh_("Anggota",H).getDataRange().getValues();
  const m={}; const c=cfg_(); const adm=check(code)==="admin";
  for(let i=1;i<a.length;i++) if(a[i][0]){
    const d=JSON.parse(a[i][1]);
    if(!adm&&c&&c.contact===false) delete d.tl;
    m[a[i][0]]=d;
  }
  const l=sh_("Riwayat",["t","json"]).getDataRange().getValues().slice(1).slice(-40).reverse().map(r=>JSON.parse(r[1]));
  return {members:m,settings:c,log:l,needCode:KODE_KELUARGA!==""};
}

function save(kind,id,data,code,dev){
  const role=check(code); if(!role) throw new Error("Kode salah");
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    if(kind==="settings"){
      if(role!=="admin") throw new Error("Hanya admin yang boleh mengubah pengaturan");
      sh_("Pengaturan",["json"]).getRange(2,1).setValue(JSON.stringify(data)); return;
    }
    const s=sh_("Anggota",H),a=s.getDataRange().getValues(); let row=0,old=null;
    for(let i=1;i<a.length;i++) if(a[i][0]===id){row=i+1; old=JSON.parse(a[i][1]);}
    if(role!=="admin"){
      if(data.deleted) throw new Error("Hanya admin yang boleh menghapus data");
      if(old&&old.deleted) throw new Error("Data sudah dihapus");
      data.pending=false; data.by=(old&&old.by)||dev;
    }
    const r=[id,JSON.stringify(data),data.name||"",data.parent||"",data.status||"",data.place||"",data.sp||"",data.b||"",data.d||"",data.note||"",data.pending?"ya":""];
    if(row)s.getRange(row,1,1,r.length).setValues([r]); else s.appendRow(r);
    addLog({t:Date.now(),act:row?"mengubah":"menambah",name:data.name||"",who:data.by||dev||""});
  }finally{lock.releaseLock();}
}

function uploadFoto(base64,nama){
  if(!FOTO_FOLDER_ID||FOTO_FOLDER_ID==="GANTI_DENGAN_FOLDER_ID") throw new Error("Folder foto belum diatur");
  const folder=DriveApp.getFolderById(FOTO_FOLDER_ID);
  const blob=Utilities.newBlob(Utilities.base64Decode(base64),"image/jpeg",(nama||"foto")+".jpg");
  const file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return file.getUrl();
}

function backupSpreadsheet(){
  const file=DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const folder=file.getParents().next();
  file.makeCopy("Backup Silsilah "+Utilities.formatDate(new Date(),"Asia/Jakarta","yyyy-MM-dd HH:mm"),folder);
}

function addLog(x){
  sh_("Riwayat",["t","json"]).appendRow([x.t||Date.now(),JSON.stringify({
    t:x.t||Date.now(),act:String(x.act||"").slice(0,30),name:String(x.name||"").slice(0,80),who:String(x.who||"").slice(0,60)
  })]);
}
