// === UBAH DUA KODE INI SEBELUM DIPAKAI (jangan dibagikan sembarangan) ===
const KODE_ADMIN = "GANTI-KODE-ADMIN";
const KODE_KELUARGA = "GANTI-KODE-KELUARGA";
// ========================================================================
const H = ["id","json","nama","orang_tua_id","status","domisili","pasangan","lahir","wafat","catatan","menunggu_persetujuan"];
function sh_(n,h){const ss=SpreadsheetApp.getActiveSpreadsheet();let s=ss.getSheetByName(n);if(!s){s=ss.insertSheet(n);s.appendRow(h)}return s}
function setup(){sh_("Anggota",H);sh_("Pengaturan",["json"]);sh_("Riwayat",["t","json"])}
function doGet(){return HtmlService.createHtmlOutputFromFile("Index").setTitle("Silsilah Keturunan Syaikh Ibrahim").addMetaTag("viewport","width=device-width, initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function cfg_(){const v=sh_("Pengaturan",["json"]).getRange(2,1).getValue();return v?JSON.parse(v):null}
function check(c){return c===KODE_ADMIN?"admin":c===KODE_KELUARGA?"keluarga":""}
function getAll(code){
  const a=sh_("Anggota",H).getDataRange().getValues(),m={},c=cfg_();
  for(let i=1;i<a.length;i++)if(a[i][0]){const d=JSON.parse(a[i][1]);if(code!==KODE_ADMIN&&c&&c.contact===false)delete d.tl;m[a[i][0]]=d}
  const l=sh_("Riwayat",["t","json"]).getDataRange().getValues().slice(1).slice(-40).reverse().map(r=>JSON.parse(r[1]));
  return {members:m,settings:c,log:l};
}
function save(kind,id,data,code,dev){
  const role=check(code);if(!role)throw new Error("Kode salah");
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    if(kind==="settings"){if(role!=="admin")throw new Error("Hanya admin yang boleh mengubah pengaturan");sh_("Pengaturan",["json"]).getRange(2,1).setValue(JSON.stringify(data));return}
    const s=sh_("Anggota",H),a=s.getDataRange().getValues(),c=cfg_()||{};let row=0,old=null;
    for(let i=1;i<a.length;i++)if(a[i][0]===id){row=i+1;old=JSON.parse(a[i][1])}
    if(role!=="admin"){
      if(old){if(!old.by||old.by!==dev||c.own===false)throw new Error("Anda hanya boleh mengubah data yang Anda tambahkan sendiri");data.pending=old.pending;data.by=old.by}
      else{if(data.deleted)throw new Error("Tidak diizinkan");data.pending=c.approve!==false;data.by=dev}
    }
    const r=[id,JSON.stringify(data),data.name||"",data.parent||"",data.status||"",data.place||"",data.sp||"",data.b||"",data.d||"",data.note||"",data.pending?"ya":""];
    if(row)s.getRange(row,1,1,r.length).setValues([r]);else s.appendRow(r);
  }finally{lock.releaseLock()}
}
function addLog(x){sh_("Riwayat",["t","json"]).appendRow([x.t,JSON.stringify({t:x.t,act:String(x.act||"").slice(0,30),name:String(x.name||"").slice(0,80),who:String(x.who||"").slice(0,60)})])}
