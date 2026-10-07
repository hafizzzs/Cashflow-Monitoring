const APP = Object.freeze({
  TZ: 'Asia/Jakarta', KEY: 'CF_V2_SPREADSHEET',
  LEDGER: 'CashflowLedgerV2', CATEGORIES: 'CashflowCategoriesV2',
  MAX: 1000000000000, IMAGE_LIMIT: 4 * 1024 * 1024
});
const COLS = ['ID','CreatedAt','Type','Amount','Fee','CategoryID','Category','Note','Date','ImageHash'];
const CAT_COLS = ['ID','Type','Name'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Jalankan setup dari script yang terikat spreadsheet.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const old = props.getProperty(APP.KEY);
    if (old && old !== ss.getId()) throw new Error('Script terhubung ke spreadsheet berbeda.');
    const ledger = ensureSheet_(ss, APP.LEDGER, COLS);
    const cats = ensureSheet_(ss, APP.CATEGORIES, CAT_COLS);
    if (cats.getLastRow() === 1) {
      const entries = [];
      ['Gaji','Penjualan','Usaha','Freelance / jasa','Bonus / komisi','Hadiah','Pengembalian dana','Lainnya']
        .forEach(n => entries.push([Utilities.getUuid(), 'IN', n]));
      entries.push(['opening', 'IN', 'Saldo awal']);
      ['Makan & minum','Transportasi','Jalan-jalan & liburan','Belanja','Tagihan','Kesehatan','Hiburan','Pendidikan','Lainnya']
        .forEach(n => entries.push([Utilities.getUuid(), 'OUT', n]));
      entries.push(['internal', 'TRANSFER', 'Transfer antar-rekening sendiri']);
      cats.getRange(2,1,entries.length,3).setValues(entries);
    }
    ledger.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    ledger.getRange('D:E').setNumberFormat('#,##0');
    ledger.getRange('I:I').setNumberFormat('@');
    ss.setSpreadsheetTimeZone(APP.TZ);
    props.setProperty(APP.KEY, ss.getId());
    SpreadsheetApp.flush();
    return 'Setup v2 selesai.';
  } finally { lock.releaseLock(); }
}
function ensureSheet_(ss, name, headers) {
  let s = ss.getSheetByName(name) || ss.insertSheet(name);
  if (!s.getLastRow()) s.getRange(1,1,1,headers.length).setValues([headers]);
  const actual = s.getRange(1,1,1,headers.length).getValues()[0];
  if (!headers.every((h,i) => actual[i] === h)) throw new Error('Header tidak sesuai: ' + name);
  s.setFrozenRows(1);
  s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f0f2f5');
  return s;
}
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.KEY);
  if (!id) throw new Error('Jalankan setup dahulu.');
  return SpreadsheetApp.openById(id);
}
function table_(name, headers) {
  const s = db_().getSheetByName(name);
  if (!s) throw new Error('Sheet tidak ditemukan: ' + name);
  const h = s.getRange(1,1,1,headers.length).getValues()[0];
  if (!headers.every((v,i) => h[i] === v)) throw new Error('Header berubah: ' + name);
  return s;
}
function rows_(s, width) {
  return s.getLastRow() > 1 ? s.getRange(2,1,s.getLastRow()-1,width).getValues() : [];
}
function categories_() {
  return rows_(table_(APP.CATEGORIES, CAT_COLS),3).map(r => ({id:String(r[0]), type:String(r[1]), name:String(r[2])}));
}
function today_() { return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'); }
function day_(value) { return value instanceof Date ? Utilities.formatDate(value,APP.TZ,'yyyy-MM-dd') : String(value); }
function text_(value) {
  const s = String(value || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Cashflow Capture')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover');
}
function getSnapshot() {
  const data = snapshot_(rows_(table_(APP.LEDGER,COLS),COLS.length));
  data.categories = categories_();
  return data;
}
function snapshot_(rows) {
  const today = today_(), month = today.slice(0,7);
  let balance=0, income=0, expense=0;
  rows.forEach(r => {
    const type=String(r[2]), amount=Number(r[3]), fee=Number(r[4]);
    if (!['IN','OUT','TRANSFER'].includes(type) || !Number.isSafeInteger(amount) || amount<=0 || !Number.isSafeInteger(fee) || fee<0)
      throw new Error('Data ledger tidak valid. Periksa spreadsheet.');
    balance += (type==='IN' ? amount : type==='OUT' ? -amount : 0) - fee;
    if (day_(r[8]).slice(0,7)===month) {
      if (type==='IN' && String(r[5])!=='opening') income+=amount;
      expense += (type==='OUT' ? amount : 0) + fee;
    }
  });
  if (![balance,income,expense].every(Number.isSafeInteger)) throw new Error('Total melampaui batas presisi.');
  return {today,month,balance,income,expense,count:rows.length,recent:rows.slice(-30).reverse().map(r => ({
    id:String(r[0]),type:String(r[2]),amount:Number(r[3]),fee:Number(r[4]),category:String(r[6]),note:String(r[7]),date:day_(r[8])
  }))};
}
function addCategory(type, name) {
  if (!['IN','OUT'].includes(type)) throw new Error('Kategori khusus hanya untuk uang masuk/keluar.');
  name=String(name || '').trim();
  if (!name || name.length>50 || /^[=+\-@]/.test(name) || /[\x00-\x1f]/.test(name)) throw new Error('Nama kategori tidak valid (maksimal 50 karakter).');
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const cats=categories_();
    const existing=cats.find(c => c.type===type && c.name.toLowerCase()===name.toLowerCase());
    if (existing) return {category:existing,categories:cats};
    const category={id:Utilities.getUuid(),type,name};
    const s=table_(APP.CATEGORIES,CAT_COLS);
    s.getRange(s.getLastRow()+1,1,1,3).setValues([[category.id,type,name]]);
    SpreadsheetApp.flush();
    return {category,categories:cats.concat(category)};
  } finally { lock.releaseLock(); }
}
function validate_(p) {
  if (!p || typeof p!=='object') throw new Error('Payload tidak valid.');
  const id=String(p.id||''), type=String(p.type||''), amount=Number(p.amount), fee=Number(p.fee);
  const categoryId=String(p.categoryId||''), date=String(p.date||''), note=String(p.note||'').trim(), hash=String(p.hash||'');
  if (!/^[a-zA-Z0-9_-]{16,100}$/.test(id)) throw new Error('ID tidak valid.');
  if (!['IN','OUT','TRANSFER'].includes(type)) throw new Error('Jenis tidak valid.');
  if (!Number.isSafeInteger(amount)||amount<=0||amount>APP.MAX) throw new Error('Nominal tidak valid.');
  if (!Number.isSafeInteger(fee)||fee<0||fee>APP.MAX) throw new Error('Biaya admin tidak valid.');
  if (note.length>200) throw new Error('Catatan maksimal 200 karakter.');
  if (hash && !/^[a-f0-9]{64}$/.test(hash)) throw new Error('Hash gambar tidak valid.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  const d=new Date(date+'T00:00:00Z');
  if (isNaN(d.getTime())||d.toISOString().slice(0,10)!==date||date<'2000-01-01'||date>today_()) throw new Error('Tanggal harus valid dan tidak di masa depan.');
  const cat=categories_().find(c => c.id===categoryId && c.type===type);
  if (!cat) throw new Error('Kategori tidak sesuai jenis transaksi.');
  if (categoryId==='opening' && fee!==0) throw new Error('Saldo awal tidak menggunakan biaya admin.');
  return {id,type,amount,fee,categoryId,category:cat.name,date,note:text_(note),hash};
}
function saveTransaction(payload) {
  let p;
  try { p=validate_(payload); } catch(e) { return {ok:false,error:e.message}; }
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const s=table_(APP.LEDGER,COLS), rows=rows_(s,COLS.length);
    const old=rows.find(r => String(r[0])===p.id);
    if (old) {
      const same=String(old[2])===p.type && Number(old[3])===p.amount && Number(old[4])===p.fee && String(old[5])===p.categoryId && String(old[7])===p.note && day_(old[8])===p.date && String(old[9])===p.hash;
      if (!same) return {ok:false,error:'ID sudah dipakai untuk data berbeda.'};
      return {ok:true,duplicate:true,data:snapshot_(rows)};
    }
    if (p.hash && rows.some(r => String(r[9])===p.hash)) return {ok:false,error:'Gambar identik sudah dicatat. Periksa riwayat.'};
    const row=[p.id,new Date(),p.type,p.amount,p.fee,p.categoryId,p.category,p.note,p.date,p.hash];
    const projected=rows.concat([row]);
    const data=snapshot_(projected);
    s.getRange(s.getLastRow()+1,1,1,COLS.length).setValues([row]);
    SpreadsheetApp.flush();
    return {ok:true,duplicate:false,data};
  } finally { lock.releaseLock(); }
}

// OCR generik: kandidat nominal wajib diperiksa pengguna.
function readReceipt(dataUrl) {
  if (typeof Drive==='undefined') throw new Error('Aktifkan Services → Drive API dahulu.');
  if (typeof dataUrl!=='string' || dataUrl.length>Math.ceil(APP.IMAGE_LIMIT*4/3)+200) throw new Error('Gambar maksimal 4 MB.');
  const m=dataUrl.match(/^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error('Gunakan gambar PNG atau JPEG.');
  const bytes=Utilities.base64Decode(m[2]);
  if (!bytes.length || bytes.length>APP.IMAGE_LIMIT) throw new Error('Ukuran gambar tidak valid.');
  const unsigned=bytes.map(b => (b+256)%256);
  const png=unsigned[0]===137&&unsigned[1]===80&&unsigned[2]===78&&unsigned[3]===71;
  const jpg=unsigned[0]===255&&unsigned[1]===216&&unsigned[2]===255;
  if ((m[1]==='image/png'&&!png)||(m[1]==='image/jpeg'&&!jpg)) throw new Error('Isi gambar tidak sesuai format.');
  const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes).map(b => ('0'+((b+256)%256).toString(16)).slice(-2)).join('');
  const duplicate=rows_(table_(APP.LEDGER,COLS),COLS.length).some(r => String(r[9])===hash);
  if (duplicate) return {duplicate:true,hash,candidates:[],text:'',warning:''};
  let docId='', cleanupFailed=false, raw='';
  try {
    const blob=Utilities.newBlob(bytes,m[1],'receipt');
    const file=Drive.Files.create({name:'CF-OCR-'+Utilities.getUuid(),mimeType:'application/vnd.google-apps.document'},blob,{ocrLanguage:'id',fields:'id'});
    docId=file.id;
    for (let i=0;i<4;i++) {
      try { raw=DocumentApp.openById(docId).getBody().getText(); } catch(e) { if (i===3) throw e; }
      if (raw.trim()) break;
      Utilities.sleep(700);
    }
    if (!raw.trim()) throw new Error('OCR tidak menghasilkan teks. Coba gambar lebih jelas atau isi manual.');
  } finally {
    if (docId) {
      try { Drive.Files.remove(docId); } catch(e) { cleanupFailed=true; }
    }
  }
  const candidates=[];
  const matches=raw.matchAll(/(?:Rp\.?|IDR)\s*([0-9]+(?:[. ][0-9]{3})*(?:,[0-9]{2})?)/gi);
  for (const match of matches) {
    const value=Number(match[1].replace(/[. ]/g,'').replace(',','.'));
    if (Number.isSafeInteger(value)&&value>0&&value<=APP.MAX&&!candidates.includes(value)) candidates.push(value);
  }
  const bank=/mandiri|livin/i.test(raw)?'Livin’ / Mandiri':/\bbtn\b|bal[eé]/i.test(raw)?'balé / BTN':'Belum dikenali';
  return {duplicate:false,hash,bank,candidates:candidates.slice(0,20),text:raw.slice(0,18000),warning:cleanupFailed?'Dokumen OCR sementara gagal dihapus. Periksa file CF-OCR di Google Drive.':''};
}
