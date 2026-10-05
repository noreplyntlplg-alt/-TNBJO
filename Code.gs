/**
 * TNBJO - ระบบออกใบสั่งจ่าย / เช็ค
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีติดตั้ง:
 * 1) สร้าง Google Sheet 1 ไฟล์
 * 2) Extensions > Apps Script แล้ววาง Code.gs นี้
 * 3) ตั้ง Script Property: SPREADSHEET_ID = ID ของ Google Sheet
 *    (ถ้าไม่ตั้ง จะใช้ Spreadsheet ที่ผูกกับสคริปต์)
 * 4) รัน setupSystem() 1 ครั้งเพื่อสร้างหัวตาราง
 * 5) Deploy > New deployment > Web app
 *    Execute as: Me
 *    Who has access: Anyone
 * 6) นำ Web app URL ไปใส่ใน API_URL ของ index.html
 */

const APP = {
  NAME: 'TNBJO',
  VERSION: '1.0.0',
  TZ: Session.getScriptTimeZone() || 'Asia/Bangkok',
  SHEETS: {
    SETTINGS: 'Settings',
    PAYEES: 'Payees',
    PAYMENTS: 'Payments',
    ITEMS: 'PaymentItems',
    AUDIT: 'AuditLog'
  },
  HEADERS: {
    Settings: ['Key','Value','Description','UpdatedAt'],
    Payees: ['ID','Code','Name','TaxId','Address','Bank','BankAccount','Phone','Email','Active','CreatedAt','UpdatedAt'],
    Payments: ['ID','DocNo','PaymentDate','PayeeId','PayeeName','Description','Amount','Vat','Wht','NetAmount','Method','BankName','BankAccount','ChequeNo','Status','Note','CreatedBy','CreatedAt','UpdatedAt'],
    PaymentItems: ['ID','PaymentId','Seq','Description','Qty','Unit','UnitPrice','Amount','CreatedAt'],
    AuditLog: ['ID','Action','Entity','EntityId','DocNo','Details','User','CreatedAt']
  }
};

function doGet(e) {
  return api_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  // รองรับ POST สำหรับเครื่องมือที่ส่งข้อมูลได้โดยไม่ติด CORS
  let p = {};
  try {
    p = JSON.parse(e.postData.contents || '{}');
  } catch (err) {
    p = e && e.parameter ? e.parameter : {};
  }
  return api_(p);
}

function api_(p) {
  try {
    const action = String(p.action || 'bootstrap');
    const payload = p.payload ? JSON.parse(p.payload) : p;
    let result;

    switch (action) {
      case 'bootstrap':
        result = bootstrap_();
        break;
      case 'setup':
        result = setupSystem();
        break;
      case 'dashboard':
        result = dashboard_(payload);
        break;
      case 'settings':
        result = getSettings_();
        break;
      case 'saveSettings':
        result = saveSettings_(payload);
        break;
      case 'payees':
        result = listPayees_(payload);
        break;
      case 'savePayee':
        result = savePayee_(payload);
        break;
      case 'deletePayee':
        result = deletePayee_(payload);
        break;
      case 'payments':
        result = listPayments_(payload);
        break;
      case 'getPayment':
        result = getPayment_(payload);
        break;
      case 'savePayment':
        result = savePayment_(payload);
        break;
      case 'deletePayment':
        result = deletePayment_(payload);
        break;
      case 'audit':
        result = listAudit_(payload);
        break;
      case 'nextDocNo':
        result = {ok:true, docNo: nextDocNo_()};
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }

    return json_(result);
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err), stack:String(err && err.stack || '')});
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('กรุณาตั้ง Script Property: SPREADSHEET_ID');
  return active;
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    const missing = headers.some((h,i) => current[i] !== h);
    if (missing) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  return sh;
}

function setupSystem() {
  Object.keys(APP.HEADERS).forEach(k => {
    const name = APP.SHEETS[k];
    ensureSheet_(name, APP.HEADERS[k]);
  });

  const settings = getSettings_();
  if (!settings.COMPANY_NAME) {
    saveSettings_({
      COMPANY_NAME:'หน่วยงาน/ส่วนราชการ',
      COMPANY_ADDRESS:'',
      COMPANY_TEL:'',
      COMPANY_TAX_ID:'',
      PREFIX:'PAY',
      CURRENCY:'บาท',
      PRINTER_SERVICE_UUID:'',
      PRINTER_CHARACTERISTIC_UUID:'',
      PRINTER_CHUNK:180,
      PRINTER_CODEPAGE:'CP874'
    });
  }
  return {ok:true, message:'สร้างโครงสร้างระบบและหัวตารางเรียบร้อย', sheets:APP.SHEETS};
}

function bootstrap_() {
  setupSystem();
  return {
    ok:true,
    app:APP,
    settings:getSettings_(),
    dashboard:dashboard_({}),
    payees:listPayees_({activeOnly:true, limit:1000}),
    payments:listPayments_({limit:100})
  };
}

function rows_(sheetName) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return values.map(r => {
    const o = {};
    headers.forEach((h,i)=>o[h]=r[i]);
    return o;
  });
}

function appendObject_(sheetName, obj) {
  const sh = ss_().getSheetByName(sheetName);
  const headers = APP.HEADERS[Object.keys(APP.SHEETS).find(k=>APP.SHEETS[k]===sheetName)] || sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function updateObject_(sheetName, id, obj) {
  const sh = ss_().getSheetByName(sheetName);
  const data = rows_(sheetName);
  const idx = data.findIndex(x => String(x.ID) === String(id));
  if (idx < 0) throw new Error('ไม่พบข้อมูล ID: '+id);
  const row = idx + 2;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const old = sh.getRange(row,1,1,headers.length).getValues()[0];
  const merged = headers.map((h,i)=>obj[h] !== undefined ? obj[h] : old[i]);
  sh.getRange(row,1,1,headers.length).setValues([merged]);
}

function uuid_() {
  return Utilities.getUuid();
}

function now_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd HH:mm:ss');
}

function getSettings_() {
  const out = {};
  rows_(APP.SHEETS.SETTINGS).forEach(r => out[String(r.Key)] = r.Value);
  return out;
}

function saveSettings_(data) {
  const sh = ss_().getSheetByName(APP.SHEETS.SETTINGS);
  Object.keys(data || {}).forEach(key => {
    if (key === 'action' || key === 'payload') return;
    const rows = rows_(APP.SHEETS.SETTINGS);
    const index = rows.findIndex(r=>String(r.Key)===String(key));
    if (index >= 0) {
      sh.getRange(index+2,2).setValue(data[key]);
      sh.getRange(index+2,4).setValue(now_());
    } else {
      sh.appendRow([key,data[key],'',now_()]);
    }
  });
  audit_('UPDATE','Settings','SETTINGS','',JSON.stringify(data),'web');
  return {ok:true,settings:getSettings_()};
}

function listPayees_(p) {
  let data = rows_(APP.SHEETS.PAYEES);
  if (p && p.activeOnly) data = data.filter(x=>String(x.Active).toLowerCase() !== 'false');
  if (p && p.q) {
    const q=String(p.q).toLowerCase();
    data=data.filter(x=>[x.Code,x.Name,x.TaxId,x.Bank,x.BankAccount].some(v=>String(v||'').toLowerCase().includes(q)));
  }
  return {ok:true,data:data.slice(0, Number(p.limit||1000))};
}

function savePayee_(p) {
  const d = p || {};
  const id = d.ID || uuid_();
  const old = rows_(APP.SHEETS.PAYEES).find(x=>String(x.ID)===String(id));
  const obj = {
    ID:id, Code:d.Code||'', Name:d.Name||'', TaxId:d.TaxId||'', Address:d.Address||'',
    Bank:d.Bank||'', BankAccount:d.BankAccount||'', Phone:d.Phone||'', Email:d.Email||'',
    Active:d.Active !== false, CreatedAt:old?old.CreatedAt:now_(), UpdatedAt:now_()
  };
  if (old) updateObject_(APP.SHEETS.PAYEES,id,obj); else appendObject_(APP.SHEETS.PAYEES,obj);
  audit_(old?'UPDATE':'CREATE','Payee',id,'',JSON.stringify(obj),'web');
  return {ok:true,data:obj};
}

function deletePayee_(p) {
  const id=String(p.ID||'');
  const sh=ss_().getSheetByName(APP.SHEETS.PAYEES);
  const rows=rows_(APP.SHEETS.PAYEES);
  const idx=rows.findIndex(x=>String(x.ID)===id);
  if(idx<0) throw new Error('ไม่พบผู้รับเงิน');
  sh.getRange(idx+2,10).setValue(false);
  sh.getRange(idx+2,12).setValue(now_());
  audit_('DISABLE','Payee',id,'','ปิดการใช้งาน','web');
  return {ok:true};
}

function listPayments_(p) {
  let data=rows_(APP.SHEETS.PAYMENTS);
  const q=String((p&&p.q)||'').toLowerCase();
  if(q) data=data.filter(x=>[x.DocNo,x.PayeeName,x.Description,x.ChequeNo,x.Status].some(v=>String(v||'').toLowerCase().includes(q)));
  if(p && p.status) data=data.filter(x=>String(x.Status)===String(p.status));
  if(p && p.dateFrom) data=data.filter(x=>String(x.PaymentDate)>=String(p.dateFrom));
  if(p && p.dateTo) data=data.filter(x=>String(x.PaymentDate)<=String(p.dateTo));
  data.sort((a,b)=>String(b.CreatedAt).localeCompare(String(a.CreatedAt)));
  return {ok:true,data:data.slice(0,Number((p&&p.limit)||200))};
}

function getPayment_(p) {
  const id=String(p.ID||'');
  const payment=rows_(APP.SHEETS.PAYMENTS).find(x=>String(x.ID)===id);
  if(!payment) throw new Error('ไม่พบใบสั่งจ่าย');
  const items=rows_(APP.SHEETS.ITEMS).filter(x=>String(x.PaymentId)===id).sort((a,b)=>Number(a.Seq)-Number(b.Seq));
  return {ok:true,data:{payment,items}};
}

function savePayment_(p) {
  const d=p||{};
  if(!d.PayeeId && !d.PayeeName) throw new Error('กรุณาระบุผู้รับเงิน');
  const amount=Number(d.Amount||0);
  if(amount<=0) throw new Error('ยอดเงินต้องมากกว่า 0');

  const id=d.ID||uuid_();
  const old=rows_(APP.SHEETS.PAYMENTS).find(x=>String(x.ID)===String(id));
  const docNo=old?old.DocNo:(d.DocNo||nextDocNo_());
  const vat=Number(d.Vat||0), wht=Number(d.Wht||0);
  const net=Number(d.NetAmount!==undefined?d.NetAmount:(amount+vat-wht));
  const obj={
    ID:id,DocNo:docNo,PaymentDate:d.PaymentDate||Utilities.formatDate(new Date(),APP.TZ,'yyyy-MM-dd'),
    PayeeId:d.PayeeId||'',PayeeName:d.PayeeName||'',Description:d.Description||'',
    Amount:amount,Vat:vat,Wht:wht,NetAmount:net,Method:d.Method||'เช็ค',
    BankName:d.BankName||'',BankAccount:d.BankAccount||'',ChequeNo:d.ChequeNo||'',
    Status:d.Status||'รอจ่าย',Note:d.Note||'',CreatedBy:d.CreatedBy||'web',
    CreatedAt:old?old.CreatedAt:now_(),UpdatedAt:now_()
  };
  if(old) updateObject_(APP.SHEETS.PAYMENTS,id,obj); else appendObject_(APP.SHEETS.PAYMENTS,obj);

  const itemSh=ss_().getSheetByName(APP.SHEETS.ITEMS);
  rows_(APP.SHEETS.ITEMS).filter(x=>String(x.PaymentId)===String(id)).reverse().forEach(x=>{
    const all=rows_(APP.SHEETS.ITEMS);
    const idx=all.findIndex(y=>String(y.ID)===String(x.ID));
    if(idx>=0) itemSh.deleteRow(idx+2);
  });
  (d.Items||[]).forEach((it,i)=>{
    const qty=Number(it.Qty||1), price=Number(it.UnitPrice||0);
    appendObject_(APP.SHEETS.ITEMS,{
      ID:uuid_(),PaymentId:id,Seq:i+1,Description:it.Description||'',
      Qty:qty,Unit:it.Unit||'',UnitPrice:price,Amount:Number(it.Amount!==undefined?it.Amount:qty*price),CreatedAt:now_()
    });
  });

  audit_(old?'UPDATE':'CREATE','Payment',id,docNo,JSON.stringify(obj),'web');
  return getPayment_({ID:id});
}

function deletePayment_(p) {
  const id=String(p.ID||'');
  const sh=ss_().getSheetByName(APP.SHEETS.PAYMENTS);
  const data=rows_(APP.SHEETS.PAYMENTS);
  const idx=data.findIndex(x=>String(x.ID)===id);
  if(idx<0) throw new Error('ไม่พบใบสั่งจ่าย');
  sh.getRange(idx+2,15).setValue('ยกเลิก');
  sh.getRange(idx+2,19).setValue(now_());
  audit_('CANCEL','Payment',id,data[idx].DocNo,'ยกเลิกเอกสาร','web');
  return {ok:true};
}

function nextDocNo_() {
  const settings=getSettings_();
  const prefix=String(settings.PREFIX||'PAY');
  const ym=Utilities.formatDate(new Date(),APP.TZ,'yyyyMM');
  const rows=rows_(APP.SHEETS.PAYMENTS);
  const re=new RegExp('^'+prefix+'-'+ym+'-(\\d+)$');
  let max=0;
  rows.forEach(r=>{
    const m=String(r.DocNo||'').match(re);
    if(m) max=Math.max(max,Number(m[1]));
  });
  return prefix+'-'+ym+'-'+String(max+1).padStart(4,'0');
}

function dashboard_() {
  const payments=rows_(APP.SHEETS.PAYMENTS);
  const active=payments.filter(x=>String(x.Status)!=='ยกเลิก');
  const total=active.reduce((s,x)=>s+Number(x.NetAmount||0),0);
  const pending=active.filter(x=>String(x.Status)==='รอจ่าย').length;
  const paid=active.filter(x=>String(x.Status)==='จ่ายแล้ว').length;
  const month=Utilities.formatDate(new Date(),APP.TZ,'yyyy-MM');
  const monthTotal=active.filter(x=>String(x.PaymentDate||'').startsWith(month)).reduce((s,x)=>s+Number(x.NetAmount||0),0);
  return {ok:true,summary:{documents:active.length,total,pending,paid,monthTotal},recent:active.sort((a,b)=>String(b.CreatedAt).localeCompare(String(a.CreatedAt))).slice(0,10)};
}

function listAudit_(p) {
  let data=rows_(APP.SHEETS.AUDIT);
  if(p&&p.q){const q=String(p.q).toLowerCase();data=data.filter(x=>JSON.stringify(x).toLowerCase().includes(q));}
  data.sort((a,b)=>String(b.CreatedAt).localeCompare(String(a.CreatedAt)));
  return {ok:true,data:data.slice(0,Number((p&&p.limit)||200))};
}

function audit_(action,entity,id,docNo,details,user) {
  appendObject_(APP.SHEETS.AUDIT,{
    ID:uuid_(),Action:action,Entity:entity,EntityId:id,DocNo:docNo||'',Details:details||'',User:user||'web',CreatedAt:now_()
  });
}

// เรียกจาก Apps Script editor ได้โดยตรง เพื่อสร้างหัวตารางอีกครั้ง
function createHeaders() {
  return setupSystem();
}
