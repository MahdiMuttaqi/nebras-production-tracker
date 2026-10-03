const PICK_QUEUE_KEY = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903";
const PICK_QUEUE_FILE_PROP = "NEBRAS_PICK_QUEUE_SHEET_ID";
const PICK_QUEUE_SHEET = "PickQueueV2";
const PICK_QUEUE_ISSUES_SHEET = "PickQueueIssues";
const PICK_QUEUE_HEADERS = [
  "OrderId","Batch","Customer","LineId","Code","Qty","Location",
  "ItemStatus","Worker","OrderStatus","Note","CreatedAt","UpdatedAt",
  "ActualQty","IssueNote","ClientActionId"
];
const PICK_QUEUE_BASE_HEADERS = PICK_QUEUE_HEADERS.slice(0,13);
// Storage stays positional and stable.  These Persian labels are presentation
// only, so changing a visible header never changes the queue's data contract.
const PICK_QUEUE_DISPLAY_HEADERS = [
  "شناسه سفارش","شماره دسته","مشتری","شناسه ردیف","کد محصول","تعداد درخواستی","موقعیت انبار",
  "وضعیت کالا","انباردار مسئول","وضعیت سفارش","توضیحات","تاریخ ثبت","تاریخ آخرین تغییر",
  "تعداد واقعی","توضیح پیگیری","شناسه عملیات موبایل"
];
const PICK_QUEUE_DISPLAY_BASE_HEADERS = PICK_QUEUE_DISPLAY_HEADERS.slice(0,13);
const PICK_QUEUE_ISSUES_HEADERS = [
  "IssueKey","OrderId","Batch","Customer","Code","RequestedQty","ActualQty",
  "ShortageQty","IssueType","IssueNote","FollowUpStatus","UpdatedAt"
];
const PICK_QUEUE_ISSUES_DISPLAY_HEADERS = [
  "شناسه پیگیری","شناسه سفارش","شماره دسته","مشتری","کد محصول","تعداد درخواستی","تعداد واقعی",
  "تعداد کسری","نوع مورد","توضیح پیگیری","وضعیت پیگیری","تاریخ آخرین تغییر"
];
// Independent, server-side print/reconciliation queue.  It is intentionally
// separate from the inventory, history and the existing Excel shortage form.
// Excel only projects these rows into its print form when the operator asks it
// to do so; no mobile click can change stock or a finalized output record.
const PICK_QUEUE_PRINT_SHEET = "PickQueuePrint";
const PICK_QUEUE_PRINT_HEADERS = [
  "PrintKey","Batch","Customer","Code","Source","RequestedQty","ActualQty","ShortageQty",
  "IssueType","IssueNote","IssueState","CheckStatus","CheckedAt","UpdatedAt"
];
const PICK_QUEUE_PRINT_DISPLAY_HEADERS = [
  "شناسه مورد چاپ","شماره دسته","مشتری","کد محصول","منبع","تعداد درخواستی","تعداد واقعی","تعداد برای چاپ",
  "نوع مورد","توضیح","وضعیت مورد","وضعیت تطبیق","زمان تطبیق","تاریخ آخرین تغییر"
];
const PICK_QUEUE_FREIGHT_SHEET = "PickQueueFreight";
const PICK_QUEUE_FREIGHT_HEADERS = ["GroupKey","Batch","Customer","FreightStatus","UpdatedAt","ClientActionId"];
const PICK_QUEUE_FREIGHT_DISPLAY_HEADERS = ["شناسه گروه","شماره دسته","مشتری","وضعیت تحویل","تاریخ آخرین تغییر","شناسه عملیات موبایل"];
const PICK_QUEUE_TIME_ZONE = "Asia/Tehran";

function pickOut_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function pickTextOut_(text) {
  return ContentService.createTextOutput(String(text == null ? "" : text));
}

function pickAuthorized_(key) {
  return String(key || "") === PICK_QUEUE_KEY;
}

function pickCodeText_(value) {
  if (value instanceof Date) {
    return String(value.getMonth() + 1).padStart(2, "0") + "-" + String(value.getFullYear());
  }
  return String(value == null ? "" : value);
}

function pickHeadersMatch_(first, headers) {
  return headers.every((header,index)=>String(first[index] || "") === header);
}

function pickFaDigits_(value) {
  return String(value == null ? "" : value).replace(/\d/g,function(d){
    return "۰۱۲۳۴۵۶۷۸۹".charAt(Number(d));
  });
}

function pickGregorianToJalali_(gy, gm, gd) {
  const monthDays=[0,31,59,90,120,151,181,212,243,273,304,334];
  const gy2=gm>2 ? gy+1 : gy;
  let days=355666 + 365*gy + Math.floor((gy2+3)/4) - Math.floor((gy2+99)/100) + Math.floor((gy2+399)/400) + gd + monthDays[gm-1];
  let jy=-1595 + 33*Math.floor(days/12053);
  days%=12053;
  jy+=4*Math.floor(days/1461);
  days%=1461;
  if (days>365) {
    jy+=Math.floor((days-1)/365);
    days=(days-1)%365;
  }
  const jm=days<186 ? 1+Math.floor(days/31) : 7+Math.floor((days-186)/30);
  const jd=1+(days<186 ? days%31 : (days-186)%30);
  return [jy,jm,jd];
}

function pickJalaliDateTime_(value) {
  const source=value instanceof Date ? value : new Date(value || Date.now());
  const parts=Utilities.formatDate(source,PICK_QUEUE_TIME_ZONE,"yyyy|M|d|HH|mm|ss").split("|").map(Number);
  const j=pickGregorianToJalali_(parts[0],parts[1],parts[2]);
  const pad=function(n){return String(n).padStart(2,"0");};
  return pickFaDigits_(j[0]+"/"+pad(j[1])+"/"+pad(j[2])+" "+pad(parts[3])+":"+pad(parts[4])+":"+pad(parts[5]));
}

function pickDateText_(value) {
  return value instanceof Date ? pickJalaliDateTime_(value) : String(value == null ? "" : value);
}

function pickSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PICK_QUEUE_FILE_PROP);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { id = ""; }
  }
  if (!id) {
    ss = SpreadsheetApp.create("Nebras Pick Queue Data");
    props.setProperty(PICK_QUEUE_FILE_PROP, ss.getId());
  }

  let sh = ss.getSheetByName(PICK_QUEUE_SHEET);
  if (!sh) sh = ss.insertSheet(PICK_QUEUE_SHEET);

  const width = Math.max(sh.getLastColumn(), PICK_QUEUE_HEADERS.length);
  const first = sh.getRange(1,1,1,width).getValues()[0];
  const legacyBaseValid = pickHeadersMatch_(first,PICK_QUEUE_BASE_HEADERS);
  const displayBaseValid = pickHeadersMatch_(first,PICK_QUEUE_DISPLAY_BASE_HEADERS);
  const hasData = sh.getLastRow() > 1;
  if (!legacyBaseValid && !displayBaseValid && hasData) throw new Error("PickQueueV2 header is not recognized; queue data was left unchanged");
  if (!pickHeadersMatch_(first,PICK_QUEUE_DISPLAY_HEADERS)) {
    sh.getRange(1,1,1,PICK_QUEUE_DISPLAY_HEADERS.length).setValues([PICK_QUEUE_DISPLAY_HEADERS]);
  }
  sh.setFrozenRows(1);

  // If the independent issues tab already exists, make its display labels
  // consistent too.  This never creates a report tab during normal listing.
  if (ss.getSheetByName(PICK_QUEUE_ISSUES_SHEET)) pickIssueSheet_(ss);
  return sh;
}

function pickDoGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const cb = p.callback || "";
    if (!pickAuthorized_(p.key)) return pickOut_({ok:false,error:"unauthorized"}, cb);
    if (p.action === "pickQueueList") return pickOut_(pickQueueList_(p), cb);
    if (p.action === "pickQueueConfirm") return pickOut_(pickQueueConfirm_(p), cb);
    if (p.action === "pickQueueSet") return pickOut_(pickQueueSet_(p), cb);
    if (p.action === "pickQueuePrintList") return pickOut_(pickQueuePrintList_(p), cb);
    if (p.action === "pickQueuePrintCheck") return pickOut_(pickQueuePrintCheck_(p), cb);
    if (p.action === "pickQueuePrintHandoff") return pickOut_(pickQueuePrintHandoff_(p), cb);
    if (p.action === "pickQueuePrintExport") return pickTextOut_(pickQueuePrintExport_(p));
    return pickOut_({ok:true,name:"صف برداشت سفارش‌های نبراس",service:"pick-queue-v13-v2"}, cb);
  } catch (err) {
    return pickOut_({ok:false,error:String(err && err.message || err)}, e && e.parameter && e.parameter.callback);
  }
}

function pickDoPost(e) {
  try {
    const type = String((e && e.postData && e.postData.type) || "");
    let p = (e && e.parameter) || {};
    if (type.indexOf("application/json") >= 0) {
      p = Object.assign({}, p, JSON.parse(e.postData.contents || "{}"));
    }
    if (!pickAuthorized_(p.key)) return pickOut_({ok:false,error:"unauthorized"});
    if (p.action === "pickQueuePut") return pickOut_(pickQueuePut_(p));
    if (p.action === "pickQueuePutMany") return pickOut_(pickQueuePutMany_(p));
    if (p.action === "pickQueueConfirm") return pickOut_(pickQueueConfirm_(p));
    if (p.action === "pickQueueSet") return pickOut_(pickQueueSet_(p));
    if (p.action === "pickQueuePrintImport") return pickOut_(pickQueuePrintImport_(p));
    if (p.action === "pickQueuePrintCheck") return pickOut_(pickQueuePrintCheck_(p));
    if (p.action === "pickQueuePrintHandoff") return pickOut_(pickQueuePrintHandoff_(p));
    return pickOut_({ok:false,error:"invalid action"});
  } catch (err) {
    return pickOut_({ok:false,error:String(err && err.message || err)});
  }
}

function pickQueuePut_(p) {
  const payload = typeof p.payload === "string" ? JSON.parse(p.payload) : p.payload;
  return pickQueuePutOrders_([payload]);
}

function pickQueuePutMany_(p) {
  const payload = typeof p.payload === "string" ? JSON.parse(p.payload) : p.payload;
  if (!payload || !Array.isArray(payload.orders)) throw new Error("invalid payload");
  return pickQueuePutOrders_(payload.orders);
}

function pickQueuePutOrders_(orders) {
  const cleanOrders = orders.filter(o=>o && o.orderId && Array.isArray(o.items));
  if (!cleanOrders.length) throw new Error("no orders");
  // The Excel sender can retry after a slow/ambiguous response.  OrderId is a
  // permanent idempotency key: never overwrite a card that already exists,
  // otherwise a retry could erase a picker’s recorded status on mobile.
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = pickSheet_();
    const lastRow=sh.getLastRow();
    const existing=new Set();
    if (lastRow>1) {
      sh.getRange(2,1,lastRow-1,1).getValues().forEach(r=>{
        const id=String(r[0]||""); if (id) existing.add(id);
      });
    }

    const now = pickJalaliDateTime_(new Date()), rows=[], alreadyPresent=[];
    let acceptedOrders=0;
    cleanOrders.forEach(order=>{
      const orderId=String(order.orderId);
      if (existing.has(orderId)) { alreadyPresent.push(orderId); return; }
      acceptedOrders++;
      order.items.forEach(it=>rows.push([
        orderId,
        String(order.batch || ""),
        String(order.customer || ""),
        String(it.lineId || ""),
        pickCodeText_(it.code),
        Number(it.qty || 0),
        String(it.location || ""),
        "در انتظار",
        "",
        "در انتظار برداشت",
        "",
        now,
        now,
        "",
        "",
        ""
      ]));
    });
    if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,PICK_QUEUE_HEADERS.length).setValues(rows);
    sh.setFrozenRows(1);
    return {ok:true,orders:acceptedOrders,items:rows.length,alreadyPresent};
  } finally {
    lock.releaseLock();
  }
}

function pickQueueConfirm_(p) {
  const raw=p.orderIds===undefined ? [] : (Array.isArray(p.orderIds) ? p.orderIds : String(p.orderIds).split("|"));
  const wanted=[...new Set(raw.map(x=>String(x||"").trim()).filter(Boolean))];
  if (!wanted.length) throw new Error("orderIds required");
  const wantedSet=new Set(wanted), present=new Set();
  const sh=pickSheet_(), lastRow=sh.getLastRow();
  if (lastRow>1) sh.getRange(2,1,lastRow-1,1).getValues().forEach(r=>{
    const id=String(r[0]||""); if (wantedSet.has(id)) present.add(id);
  });
  return {ok:true,present:wanted.filter(id=>present.has(id)),missing:wanted.filter(id=>!present.has(id))};
}

function pickIssueSheet_(ss) {
  let sh=ss.getSheetByName(PICK_QUEUE_ISSUES_SHEET);
  if (!sh) sh=ss.insertSheet(PICK_QUEUE_ISSUES_SHEET);
  const width=Math.max(sh.getLastColumn(),PICK_QUEUE_ISSUES_HEADERS.length);
  const first=sh.getRange(1,1,1,width).getValues()[0];
  const hasData=sh.getLastRow()>1;
  const legacyValid=pickHeadersMatch_(first,PICK_QUEUE_ISSUES_HEADERS);
  const displayValid=pickHeadersMatch_(first,PICK_QUEUE_ISSUES_DISPLAY_HEADERS);
  if (!legacyValid && !displayValid && hasData) throw new Error("PickQueueIssues header is not recognized; issue data was left unchanged");
  if (!displayValid) sh.getRange(1,1,1,PICK_QUEUE_ISSUES_DISPLAY_HEADERS.length).setValues([PICK_QUEUE_ISSUES_DISPLAY_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function pickPrintSheet_(ss) {
  let sh=ss.getSheetByName(PICK_QUEUE_PRINT_SHEET);
  if (!sh) sh=ss.insertSheet(PICK_QUEUE_PRINT_SHEET);
  const width=Math.max(sh.getLastColumn(),PICK_QUEUE_PRINT_HEADERS.length);
  const first=sh.getRange(1,1,1,width).getValues()[0];
  const hasData=sh.getLastRow()>1;
  const legacyValid=pickHeadersMatch_(first,PICK_QUEUE_PRINT_HEADERS);
  const displayValid=pickHeadersMatch_(first,PICK_QUEUE_PRINT_DISPLAY_HEADERS);
  if (!legacyValid && !displayValid && hasData) throw new Error("PickQueuePrint header is not recognized; print data was left unchanged");
  if (!displayValid) sh.getRange(1,1,1,PICK_QUEUE_PRINT_DISPLAY_HEADERS.length).setValues([PICK_QUEUE_PRINT_DISPLAY_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function pickExistingPrintSheet_(ss) {
  return ss.getSheetByName(PICK_QUEUE_PRINT_SHEET) ? pickPrintSheet_(ss) : null;
}

function pickFreightSheet_(ss) {
  let sh=ss.getSheetByName(PICK_QUEUE_FREIGHT_SHEET);
  if (!sh) sh=ss.insertSheet(PICK_QUEUE_FREIGHT_SHEET);
  const width=Math.max(sh.getLastColumn(),PICK_QUEUE_FREIGHT_HEADERS.length);
  const first=sh.getRange(1,1,1,width).getValues()[0];
  const hasData=sh.getLastRow()>1;
  const legacyValid=pickHeadersMatch_(first,PICK_QUEUE_FREIGHT_HEADERS);
  const displayValid=pickHeadersMatch_(first,PICK_QUEUE_FREIGHT_DISPLAY_HEADERS);
  if (!legacyValid && !displayValid && hasData) throw new Error("PickQueueFreight header is not recognized; freight data was left unchanged");
  if (!displayValid) sh.getRange(1,1,1,PICK_QUEUE_FREIGHT_DISPLAY_HEADERS.length).setValues([PICK_QUEUE_FREIGHT_DISPLAY_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function pickExistingFreightSheet_(ss) {
  return ss.getSheetByName(PICK_QUEUE_FREIGHT_SHEET) ? pickFreightSheet_(ss) : null;
}

function pickNumber_(value) {
  const n=Number(value);
  return Number.isFinite(n) ? n : 0;
}

function pickKeyPart_(value) {
  return encodeURIComponent(String(value == null ? "" : value).trim());
}

function pickPrintGroupKey_(batch,customer) {
  return "G|"+pickKeyPart_(batch)+"|"+pickKeyPart_(customer);
}

function pickPrintRowKey_(batch,customer,code) {
  return pickPrintGroupKey_(batch,customer)+"|"+pickKeyPart_(code);
}

function pickInitialPrintKey_(batch,customer,code) {
  return "INITIAL|"+pickKeyPart_(batch)+"|"+pickKeyPart_(customer)+"|"+pickKeyPart_(code);
}

function pickMobilePrintKey_(orderId,code) {
  return "MOBILE|"+pickKeyPart_(orderId)+"|"+pickKeyPart_(code);
}

function pickPrintBusinessEquals_(row,record) {
  return String(row[1]||"")===record.batch &&
    String(row[2]||"")===record.customer &&
    pickCodeText_(row[3])===record.code &&
    String(row[4]||"")===record.source &&
    pickNumber_(row[5])===record.requestedQty &&
    pickNumber_(row[6])===record.actualQty &&
    pickNumber_(row[7])===record.shortageQty &&
    String(row[8]||"")===record.issueType &&
    String(row[9]||"")===record.issueNote &&
    String(row[10]||"")===record.issueState;
}

function pickQueueUpsertPrintRecords_(ss,records) {
  const clean=(records||[]).filter(r=>r && r.printKey && r.batch && r.customer && r.code).map(r=>({
    printKey:String(r.printKey), batch:String(r.batch), customer:String(r.customer), code:pickCodeText_(r.code),
    source:String(r.source||""), requestedQty:pickNumber_(r.requestedQty), actualQty:pickNumber_(r.actualQty),
    shortageQty:Math.max(0,pickNumber_(r.shortageQty)), issueType:String(r.issueType||""),
    issueNote:String(r.issueNote||"").trim(), issueState:String(r.issueState||"نیازمند پیگیری")
  }));
  if (!clean.length) return {changed:0};
  const sh=pickPrintSheet_(ss), lastRow=sh.getLastRow();
  const values=lastRow>1 ? sh.getRange(2,1,lastRow-1,PICK_QUEUE_PRINT_HEADERS.length).getValues() : [];
  const index={};
  values.forEach((row,i)=>{const key=String(row[0]||""); if (key) index[key]=i;});
  const now=pickJalaliDateTime_(new Date());
  let changed=0;
  clean.forEach(record=>{
    const oldIndex=index[record.printKey];
    if (oldIndex===undefined) {
      values.push([
        record.printKey,record.batch,record.customer,record.code,record.source,record.requestedQty,record.actualQty,record.shortageQty,
        record.issueType,record.issueNote,record.issueState,record.issueState==="رفع شد"?"رفع شد":"در انتظار تطبیق","",now
      ]);
      index[record.printKey]=values.length-1;
      changed++;
      return;
    }
    const old=values[oldIndex];
    if (pickPrintBusinessEquals_(old,record)) return;
    const active=record.issueState!=="رفع شد";
    values[oldIndex]=[
      record.printKey,record.batch,record.customer,record.code,record.source,record.requestedQty,record.actualQty,record.shortageQty,
      record.issueType,record.issueNote,record.issueState,active?"در انتظار تطبیق":"رفع شد","",now
    ];
    changed++;
  });
  if (changed) sh.getRange(2,1,values.length,PICK_QUEUE_PRINT_HEADERS.length).setValues(values);
  return {changed:changed};
}

function pickQueueReadFreightMap_(ss) {
  const sh=pickExistingFreightSheet_(ss);
  const map={};
  if (!sh || sh.getLastRow()<2) return map;
  sh.getRange(2,1,sh.getLastRow()-1,PICK_QUEUE_FREIGHT_HEADERS.length).getValues().forEach(row=>{
    const key=String(row[0]||"");
    if (key) map[key]={status:String(row[3]||""),updatedAt:String(row[4]||"")};
  });
  return map;
}

function pickQueueBuildPrintList_(ss,batchFilter) {
  const sh=pickExistingPrintSheet_(ss);
  if (!sh || sh.getLastRow()<2) return {customers:[],totalRows:0};
  const freight=pickQueueReadFreightMap_(ss), groups={};
  const values=sh.getRange(2,1,sh.getLastRow()-1,PICK_QUEUE_PRINT_HEADERS.length).getValues();
  values.forEach(row=>{
    const batch=String(row[1]||""),customer=String(row[2]||""),code=pickCodeText_(row[3]);
    const issueState=String(row[10]||""),issueType=String(row[8]||"");
    const shortageQty=Math.max(0,pickNumber_(row[7]));
    if (!batch || !customer || !code || issueState==="رفع شد") return;
    if (batchFilter && batch!==batchFilter) return;
    // A zero-quantity mismatch is still useful for reconciliation, even though
    // it is not added to the physical production/print quantity.
    if (shortageQty<=0 && issueType!=="مغایرت") return;
    const groupKey=pickPrintGroupKey_(batch,customer),rowKey=pickPrintRowKey_(batch,customer,code);
    if (!groups[groupKey]) groups[groupKey]={
      groupKey:groupKey,batch:batch,customer:customer,rows:{},lastUpdatedAt:""
    };
    const group=groups[groupKey];
    if (!group.rows[rowKey]) group.rows[rowKey]={
      rowKey:rowKey,code:code,requestedQty:0,actualQty:0,shortageQty:0,sources:[],issueTypes:[],notes:[],checked:true
    };
    const target=group.rows[rowKey];
    target.requestedQty+=pickNumber_(row[5]);
    target.actualQty+=pickNumber_(row[6]);
    target.shortageQty+=shortageQty;
    const source=String(row[4]||"");
    const note=String(row[9]||"").trim();
    if (source && target.sources.indexOf(source)<0) target.sources.push(source);
    if (issueType && target.issueTypes.indexOf(issueType)<0) target.issueTypes.push(issueType);
    if (note && target.notes.indexOf(note)<0) target.notes.push(note);
    if (String(row[11]||"")!=="تأیید شد") target.checked=false;
    const updatedAt=String(row[13]||"");
    if (updatedAt>group.lastUpdatedAt) group.lastUpdatedAt=updatedAt;
  });
  const customers=[];
  Object.keys(groups).forEach(groupKey=>{
    const group=groups[groupKey], handoff=freight[groupKey];
    if (handoff && handoff.status==="تحویل باربری" && String(handoff.updatedAt||"")>=String(group.lastUpdatedAt||"")) return;
    const rows=Object.keys(group.rows).map(key=>group.rows[key]).sort((a,b)=>String(a.code).localeCompare(String(b.code)));
    customers.push({
      groupKey:group.groupKey,batch:group.batch,customer:group.customer,rows:rows,
      allChecked:rows.length>0 && rows.every(x=>x.checked),lastUpdatedAt:group.lastUpdatedAt
    });
  });
  customers.sort((a,b)=>String(b.batch).localeCompare(String(a.batch)) || String(a.customer).localeCompare(String(b.customer)));
  return {customers:customers,totalRows:customers.reduce((n,c)=>n+c.rows.length,0)};
}

function pickQueuePrintImport_(p) {
  const payload=typeof p.payload==="string" ? JSON.parse(p.payload) : p.payload;
  const batch=String((payload&&payload.batch)||p.batch||"").trim();
  const items=payload&&Array.isArray(payload.items) ? payload.items : [];
  if (!batch) throw new Error("batch required");
  const merged={};
  items.forEach(item=>{
    const customer=String(item&&item.customer||"").trim(),code=pickCodeText_(item&&item.code).trim();
    const qty=Math.max(0,pickNumber_(item&&(item.qty!==undefined?item.qty:item.shortageQty)));
    if (!customer || !code || qty<=0) return;
    const key=customer+"\u0001"+code;
    if (!merged[key]) merged[key]={customer:customer,code:code,qty:0,notes:[]};
    merged[key].qty+=qty;
    const note=String(item&&item.note||"").trim();
    if (note && merged[key].notes.indexOf(note)<0) merged[key].notes.push(note);
  });
  const records=Object.keys(merged).map(key=>{
    const item=merged[key];
    return {
      printKey:pickInitialPrintKey_(batch,item.customer,item.code),batch:batch,customer:item.customer,code:item.code,
      source:"اکسل: کسری و ناموجود",requestedQty:item.qty,actualQty:0,shortageQty:item.qty,
      issueType:"کسری/ناموجود",issueNote:item.notes.join(" | "),issueState:"نیازمند پیگیری"
    };
  });
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error("queue busy; retry");
  try {
    const result=pickQueueUpsertPrintRecords_(pickSheet_().getParent(),records);
    return {ok:true,imported:records.length,changed:result.changed,batch:batch};
  } finally {
    lock.releaseLock();
  }
}

function pickQueuePrintList_(p) {
  const result=pickQueueBuildPrintList_(pickSheet_().getParent(),String(p.batch||"").trim());
  return {ok:true,customers:result.customers,totalRows:result.totalRows};
}

function pickQueuePrintCheck_(p) {
  const groupKey=String(p.groupKey||""),rowKey=String(p.rowKey||"");
  const checked=String(p.checked||"")==="1" || p.checked===true;
  if (!groupKey || !rowKey) throw new Error("groupKey and rowKey required");
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error("queue busy; retry");
  try {
    const ss=pickSheet_().getParent(), sh=pickExistingPrintSheet_(ss);
    if (!sh || sh.getLastRow()<2) throw new Error("print row not found");
    const range=sh.getRange(2,1,sh.getLastRow()-1,PICK_QUEUE_PRINT_HEADERS.length), values=range.getValues();
    const target=[];
    values.forEach((row,i)=>{
      if (String(row[10]||"")==="رفع شد") return;
      if (pickPrintGroupKey_(row[1],row[2])===groupKey && pickPrintRowKey_(row[1],row[2],row[3])===rowKey) target.push(i);
    });
    if (!target.length) throw new Error("print row not found");
    const now=pickJalaliDateTime_(new Date());
    target.forEach(i=>{
      values[i][11]=checked?"تأیید شد":"در انتظار تطبیق";
      values[i][12]=checked?now:"";
      values[i][13]=now;
    });
    range.setValues(values);
    return {ok:true,changed:target.length,checked:checked,clientActionId:String(p.clientActionId||"")};
  } finally {
    lock.releaseLock();
  }
}

function pickQueuePrintHandoff_(p) {
  const groupKey=String(p.groupKey||"");
  if (!groupKey) throw new Error("groupKey required");
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error("queue busy; retry");
  try {
    const ss=pickSheet_().getParent(), list=pickQueueBuildPrintList_(ss,"");
    const group=list.customers.find(x=>String(x.groupKey)===groupKey);
    if (!group) {
      const old=pickQueueReadFreightMap_(ss)[groupKey];
      if (old && old.status==="تحویل باربری") return {ok:true,alreadyDone:true,clientActionId:String(p.clientActionId||"")};
      throw new Error("customer print list not found");
    }
    if (!group.allChecked) throw new Error("all print rows must be checked before freight handoff");
    const sh=pickFreightSheet_(ss),lastRow=sh.getLastRow();
    const values=lastRow>1 ? sh.getRange(2,1,lastRow-1,PICK_QUEUE_FREIGHT_HEADERS.length).getValues() : [];
    const index=values.findIndex(row=>String(row[0]||"")===groupKey);
    const record=[groupKey,group.batch,group.customer,"تحویل باربری",pickJalaliDateTime_(new Date()),String(p.clientActionId||"")];
    if (index>=0) sh.getRange(index+2,1,1,record.length).setValues([record]);
    else sh.getRange(sh.getLastRow()+1,1,1,record.length).setValues([record]);
    return {ok:true,groupKey:groupKey,clientActionId:String(p.clientActionId||"")};
  } finally {
    lock.releaseLock();
  }
}

function pickTsvText_(value) {
  return String(value == null ? "" : value).replace(/[\t\r\n]/g," ");
}

function pickQueuePrintExport_(p) {
  const batch=String(p.batch||"").trim();
  if (!batch) throw new Error("batch required");
  const list=pickQueueBuildPrintList_(pickSheet_().getParent(),batch);
  const lines=["Batch\tCustomer\tCode\tQty\tSource\tIssueType\tNote"];
  list.customers.forEach(customer=>customer.rows.forEach(row=>{
    // Only positive shortages become a production/print quantity.  A zero
    // mismatch stays visible in the mobile reconciliation checklist instead.
    if (pickNumber_(row.shortageQty)<=0) return;
    lines.push([
      customer.batch,customer.customer,row.code,row.shortageQty,row.sources.join(" | "),
      row.issueTypes.join(" | "),row.notes.join(" | ")
    ].map(pickTsvText_).join("\t"));
  }));
  return lines.join("\r\n");
}

function pickQueueSyncIssueReport_(queueSheet, values, orderId, targetIndexes, itemStatus, now) {
  if (!targetIndexes.length) return;
  const first=values[targetIndexes[0]], code=pickCodeText_(first[4]);
  const issueKey=String(orderId)+"|"+String(code);
  const isIssue=itemStatus==="ناقص" || itemStatus==="مغایرت";
  const ss=queueSheet.getParent();
  let issueSheet=ss.getSheetByName(PICK_QUEUE_ISSUES_SHEET);
  // Do not create or query a reporting sheet for normal picked rows.
  if (!issueSheet && !isIssue) return;
  issueSheet=pickIssueSheet_(ss);
  const lastRow=issueSheet.getLastRow();
  const report=lastRow>1 ? issueSheet.getRange(2,1,lastRow-1,PICK_QUEUE_ISSUES_HEADERS.length).getValues() : [];
  const existingIndex=report.findIndex(r=>String(r[0]||"")===issueKey);
  if (!isIssue && existingIndex<0) return;

  const requestedQty=targetIndexes.reduce((sum,i)=>sum+Number(values[i][5]||0),0);
  const actualQty=first[13] === "" || first[13] == null ? requestedQty : Number(first[13]||0);
  const previous=existingIndex>=0 ? report[existingIndex] : null;
  const issueNote=isIssue ? String(first[14]||"").trim() : String(previous?.[9]||"").trim();
  const record=[
    issueKey, String(orderId), String(first[1]||""), String(first[2]||""), code,
    requestedQty, actualQty, Math.max(0,requestedQty-actualQty),
    isIssue ? itemStatus : "رفع شد", issueNote,
    isIssue ? "نیازمند پیگیری" : "رفع شد", now
  ];
  if (existingIndex>=0) issueSheet.getRange(existingIndex+2,1,1,record.length).setValues([record]);
  else issueSheet.getRange(issueSheet.getLastRow()+1,1,1,record.length).setValues([record]);

  // The same mobile issue is mirrored into a separate print/reconciliation
  // queue.  This keeps the existing issue report intact while allowing the
  // customer checklist and the later Excel print refresh to use one source.
  pickQueueUpsertPrintRecords_(ss,[{
    printKey:pickMobilePrintKey_(orderId,code),batch:String(first[1]||""),customer:String(first[2]||""),code:code,
    source:"موبایل",requestedQty:requestedQty,actualQty:actualQty,shortageQty:Math.max(0,requestedQty-actualQty),
    issueType:isIssue ? itemStatus : "رفع شد",issueNote:issueNote,issueState:isIssue ? "نیازمند پیگیری" : "رفع شد"
  }]);
}

function pickQueueList_(p) {
  const data = pickSheet_().getDataRange().getValues();
  const map = {};
  for (let i=1;i<data.length;i++) {
    const r=data[i], id=String(r[0]||"");
    if (!id) continue;
    if (!map[id]) map[id]={
      orderId:id,batch:String(r[1]||""),customer:String(r[2]||""),
      worker:String(r[8]||""),orderStatus:String(r[9]||"در انتظار برداشت"),
      note:String(r[10]||""),createdAt:pickDateText_(r[11]),
      items:[]
    };
    // Keep one mobile card per product code.  The queue itself still retains
    // its physical-allocation rows, so no stock/history data is changed.
    const code=pickCodeText_(r[4]);
    const itemKey=String(code||"").trim();
    if (!map[id]._itemMap) map[id]._itemMap={};
    if (!map[id]._itemMap[itemKey]) {
      map[id]._itemMap[itemKey]={
        lineId:"GROUP:"+itemKey, code:code, qty:0, locations:[], statuses:[], actualQty:0, hasActual:false, issueNotes:[]
      };
    }
    const item=map[id]._itemMap[itemKey];
    item.qty+=Number(r[5]||0);
    const loc=String(r[6]||"").trim();
    if (loc && item.locations.indexOf(loc)<0) item.locations.push(loc);
    item.statuses.push(String(r[7]||"در انتظار"));
    if (r[13] !== "" && r[13] != null) { item.actualQty += Number(r[13]||0); item.hasActual=true; }
    const issueNote=String(r[14]||"").trim();
    if (issueNote && item.issueNotes.indexOf(issueNote)<0) item.issueNotes.push(issueNote);
  }
  Object.keys(map).forEach(id=>{
    const order=map[id];
    order.items=Object.keys(order._itemMap).map(key=>{
      const item=order._itemMap[key];
      const uniqueStatuses=[...new Set(item.statuses)];
      return {
        lineId:item.lineId,
        code:item.code,
        qty:item.qty,
        location:item.locations.join(" | "),
        status:uniqueStatuses.length===1 ? uniqueStatuses[0] : "در انتظار",
        actualQty:item.actualQty,
        hasActual:item.hasActual,
        issueNote:item.issueNotes.join(" | ")
      };
    });
    delete order._itemMap;
  });
  let orders=Object.values(map);
  if (String(p.includeDone||"") !== "1") orders=orders.filter(x=>x.orderStatus!=="تحویل بسته‌بندی");
  // Batch is an ASCII timestamp from Excel.  It keeps order deterministic
  // even though the visible date columns use Persian calendar text.
  orders.sort((a,b)=>String(b.batch).localeCompare(String(a.batch)) || String(b.orderId).localeCompare(String(a.orderId)));
  return {ok:true,orders};
}

function pickQueuePackageEligibility_(values,matches) {
  const byCode={};
  matches.forEach(i=>{
    const row=values[i],code=pickCodeText_(row[4]);
    if (!byCode[code]) byCode[code]={requestedQty:0,actualQty:0,hasActual:false,statuses:[]};
    const item=byCode[code];
    item.requestedQty+=pickNumber_(row[5]);
    if (row[13]!=="" && row[13]!=null) { item.actualQty+=pickNumber_(row[13]); item.hasActual=true; }
    item.statuses.push(String(row[7]||"در انتظار"));
  });
  let unresolved=0,totalActual=0;
  Object.keys(byCode).forEach(code=>{
    const item=byCode[code],statuses=[...new Set(item.statuses)];
    const status=statuses.length===1 ? statuses[0] : "در انتظار";
    if (status!=="برداشت شد" && status!=="ناقص" && status!=="مغایرت") { unresolved++; return; }
    // Legacy picked rows may not carry ActualQty.  A picked status then means
    // its requested quantity was physically found; issue rows without an
    // actual quantity deliberately count as zero.
    totalActual+=status==="برداشت شد" && !item.hasActual ? item.requestedQty : item.actualQty;
  });
  return {unresolved:unresolved,totalActual:totalActual};
}

function pickQueueSet_(p) {
  const orderId=String(p.orderId||""), lineId=String(p.lineId||"");
  if (!orderId) throw new Error("orderId required");

  const allowedItem=new Set(["در انتظار","برداشت شد","ناقص","مغایرت"]);
  const allowedOrder=new Set(["در انتظار برداشت","در حال برداشت","آماده بسته‌بندی","تحویل بسته‌بندی"]);
  // Keep every mobile click short: one read and one atomic write instead of
  // many per-cell writes.  The same lock is used by Excel order uploads.
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error("queue busy; retry");
  try {
    const sh=pickSheet_();
    const lastRow=sh.getLastRow();
    if (lastRow < 2) throw new Error("order not found");

    const range=sh.getRange(2,1,lastRow-1,PICK_QUEUE_HEADERS.length);
    const values=range.getValues();
    const matches=[];
    const groupedCode=lineId.indexOf("GROUP:")===0 ? lineId.substring(6) : "";
    const groupedLineIndexes=[];
    let lineFound=!lineId;

    for (let i=0;i<values.length;i++) {
      const r=values[i];
      if (String(r[0]||"")!==orderId) continue;
      matches.push(i);
      if (groupedCode) {
        if (pickCodeText_(r[4])===groupedCode) {
          lineFound=true;
          groupedLineIndexes.push(i);
        }
      } else if (lineId && String(r[3]||"")===lineId) {
        lineFound=true;
        groupedLineIndexes.push(i);
      }
    }
    if (!matches.length) throw new Error("order not found");
    if (!lineFound) throw new Error("item not found");

    const now=pickJalaliDateTime_(new Date());
    let changedIndexes=matches;
    if (p.worker !== undefined) {
      const v=String(p.worker||"");
      matches.forEach(i=>{ values[i][8]=v; values[i][12]=now; });
    }
    if (p.orderStatus !== undefined) {
      const v=String(p.orderStatus||"");
      if (!allowedOrder.has(v)) throw new Error("invalid order status");
      if (v==="آماده بسته‌بندی") {
        const readiness=pickQueuePackageEligibility_(values,matches);
        if (readiness.unresolved>0) throw new Error("ابتدا وضعیت همه ردیف‌ها را مشخص کنید.");
      }
      if (v==="تحویل بسته‌بندی") {
        const ready=matches.every(i=>String(values[i][9]||"") === "آماده بسته‌بندی");
        if (!ready) throw new Error("ابتدا سفارش را آماده بسته‌بندی کنید.");
      }
      matches.forEach(i=>{ values[i][9]=v; values[i][12]=now; });
    }
    if (p.note !== undefined) {
      const v=String(p.note||"");
      matches.forEach(i=>{ values[i][10]=v; values[i][12]=now; });
    }
    if (lineId && p.itemStatus !== undefined) {
      const v=String(p.itemStatus||"");
      if (!allowedItem.has(v)) throw new Error("invalid item status");
      if (!groupedLineIndexes.length) throw new Error("item not found");
      let actualQty;
      if (p.actualQty !== undefined && String(p.actualQty).trim() !== "") {
        actualQty=Number(p.actualQty);
        const requestedQty=groupedLineIndexes.reduce((sum,i)=>sum+Number(values[i][5]||0),0);
        if (!Number.isFinite(actualQty) || actualQty<0 || actualQty>requestedQty) throw new Error("invalid actual quantity");
      }
      const issueNote=p.issueNote === undefined ? undefined : String(p.issueNote||"").trim();
      const actionId=p.clientActionId === undefined ? undefined : String(p.clientActionId||"").trim();
      groupedLineIndexes.forEach((i,index)=>{
        values[i][7]=v;
        // A grouped mobile card may contain multiple physical allocation rows.
        // Save the total actual quantity/note once, then list_ aggregates it.
        if (actualQty !== undefined) values[i][13]=index===0 ? actualQty : "";
        if (issueNote !== undefined) values[i][14]=index===0 ? issueNote : "";
        if (actionId !== undefined) values[i][15]=actionId;
        values[i][12]=now;
      });
      changedIndexes=groupedLineIndexes;
    }

    range.setValues(values);
    if (lineId && p.itemStatus !== undefined) {
      pickQueueSyncIssueReport_(sh,values,orderId,groupedLineIndexes,String(p.itemStatus||""),now);
    }
    return {ok:true,changed:changedIndexes.length,clientActionId:String(p.clientActionId||"")};
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) { return pickDoGet(e); }
function doPost(e) { return pickDoPost(e); }

// Independent queue only. Never reads or writes InventoryBankTable.
