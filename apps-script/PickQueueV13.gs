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
const PICK_QUEUE_TIME_ZONE = "Asia/Tehran";

function pickOut_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
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
