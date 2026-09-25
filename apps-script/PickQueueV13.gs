const PICK_QUEUE_KEY = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903";
const PICK_QUEUE_FILE_PROP = "NEBRAS_PICK_QUEUE_SHEET_ID";
const PICK_QUEUE_SHEET = "PickQueue";
const PICK_QUEUE_HEADERS = [
  "OrderId","Batch","Customer","LineId","Code","Qty","Location",
  "ItemStatus","Worker","OrderStatus","Note","CreatedAt","UpdatedAt"
];

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
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,PICK_QUEUE_HEADERS.length).setValues([PICK_QUEUE_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function pickRows_() {
  const sh = pickSheet_();
  const values = sh.getDataRange().getValues();
  return {sh, values};
}

function pickDoGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const cb = p.callback || "";
    if (!pickAuthorized_(p.key)) return pickOut_({ok:false,error:"unauthorized"}, cb);
    if (p.action === "pickQueueList") return pickOut_(pickQueueList_(p), cb);
    if (p.action === "pickQueueSet") return pickOut_(pickQueueSet_(p), cb);
    return pickOut_({ok:true,name:"سامانه پیگیری تولید نبراس",service:"pick-queue-v13"}, cb);
  } catch (err) {
    return pickOut_({ok:false,error:String(err && err.message || err)}, e && e.parameter && e.parameter.callback);
  }
}

function pickDoPost(e) {
  try {
    const type = String((e && e.postData && e.postData.type) || "");
    let p = (e && e.parameter) || {};
    if (type.indexOf("application/json") >= 0) {
      const q = JSON.parse(e.postData.contents || "{}");
      p = Object.assign({}, p, q);
    }
    if (!pickAuthorized_(p.key)) return pickOut_({ok:false,error:"unauthorized"});
    if (p.action === "pickQueuePut") return pickOut_(pickQueuePut_(p));
    if (p.action === "pickQueueSet") return pickOut_(pickQueueSet_(p));
    return pickOut_({ok:false,error:"invalid action"});
  } catch (err) {
    return pickOut_({ok:false,error:String(err && err.message || err)});
  }
}

function pickQueuePut_(p) {
  const payload = typeof p.payload === "string" ? JSON.parse(p.payload) : p.payload;
  if (!payload || !payload.orderId || !Array.isArray(payload.items)) throw new Error("invalid payload");
  const sh = pickSheet_();
  const data = sh.getDataRange().getValues();
  const keep = [PICK_QUEUE_HEADERS];
  for (let i=1;i<data.length;i++) if (String(data[i][0]) !== String(payload.orderId)) keep.push(data[i]);
  const now = new Date();
  payload.items.forEach(it => keep.push([
    String(payload.orderId), String(payload.batch || ""), String(payload.customer || ""),
    String(it.lineId || ""), String(it.code || ""), Number(it.qty || 0), String(it.location || ""),
    "در انتظار", "", "در انتظار برداشت", "", now, now
  ]));
  sh.clearContents();
  sh.getRange(1,1,keep.length,PICK_QUEUE_HEADERS.length).setValues(keep);
  sh.setFrozenRows(1);
  return {ok:true,orderId:String(payload.orderId),items:payload.items.length};
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
      note:String(r[10]||""),createdAt:r[11] instanceof Date?r[11].toISOString():String(r[11]||""),
      items:[]
    };
    map[id].items.push({
      lineId:String(r[3]||""),code:String(r[4]||""),qty:Number(r[5]||0),
      location:String(r[6]||""),status:String(r[7]||"در انتظار")
    });
  }
  let orders=Object.values(map);
  if (String(p.includeDone||"") !== "1") orders=orders.filter(x=>x.orderStatus!=="تحویل بسته‌بندی");
  orders.sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
  return {ok:true,orders};
}

function pickQueueSet_(p) {
  const orderId=String(p.orderId||""), lineId=String(p.lineId||"");
  if (!orderId) throw new Error("orderId required");
  const allowedItem=new Set(["در انتظار","برداشت شد","ناقص","مغایرت"]);
  const allowedOrder=new Set(["در انتظار برداشت","در حال برداشت","آماده بسته‌بندی","تحویل بسته‌بندی"]);
  const {sh,values}=pickRows_();
  let changed=0;
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0])!==orderId) continue;
    if (p.worker !== undefined) values[i][8]=String(p.worker||"");
    if (p.orderStatus !== undefined) {
      const s=String(p.orderStatus||"");
      if (!allowedOrder.has(s)) throw new Error("invalid order status");
      values[i][9]=s;
    }
    if (p.note !== undefined) values[i][10]=String(p.note||"");
    if (lineId && String(values[i][3])===lineId && p.itemStatus !== undefined) {
      const s=String(p.itemStatus||"");
      if (!allowedItem.has(s)) throw new Error("invalid item status");
      values[i][7]=s;
    }
    values[i][12]=new Date();
    changed++;
  }
  if (!changed) throw new Error("order not found");
  sh.getRange(1,1,values.length,PICK_QUEUE_HEADERS.length).setValues(values);
  return {ok:true,changed};
}

/*
  Integration:
  In the standalone Apps Script project, route your existing doGet/doPost to these helpers:

  function doGet(e){ return pickDoGet(e); }
  function doPost(e){ return pickDoPost(e); }

  This queue is independent. It never reads or writes InventoryBankTable.
*/
