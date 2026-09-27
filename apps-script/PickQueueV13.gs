const PICK_QUEUE_KEY = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903";
const PICK_QUEUE_FILE_PROP = "NEBRAS_PICK_QUEUE_SHEET_ID";
const PICK_QUEUE_SHEET = "PickQueueV2";
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

function pickCodeText_(value) {
  if (value instanceof Date) {
    return String(value.getMonth() + 1).padStart(2, "0") + "-" + String(value.getFullYear());
  }
  return String(value == null ? "" : value);
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

  const first = sh.getRange(1,1,1,PICK_QUEUE_HEADERS.length).getValues()[0];
  const validHeader = PICK_QUEUE_HEADERS.every((h,i)=>String(first[i]||"")===h);
  if (!validHeader) {
    sh.clearContents();
    sh.getRange(1,1,1,PICK_QUEUE_HEADERS.length).setValues([PICK_QUEUE_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function pickDoGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const cb = p.callback || "";
    if (!pickAuthorized_(p.key)) return pickOut_({ok:false,error:"unauthorized"}, cb);
    if (p.action === "pickQueueList") return pickOut_(pickQueueList_(p), cb);
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

  const sh = pickSheet_();
  const data = sh.getDataRange().getValues();
  const replacing = new Set(cleanOrders.map(o=>String(o.orderId)));
  const keep = [PICK_QUEUE_HEADERS];

  for (let i=1;i<data.length;i++) {
    if (!replacing.has(String(data[i][0]||""))) keep.push(data[i]);
  }

  const now = new Date();
  let itemCount = 0;
  cleanOrders.forEach(order=>{
    order.items.forEach(it=>{
      itemCount++;
      keep.push([
        String(order.orderId),
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
        now
      ]);
    });
  });

  // IMPORTANT: no setNumberFormat calls here.
  // Google Sheets "typed columns" reject programmatic number-format changes.
  // We write plain values only, so the queue works whether or not another sheet uses tables.
  sh.clearContents();
  sh.getRange(1,1,keep.length,PICK_QUEUE_HEADERS.length).setValues(keep);
  sh.setFrozenRows(1);

  return {ok:true,orders:cleanOrders.length,items:itemCount};
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
      lineId:String(r[3]||""),code:pickCodeText_(r[4]),qty:Number(r[5]||0),
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
  const sh=pickSheet_();
  const lastRow=sh.getLastRow();
  if (lastRow < 2) throw new Error("order not found");

  const values=sh.getRange(2,1,lastRow-1,PICK_QUEUE_HEADERS.length).getValues();
  const matches=[];
  let lineFound=!lineId;

  for (let i=0;i<values.length;i++) {
    const r=values[i];
    if (String(r[0]||"")!==orderId) continue;
    matches.push(i+2);
    if (lineId && String(r[3]||"")===lineId) lineFound=true;
  }
  if (!matches.length) throw new Error("order not found");
  if (!lineFound) throw new Error("item not found");

  const now=new Date();

  if (p.worker !== undefined) {
    const v=String(p.worker||"");
    matches.forEach(row=>sh.getRange(row,9).setValue(v));
  }
  if (p.orderStatus !== undefined) {
    const v=String(p.orderStatus||"");
    if (!allowedOrder.has(v)) throw new Error("invalid order status");
    matches.forEach(row=>sh.getRange(row,10).setValue(v));
  }
  if (p.note !== undefined) {
    const v=String(p.note||"");
    matches.forEach(row=>sh.getRange(row,11).setValue(v));
  }
  if (lineId && p.itemStatus !== undefined) {
    const v=String(p.itemStatus||"");
    if (!allowedItem.has(v)) throw new Error("invalid item status");
    for (let i=0;i<values.length;i++) {
      const r=values[i];
      if (String(r[0]||"")===orderId && String(r[3]||"")===lineId) {
        sh.getRange(i+2,8).setValue(v);
        sh.getRange(i+2,13).setValue(now);
        break;
      }
    }
  }
  if (p.worker !== undefined || p.orderStatus !== undefined || p.note !== undefined) {
    matches.forEach(row=>sh.getRange(row,13).setValue(now));
  }

  return {ok:true,changed:matches.length};
}

function doGet(e) { return pickDoGet(e); }
function doPost(e) { return pickDoPost(e); }

// Independent queue only. Never reads or writes InventoryBankTable.
