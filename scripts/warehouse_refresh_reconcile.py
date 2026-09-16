from pathlib import Path

p = Path('web/warehouse-pilot/index.html')
s = p.read_text(encoding='utf-8')

def one(old, new, label):
    global s
    n = s.count(old)
    if n != 1:
        raise SystemExit(f'{label}: expected 1 match, found {n}')
    s = s.replace(old, new, 1)

marker = "function syncSummary(){const unsent=db.history.filter(h=>h.sync!=='synced').length,waiting=db.history.filter(h=>h.sync==='synced'&&h.excelStatus!=='اعمال شد').length;return{unsent,waiting}}"
helpers = """function excelAppliedStatus(v){return String(v||'').trim()==='اعمال شد'}
function excelRejectedStatus(v){const x=String(v||'');return x.includes('اعمال نشد')||x.includes('نیازمند بررسی')}
function operationNeedsReplay(h){if(h.sync!=='synced')return true;return !excelAppliedStatus(h.excelStatus)&&!excelRejectedStatus(h.excelStatus)}
async function reconcileStatusesBeforeSnapshot(){
 if(!db||!Array.isArray(db.history)||!navigator.onLine)return;
 const candidates=db.history.filter(h=>h.sync==='synced'&&!excelAppliedStatus(h.excelStatus)).slice(0,60);
 for(let i=0;i<candidates.length;i+=6){
  const batch=candidates.slice(i,i+6);
  await Promise.allSettled(batch.map(async h=>{
   const status=await jsonpStatus(h.id);
   if(status?.found){
    h.excelStatus=status.excelStatus||h.excelStatus||'در انتظار';
    if(excelAppliedStatus(h.excelStatus)&&!h.confirmedAt)h.confirmedAt=new Date().toISOString();
   }
  }));
 }
}
function syncSummary(){const unsent=db.history.filter(h=>h.sync!=='synced').length,waiting=db.history.filter(h=>h.sync==='synced'&&!excelAppliedStatus(h.excelStatus)&&!excelRejectedStatus(h.excelStatus)).length,review=db.history.filter(h=>h.sync==='synced'&&excelRejectedStatus(h.excelStatus)).length;return{unsent,waiting,review}}"""
one(marker, helpers, 'status helpers and summary')

one(
    "function showSyncSummary(prefix=''){const s=syncSummary();if(s.unsent)syncState(`${prefix}${fa(s.unsent)} عملیات در صف ارسال`,false);else if(s.waiting)syncState(`${prefix}${fa(s.waiting)} عملیات رسیده؛ منتظر اکسل`,true);else syncState(prefix?'به‌روز از اکسل':'همگام',true)}",
    "function showSyncSummary(prefix=''){const s=syncSummary();if(s.unsent)syncState(`${prefix}${fa(s.unsent)} عملیات در صف ارسال`,false);else if(s.waiting)syncState(`${prefix}${fa(s.waiting)} عملیات رسیده؛ منتظر اکسل`,true);else if(s.review)syncState(`${prefix}${fa(s.review)} عملیات نیازمند بررسی`,false);else syncState(prefix?'به‌روز از اکسل':'همگام',true)}",
    'summary display'
)

one(
    "async function refreshFromExcel(){const btn=document.querySelector('#refreshExcel');if(!navigator.onLine){toast('اینترنت در دسترس نیست.');return}btn.disabled=true;btn.textContent='در حال دریافت…';syncState('دریافت مستقیم از اکسل…',false);try{const snapshot=await jsonpSnapshot();",
    "async function refreshFromExcel(){const btn=document.querySelector('#refreshExcel');if(!navigator.onLine){toast('اینترنت در دسترس نیست.');return}btn.disabled=true;btn.textContent='در حال دریافت…';syncState('بررسی وضعیت عملیات و دریافت اکسل…',false);try{await reconcileStatusesBeforeSnapshot();const snapshot=await jsonpSnapshot();",
    'reconcile before snapshot'
)

one(
    "const localHistory=db.history.slice(),pending=localHistory.filter(h=>h.sync!=='synced'||h.excelStatus!=='اعمال شد');",
    "const localHistory=db.history.slice(),pending=localHistory.filter(operationNeedsReplay);",
    'only true pending operations replay over Excel snapshot'
)

one(
    "function operationStatus(h){if(h.sync!=='synced')return'در حال ارسال';if(h.excelStatus==='اعمال شد')return'اعمال شد ✓';return'منتظر اکسل'}",
    "function operationStatus(h){if(h.sync!=='synced')return'در حال ارسال';if(excelAppliedStatus(h.excelStatus))return'اعمال شد ✓';if(excelRejectedStatus(h.excelStatus))return'نیازمند بررسی ⚠';return'منتظر اکسل'}",
    'history status display'
)

p.write_text(s, encoding='utf-8')
