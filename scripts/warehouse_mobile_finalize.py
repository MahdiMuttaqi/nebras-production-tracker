from pathlib import Path

path = Path('web/warehouse-pilot/index.html')
s = path.read_text(encoding='utf-8')

def replace_once(old, new, label):
    global s
    count = s.count(old)
    if count != 1:
        raise SystemExit(f'{label}: expected exactly 1 match, found {count}')
    s = s.replace(old, new, 1)

# Explicit logout first, so the successful-login patch below cannot create a second ambiguous match.
replace_once(
    " sessionStorage.removeItem(SESSION_PASSWORD_KEY);\n clearTimeout(syncTimer);",
    " localStorage.removeItem(SESSION_PASSWORD_KEY);sessionStorage.removeItem(SESSION_PASSWORD_KEY);\n clearTimeout(syncTimer);",
    'explicit logout clears trusted login'
)
replace_once(
    "sessionStorage.setItem(SESSION_PASSWORD_KEY,password);",
    "localStorage.setItem(SESSION_PASSWORD_KEY,password);sessionStorage.removeItem(SESSION_PASSWORD_KEY);",
    'persist successful login'
)
replace_once(
    "const rememberedPassword=sessionStorage.getItem(SESSION_PASSWORD_KEY);\nif(rememberedPassword){document.querySelector('#managerPassword').value=rememberedPassword;setTimeout(unlock,0)}",
    "const rememberedPassword=localStorage.getItem(SESSION_PASSWORD_KEY)||sessionStorage.getItem(SESSION_PASSWORD_KEY);\nif(rememberedPassword){localStorage.setItem(SESSION_PASSWORD_KEY,rememberedPassword);sessionStorage.removeItem(SESSION_PASSWORD_KEY);document.querySelector('#managerPassword').value=rememberedPassword;setTimeout(unlock,0)}",
    'restore trusted login'
)

replace_once(
    "function operationStatus(h){if(h.sync!=='synced')return'در صف ارسال به سرور';if(h.excelStatus==='اعمال شد')return'اعمال‌شده در اکسل';return'رسیده به سرور؛ منتظر اکسل'}",
    "function operationStatus(h){if(h.sync!=='synced')return'در حال ارسال';if(h.excelStatus==='اعمال شد')return'اعمال شد ✓';return'منتظر اکسل'}",
    'three clear operation states'
)

marker = "function syncSummary(){const unsent=db.history.filter(h=>h.sync!=='synced').length,waiting=db.history.filter(h=>h.sync==='synced'&&h.excelStatus!=='اعمال شد').length;return{unsent,waiting}}"
cleanup = """const CONFIRMED_HISTORY_MAX_AGE_MS=7*24*60*60*1000;
const CONFIRMED_HISTORY_KEEP=100;
function cleanupConfirmedHistory(){
 if(!db||!Array.isArray(db.history))return false;
 const nowMs=Date.now();let confirmedSeen=0,changed=false;const next=[];
 for(const h of db.history){
  const applied=h.sync==='synced'&&h.excelStatus==='اعمال شد';
  if(!applied){next.push(h);continue}
  confirmedSeen++;
  const base=Date.parse(h.confirmedAt||h.syncedAt||h.at||'');
  const tooOld=Number.isFinite(base)&&nowMs-base>CONFIRMED_HISTORY_MAX_AGE_MS;
  const tooMany=confirmedSeen>CONFIRMED_HISTORY_KEEP;
  if(tooOld||tooMany){changed=true;continue}
  next.push(h)
 }
 if(changed)db.history=next;
 return changed
}
""" + marker
replace_once(marker, cleanup, 'confirmed history cleanup')

replace_once(
    "function save(){render();encryptJson(db,cryptoKey).then(x=>{localStorage.setItem(SECURE_KEY,JSON.stringify(x));scheduleSync(150)}).catch(()=>toast('ذخیره امن انجام نشد.'))}",
    "function save(){cleanupConfirmedHistory();render();encryptJson(db,cryptoKey).then(x=>{localStorage.setItem(SECURE_KEY,JSON.stringify(x));scheduleSync(150)}).catch(()=>toast('ذخیره امن انجام نشد.'))}",
    'cleanup before safe save'
)

replace_once(
    "const next=status.excelStatus||'در انتظار';\n    if(h.excelStatus!==next){h.excelStatus=next;changed=true}",
    "const next=status.excelStatus||'در انتظار';\n    if(h.excelStatus!==next){h.excelStatus=next;changed=true}\n    if(next==='اعمال شد'&&!h.confirmedAt){h.confirmedAt=new Date().toISOString();changed=true}",
    'stamp confirmed status polling'
)
replace_once(
    "if(changed)await encryptJson(db,cryptoKey).then(x=>localStorage.setItem(SECURE_KEY,JSON.stringify(x))).catch(()=>{});",
    "if(cleanupConfirmedHistory())changed=true;\n  if(changed)await encryptJson(db,cryptoKey).then(x=>localStorage.setItem(SECURE_KEY,JSON.stringify(x))).catch(()=>{});",
    'cleanup after central confirmation'
)
replace_once(
    "h.sync='synced';h.excelStatus=status.excelStatus||'در انتظار';h.syncedAt=new Date().toISOString();return true",
    "h.sync='synced';h.excelStatus=status.excelStatus||'در انتظار';h.syncedAt=new Date().toISOString();if(h.excelStatus==='اعمال شد'&&!h.confirmedAt)h.confirmedAt=new Date().toISOString();return true",
    'stamp confirmed immediate push status'
)

path.write_text(s, encoding='utf-8')
