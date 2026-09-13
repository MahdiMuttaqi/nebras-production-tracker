// Isolated UI stability helpers.
// Keeps transfers instant, blocks stale list repaint, preserves open histories,
// and hides only system-generated registration notes.

let nebrasTransitionRevision=0;
const nebrasOriginalCall=call;
const nebrasOriginalAdvance=advance;

function nebrasIsTypedColumnFormatError(err){
  const m=String(err?.message||err||"").toLowerCase();
  return m.includes("typed column")&&m.includes("number format");
}
function nebrasDelay(ms){return new Promise(ok=>setTimeout(ok,ms))}

// Reject only list responses that started before a transfer and returned afterwards.
call=async function(action,data={}){
  if(action!=="list")return nebrasOriginalCall(action,data);
  const revisionAtStart=nebrasTransitionRevision;
  const result=await nebrasOriginalCall(action,data);
  if(revisionAtStart!==nebrasTransitionRevision){
    const err=new Error("stale_list_response");
    err.code="STALE_LIST";
    throw err;
  }
  return result;
};

// Fast optimistic transfer. Move the card on screen immediately; network confirmation continues in background.
advance=async function(id,route="",tailor=""){
  const key=String(id);
  if(busyIds.has(key))return;
  const j=all.find(x=>String(x.id)===key);
  if(!j)return;

  const from=j.stage;
  const to=localNext(j,route,tailor);
  const stamp=faNow();
  const event={id:Date.now(),fromStage:from,toStage:to,note:"",createdAt:stamp};

  nebrasTransitionRevision++;
  busyIds.add(key);
  localMove(j,to,event);

  // Critical: paint the optimistic move NOW instead of waiting for Apps Script.
  render();

  try{
    await nebrasOriginalCall("advance",{id,route,tailor});
    syncSoon();
  }catch(e){
    if(e?.code==="NO_RESPONSE"||nebrasIsTypedColumnFormatError(e)){
      // Keep the already-applied optimistic move and reconcile in background.
      syncSoon();
    }else{
      alert(e?.message||"انتقال سفارش انجام نشد");
      await load(true);
    }
  }finally{
    busyIds.delete(key);
    render();
  }
};

// Preserve open/closed transfer-history panels across live refresh renders.
const nebrasOpenHistoryIds=new Set();
function nebrasVisibleJobs(){
  const q=$("#search")?.value.trim().toLowerCase()||"";
  const active=all.filter(j=>j.stage!=="completed");
  const archive=all.filter(j=>j.stage==="completed");
  const source=me?.role==="admin"?(orderView==="archive"?archive:active):all;
  return source.filter(j=>(j.code+" "+j.customer+" "+roleLabel(j.stage)+" "+(j.events||[]).map(e=>e.note||"").join(" ")).toLowerCase().includes(q));
}
function nebrasBindHistoryState(){
  const jobs=nebrasVisibleJobs();
  document.querySelectorAll("#jobs article.job").forEach((article,index)=>{
    const details=article.querySelector("details.transferHistory");
    const job=jobs[index];
    if(!details||!job)return;
    const key=String(job.id);
    details.dataset.jobId=key;
    details.open=nebrasOpenHistoryIds.has(key);
    details.addEventListener("toggle",()=>{
      if(details.open)nebrasOpenHistoryIds.add(key);
      else nebrasOpenHistoryIds.delete(key);
    });
  });
}

function nebrasSanitizeSystemNotes(){
  try{
    (all||[]).forEach(job=>{
      (job.events||[]).forEach(ev=>{
        const note=String(ev.note||"").trim();
        if(note==="ثبت سفارش"||/^ارسال گروهی از اکسل(?:\s|$)/.test(note))ev.note="";
      });
    });
  }catch(_){}
}

const nebrasBaseRender=render;
render=function(){
  document.querySelectorAll("details.transferHistory[data-job-id]").forEach(details=>{
    const key=details.dataset.jobId;
    if(!key)return;
    if(details.open)nebrasOpenHistoryIds.add(key);
    else nebrasOpenHistoryIds.delete(key);
  });
  nebrasSanitizeSystemNotes();
  nebrasBaseRender();
  nebrasBindHistoryState();
};

// Manual-order submit guard only. This does not affect transfer-button speed.
function nebrasOrderSignature(code,quantity,customer){
  return String(code||"").trim()+"|"+String(+quantity||0)+"|"+String(customer||"").trim();
}
function nebrasCountMatchingOrders(signature,jobs=all){
  try{return (jobs||[]).filter(j=>nebrasOrderSignature(j.code,j.quantity,j.customer)===signature).length}catch(_){return 0}
}
async function nebrasReadFreshListDirect(){try{return await nebrasOriginalCall("list")}catch(_){return null}}
async function nebrasVerifyManualOrder(signature,beforeCount){
  for(const ms of [400,900,1600]){
    await nebrasDelay(ms);
    const d=await nebrasReadFreshListDirect();
    if(!d?.ok)continue;
    if(nebrasCountMatchingOrders(signature,d.jobs||[])>beforeCount){
      me=d.user;all=(d.jobs||[]).map(normalizeJob);tailorUsers=d.tailors||[];savePageCache();showApp();
      return true;
    }
  }
  return false;
}

const nebrasJobForm=document.querySelector("#jobForm");
if(nebrasJobForm){
  nebrasJobForm.addEventListener("submit",async function(e){
    e.preventDefault();
    e.stopImmediatePropagation();
    const b=document.querySelector("#saveJob");
    if(!b||b.disabled)return;
    const code=document.querySelector("#code").value.trim();
    const quantity=+document.querySelector("#quantity").value;
    const customer=document.querySelector("#customer").value.trim();
    const signature=nebrasOrderSignature(code,quantity,customer);
    const beforeCount=nebrasCountMatchingOrders(signature);
    b.disabled=true;
    try{
      await nebrasOriginalCall("add",{code,quantity,customer});
      document.querySelector("#jobDialog").close();
      e.target.reset();
      await load(true);
    }catch(x){
      if(x?.code==="NO_RESPONSE"||nebrasIsTypedColumnFormatError(x)){
        const created=await nebrasVerifyManualOrder(signature,beforeCount);
        if(created){document.querySelector("#jobDialog").close();e.target.reset()}
        else alert("ثبت سفارش تأیید نشد. لطفاً وضعیت اینترنت را بررسی و دوباره تلاش کنید.");
      }else alert(x?.message||"ثبت سفارش انجام نشد");
    }finally{b.disabled=false}
  },true);
}

nebrasSanitizeSystemNotes();
nebrasBindHistoryState();
