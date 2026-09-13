// Isolated UI stability helpers. Keeps stale list responses from repainting an older stage
// and preserves opened transfer-history panels across automatic renders.
let nebrasTransitionRevision=0;
const nebrasOriginalCall=call;
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
const nebrasOriginalAdvance=advance;
advance=async function(id,route="",tailor=""){
  nebrasTransitionRevision++;
  return nebrasOriginalAdvance(id,route,tailor);
};

// Keep the user's open/closed history choice stable when the 10-second live refresh rerenders cards.
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
const nebrasOriginalRender=render;
render=function(){
  // Capture the current state immediately before DOM replacement as an extra safeguard.
  document.querySelectorAll("details.transferHistory[data-job-id]").forEach(details=>{
    const key=details.dataset.jobId;
    if(!key)return;
    if(details.open)nebrasOpenHistoryIds.add(key);
    else nebrasOpenHistoryIds.delete(key);
  });
  nebrasOriginalRender();
  nebrasBindHistoryState();
};
// Bind cards that may have been rendered from cache before this helper loaded.
nebrasBindHistoryState();

// Manual-order safeguards only. These do not change transfer, tailoring, archive or user-management logic.
// Hide the system-generated registration note while keeping the actual new->plotter transfer event.
function nebrasSanitizeManualEntryNotes(){
  try{
    (all||[]).forEach(job=>{
      (job.events||[]).forEach(ev=>{
        if(String(ev.note||"").trim()==="ثبت سفارش")ev.note="";
      });
    });
  }catch(_){}
}
const nebrasEntryOriginalRender=render;
render=function(){
  nebrasSanitizeManualEntryNotes();
  return nebrasEntryOriginalRender();
};

function nebrasOrderSignature(code,quantity,customer){
  return String(code||"").trim()+"|"+String(+quantity||0)+"|"+String(customer||"").trim();
}
function nebrasCountMatchingOrders(signature){
  try{return (all||[]).filter(j=>nebrasOrderSignature(j.code,j.quantity,j.customer)===signature).length}catch(_){return 0}
}
function nebrasDelay(ms){return new Promise(ok=>setTimeout(ok,ms))}
async function nebrasVerifyManualOrder(signature,beforeCount){
  for(const ms of [700,1300,2200]){
    await nebrasDelay(ms);
    await load(true);
    if(nebrasCountMatchingOrders(signature)>beforeCount)return true;
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
      await call("add",{code,quantity,customer});
      document.querySelector("#jobDialog").close();
      e.target.reset();
      await load(true);
    }catch(x){
      if(x&&x.code==="NO_RESPONSE"){
        const created=await nebrasVerifyManualOrder(signature,beforeCount);
        if(created){
          document.querySelector("#jobDialog").close();
          e.target.reset();
        }else{
          alert("پاسخ قطعی سامانه دریافت نشد و ثبت سفارش هم تأیید نشد. لطفاً وضعیت اینترنت را بررسی و سپس دوباره تلاش کنید.");
        }
      }else{
        alert(x?.message||"ثبت سفارش انجام نشد");
      }
    }finally{
      b.disabled=false;
    }
  },true);
}

nebrasSanitizeManualEntryNotes();
render();
