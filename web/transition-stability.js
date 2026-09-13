// Isolated UI stability and recovery helpers.
// Goals:
// 1) Prevent stale list responses from repainting an older stage.
// 2) Preserve opened transfer-history panels across automatic renders.
// 3) Treat known non-critical Google Sheets typed-column formatting failures as uncertain responses,
//    then verify the real server state before deciding success/failure.
// 4) Never suppress an error unless the requested operation is independently confirmed.

let nebrasTransitionRevision=0;
const nebrasOriginalCall=call;

function nebrasIsTypedColumnFormatError(err){
  const m=String(err?.message||err||"").toLowerCase();
  return m.includes("typed column")&&m.includes("number format");
}
function nebrasDelay(ms){return new Promise(ok=>setTimeout(ok,ms))}

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

function nebrasApplyListSnapshot(d){
  if(!d||!d.ok)return false;
  me=d.user;
  all=(d.jobs||[]).map(normalizeJob);
  tailorUsers=d.tailors||[];
  savePageCache();
  showApp();
  return true;
}

async function nebrasReadFreshListDirect(){
  try{return await nebrasOriginalCall("list")}catch(_){return null}
}

async function nebrasVerifyStage(id,expectedStage){
  for(const ms of [350,800,1600]){
    await nebrasDelay(ms);
    const d=await nebrasReadFreshListDirect();
    if(!d?.ok)continue;
    const j=(d.jobs||[]).find(x=>String(x.id)===String(id));
    if(j&&stageKey(j.stage)===expectedStage){
      nebrasApplyListSnapshot(d);
      return true;
    }
  }
  return false;
}

// Stable transfer implementation. We deliberately do not use the original advance handler here,
// because that handler alerts immediately on every backend error. This version verifies reality first.
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

  try{
    await nebrasOriginalCall("advance",{id,route,tailor});
    syncSoon();
  }catch(e){
    if(e?.code==="NO_RESPONSE"||nebrasIsTypedColumnFormatError(e)){
      const confirmed=await nebrasVerifyStage(id,to);
      if(!confirmed){
        alert(e?.code==="NO_RESPONSE"?"پاسخ سامانه دریافت نشد و انتقال نیز تأیید نشد. لطفاً دوباره وضعیت سفارش را بررسی کنید.":(e?.message||"انتقال سفارش تأیید نشد"));
        await load(true);
      }
    }else{
      alert(e?.message||"انتقال سفارش انجام نشد");
      await load(true);
    }
  }finally{
    busyIds.delete(key);
    render();
  }
};

// Keep the user's open/closed history choice stable when the live refresh rerenders cards.
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
  document.querySelectorAll("details.transferHistory[data-job-id]").forEach(details=>{
    const key=details.dataset.jobId;
    if(!key)return;
    if(details.open)nebrasOpenHistoryIds.add(key);
    else nebrasOpenHistoryIds.delete(key);
  });
  nebrasOriginalRender();
  nebrasBindHistoryState();
};
nebrasBindHistoryState();

// Manual-order safeguards.
// Hide the system-generated registration note while preserving the real new->plotter transfer event.
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
function nebrasCountMatchingOrders(signature,jobs=all){
  try{return (jobs||[]).filter(j=>nebrasOrderSignature(j.code,j.quantity,j.customer)===signature).length}catch(_){return 0}
}
async function nebrasVerifyManualOrder(signature,beforeCount){
  for(const ms of [350,800,1600,2600]){
    await nebrasDelay(ms);
    const d=await nebrasReadFreshListDirect();
    if(!d?.ok)continue;
    const count=nebrasCountMatchingOrders(signature,d.jobs||[]);
    if(count>beforeCount){
      nebrasApplyListSnapshot(d);
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
        if(created){
          document.querySelector("#jobDialog").close();
          e.target.reset();
        }else{
          alert(x?.code==="NO_RESPONSE"?"پاسخ قطعی سامانه دریافت نشد و ثبت سفارش هم تأیید نشد. لطفاً وضعیت اینترنت را بررسی و سپس دوباره تلاش کنید.":(x?.message||"ثبت سفارش تأیید نشد"));
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
