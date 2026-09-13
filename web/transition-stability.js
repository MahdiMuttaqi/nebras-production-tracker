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
