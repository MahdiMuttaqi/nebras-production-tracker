// Isolated transition stability guard. Keeps stale list responses from repainting an older stage.
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
