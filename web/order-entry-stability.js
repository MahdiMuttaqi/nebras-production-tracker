// Isolated safeguards for manual order entry.
// 1) Hide only the system note "ثبت سفارش" from cards/history notes.
// 2) If the POST response is uncertain but the order was actually created, verify via live list instead of showing a false error.

function nebrasSanitizeManualEntryNotes(){
  try{
    (all||[]).forEach(job=>{
      (job.events||[]).forEach(ev=>{
        if(String(ev.note||"").trim()==="ثبت سفارش") ev.note="";
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
  try{
    return (all||[]).filter(j=>nebrasOrderSignature(j.code,j.quantity,j.customer)===signature).length;
  }catch(_){return 0}
}
function nebrasDelay(ms){return new Promise(ok=>setTimeout(ok,ms))}
async function nebrasVerifyManualOrder(signature,beforeCount){
  const waits=[700,1300,2200];
  for(const ms of waits){
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

// Sanitize cards already rendered before this helper loaded.
nebrasSanitizeManualEntryNotes();
render();
