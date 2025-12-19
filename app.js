/* =========================================================
   Utilities
========================================================= */
const $ = (id)=>document.getElementById(id);
const qsa = (sel)=>Array.from(document.querySelectorAll(sel));
const safeNum = (v)=> Number(v||0);
const clamp = (x,min,max)=>Math.min(max,Math.max(min,x));
const safeDiv = (a,b)=> (safeNum(b)===0? 0 : safeNum(a)/safeNum(b));
const fmt$ = (n)=> {
  n=safeNum(n);
  const sign = n<0 ? "-" : "";
  n=Math.abs(n);
  return sign + n.toLocaleString(undefined,{style:"currency",currency:"USD",maximumFractionDigits:0});
};
const fmtPct = (x,dp=2)=> (safeNum(x)).toFixed(dp) + "%";
const fmtX = (x,dp=2)=> (safeNum(x)).toFixed(dp) + "×";
const todayISO = ()=> new Date().toISOString().slice(0,10);
function escapeHtml(s){
  return String(s ?? "").replace(/[&<>"]/g, (c)=>({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;"
  }[c]));
}

function pmt(ratePerPeriod, nper, pv){
  const r=ratePerPeriod;
  if(nper<=0) return 0;
  if(Math.abs(r)<1e-9) return pv/nper;
  return (pv*r)/(1-Math.pow(1+r,-nper));
}
function annDebtService(loanAmt, noteRatePct, amortYears){
  const r=(noteRatePct/100)/12;
  const n=Math.max(1, Math.round(amortYears*12));
  return pmt(r,n,loanAmt)*12;
}
function annInterestOnly(loanAmt, noteRatePct){
  return loanAmt*(noteRatePct/100);
}
function tagFor(val, goodFn, warnFn){
  if(goodFn(val)) return {cls:"good", label:"Good"};
  if(warnFn(val)) return {cls:"warn", label:"Watch"};
  return {cls:"bad", label:"Weak"};
}
function kpiCard(title, value, tag, note=""){
  return `
    <div class="kpi">
      <div class="t">${escapeHtml(title)}</div>
      <div class="v">${value}</div>
      <div class="s">
        <span class="tag ${tag.cls}">
          <span class="dot ${tag.cls}"></span>${tag.label}${note?` · <span class="muted">${escapeHtml(note)}</span>`:""}
        </span>
      </div>
    </div>`;
}
function productLabel(p){
  const map={
    bridge:"Bridge (Transitional)",
    fixflip:"Fix‑and‑Flip",
    construction:"Ground‑Up Construction",
    commercial_bridge:"Commercial Bridge",
    land:"Land Acquisition",
    second_lien:"Second Lien / Mezz",
    dscr:"DSCR / Investment Property"
  };
  return map[p]||p;
}

/* =========================================================
   Product Defaults (you can tune)
========================================================= */
const productDefaults = {
  bridge:            { termMonths:12, ioMonths:12, noteRate:12.0, amortYears:30, origPoints:2.0, lienPos:"1st" },
  fixflip:           { termMonths:12, ioMonths:12, noteRate:12.5, amortYears:30, origPoints:2.0, lienPos:"1st" },
  construction:      { termMonths:18, ioMonths:18, noteRate:12.5, amortYears:30, origPoints:2.0, lienPos:"1st" },
  commercial_bridge: { termMonths:18, ioMonths:18, noteRate:11.5, amortYears:30, origPoints:1.5, lienPos:"1st" },
  land:              { termMonths:18, ioMonths:18, noteRate:12.0, amortYears:30, origPoints:2.0, lienPos:"1st" },
  second_lien:       { termMonths:12, ioMonths:12, noteRate:14.0, amortYears:30, origPoints:2.5, lienPos:"2nd" },
  dscr:              { termMonths:60, ioMonths:0,  noteRate:8.5,  amortYears:30, origPoints:1.0, lienPos:"1st" }
};
function applyDefaults(prod){
  const d = productDefaults[prod]; if(!d) return;
  $("termMonths").value = d.termMonths;
  $("ioMonths").value = d.ioMonths;
  $("noteRate").value = d.noteRate.toFixed(2);
  $("amortYears").value = d.amortYears;
  $("origPoints").value = d.origPoints.toFixed(2);
  $("lienPos").value = d.lienPos;
}

/* =========================================================
   Dynamic Tables (Rent Roll, Global, Draws)
========================================================= */
let rentRoll = [];
let globProps = [];
let draws = [];

function rrRow(i, r){
  return `<tr>
    <td><input data-rr="tenant" data-i="${i}" value="${escapeHtml(r.tenant||"")}" placeholder="Tenant/Unit"/></td>
    <td class="right"><input data-rr="rent" data-i="${i}" type="number" step="10" value="${safeNum(r.rent||0)}"/></td>
    <td class="right"><input data-rr="exp" data-i="${i}" value="${escapeHtml(r.exp||"")}" placeholder="MM/YY"/></td>
    <td><input data-rr="notes" data-i="${i}" value="${escapeHtml(r.notes||"")}" placeholder="Notes"/></td>
    <td class="right"><button class="btn" onclick="delRR(${i})">Del</button></td>
  </tr>`;
}
function renderRR(){
  const tb = $("rentRollTbl").querySelector("tbody");
  tb.innerHTML = rentRoll.map((r,i)=>rrRow(i,r)).join("") || `<tr><td colspan="5" class="small">No rows yet.</td></tr>`;
  qsa("[data-rr]").forEach(el=>{
    el.addEventListener("input", ()=>{
      const i=Number(el.dataset.i), k=el.dataset.rr;
      rentRoll[i][k] = (k==="rent")? safeNum(el.value): el.value;
      persistDraft();
    });
  });
}
function delRR(i){ rentRoll.splice(i,1); renderRR(); calc(); }
function addRR(){
  rentRoll.push({tenant:"",rent:0,exp:"",notes:""});
  renderRR();
}
function sumRR(){
  const totalMonthly = rentRoll.reduce((s,r)=>s+safeNum(r.rent),0);
  $("grossRent").value = Math.round(totalMonthly*12);
  toast("Gross Scheduled Rent updated from rent roll.");
  calc();
}

function globRow(i, r){
  return `<tr>
    <td><input data-g="name" data-i="${i}" value="${escapeHtml(r.name||"")}" placeholder="Property"/></td>
    <td class="right"><input data-g="noi" data-i="${i}" type="number" step="100" value="${safeNum(r.noi||0)}"/></td>
    <td class="right"><input data-g="ds" data-i="${i}" type="number" step="100" value="${safeNum(r.ds||0)}"/></td>
    <td><input data-g="notes" data-i="${i}" value="${escapeHtml(r.notes||"")}" placeholder="Notes"/></td>
    <td class="right"><button class="btn" onclick="delGlob(${i})">Del</button></td>
  </tr>`;
}
function renderGlob(){
  const tb = $("globTbl").querySelector("tbody");
  tb.innerHTML = globProps.map((r,i)=>globRow(i,r)).join("") || `<tr><td colspan="5" class="small">No properties yet.</td></tr>`;
  qsa("[data-g]").forEach(el=>{
    el.addEventListener("input", ()=>{
      const i=Number(el.dataset.i), k=el.dataset.g;
      globProps[i][k] = (k==="noi"||k==="ds")? safeNum(el.value): el.value;
      persistDraft();
    });
  });
}
function delGlob(i){ globProps.splice(i,1); renderGlob(); calc(); }
function addGlob(){
  globProps.push({name:"",noi:0,ds:0,notes:""});
  renderGlob();
}
function sumGlob(){
  const noi = globProps.reduce((s,r)=>s+safeNum(r.noi),0);
  const ds  = globProps.reduce((s,r)=>s+safeNum(r.ds),0);
  $("globalNoi").value = Math.round(noi);
  $("globalDebtSvc").value = Math.round(ds);
  toast("Global NOI / Debt Service set from table.");
  calc();
}

function drawRow(i, r){
  return `<tr>
    <td class="nowrap"><input data-d="mo" data-i="${i}" type="number" min="1" step="1" value="${safeNum(r.mo||1)}"/></td>
    <td class="right"><input data-d="amt" data-i="${i}" type="number" step="1000" value="${safeNum(r.amt||0)}"/></td>
    <td><input data-d="ms" data-i="${i}" value="${escapeHtml(r.ms||"")}" placeholder="Milestone"/></td>
    <td class="right"><button class="btn" onclick="delDraw(${i})">Del</button></td>
  </tr>`;
}
function renderDraws(){
  const tb = $("drawTbl").querySelector("tbody");
  tb.innerHTML = draws.map((r,i)=>drawRow(i,r)).join("") || `<tr><td colspan="4" class="small">No draw rows yet.</td></tr>`;
  qsa("[data-d]").forEach(el=>{
    el.addEventListener("input", ()=>{
      const i=Number(el.dataset.i), k=el.dataset.d;
      draws[i][k] = (k==="amt"||k==="mo")? safeNum(el.value): el.value;
      persistDraft();
    });
  });
}
function delDraw(i){ draws.splice(i,1); renderDraws(); calc(); }
function addDraw(){ draws.push({mo:1, amt:0, ms:""}); renderDraws(); }
function evenDraw(){
  // Auto-even draws across constMonths (or stabMonths if defined)
  const loanAmt = safeNum($("loanAmount").value);
  const months = Math.max(1, safeNum($("constMonths").value||0) || safeNum($("stabMonths").value||0) || 6);
  draws = [];
  const per = loanAmt / months;
  for(let m=1;m<=months;m++) draws.push({mo:m, amt:Math.round(per/1000)*1000, ms:""});
  renderDraws();
  toast("Even draw schedule generated.");
  calc();
}

/* =========================================================
   Persistence: localStorage + JSON
========================================================= */
const STORAGE_KEY="uw_tool_v2";
function collectState(){
  const ids = qsa("input,select,textarea").map(el=>el.id).filter(Boolean);
  const data={}; ids.forEach(id=>{ data[id] = $(id).value; });
  data.__tables = { rentRoll, globProps, draws };
  return data;
}
function applyState(data){
  if(!data) return;
  Object.keys(data).forEach(k=>{
    if(k==="__tables") return;
    const el=$(k);
    if(el) el.value = data[k];
  });
  if(data.__tables){
    rentRoll = data.__tables.rentRoll || [];
    globProps = data.__tables.globProps || [];
    draws = data.__tables.draws || [];
    renderRR(); renderGlob(); renderDraws();
  }
}
function persistDraft(){
  try{ localStorage.setItem(STORAGE_KEY+"_draft", JSON.stringify({ts:Date.now(), data:collectState()})); }catch(e){}
}
function save(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify({ts:Date.now(), data:collectState()}));
  toast("Saved.");
}
function load(){
  const raw = localStorage.getItem(STORAGE_KEY) || localStorage.getItem(STORAGE_KEY+"_draft");
  if(!raw) return false;
  try{
    const payload=JSON.parse(raw);
    applyState(payload.data||{});
    return true;
  }catch(e){ return false; }
}
function exportJSON(){
  const payload={version:"uw_tool_v2", exportedAt:new Date().toISOString(), data:collectState()};
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  const nm=($("dealName").value||"underwriting").replace(/[^a-z0-9]+/gi,"_").toLowerCase();
  a.href=url; a.download=nm+"_uw_v2.json";
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}
function importJSON(){
  const inp=document.createElement("input");
  inp.type="file"; inp.accept="application/json";
  inp.onchange=async ()=>{
    const f=inp.files[0]; if(!f) return;
    const txt=await f.text();
    try{
      const payload=JSON.parse(txt);
      applyState(payload.data||payload);
      toast("Imported.");
      calc();
    }catch(e){ alert("Invalid JSON."); }
  };
  inp.click();
}
function clearAll(){
  if(!confirm("Clear all inputs and local saved state?")) return;
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(STORAGE_KEY+"_draft");
  qsa("input,select,textarea").forEach(el=>{
    if(el.type==="number") el.value = 0;
    else if(el.type==="date") el.value = "";
    else if(el.tagName==="SELECT") el.selectedIndex = 0;
    else el.value = "";
  });
  rentRoll=[]; globProps=[]; draws=[];
  renderRR(); renderGlob(); renderDraws();
  // restore some sane defaults
  $("uwDate").value = todayISO();
  $("cofPctFunded").value = 80;
  $("cofRate").value = 8.00;
  $("policyMinDSCR").value = 1.20;
  $("targetROA").value = 10.0;
  $("targetCoC").value = 18.0;
  applyDefaults($("loanProduct").value);
  toast("Cleared.");
  calc();
}

/* =========================================================
   UI helpers
========================================================= */
function toast(msg){
  const t=document.createElement("div");
  t.textContent=msg;
  t.style.cssText="position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,.78);border:1px solid rgba(255,255,255,.18);color:white;padding:10px 12px;border-radius:12px;font-size:12px;z-index:9999";
  document.body.appendChild(t);
  setTimeout(()=>t.remove(), 1300);
}
function showPage(id){
  qsa(".page").forEach(p=>p.classList.remove("active"));
  $(id).classList.add("active");
  qsa("#nav button").forEach(b=>{
    b.classList.toggle("active", b.dataset.page===id);
  });
  window.scrollTo({top:0, behavior:"instant"});
}

/* =========================================================
   Core Calculations
========================================================= */
let last = {}; // store last calc outputs for memo/exports
function calc(){
  // Deal header
  const deal = $("dealName").value.trim() || "—";
  const prod = $("loanProduct").value;
  $("hdrDeal").textContent = deal;
  $("hdrProd").textContent = productLabel(prod);

  // Ensure source loan default mirrors loan amount
  if(safeNum($("srcLoan").value)===0 && safeNum($("loanAmount").value)>0){
    $("srcLoan").value = $("loanAmount").value;
  }

  // Pull core inputs
  const loanAmt = safeNum($("loanAmount").value);
  const termMo = Math.max(1, safeNum($("termMonths").value));
  const ioMo = safeNum($("ioMonths").value);
  const noteRate = safeNum($("noteRate").value);
  const amortY = Math.max(1, safeNum($("amortYears").value));
  const points = safeNum($("origPoints").value);
  const otherFees = safeNum($("otherFees").value);
  const exitFee = safeNum($("exitFee").value);
  const cof = safeNum($("cofRate").value);
  const cofPct = safeNum($("cofPctFunded").value)/100;

  const lienPos = $("lienPos").value;
  const seniorAhead = safeNum($("seniorAhead").value);

  // Values
  const asIs = safeNum($("asIsValue").value);
  const arv = safeNum($("arvValue").value) || asIs;
  const purchase = safeNum($("purchasePrice").value);

  // Uses
  const borClosing = safeNum($("borClosing").value);
  const rehab = safeNum($("rehabBudget").value);
  const soft = safeNum($("softCosts").value);
  const contPct = safeNum($("contPct").value)/100;
  const contingency = rehab * contPct;
  const estCarry = safeNum($("estCarry").value);
  const leasingCosts = safeNum($("leasingCosts").value);
  const totalUsesOverride = safeNum($("totalUses").value);
  const totalUses = totalUsesOverride>0 ? totalUsesOverride : (purchase+borClosing+rehab+soft+contingency+estCarry+leasingCosts);
  const loanSrc = safeNum($("srcLoan").value) || loanAmt;
  const equity = safeNum($("sponsorEquity").value);
  const otherFin = safeNum($("otherFin").value);
  const sources = loanSrc + equity + otherFin;
  const gap = sources - totalUses;

  // Historical Ops
  const grossRent = safeNum($("grossRent").value);
  const otherInc = safeNum($("otherIncome").value);
  const vacPct = safeNum($("vacPct").value)/100;
  const mgmtPct = safeNum($("mgmtPct").value)/100;

  const taxes = safeNum($("taxes").value);
  const ins = safeNum($("insurance").value);
  const repairs = safeNum($("repairs").value);
  const utils = safeNum($("utilities").value);
  const payroll = safeNum($("payroll").value);
  const otherOpex = safeNum($("otherOpex").value);
  const repl = safeNum($("replReserves").value);
  const normAdj = safeNum($("normAdj").value);

  const grossIncome = grossRent + otherInc;
  const vacLoss = grossIncome * vacPct;
  const egi = grossIncome - vacLoss;
  const mgmt = egi * mgmtPct;
  const opEx = taxes + ins + repairs + utils + payroll + otherOpex + mgmt + repl + Math.max(0,normAdj);
  const noi = egi - opEx;

  // Pro Forma
  const stbGrossRent = safeNum($("stbGrossRent").value);
  const stbOtherIncome = safeNum($("stbOtherIncome").value);
  const stbVacPct = safeNum($("stbVacPct").value)/100;
  const badDebtPct = safeNum($("badDebtPct").value)/100;
  const rentG = safeNum($("rentGrowth").value)/100;
  const expG = safeNum($("expGrowth").value)/100;

  const stbTaxes = safeNum($("stbTaxes").value) || taxes;
  const stbIns = safeNum($("stbIns").value) || ins;
  const stbRep = safeNum($("stbRepairs").value) || repairs;
  const stbUtl = safeNum($("stbUtils").value) || utils;
  const stbOther = safeNum($("stbOtherOpex").value) || (payroll + otherOpex);
  const stbRepl = safeNum($("stbRepl").value) || repl;
  const stbCapex = safeNum($("stbCapex").value);
  const stbMgmtPct = safeNum($("stbMgmtPct").value)/100;

  const stbGI = stbGrossRent + stbOtherIncome;
  const stbVacLoss = stbGI * stbVacPct;
  const stbBadDebt = stbGI * badDebtPct;
  const stbEGI = stbGI - stbVacLoss - stbBadDebt;
  const stbMgmt = stbEGI * stbMgmtPct;
  const stbOpEx = stbTaxes + stbIns + stbRep + stbUtl + stbOther + stbMgmt + stbRepl + stbCapex;
  const stbNOI = stbEGI - stbOpEx;

  // Debt service (current loan)
  const annIO = annInterestOnly(loanAmt, noteRate);
  const annAm = annDebtService(loanAmt, noteRate, amortY);
  const dsAnnual = (prod==="dscr" && ioMo===0) ? annAm : (ioMo>=termMo ? annIO : annAm);

  // Ratios
  const ltv = safeDiv(loanAmt, asIs)*100;
  const ltvStb = safeDiv(loanAmt, arv)*100;
  const totalDebt = loanAmt + (lienPos!=="1st" ? seniorAhead : 0);
  const cltv = safeDiv(totalDebt, asIs)*100;
  const cltvStb = safeDiv(totalDebt, arv)*100;
  const ltc = safeDiv(loanAmt, totalUses)*100;

  const dscrHist = safeDiv(noi, dsAnnual);
  const dscrStb = safeDiv(stbNOI, dsAnnual);

  const dyHist = safeDiv(noi, loanAmt)*100;
  const dyStb = safeDiv(stbNOI, loanAmt)*100;

  const breakeven = safeDiv((opEx + dsAnnual), grossIncome)*100;

  // Exit / takeout
  const exitCap = safeNum($("exitCap").value)/100;
  const capBuf = safeNum($("exitCapBuffer").value)/10000;
  const exitCapCons = exitCap + capBuf;
  const saleCostPct = safeNum($("saleCostPct").value)/100;
  const projYears = Math.max(1, Math.round(safeNum($("projYears").value)));
  const takeoutRate = safeNum($("takeoutRate").value);
  const takeoutAmort = Math.max(1, safeNum($("takeoutAmort").value));
  const takeoutDSCR = safeNum($("takeoutDSCR").value);

  // Projections
  let proj = [];
  let yRent = stbGrossRent;
  let yOtherInc = stbOtherIncome;
  let yOpEx = stbOpEx;
  for(let y=1;y<=projYears;y++){
    if(y>1){ yRent *= (1+rentG); yOtherInc *= (1+rentG*0.5); yOpEx *= (1+expG); }
    const yGI = yRent + yOtherInc;
    const yEGI = yGI * (1 - stbVacPct - badDebtPct);
    const yNOI = yEGI - yOpEx;
    proj.push({year:y, gross:yGI, egi:yEGI, opex:yOpEx, noi:yNOI, ds:dsAnnual, dscr:safeDiv(yNOI, dsAnnual)});
  }
  const exitNOI = proj.length ? proj[proj.length-1].noi : stbNOI;
  const exitValue = exitCapCons>0 ? (exitNOI/exitCapCons) : 0;
  const netSale = exitValue * (1 - saleCostPct);
  // payoff estimate: principal + (no amort assumption beyond dsAnnual calc; keep conservative)
  const payoff = totalDebt;
  const saleCoverage = safeDiv(netSale, payoff);
  const impliedSaleLoss = Math.max(0, payoff - netSale);

  // Takeout feasibility: compute max takeout loan based on DSCR requirement at takeout rate/amort
  const takeoutDSper$ = annDebtService(1_000_000, takeoutRate, takeoutAmort) / 1_000_000;
  const maxTakeout = (takeoutDSper$>0) ? (exitNOI / (takeoutDSCR * takeoutDSper$)) : 0;
  const takeoutLTV = safeDiv(maxTakeout, exitValue)*100;

  // Lender economics (simplified)
  const termYears = termMo/12;
  const grossInt = loanAmt*(noteRate/100)*termYears;
  const ptsIncome = loanAmt*(points/100);
  const income = grossInt + ptsIncome + otherFees + exitFee;
  const cofCost = (loanAmt*cofPct)*(cof/100)*termYears;
  const netLender = income - cofCost;
  const roa = safeDiv(netLender, loanAmt)*100;
  const lenderEquity = loanAmt*(1-cofPct);
  const coc = safeDiv(netLender, lenderEquity)*100;

  // Construction carry estimation
  const interestReserve = safeNum($("interestReserve").value);
  const stabMonths = safeNum($("stabMonths").value) || (safeNum($("constMonths").value)+safeNum($("leaseupMonths").value));
  const carryBufferMo = safeNum($("carryBufferMo").value);
  const avgUtilPct = safeNum($("avgUtilPct").value)/100;
  const leaseupDef = safeNum($("leaseupDef").value);
  const carryMonths = Math.max(0, stabMonths + carryBufferMo);
  const avgBal = loanAmt * clamp(avgUtilPct,0,1);
  const carryInterest = avgBal*(noteRate/100)*(carryMonths/12);
  const carryDeficit = leaseupDef*carryMonths;
  const carryNeed = carryInterest + carryDeficit;
  const carryGap = carryNeed - interestReserve;

  // Stress tests
  const stRentDown = safeNum($("stRentDown").value)/100;
  const stVacUp = safeNum($("stVacUp").value)/100;
  const stCapBps = safeNum($("stCapBps").value)/10000;
  const stRateBps = safeNum($("stRateBps").value)/10000;
  const stCostOver = safeNum($("stCostOver").value)/100;
  const stDelayMo = safeNum($("stDelayMo").value);

  const stGrossRent = stbGrossRent*(1-stRentDown);
  const stGI = stGrossRent + stbOtherIncome;
  const stVac = clamp(stbVacPct + stVacUp, 0, 0.50);
  const stEGI = stGI*(1 - stVac - badDebtPct);
  const stOpEx = stbOpEx*(1+expG); // mild expense shock
  const stNOI = stEGI - stOpEx;

  const stRate = noteRate + (stRateBps*100);
  const stDS = annInterestOnly(loanAmt, stRate); // conservative sensitivity
  const stDSCR = safeDiv(stNOI, stDS);

  const stExitCap = exitCapCons + stCapBps;
  const stExitValue = stExitCap>0 ? (stNOI/stExitCap) : 0;
  const stNetSale = stExitValue*(1-saleCostPct);
  const stSaleCoverage = safeDiv(stNetSale, payoff);
  const stImpliedLoss = Math.max(0, payoff - stNetSale);

  const stHard = rehab*(1+stCostOver);
  const stUses = (totalUsesOverride>0? totalUsesOverride : (purchase+borClosing+stHard+soft+(stHard*contPct)+estCarry+leasingCosts));
  const stLTC = safeDiv(loanAmt, stUses)*100;

  // Forced sale
  const forcedDisc = safeNum($("forcedDisc").value)/100;
  const workoutCost = safeNum($("workoutCost").value)/100;
  const forcedValue = asIs*(1-forcedDisc);
  const netForced = forcedValue*(1-workoutCost);
  const forcedCoverage = safeDiv(netForced, payoff);
  const forcedLoss = Math.max(0, payoff - netForced);

  // Sponsor / PFS
  const pfsCash = safeNum($("pfsCash").value);
  const pfsMarket = safeNum($("pfsMarket").value);
  const pfsOtherLiq = safeNum($("pfsOtherLiq").value);
  const pfsReEq = safeNum($("pfsReEq").value);
  const pfsBiz = safeNum($("pfsBiz").value);
  const pfsOther = safeNum($("pfsOther").value);
  const pfsLiab = safeNum($("pfsLiab").value);
  const contLiab = safeNum($("contLiab").value);
  const liqHaircut = safeNum($("liqHaircut").value)/100;
  const reHaircut = safeNum($("reHaircut").value)/100;

  const grossLiq = pfsCash + pfsMarket + pfsOtherLiq;
  const adjLiq = grossLiq*(1-liqHaircut);
  const adjRe = pfsReEq*(1-reHaircut);
  const nw = (grossLiq + pfsReEq + pfsBiz + pfsOther) - pfsLiab;
  const adjNW = (adjLiq + adjRe + pfsBiz + pfsOther) - (pfsLiab + contLiab*0.5); // haircutted contingent
  const liqToLoan = safeDiv(adjLiq, loanAmt);

  // Global
  const globalNoi = safeNum($("globalNoi").value);
  const globalDebtSvc = safeNum($("globalDebtSvc").value);
  const living = safeNum($("living").value);
  const otherDebt = safeNum($("otherDebt").value);
  const globalDSCR = safeDiv((globalNoi - living), (globalDebtSvc + otherDebt));
  const globalWithDeal = safeDiv((globalNoi - living), (globalDebtSvc + otherDebt + dsAnnual));
  const globVacSens = safeNum($("globVacSens").value)/100;
  const globRateBps = safeNum($("globRateBps").value)/10000;
  const globalNoiStress = globalNoi*(1-globVacSens);
  const globalDebtStress = globalDebtSvc*(1 + (globRateBps*2)); // rough
  const globalStressDSCR = safeDiv((globalNoiStress - living), (globalDebtStress + otherDebt + stDS));

  // Liquidity burn under delay stress (property deficit * delay)
  const stMonthlyDef = Math.max(0, (stDS - stNOI)/12);
  const liqBurn = stMonthlyDef * stDelayMo;
  const liqAfterBurn = Math.max(0, adjLiq - liqBurn);

  // Market / diligence scoring
  const liqScore = safeNum($("liqScore").value);
  const tts = safeNum($("tts").value);
  const pipeline = $("pipeline").value;
  const mktVacTrend = $("mktVacTrend").value;

  // Diligence readiness
  const titleOk = $("titleOk").value;
  const zoning = $("zoning").value;
  const phaseRes = $("phaseRes").value;
  const insOk = $("insOk").value;
  const appRev = $("appRev").value;

  // Risk rating (weighted)
  const wLev = safeNum($("wLev").value);
  const wCF = safeNum($("wCF").value);
  const wSp = safeNum($("wSp").value);
  const wCol = safeNum($("wCol").value);
  const wSum = Math.max(1, wLev+wCF+wSp+wCol);

  const levScore = scoreLeverage({prod, ltv, cltv, ltc, lienPos});
  const cfScore  = scoreCashflow({prod, dscrStb, stDSCR, dyStb});
  const spScore  = scoreSponsor({adjLiq, adjNW, liqToLoan, globalWithDeal, liqAfterBurn});
  const colScore = scoreCollateral({liqScore, tts, pipeline, mktVacTrend});

  const composite = (levScore*wLev + cfScore*wCF + spScore*wSp + colScore*wCol)/wSum;
  const rating = toRating(composite);
  const rec = recommend({prod, dscrStb, stDSCR, ltv, ltc, forcedCoverage, globalWithDeal, rating, titleOk, zoning, phaseRes, insOk, appRev});

  $("hdrRec").textContent = rec.status;
  $("hdrRisk").textContent = `${rating.code} (${Math.round(composite)})`;

  // Build Auto Conditions (baseline)
  const autoConds = buildConditions({prod, recourse:$("recourse").value, lienPos, titleOk, zoning, phaseRes, insOk, appRev, permits:$("permits").value});
  if(!$("conds").value.trim()) $("conds").value = autoConds.join("\n");

  // Render all page KPIs
  renderCompleteness({loanAmt, asIs, stbGrossRent, pfsCash, takeoutRate, titleOk});
  renderPayments({loanAmt, noteRate, amortY, termMo, ioMo, dsAnnual, annIO, annAm});
  renderEconomics({income, cofCost, netLender, roa, coc, targetROA:safeNum($("targetROA").value), targetCoC:safeNum($("targetCoC").value)});
  renderSU({sources, totalUses, gap, contingency, contPct, stUses, stLTC});
  renderMarket({liqScore, tts, pipeline, mktVacTrend});
  renderHist({grossIncome, egi, opEx, noi, dsAnnual, dscrHist, breakeven});
  renderPF({stbGI, stbEGI, stbOpEx, stbNOI, dsAnnual, dscrStb, dyStb, exitValue, netSale, saleCoverage, impliedSaleLoss, maxTakeout, takeoutLTV});
  renderProj(proj);
  renderCarry({carryMonths, avgBal, carryInterest, carryDeficit, carryNeed, interestReserve, carryGap});
  renderSponsor({grossLiq, adjLiq, nw, adjNW, liqToLoan, liqAfterBurn, liqBurn});
  renderGlobal({globalDSCR, globalWithDeal, globalStressDSCR});
  renderGlobalSensitivity({globalNoi, globalDebtSvc, living, otherDebt, dsAnnual});
  renderStress({stNOI, stDS, stDSCR, stExitValue, stNetSale, stSaleCoverage, stImpliedLoss, forcedValue, netForced, forcedCoverage, forcedLoss, stLTC});
  renderMatrix({stbNOI, loanAmt, noteRate, exitCapCons});
  renderStressNarr({stDSCR, stSaleCoverage, forcedCoverage, liqAfterBurn});
  renderDiligence({titleOk, zoning, phaseRes, insOk, appRev});
  renderRating({levScore, cfScore, spScore, colScore, composite, rating});
  renderPricingGuide({rating});
  renderRatingNarr({rating, rec, ltv, ltc, dscrStb, dyStb, forcedCoverage, globalWithDeal});
  renderMonitoringNarr({rec, covDSCR:safeNum($("covDSCR").value), covLTV:safeNum($("covLTV").value), sweep:safeNum($("sweepDSCR").value)});
  renderOutputs({deal, prod, rec, rating, loanAmt, termMo, ioMo, noteRate, points, ltv, ltvStb, cltv, cltvStb, ltc,
                 noi, stbNOI, dsAnnual, dscrHist, dscrStb, stDSCR, dyStb, breakeven,
                 exitValue, netSale, saleCoverage, impliedSaleLoss, maxTakeout, takeoutLTV,
                 forcedCoverage, forcedLoss, globalWithDeal, adjLiq, adjNW, liqAfterBurn, liqBurn, autoConds});
  renderQuality({deal, loanAmt, asIs, stbNOI, takeoutDSCR, maxTakeout, titleOk, zoning, phaseRes, insOk, appRev, rating, rec, rentComps:$("rentComps").value, saleComps:$("saleComps").value, txNarr:$("txNarr").value});

  // Save last calc snapshot
  last = { deal, prod, loanAmt, termMo, ioMo, noteRate, points, otherFees, exitFee, dsAnnual, noi, stbNOI, dscrStb, dyStb, ltv, ltc, exitValue, saleCoverage, forcedCoverage, composite, rating, rec };
  persistDraft();
}

/* =========================================================
   Scoring + Recommendation
========================================================= */
function scoreLeverage(m){
  // 0 best -> 100 worst
  const l = (m.lienPos==="1st") ? m.ltv : m.cltv;
  let s = 0;
  // Product baseline (higher risk)
  const base = { dscr:10, commercial_bridge:14, bridge:16, fixflip:20, construction:24, land:26, second_lien:30 }[m.prod] ?? 18;
  s += base;

  // Leverage
  if(l<=60) s += 10;
  else if(l<=70) s += 18;
  else if(l<=80) s += 30;
  else if(l<=90) s += 44;
  else s += 56;

  // LTC
  if(m.ltc<=60) s += 10;
  else if(m.ltc<=70) s += 18;
  else if(m.ltc<=80) s += 30;
  else if(m.ltc<=90) s += 42;
  else s += 52;

  return clamp(s,0,100);
}
function scoreCashflow(m){
  let s=0;
  // DSCR
  if(m.prod==="fixflip"||m.prod==="construction"||m.prod==="land"){
    // value-add: DSCR still matters but less determinative
    if(m.dscrStb>=1.20) s+=18;
    else if(m.dscrStb>=1.10) s+=28;
    else if(m.dscrStb>=1.00) s+=40;
    else s+=55;
  } else {
    if(m.dscrStb>=1.35) s+=12;
    else if(m.dscrStb>=1.20) s+=22;
    else if(m.dscrStb>=1.10) s+=36;
    else if(m.dscrStb>=1.00) s+=52;
    else s+=68;
  }
  // Stress DSCR
  if(m.stDSCR>=1.10) s+=10;
  else if(m.stDSCR>=1.00) s+=18;
  else if(m.stDSCR>=0.90) s+=28;
  else s+=40;

  // Debt Yield
  if(m.dyStb>=12) s+=10;
  else if(m.dyStb>=10) s+=16;
  else if(m.dyStb>=8) s+=26;
  else s+=38;

  return clamp(s,0,100);
}
function scoreSponsor(m){
  let s=0;
  if(m.adjNW<=0) s+=55;
  else if(m.adjNW<500000) s+=42;
  else if(m.adjNW<2000000) s+=32;
  else s+=22;

  if(m.adjLiq<50000) s+=40;
  else if(m.adjLiq<250000) s+=30;
  else if(m.adjLiq<750000) s+=22;
  else s+=14;

  // Liquidity relative to loan
  if(m.liqToLoan>=0.25) s+=10;
  else if(m.liqToLoan>=0.10) s+=18;
  else if(m.liqToLoan>=0.05) s+=26;
  else s+=34;

  // Global DSCR
  if(m.globalWithDeal>=1.25) s+=10;
  else if(m.globalWithDeal>=1.15) s+=16;
  else if(m.globalWithDeal>=1.05) s+=24;
  else s+=34;

  // Liquidity after burn
  if(m.liqAfterBurn>=250000) s+=8;
  else if(m.liqAfterBurn>=75000) s+=14;
  else s+=22;

  return clamp(s,0,100);
}
function scoreCollateral(m){
  let s=0;
  // liquidity score 1 best -> 5 worst
  s += (m.liqScore-1)*12; // 0..48
  if(m.tts<=3) s+=8;
  else if(m.tts<=6) s+=14;
  else if(m.tts<=9) s+=22;
  else s+=30;

  if(m.pipeline==="High") s+=10;
  else if(m.pipeline==="Moderate") s+=6;

  if(m.mktVacTrend==="Softening") s+=10;
  else if(m.mktVacTrend==="Stable") s+=6;

  return clamp(s,0,100);
}
function toRating(score){
  // Example: 1-10 style
  if(score<=20) return {code:"1", desc:"Low Risk"};
  if(score<=30) return {code:"2", desc:"Low‑Moderate"};
  if(score<=40) return {code:"3", desc:"Moderate"};
  if(score<=50) return {code:"4", desc:"Moderate‑Elevated"};
  if(score<=60) return {code:"5", desc:"Elevated"};
  if(score<=70) return {code:"6", desc:"High"};
  if(score<=80) return {code:"7", desc:"Very High"};
  return {code:"8", desc:"Workout / Special Mention"};
}
function recommend(m){
  // Conservative gating
  const minDSCR = safeNum($("policyMinDSCR").value);
  let status="Approve";
  const rationale=[];
  if(m.rating.code>="7"){ status="Decline"; rationale.push("Composite risk rating exceeds tolerance."); }
  else if(m.rating.code>="6"){ status="Approve with Conditions"; rationale.push("High risk rating; tighten leverage/reserves/controls."); }
  else if(m.rating.code>="5"){ status="Approve with Conditions"; rationale.push("Elevated risk rating; require mitigants."); }

  if(m.prod==="dscr" && m.dscrStb < Math.max(minDSCR,1.15)){ status="Decline"; rationale.push("Stabilized DSCR below policy minimum for cash‑flow lending."); }
  if(m.ltv>90 && m.prod!=="land"){ status="Decline"; rationale.push("As‑is leverage exceeds policy cap."); }
  if(m.ltc>90){ status="Approve with Conditions"; rationale.push("High LTC; require additional equity or reserves."); }
  if(m.forcedCoverage<0.90){ status="Approve with Conditions"; rationale.push("Forced sale proceeds may not cover total debt net of workout costs."); }
  if(m.globalWithDeal<1.05){ status="Approve with Conditions"; rationale.push("Weak global cash flow reduces sponsor capacity under stress."); }

  // Diligence gating (convert to conditions, not always immediate decline)
  const diligenceIssues=[];
  if(m.titleOk==="No") diligenceIssues.push("Unacceptable title exceptions.");
  if(m.zoning==="No") diligenceIssues.push("Zoning / use not confirmed.");
  if(m.phaseRes==="RECs Identified") diligenceIssues.push("Environmental RECs identified; require resolution/Phase II where applicable.");
  if(m.insOk==="No") diligenceIssues.push("Insurance not adequate.");
  if(m.appRev==="No") diligenceIssues.push("Appraisal not reviewed.");
  if(diligenceIssues.length){
    if(status==="Approve") status="Approve with Conditions";
    rationale.push("Diligence items outstanding: " + diligenceIssues.join(" "));
  }

  return {status, rationale};
}
function buildConditions(m){
  const conds=[];
  conds.push("Satisfactory third‑party valuation (Appraisal/BPO) supporting as‑is and stabilized assumptions.");
  conds.push("Title policy (loan policy) with required endorsements; no unacceptable exceptions; lender-approved settlement agent.");
  conds.push("Evidence of hazard/wind/flood (as applicable) with lender as mortgagee/additional insured; builder’s risk for construction.");
  if(m.prod==="construction"||m.prod==="fixflip"){
    conds.push("Third‑party budget review; executed GC contract; lender-approved draw/inspection protocol with lien waivers each draw.");
    conds.push("Contingency held lender-controlled to cover overruns; minimum contingency per policy.");
    if(m.permits!=="yes") conds.push("Evidence of permits/entitlements sufficient to commence work prior to first draw.");
  }
  if(m.prod==="bridge"||m.prod==="commercial_bridge"||m.prod==="construction"){
    conds.push("Interest reserve funded/held back as needed to cover projected carry through stabilization (including buffer).");
  }
  if(m.recourse!=="non") conds.push("Personal guaranty from financially capable sponsor; verification of liquidity and contingent liabilities.");
  conds.push("Ongoing reporting: monthly rent roll, quarterly operating statement, annual taxes/insurance evidence, and material lease updates.");
  if(m.phaseRes==="RECs Identified") conds.push("Environmental: address RECs per Phase I; Phase II / remediation if required.");
  if(m.zoning==="No") conds.push("Zoning/use confirmation and lender acceptance prior to closing.");
  return conds;
}

/* =========================================================
   Renderers
========================================================= */
function renderCompleteness(m){
  // Core 6: loan amount, as-is value, stabilized rent, sponsor cash, takeout rate, title ok
  const k=[];
  k.push(kpiCard("Loan Amount", fmt$(m.loanAmt), tagFor(m.loanAmt, x=>x>0, x=>x>0), "required"));
  k.push(kpiCard("As‑Is Value", fmt$(m.asIs), tagFor(m.asIs, x=>x>0, x=>x>0), "required"));
  k.push(kpiCard("Stabilized Rent", fmt$(m.stbGrossRent), tagFor(m.stbGrossRent, x=>x>0, x=>x>0), "required"));
  k.push(kpiCard("Sponsor Cash (PFS)", fmt$(m.pfsCash), tagFor(m.pfsCash, x=>x>0, x=>x>0), "required"));
  k.push(kpiCard("Takeout Rate", fmtPct(m.takeoutRate,2), tagFor(m.takeoutRate, x=>x>0, x=>x>0), "required"));
  k.push(kpiCard("Title Acceptable", escapeHtml(m.titleOk||"Unknown"), tagFor(m.titleOk, x=>x==="Yes", x=>x!=="No"), "required"));
  $("kpiCompleteness").innerHTML = k.join("");
}
function renderPayments(m){
  const k=[];
  const mIO = m.annIO/12;
  const mAm = m.annAm/12;
  k.push(kpiCard("Monthly IO Pmt", fmt$(mIO), tagFor(mIO, x=>x>0, x=>x>0)));
  k.push(kpiCard("Monthly Amort Pmt", fmt$(mAm), tagFor(mAm, x=>x>0, x=>x>0)));
  k.push(kpiCard("Annual Debt Service (UW)", fmt$(m.dsAnnual), tagFor(m.dsAnnual, x=>x>0, x=>x>0)));
  k.push(kpiCard("IO vs Term", `${Math.min(m.ioMo,m.termMo)} / ${m.termMo} mo`, tagFor(m.ioMo, x=>x>=m.termMo, x=>x>0), "IO months"));
  $("kpiPayments").innerHTML = k.join("");
}
function renderEconomics(m){
  const k=[];
  k.push(kpiCard("Total Lender Income", fmt$(m.income), tagFor(m.income, x=>x>0, x=>x>0)));
  k.push(kpiCard("Cost of Funds", fmt$(m.cofCost), tagFor(m.cofCost, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Net Income (Term)", fmt$(m.netLender), tagFor(m.netLender, x=>x>0, x=>x>=0)));
  k.push(kpiCard("ROA (Term)", fmtPct(m.roa,2), tagFor(m.roa, x=>x>=m.targetROA, x=>x>=m.targetROA*0.85), `target ${m.targetROA}%`));
  k.push(kpiCard("CoC (Equity)", fmtPct(m.coc,2), tagFor(m.coc, x=>x>=m.targetCoC, x=>x>=m.targetCoC*0.85), `target ${m.targetCoC}%`));
  $("kpiEconomics").innerHTML = k.join("");
}
function renderSU(m){
  const k=[];
  k.push(kpiCard("Total Uses", fmt$(m.totalUses), tagFor(m.totalUses, x=>x>0, x=>x>0)));
  k.push(kpiCard("Total Sources", fmt$(m.sources), tagFor(m.sources, x=>x>0, x=>x>0)));
  k.push(kpiCard("Sources‑Uses Gap", fmt$(m.gap), tagFor(m.gap, x=>Math.abs(x)<=100, x=>Math.abs(x)<=5000), "must ≈ 0"));
  k.push(kpiCard("Contingency", fmt$(m.contingency), tagFor(m.contingency, x=>x>=0.10*safeNum($("rehabBudget").value), x=>x>=0.05*safeNum($("rehabBudget").value)), "rule of thumb"));
  k.push(kpiCard("Stress LTC", fmtPct(m.stLTC,1), tagFor(m.stLTC, x=>x<=80, x=>x<=90), "overrun case"));
  $("kpiSU").innerHTML = k.join("");

  const box=$("gapBox");
  if(Math.abs(m.gap)>100){
    box.style.display="block";
    box.innerHTML = `<strong>Gap Detected:</strong> Sources (${fmt$(m.sources)}) do not equal Uses (${fmt$(m.totalUses)}). The gap is ${fmt$(m.gap)}. <br><br>
    <em>Institutional note:</em> A deal with an unexplained gap is not ready for approval. Identify the exact source or reduce uses.`;
  }else{
    box.style.display="none";
  }
}
function renderMarket(m){
  const k=[];
  k.push(kpiCard("Liquidity Score", `${m.liqScore}/5`, tagFor(m.liqScore, x=>x<=2, x=>x<=3), "lower is better"));
  k.push(kpiCard("Time‑to‑Sell", `${m.tts} mo`, tagFor(m.tts, x=>x<=6, x=>x<=9)));
  k.push(kpiCard("Supply Pipeline", escapeHtml(m.pipeline), tagFor(m.pipeline, x=>x==="Low", x=>x!=="High")));
  k.push(kpiCard("Vacancy Trend", escapeHtml(m.mktVacTrend), tagFor(m.mktVacTrend, x=>x==="Improving", x=>x==="Stable")));
  $("kpiMarket").innerHTML = k.join("");
}
function renderHist(m){
  const k=[];
  k.push(kpiCard("Gross Income", fmt$(m.grossIncome), tagFor(m.grossIncome, x=>x>0, x=>x>0)));
  k.push(kpiCard("Effective Gross Income", fmt$(m.egi), tagFor(m.egi, x=>x>0, x=>x>0)));
  k.push(kpiCard("Operating Expenses", fmt$(m.opEx), tagFor(m.opEx, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("NOI (Normalized)", fmt$(m.noi), tagFor(m.noi, x=>x>0, x=>x>=0)));
  k.push(kpiCard("DSCR (In‑Place)", fmtX(m.dscrHist,2), tagFor(m.dscrHist, x=>x>=1.20, x=>x>=1.10)));
  k.push(kpiCard("Breakeven Ratio", fmtPct(m.breakeven,1), tagFor(m.breakeven, x=>x<=80, x=>x<=90)));
  $("kpiHist").innerHTML = k.join("");
}
function renderPF(m){
  const k=[];
  k.push(kpiCard("Stabilized NOI", fmt$(m.stbNOI), tagFor(m.stbNOI, x=>x>0, x=>x>=0)));
  k.push(kpiCard("DSCR (Stabilized)", fmtX(m.dscrStb,2), tagFor(m.dscrStb, x=>x>=safeNum($("policyMinDSCR").value), x=>x>=safeNum($("policyMinDSCR").value)*0.92)));
  k.push(kpiCard("Debt Yield (Stb)", fmtPct(m.dyStb,2), tagFor(m.dyStb, x=>x>=10, x=>x>=8)));
  k.push(kpiCard("Exit Value (Cons.)", fmt$(m.exitValue), tagFor(m.exitValue, x=>x>0, x=>x>0)));
  k.push(kpiCard("Sale Coverage", fmtX(m.saleCoverage,2), tagFor(m.saleCoverage, x=>x>=1.0, x=>x>=0.9)));
  k.push(kpiCard("Implied Sale Loss", fmt$(m.impliedSaleLoss), tagFor(m.impliedSaleLoss, x=>x<=0, x=>x<=0)));
  k.push(kpiCard("Max Takeout Loan", fmt$(m.maxTakeout), tagFor(m.maxTakeout, x=>x>=safeNum($("loanAmount").value), x=>x>=safeNum($("loanAmount").value)*0.9)));
  k.push(kpiCard("Takeout LTV", fmtPct(m.takeoutLTV,1), tagFor(m.takeoutLTV, x=>x<=70, x=>x<=75)));
  $("kpiPF").innerHTML = k.join("");
}
function renderProj(proj){
  const tbl = `
    <table class="table">
      <thead><tr><th>Year</th><th class="right">Gross</th><th class="right">EGI</th><th class="right">OpEx</th><th class="right">NOI</th><th class="right">Debt Svc</th><th class="right">DSCR</th></tr></thead>
      <tbody>
        ${proj.map(r=>`
          <tr>
            <td class="mono">${r.year}</td>
            <td class="mono right">${fmt$(r.gross)}</td>
            <td class="mono right">${fmt$(r.egi)}</td>
            <td class="mono right">${fmt$(r.opex)}</td>
            <td class="mono right">${fmt$(r.noi)}</td>
            <td class="mono right">${fmt$(r.ds)}</td>
            <td class="mono right">${fmtX(r.dscr,2)}</td>
          </tr>`).join("")}
      </tbody>
    </table>`;
  $("projWrap").innerHTML = tbl;
}
function renderCarry(m){
  const k=[];
  k.push(kpiCard("Carry Months", `${m.carryMonths} mo`, tagFor(m.carryMonths, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Avg Balance", fmt$(m.avgBal), tagFor(m.avgBal, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Est. Interest Carry", fmt$(m.carryInterest), tagFor(m.carryInterest, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Est. Op Deficit", fmt$(m.carryDeficit), tagFor(m.carryDeficit, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Total Carry Need", fmt$(m.carryNeed), tagFor(m.carryNeed, x=>x>=0, x=>x>=0)));
  k.push(kpiCard("Interest Reserve", fmt$(safeNum($("interestReserve").value)), tagFor(safeNum($("interestReserve").value), x=>x>=m.carryNeed, x=>x>=m.carryNeed*0.8)));
  $("kpiCarry").innerHTML = k.join("");

  const w=$("carryWarn");
  if(m.carryGap>0){
    w.style.display="block";
    w.innerHTML = `<strong>Reserve Gap:</strong> Estimated carry need ${fmt$(m.carryNeed)} exceeds interest reserve ${fmt$(safeNum($("interestReserve").value))} by ${fmt$(m.carryGap)}.
      <br><br>Mitigants: (i) increase reserve, (ii) reduce leverage, (iii) shorten stabilization timeline, (iv) require sponsor liquidity covenant/top-up, or (v) require principal curtailment at milestones.`;
  }else{
    w.style.display="none";
  }
}
function renderSponsor(m){
  const k=[];
  k.push(kpiCard("Gross Liquidity", fmt$(m.grossLiq), tagFor(m.grossLiq, x=>x>=250000, x=>x>=75000)));
  k.push(kpiCard("Adj. Liquidity", fmt$(m.adjLiq), tagFor(m.adjLiq, x=>x>=250000, x=>x>=75000), "haircutted"));
  k.push(kpiCard("Net Worth", fmt$(m.nw), tagFor(m.nw, x=>x>=2000000, x=>x>=500000)));
  k.push(kpiCard("Adj. Net Worth", fmt$(m.adjNW), tagFor(m.adjNW, x=>x>=2000000, x=>x>=500000), "haircutted"));
  k.push(kpiCard("Liquidity / Loan", fmtPct(m.liqToLoan*100,1), tagFor(m.liqToLoan, x=>x>=0.10, x=>x>=0.05)));
  k.push(kpiCard("Liquidity After Delay", fmt$(m.liqAfterBurn), tagFor(m.liqAfterBurn, x=>x>=250000, x=>x>=75000)));
  $("kpiSponsor").innerHTML = k.join("");

  const w=$("sponsorWarn");
  if(m.liqAfterBurn<75000 && safeNum($("loanAmount").value)>0){
    w.style.display="block";
    w.innerHTML = `<strong>Sponsor support risk:</strong> Liquidity after delay burn is estimated at ${fmt$(m.liqAfterBurn)}. Consider requiring: larger reserves, additional collateral, or tighter leverage.`;
  }else{
    w.style.display="none";
  }
}
function renderGlobal(m){
  const k=[];
  k.push(kpiCard("Global DSCR", fmtX(m.globalDSCR,2), tagFor(m.globalDSCR, x=>x>=1.20, x=>x>=1.10)));
  k.push(kpiCard("Global DSCR (w/ Deal)", fmtX(m.globalWithDeal,2), tagFor(m.globalWithDeal, x=>x>=1.15, x=>x>=1.05)));
  k.push(kpiCard("Global Stress DSCR", fmtX(m.globalStressDSCR,2), tagFor(m.globalStressDSCR, x=>x>=1.00, x=>x>=0.90)));
  $("kpiGlobal").innerHTML = k.join("");
}
function renderGlobalSensitivity(m){
  const baseNoi=safeNum(m.globalNoi), baseDs=safeNum(m.globalDebtSvc), living=safeNum(m.living), other=safeNum(m.otherDebt), deal=safeNum(m.dsAnnual);
  const vac = [0, .05, .10];
  const rate = [0, .10, .20];
  let html=`<table class="table"><thead><tr><th></th>${rate.map(r=>`<th class="right">DS +${Math.round(r*100)}%</th>`).join("")}</tr></thead><tbody>`;
  vac.forEach(v=>{
    html += `<tr><th>NOI -${Math.round(v*100)}%</th>`;
    rate.forEach(r=>{
      const noi = baseNoi*(1-v) - living;
      const ds  = baseDs*(1+r) + other + deal;
      const g = safeDiv(noi, ds);
      html += `<td class="mono right">${fmtX(g,2)}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody></table>`;
  $("globSens").innerHTML = html;
}
function renderStress(m){
  const k=[];
  k.push(kpiCard("Stress NOI", fmt$(m.stNOI), tagFor(m.stNOI, x=>x>0, x=>x>=0)));
  k.push(kpiCard("Stress Debt Service", fmt$(m.stDS), tagFor(m.stDS, x=>x>0, x=>x>0)));
  k.push(kpiCard("Stress DSCR", fmtX(m.stDSCR,2), tagFor(m.stDSCR, x=>x>=1.00, x=>x>=0.90)));
  k.push(kpiCard("Stress Exit Value", fmt$(m.stExitValue), tagFor(m.stExitValue, x=>x>0, x=>x>0)));
  k.push(kpiCard("Stress Sale Coverage", fmtX(m.stSaleCoverage,2), tagFor(m.stSaleCoverage, x=>x>=1.0, x=>x>=0.9)));
  k.push(kpiCard("Stress Implied Loss", fmt$(m.stImpliedLoss), tagFor(m.stImpliedLoss, x=>x<=0, x=>x<=0)));
  k.push(kpiCard("Forced Coverage", fmtX(m.forcedCoverage,2), tagFor(m.forcedCoverage, x=>x>=1.0, x=>x>=0.9)));
  k.push(kpiCard("Forced Loss", fmt$(m.forcedLoss), tagFor(m.forcedLoss, x=>x<=0, x=>x<=0)));
  k.push(kpiCard("Stress LTC", fmtPct(m.stLTC,1), tagFor(m.stLTC, x=>x<=80, x=>x<=90)));
  $("kpiStress").innerHTML = k.join("");
}
function renderMatrix(m){
  const noi = safeNum(m.stbNOI);
  const loan = safeNum(m.loanAmt);
  const rate = safeNum(m.noteRate)/100;
  const baseCap = safeNum(m.exitCapCons);
  const capSteps = [-0.005, 0, 0.005, 0.01]; // +/- 50-100 bps
  const noiSteps = [-0.10, 0, 0.10]; // NOI down/up
  let html=`<table class="table"><thead><tr><th>NOI \\ Cap</th>${capSteps.map(c=>`<th class="right">${fmtPct((baseCap+c)*100,2)}</th>`).join("")}</tr></thead><tbody>`;
  noiSteps.forEach(nc=>{
    const n = noi*(1+nc);
    html += `<tr><th>NOI ${nc<0?nc*100:("+"+nc*100)}%</th>`;
    capSteps.forEach(c=>{
      const cap = baseCap + c;
      const val = cap>0? n/cap : 0;
      const cover = safeDiv(val, loan);
      html += `<td class="mono right">${fmtX(cover,2)}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody></table>
    <div class="footerNote">Table shows <strong>Value / Loan</strong> under NOI & cap shifts. Below 1.00× indicates principal impairment risk unless sponsor support exists.</div>`;
  $("sensMatrix").innerHTML = html;
}
function renderStressNarr(m){
  const lines=[];
  if(m.stDSCR<1.0) lines.push("Under the defined downside case, stressed cash flow does not fully cover debt service, indicating reliance on sponsor support/reserves.");
  else lines.push("Under the defined downside case, stressed cash flow remains near/above debt service, providing some survivability buffer.");
  if(m.stSaleCoverage<1.0) lines.push("A stressed sale scenario implies potential principal impairment net of sales costs; structure should be tightened or additional support required.");
  if(m.forcedCoverage<0.9) lines.push("Forced sale analysis indicates elevated loss severity risk; this is a key private-credit failure mode.");
  if(m.liqAfterBurn<75000) lines.push("Sponsor liquidity after delay burn appears limited; require reserves/top-up covenants or reduce leverage.");
  $("stressNarr").innerHTML = `<ul>${lines.map(x=>`<li>${escapeHtml(x)}</li>`).join("")}</ul>`;
}
function renderDiligence(m){
  const k=[];
  k.push(kpiCard("Title OK?", escapeHtml(m.titleOk), tagFor(m.titleOk, x=>x==="Yes", x=>x==="Unknown")));
  k.push(kpiCard("Zoning Verified?", escapeHtml(m.zoning), tagFor(m.zoning, x=>x==="Yes", x=>x==="Unknown")));
  k.push(kpiCard("Phase I Result", escapeHtml(m.phaseRes), tagFor(m.phaseRes, x=>x==="Clean"||x==="Not Applicable", x=>x==="Pending")));
  k.push(kpiCard("Insurance OK?", escapeHtml(m.insOk), tagFor(m.insOk, x=>x==="Yes", x=>x==="Unknown")));
  k.push(kpiCard("Appraisal Reviewed?", escapeHtml(m.appRev), tagFor(m.appRev, x=>x==="Yes", x=>x==="No")));
  $("kpiDiligence").innerHTML = k.join("");

  const warn=[];
  if(m.titleOk==="No") warn.push("Title exceptions unacceptable.");
  if(m.zoning==="No") warn.push("Zoning/use not confirmed.");
  if(m.phaseRes==="RECs Identified") warn.push("Environmental RECs identified.");
  if(m.insOk==="No") warn.push("Insurance inadequate.");
  if(m.appRev==="No") warn.push("Appraisal not reviewed.");
  const w=$("dilWarn");
  if(warn.length){
    w.style.display="block";
    w.innerHTML = `<strong>Diligence Flags:</strong> ${warn.map(escapeHtml).join(" ")}<br><br>Recommendation: condition approval on satisfactory resolution and document review.`;
  }else{
    w.style.display="none";
  }
}
function renderRating(m){
  const k=[];
  k.push(kpiCard("Leverage Score", `${Math.round(m.levScore)}`, tagFor(m.levScore, x=>x<=35, x=>x<=55)));
  k.push(kpiCard("Cash Flow Score", `${Math.round(m.cfScore)}`, tagFor(m.cfScore, x=>x<=35, x=>x<=55)));
  k.push(kpiCard("Sponsor Score", `${Math.round(m.spScore)}`, tagFor(m.spScore, x=>x<=35, x=>x<=55)));
  k.push(kpiCard("Collateral Score", `${Math.round(m.colScore)}`, tagFor(m.colScore, x=>x<=35, x=>x<=55)));
  k.push(kpiCard("Composite", `${Math.round(m.composite)}`, tagFor(m.composite, x=>x<=40, x=>x<=60), `rating ${m.rating.code}`));
  k.push(kpiCard("Rating", `${m.rating.code} – ${escapeHtml(m.rating.desc)}`, tagFor(Number(m.rating.code), x=>x<=4, x=>x<=6)));
  $("kpiRating").innerHTML = k.join("");
}
function renderPricingGuide(m){
  // simplistic guide (replace with your matrix)
  const r = Number(m.rating.code);
  const baseSpread = 6.0 + (r*0.5); // placeholder
  const pts = 1.0 + Math.max(0,(r-3))*0.25;
  const html = [
    kpiCard("Suggested Spread (over COF)", fmtPct(baseSpread,2), tagFor(r, x=>x<=4, x=>x<=6), "advisory"),
    kpiCard("Suggested Points", fmtPct(pts,2), tagFor(r, x=>x<=4, x=>x<=6), "advisory"),
    kpiCard("Structure Bias", (r>=6?"Tighten":"Standard"), tagFor(r, x=>x<=4, x=>x<=6), "LTV/LTC/reserves"),
    kpiCard("Cash Mgmt", (r>=6?"Required":"As Needed"), tagFor(r, x=>x<=4, x=>x<=6), "lockbox triggers"),
  ].join("");
  $("kpiPricingGuide").innerHTML = html;
}
function renderRatingNarr(m){
  const lines=[];
  lines.push(`Composite rating is <strong>${m.rating.code}</strong> (${Math.round(m.rating?.code?Number(m.rating.code):0)}), based on leverage, cash flow, sponsor strength, and collateral liquidity.`);
  lines.push(`Key quantitative metrics: LTV ${fmtPct(m.ltv,1)}, LTC ${fmtPct(m.ltc,1)}, DSCR (stabilized) ${fmtX(m.dscrStb,2)}, debt yield ${fmtPct(m.dyStb,2)}.`);
  if(m.forcedCoverage<1.0) lines.push(`Downside: forced sale coverage is ${fmtX(m.forcedCoverage,2)}, indicating ${m.forcedCoverage<0.9?"elevated":"moderate"} loss severity risk.`);
  if(m.globalWithDeal<1.10) lines.push(`Sponsor global DSCR with deal is ${fmtX(m.globalWithDeal,2)}, indicating limited capacity to support the loan under broader stress.`);
  if(m.rec.status!=="Approve") lines.push(`Recommendation is <strong>${escapeHtml(m.rec.status)}</strong> subject to conditions and structure tightening due to identified risks.`);
  $("ratingNarr").innerHTML = `<ul>${lines.map(x=>`<li>${x}</li>`).join("")}</ul>`;
}
function renderMonitoringNarr(m){
  const lines=[];
  lines.push(`Ongoing monitoring should include monthly rent roll: <strong>${escapeHtml($("monRR").value)}</strong>, quarterly financials: <strong>${escapeHtml($("qFin").value)}</strong>, and annual financials: <strong>${escapeHtml($("annFin").value)}</strong>.`);
  lines.push(`Covenants: minimum DSCR ${fmtX(m.covDSCR,2)}; maximum LTV ${fmtPct(m.covLTV,1)}; cash sweep trigger ${fmtX(m.sweep,2)}.`);
  if(m.rec.status!=="Approve") lines.push("Given elevated risks, require more frequent reporting and reserve top-ups if performance deteriorates.");
  $("monNarr").innerHTML = `<ul>${lines.map(x=>`<li>${x}</li>`).join("")}</ul>`;
}
function renderOutputs(o){
  // Exec KPIs
  const k=[];
  const minDSCR = safeNum($("policyMinDSCR").value);
  k.push(kpiCard("Recommendation", escapeHtml(o.rec.status), tagFor(o.rec.status, x=>x==="Approve", x=>x==="Approve with Conditions")));
  k.push(kpiCard("Risk Rating", `${o.rating.code} – ${escapeHtml(o.rating.desc)}`, tagFor(Number(o.rating.code), x=>x<=4, x=>x<=6)));
  k.push(kpiCard("Loan Amount", fmt$(o.loanAmt), tagFor(o.loanAmt, x=>x>0, x=>x>0)));
  k.push(kpiCard("Term / IO", `${o.termMo} / ${o.ioMo} mo`, tagFor(o.ioMo, x=>x>=o.termMo, x=>x>0)));
  k.push(kpiCard("Rate / Points", `${o.noteRate.toFixed(2)}% / ${o.points.toFixed(2)} pts`, tagFor(o.noteRate, x=>x>0, x=>x>0)));
  k.push(kpiCard("LTV (As‑Is)", fmtPct(o.ltv,1), tagFor(o.ltv, x=>x<=70, x=>x<=80)));
  k.push(kpiCard("LTC", fmtPct(o.ltc,1), tagFor(o.ltc, x=>x<=70, x=>x<=80)));
  k.push(kpiCard("DSCR (Stb)", fmtX(o.dscrStb,2), tagFor(o.dscrStb, x=>x>=minDSCR, x=>x>=minDSCR*0.92)));
  k.push(kpiCard("Debt Yield", fmtPct(o.dyStb,2), tagFor(o.dyStb, x=>x>=10, x=>x>=8)));
  k.push(kpiCard("Stress DSCR", fmtX(o.stDSCR,2), tagFor(o.stDSCR, x=>x>=1.00, x=>x>=0.90)));
  k.push(kpiCard("Forced Coverage", fmtX(o.forcedCoverage,2), tagFor(o.forcedCoverage, x=>x>=1.00, x=>x>=0.90)));
  k.push(kpiCard("Global DSCR (w/ deal)", fmtX(o.globalWithDeal,2), tagFor(o.globalWithDeal, x=>x>=1.15, x=>x>=1.05)));
  $("kpiExec").innerHTML = k.join("");

  // Risks/mitigants
  const risks=[], mits=[];
  if(o.ltv>80) risks.push("Elevated leverage increases loss severity risk.");
  else mits.push("Leverage within moderate range.");
  if(o.ltc>80) risks.push("High LTC reduces buffer for overruns/soft cost creep.");
  else mits.push("Sponsor has meaningful equity at risk (LTC acceptable).");
  if(o.dscrStb<minDSCR) risks.push("Stabilized DSCR below policy; takeout risk elevated.");
  else mits.push("Stabilized DSCR provides cushion.");
  if(o.stDSCR<1.0) risks.push("Stress DSCR < 1.00× implies potential payment shortfall under downside case.");
  else mits.push("Stress case remains near/above break-even.");
  if(o.forcedCoverage<0.90) risks.push("Forced sale analysis suggests potential principal impairment net of workout costs.");
  else mits.push("Forced sale proceeds appear sufficient to cover debt.");
  if(o.globalWithDeal<1.05) risks.push("Weak global DSCR indicates limited sponsor backstop under stress.");
  else mits.push("Global cash flow indicates capacity to support debt.");

  $("outRisks").innerHTML = (risks.length?risks:["No material risks identified beyond normal market volatility (based on provided inputs)."]).map(x=>`<li>${escapeHtml(x)}</li>`).join("");
  $("outMitigants").innerHTML = (mits.length?mits:["Mitigants depend on final structure: reserves, draw controls, lower leverage, and verified exit."]).map(x=>`<li>${escapeHtml(x)}</li>`).join("");

  // Exit summary
  const primary = (o.prod==="fixflip") ? "Sale upon renovation completion." :
                  (o.prod==="construction") ? "Sale or refinance upon CO and stabilization." :
                  (o.prod==="dscr") ? "Hold to maturity; refinance only if terms favorable." :
                  "Refinance to permanent debt upon stabilization; sale as secondary.";
  $("outExit").innerHTML = `
    <div class="small">
      <div><span class="chip">Primary Exit</span> ${escapeHtml(primary)}</div>
      <div style="height:6px"></div>
      <div><span class="chip">Exit Value</span> ${fmt$(o.exitValue)} (net sale ${fmt$(o.netSale)}; sale coverage ${fmtX(o.saleCoverage,2)})</div>
      <div style="height:6px"></div>
      <div><span class="chip">Takeout</span> Max takeout ${fmt$(o.maxTakeout)}; takeout LTV ${fmtPct(o.takeoutLTV,1)}</div>
    </div>`;

  // Conditions list
  $("outConds").innerHTML = o.autoConds.map(x=>`<li>${escapeHtml(x)}</li>`).join("");

  // Memo (structured, longer)
  $("memo").innerHTML = buildMemo(o);
}
function buildMemo(o){
  const date = new Date().toLocaleDateString();
  const deal = escapeHtml($("dealName").value||"—");
  const uw = escapeHtml($("underwriter").value||"—");
  const entity = escapeHtml($("borrowerEntity").value||"—");
  const guarantors = escapeHtml($("guarantors").value||"—");
  const addr = escapeHtml($("propAddr").value||"—");
    const submarket = escapeHtml($("submarket").value||"—");
  const assetType = escapeHtml($("assetType").value||"—");
  const purpose = escapeHtml($("loanPurpose").value||"—");
  const narrative = escapeHtml($("txNarr").value||"");
  const lienPos = escapeHtml($("lienPos").value||"—");
  const recourse = escapeHtml($("recourse").value||"—");

  const asIsVal = safeNum($("asIsValue").value);
  const arvVal  = safeNum($("arvValue").value) || asIsVal;

  const srcs = safeNum($("srcLoan").value) + safeNum($("sponsorEquity").value) + safeNum($("otherFin").value);
  const usesOverride = safeNum($("totalUses").value);
  const purchase = safeNum($("purchasePrice").value);
  const borClosing = safeNum($("borClosing").value);
  const rehab = safeNum($("rehabBudget").value);
  const soft = safeNum($("softCosts").value);
  const contPct = safeNum($("contPct").value)/100;
  const contingency = rehab*contPct;
  const estCarry = safeNum($("estCarry").value);
  const leasing = safeNum($("leasingCosts").value);
  const uses = (usesOverride>0) ? usesOverride : (purchase+borClosing+rehab+soft+contingency+estCarry+leasing);

  const titleOk = escapeHtml($("titleOk").value||"Unknown");
  const zoning = escapeHtml($("zoning").value||"Unknown");
  const phaseRes = escapeHtml($("phaseRes").value||"Pending");
  const insOk = escapeHtml($("insOk").value||"Unknown");
  const appRev = escapeHtml($("appRev").value||"No");

  const strengths = escapeHtml($("uwStrengths").value||"");
  const weaknesses = escapeHtml($("uwWeaknesses").value||"");
  const questions = escapeHtml($("uwQuestions").value||"");
  const talking = escapeHtml($("uwTalking").value||"");

  const risks = [];
  const mits = [];

  const minDSCR = safeNum($("policyMinDSCR").value);
  if(o.ltv>80) risks.push("Elevated leverage increases potential loss severity and reduces refinance flexibility.");
  else mits.push("Leverage within a moderate range based on current inputs.");
  if(o.ltc>80) risks.push("High LTC reduces buffer for cost overruns and execution delays.");
  else mits.push("Meaningful sponsor equity at risk supports loss-absorption.");
  if(o.dscrStb<minDSCR) risks.push("Stabilized DSCR below policy indicates elevated takeout and cash-flow risk.");
  else mits.push("Stabilized DSCR meets/exceeds minimum policy threshold.");
  if(o.stDSCR<1.00) risks.push("Downside case produces DSCR < 1.00× indicating potential payment shortfall absent reserves/sponsor support.");
  else mits.push("Downside case remains near/above break-even.");
  if(o.forcedCoverage<0.90) risks.push("Forced-sale coverage suggests potential principal impairment net of workout costs.");
  else mits.push("Forced-sale coverage indicates limited principal impairment risk.");
  if(o.globalWithDeal<1.05) risks.push("Sponsor global cash flow is thin; limited capacity to support the deal through stress.");
  else mits.push("Global cash flow indicates capacity to support the debt under moderate stress.");

  const conds = ($("conds").value||"").split("\n").map(s=>s.trim()).filter(Boolean);

  const section = (title, bodyHtml)=>`
    <div class="card" style="margin-bottom:12px">
      <h3 style="margin-top:0">${title}</h3>
      ${bodyHtml}
    </div>`;

  const bullets = (arr)=> arr.length ? `<ul class="small">${arr.map(x=>`<li>${x}</li>`).join("")}</ul>` : `<div class="small">—</div>`;

  const fmtMaybe = (n)=> (n>0 ? fmt$(n) : "—");

  // Committee-ready memo structure
  let memo = "";

  memo += section("1. Executive Credit Summary",
    `<div class="small">
      <div><span class="chip">Deal</span> <strong>${deal}</strong> · <span class="chip">Date</span> ${date} · <span class="chip">Underwriter</span> ${uw}</div>
      <div style="height:8px"></div>
      <table class="table">
        <tbody>
          <tr><th>Borrower</th><td>${entity}</td><th>Guarantor(s)</th><td>${guarantors}</td></tr>
          <tr><th>Collateral</th><td>${propAddr}</td><th>Product</th><td>${productLabel(o.prod)}</td></tr>
          <tr><th>Purpose</th><td>${purpose}</td><th>Structure</th><td>${lienPos} · ${recourse}</td></tr>
          <tr><th>Recommendation</th><td><strong>${escapeHtml(o.rec.status)}</strong></td><th>Risk Rating</th><td><strong>${o.rating.code}</strong> – ${escapeHtml(o.rating.desc)}</td></tr>
        </tbody>
      </table>
      ${o.rec.rationale?.length ? `<div style="margin-top:8px" class="warnbox"><strong>Rationale:</strong> ${escapeHtml(o.rec.rationale.join(" "))}</div>` : ""}
    </div>`);

  memo += section("2. Transaction Overview",
    `<div class="small">
      <div><strong>Business plan narrative:</strong></div>
      <div style="margin-top:6px">${narrative ? narrative : "<em>No narrative provided.</em>"}</div>
      <div class="hr"></div>
      <table class="table">
        <tbody>
          <tr><th>Loan Amount</th><td class="mono right">${fmt$(o.loanAmt)}</td><th>Term / IO</th><td class="mono right">${o.termMo} / ${o.ioMo} mo</td></tr>
          <tr><th>Rate / Points</th><td class="mono right">${o.noteRate.toFixed(2)}% / ${o.points.toFixed(2)} pts</td><th>Est. Debt Service (annual)</th><td class="mono right">${fmt$(o.dsAnnual)}</td></tr>
        </tbody>
      </table>
    </div>`);

  memo += section("3. Sources & Uses",
    `<div class="small">
      <table class="table">
        <thead><tr><th>Uses</th><th class="right">Amount</th><th>Sources</th><th class="right">Amount</th></tr></thead>
        <tbody>
          <tr><td>Purchase</td><td class="mono right">${fmtMaybe(purchase)}</td><td>Loan Proceeds</td><td class="mono right">${fmtMaybe(safeNum($("srcLoan").value))}</td></tr>
          <tr><td>Borrower Closing</td><td class="mono right">${fmtMaybe(borClosing)}</td><td>Sponsor Equity</td><td class="mono right">${fmtMaybe(safeNum($("sponsorEquity").value))}</td></tr>
          <tr><td>Hard Costs</td><td class="mono right">${fmtMaybe(rehab)}</td><td>Other Financing</td><td class="mono right">${fmtMaybe(safeNum($("otherFin").value))}</td></tr>
          <tr><td>Soft Costs</td><td class="mono right">${fmtMaybe(soft)}</td><td></td><td></td></tr>
          <tr><td>Contingency</td><td class="mono right">${fmtMaybe(contingency)}</td><td></td><td></td></tr>
          <tr><td>Interest Carry (est.)</td><td class="mono right">${fmtMaybe(estCarry)}</td><td></td><td></td></tr>
          <tr><td>Leasing / TI / LC</td><td class="mono right">${fmtMaybe(leasing)}</td><td></td><td></td></tr>
          <tr><th>Total Uses</th><th class="mono right">${fmt$(uses)}</th><th>Total Sources</th><th class="mono right">${fmt$(srcs)}</th></tr>
        </tbody>
      </table>
      ${Math.abs(srcs-uses)>100 ? `<div class="badbox"><strong>Gap:</strong> Sources and uses do not reconcile. Address prior to approval/closing.</div>` : ""}
    </div>`);

  memo += section("4. Collateral & Market",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>Address</th><td>${propAddr}</td><th>Submarket</th><td>${submarket}</td></tr>
          <tr><th>Asset Type</th><td>${assetType}</td><th>Liquidity (1–5)</th><td class="mono">${escapeHtml($("liqScore").value||"—")}</td></tr>
          <tr><th>As-Is Value</th><td class="mono right">${fmt$(asIsVal)}</td><th>Stabilized/ARV</th><td class="mono right">${fmt$(arvVal)}</td></tr>
        </tbody>
      </table>
      <div class="hr"></div>
      <div><strong>Market support:</strong></div>
      <div style="margin-top:6px"><em>Rent comps:</em> ${escapeHtml($("rentComps").value||"—")}</div>
      <div style="margin-top:6px"><em>Sale comps:</em> ${escapeHtml($("saleComps").value||"—")}</div>
    </div>`);

  memo += section("5. Historical Performance (if applicable)",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>NOI (Normalized)</th><td class="mono right">${fmt$(o.noi)}</td><th>DSCR (In-Place)</th><td class="mono right">${fmtX(o.dscrHist,2)}</td></tr>
          <tr><th>Breakeven</th><td class="mono right">${fmtPct(o.breakeven,1)}</td><th>Debt Yield (In-Place)</th><td class="mono right">${fmtPct(safeDiv(o.noi,o.loanAmt)*100,2)}</td></tr>
        </tbody>
      </table>
      <div class="footerNote">Historical results should be tied to bank statements / operating statements; normalize vacancy, management, and reserves.</div>
    </div>`);

  memo += section("6. Underwritten Pro Forma & Exit",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>Stabilized NOI</th><td class="mono right">${fmt$(o.stbNOI)}</td><th>DSCR (Stabilized)</th><td class="mono right">${fmtX(o.dscrStb,2)}</td></tr>
          <tr><th>Debt Yield (Stabilized)</th><td class="mono right">${fmtPct(o.dyStb,2)}</td><th>Exit Value (Cons.)</th><td class="mono right">${fmt$(o.exitValue)}</td></tr>
          <tr><th>Sale Coverage</th><td class="mono right">${fmtX(o.saleCoverage,2)}</td><th>Max Takeout</th><td class="mono right">${fmt$(o.maxTakeout)}</td></tr>
          <tr><th>Takeout LTV</th><td class="mono right">${fmtPct(o.takeoutLTV,1)}</td><th>Implied Sale Loss</th><td class="mono right">${fmt$(o.impliedSaleLoss)}</td></tr>
        </tbody>
      </table>
      <div class="footerNote">Exit value uses cap rate + buffer. Takeout sizing uses DSCR requirement and takeout rate/amort assumptions.</div>
    </div>`);

  memo += section("7. Stress Testing & Loss Severity",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>Stress DSCR</th><td class="mono right">${fmtX(o.stDSCR,2)}</td><th>Forced Sale Coverage</th><td class="mono right">${fmtX(o.forcedCoverage,2)}</td></tr>
          <tr><th>Forced Sale Loss</th><td class="mono right">${fmt$(o.forcedLoss)}</td><th>Committee View</th><td>${(o.forcedCoverage<1.0||o.stDSCR<1.0)?"Elevated downside; structure/mitigants required.":"Downside appears manageable based on inputs."}</td></tr>
        </tbody>
      </table>
      <div class="footerNote">Forced sale is a practical private-credit loss path; underwrite buyer pool and time-to-exit conservatively.</div>
    </div>`);

  memo += section("8. Sponsor / Guarantor Support",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>Adj. Liquidity</th><td class="mono right">${fmt$(o.adjLiq)}</td><th>Adj. Net Worth</th><td class="mono right">${fmt$(o.adjNW)}</td></tr>
          <tr><th>Liquidity After Delay Burn</th><td class="mono right">${fmt$(o.liqAfterBurn)}</td><th>Est. Delay Burn</th><td class="mono right">${fmt$(o.liqBurn)}</td></tr>
          <tr><th>Global DSCR (w/ deal)</th><td class="mono right">${fmtX(o.globalWithDeal,2)}</td><th>Recourse</th><td>${recourse}</td></tr>
        </tbody>
      </table>
      <div class="footerNote">Sponsor metrics are conservative (haircuts). For institutional approval, attach full PFS and schedule of RE owned.</div>
    </div>`);

  memo += section("9. Collateral & Legal / Diligence",
    `<div class="small">
      <table class="table">
        <tbody>
          <tr><th>Title Acceptable</th><td>${titleOk}</td><th>Zoning Verified</th><td>${zoning}</td></tr>
          <tr><th>Environmental</th><td>${phaseRes}</td><th>Insurance Adequate</th><td>${insOk}</td></tr>
          <tr><th>Appraisal Reviewed</th><td>${appRev}</td><th></th><td></td></tr>
        </tbody>
      </table>
      <div style="margin-top:8px"><strong>Diligence notes / conditions:</strong></div>
      <div style="margin-top:6px">${escapeHtml($("legalConds").value||"—")}</div>
    </div>`);

  memo += section("10. Key Risks & Mitigants",
    `<div class="small">
      <div class="grid2" style="gap:10px">
        <div><strong>Risks</strong>${bullets(risks)}</div>
        <div><strong>Mitigants</strong>${bullets(mits)}</div>
      </div>
    </div>`);

  memo += section("11. Conditions / Covenants / Monitoring",
    `<div class="small">
      ${conds.length ? `<ul class="small">${conds.map(c=>`<li>${escapeHtml(c)}</li>`).join("")}</ul>` : "<em>No conditions provided.</em>"}
      <div class="hr"></div>
      <div><strong>Reporting:</strong> Monthly rent roll: ${escapeHtml($("monRR").value)} · Quarterly financials: ${escapeHtml($("qFin").value)} · Annual financials: ${escapeHtml($("annFin").value)} · Construction reporting: ${escapeHtml($("constRpt").value)}</div>
    </div>`);

  memo += section("12. Underwriter Notes (for committee discussion)",
    `<div class="small">
      <div><strong>Strengths:</strong></div><div style="margin-top:6px">${strengths||"—"}</div>
      <div style="height:10px"></div>
      <div><strong>Weaknesses:</strong></div><div style="margin-top:6px">${weaknesses||"—"}</div>
      <div style="height:10px"></div>
      <div><strong>Key Questions:</strong></div><div style="margin-top:6px">${questions||"—"}</div>
      <div style="height:10px"></div>
      <div><strong>Talking Points:</strong></div><div style="margin-top:6px">${talking||"—"}</div>
    </div>`);

  return memo;
}

/* =========================================================
   File Quality / Feedback Engine
========================================================= */
function renderQuality(m){
  const flags=[];
  const tips=[];

  // Missing core items
  if(!m.deal || m.deal==="—") flags.push({lvl:"bad", txt:"Deal name is blank (audit trail / exports will be weak)."});
  if(m.loanAmt<=0) flags.push({lvl:"bad", txt:"Loan amount is not entered; downstream metrics are not meaningful."});
  if(m.asIs<=0) flags.push({lvl:"bad", txt:"As-is value is missing; leverage cannot be validated."});
  if(m.stbNOI<=0) flags.push({lvl:"warn", txt:"Stabilized NOI is missing/zero; takeout and exit math will be unreliable."});
  if(!m.txNarr || m.txNarr.trim().length<120) flags.push({lvl:"warn", txt:"Transaction narrative is thin. Committee memos should clearly state purpose, plan, and exit."});
  if(!m.rentComps || m.rentComps.trim().length<60) flags.push({lvl:"warn", txt:"Rent comp support is thin. Stabilized rents should be defensible."});
  if(!m.saleComps || m.saleComps.trim().length<60) flags.push({lvl:"warn", txt:"Sale comp / cap rate support is thin. Exit assumptions should be supported."});

  // Diligence readiness
  if(m.titleOk==="No") flags.push({lvl:"bad", txt:"Title marked not acceptable. This is typically a gating issue."});
  if(m.zoning==="No") flags.push({lvl:"bad", txt:"Zoning/use not verified. Confirm prior to closing."});
  if(m.phaseRes==="RECs Identified") flags.push({lvl:"warn", txt:"Environmental RECs identified. Confirm mitigations / Phase II where needed."});
  if(m.insOk==="No") flags.push({lvl:"bad", txt:"Insurance not adequate. This is a closing requirement."});
  if(m.appRev==="No") flags.push({lvl:"warn", txt:"Appraisal not reviewed. Provide internal review and reconcile to underwriting."});

  // Credit quality suggestions
  if(m.rec.status!=="Approve"){
    tips.push("Translate each key risk into an enforceable mitigant: lower leverage, reserves, covenants, cash management, or additional collateral.");
    tips.push("Tighten assumptions rather than trying to 'solve' with pricing. Private credit wins by avoiding loss severity, not by maximizing coupon.");
  }
  if(m.rating.code>="6"){
    tips.push("For high-risk grades, consider: (i) lender-controlled contingency, (ii) interest reserve sized to stabilization + buffer, (iii) PG + liquidity covenant, (iv) milestone-based curtailments.");
  }
  if(m.maxTakeout<m.loanAmt && m.loanAmt>0){
    tips.push("Exit feasibility: permanent takeout (by DSCR) does not size to repay the loan. Require lower loan amount, verified sale exit, or additional equity.");
  }
  if(m.forcedCoverage<1.0){
    tips.push("Loss severity: forced-sale proceeds may not repay debt. Reduce leverage and/or require additional collateral, guarantees, and reserves.");
  }

  const renderFlag = (f)=>{
    const cls = f.lvl==="bad" ? "badbox" : (f.lvl==="warn" ? "warnbox" : "callout");
    return `<div class="${cls}" style="margin-bottom:8px"><strong>${f.lvl.toUpperCase()}:</strong> ${escapeHtml(f.txt)}</div>`;
  };

  const html = `
    ${flags.length ? flags.map(renderFlag).join("") : `<div class="callout"><strong>Good:</strong> No major file-quality flags detected based on current inputs.</div>`}
    <div class="hr"></div>
    <div class="callout">
      <strong>Suggested Improvements:</strong>
      <ul>
        ${(tips.length?tips:[
          "Add exhibits: rent roll, T-12, appraisal/BPO, title, insurance binder, budget, sponsor PFS, and exit comps.",
          "If value-add/construction: attach GC contract, schedule, draw controls, and third-party budget review."
        ]).map(t=>`<li>${escapeHtml(t)}</li>`).join("")}
      </ul>
    </div>`;
  $("qualityFlags").innerHTML = html;
}

/* =========================================================
   Initialization / Events
========================================================= */
function wire(){
  // Navigation
  qsa("#nav button").forEach(b=>{
    b.addEventListener("click", ()=> showPage(b.dataset.page));
  });

  // Buttons
  $("btnRecalc").addEventListener("click", calc);
  $("btnSave").addEventListener("click", save);
  $("btnExport").addEventListener("click", exportJSON);
  $("btnImport").addEventListener("click", importJSON);
  $("btnClear").addEventListener("click", clearAll);
  $("btnPrint").addEventListener("click", ()=>{ showPage("p_outputs"); setTimeout(()=>window.print(), 150); });

  // Tables
  $("addRR").addEventListener("click", addRR);
  $("sumRR").addEventListener("click", sumRR);
  $("addGlob").addEventListener("click", addGlob);
  $("sumGlob").addEventListener("click", sumGlob);
  $("addDraw").addEventListener("click", addDraw);
  $("evenDraw").addEventListener("click", evenDraw);

  // Product change -> defaults
  $("loanProduct").addEventListener("change", (e)=>{
    applyDefaults(e.target.value);
    calc();
  });

  // Auto recalc on input changes (debounced)
  let t=null;
  qsa("input,select,textarea").forEach(el=>{
    el.addEventListener("input", ()=>{
      clearTimeout(t);
      t=setTimeout(calc, 120);
    });
    el.addEventListener("change", ()=>{
      clearTimeout(t);
      t=setTimeout(calc, 120);
    });
  });
}

function boot(){
  // Default dates
  if(!$("uwDate").value) $("uwDate").value = todayISO();
  if(!$("closeDate").value) $("closeDate").value = "";

  // Load saved state if present
  const loaded = load();
  if(!loaded){
    applyDefaults($("loanProduct").value);
  }

  // Render empty tables if needed
  renderRR(); renderGlob(); renderDraws();

  // Wire events and calc
  wire();
  calc();
}

boot();
/* ===== Inputs vs Outputs Mode ===== */
(function () {
  const MODE_KEY = "uo_mode";

  const INPUT_PAGES = new Set([
    "p_deal",
    "p_structure",
    "p_sources",
    "p_property",
    "p_hist",
    "p_proforma",
    "p_construction",
    "p_sponsor",
    "p_global",
    "p_stress",
    "p_legal",
    "p_rating",
    "p_conditions"
  ]);

  const OUTPUT_PAGES = new Set([
    "p_outputs",
    "p_feedback"
  ]);

  function setMode(mode) {
    mode = (mode === "outputs") ? "outputs" : "inputs";
    const isInputs = mode === "inputs";

    // Badge + button styling
    const badge = document.getElementById("modeBadge");
    const bIn = document.getElementById("btnModeInputs");
    const bOut = document.getElementById("btnModeOutputs");
    if (badge) badge.textContent = "Mode: " + (isInputs ? "Inputs" : "Outputs");
    if (bIn) bIn.classList.toggle("active", isInputs);
    if (bOut) bOut.classList.toggle("active", !isInputs);

    // Sidebar: hide/show ONLY the real page nav buttons
    document.querySelectorAll("#nav button[data-page]").forEach((btn) => {
      const pageId = btn.getAttribute("data-page");
      const show = isInputs ? INPUT_PAGES.has(pageId) : OUTPUT_PAGES.has(pageId);
      btn.style.display = show ? "" : "none";
      btn.classList.remove("active");
    });

    // Pages: hide/show by ID
    document.querySelectorAll("section.page").forEach((sec) => {
      const show = isInputs ? INPUT_PAGES.has(sec.id) : OUTPUT_PAGES.has(sec.id);
      sec.style.display = show ? "" : "none";
      sec.classList.remove("active");
    });

    // Force a valid page visible in this mode
    const target = isInputs ? "p_deal" : "p_outputs";
    const btn = document.querySelector(`#nav button[data-page="${target}"]`);
    if (btn) {
      btn.classList.add("active");
    }
    const page = document.getElementById(target);
    if (page) {
      page.classList.add("active");
      page.style.display = "";
    }

    // Persist
    try { localStorage.setItem(MODE_KEY, mode); } catch (e) {}
  }

  document.addEventListener("DOMContentLoaded", function () {
    const bIn = document.getElementById("btnModeInputs");
    const bOut = document.getElementById("btnModeOutputs");

    if (bIn) bIn.addEventListener("click", () => setMode("inputs"));
    if (bOut) bOut.addEventListener("click", () => setMode("outputs"));

    let saved = "inputs";
    try { saved = localStorage.getItem(MODE_KEY) || "inputs"; } catch (e) {}
    setMode(saved);
  });
})();
