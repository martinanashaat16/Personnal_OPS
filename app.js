/* PersonnelOps — main application logic */
/* ─── DATA ─── */
const TASKS=[
  {id:'p1',phase:'Preparations',label:'Onboarding on Apex',tag:''},
  {id:'p2',phase:'Preparations',label:'Signed the contract',tag:''},
  {id:'p3',phase:'Preparations',label:'Bus Arranged',tag:''},
  {id:'p4',phase:'Preparations',label:'Onboarding Email',tag:''},
  {id:'p5',phase:'Preparations',label:'Documents Completed',tag:''},
  {id:'p6',phase:'Preparations',label:'Data Submitted on Oracle',tag:''},
  {id:'p7',phase:'Preparations',label:'Filled Bank application to open an account',tag:''},
  {id:'p8',phase:'Preparations',label:'Open Time Management on apex',tag:''},
  {id:'p9',phase:'Preparations',label:'Received Access Card',tag:''},
  {id:'p10',phase:'Preparations',label:'File got archived',tag:''},
  {id:'p11',phase:'Preparations',label:'Welcome Onboard Announcement',tag:''},
  {id:'p12',phase:'Preparations',label:'Induction Session',tag:'',noCount:true},
  {id:'p13',phase:'Preparations',label:'Laptop',tag:'',type:'eligibility'},
  {id:'p14',phase:'Preparations',label:'Mobile Line',tag:'',type:'eligibility'},
  {id:'p15',phase:'Preparations',label:'Policies Acknowledged',tag:''},
  {id:'p16',phase:'Preparations',label:'Received Medical Card',tag:''},
  {id:'p17',phase:'Preparations',label:'Received Visa Card',tag:''},
  {id:'p18',phase:'Preparations',label:'Attend the Quarterly Onboarding session',tag:'',type:'yesno'},
];
// Tasks flagged noCount (Induction Session) stay on the checklist but are left out of every % and x/y count
const COUNTED_TASKS=TASKS.filter(t=>!t.noCount);
const PHASES=['Preparations'];
const PICONS={
  'Preparations':'<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"/></svg>',
};

let users     = JSON.parse(localStorage.getItem('hr3_users')||'[]');
let hires     = JSON.parse(localStorage.getItem('hr3_hires')||'[]');
let contracts = JSON.parse(localStorage.getItem('hr3_contracts')||'[]');
let offboards = JSON.parse(localStorage.getItem('hr3_offboards')||'[]');
let obExtraCols = JSON.parse(localStorage.getItem('hr3_ob_extra_cols')||'[]'); // [{key,label}] extra columns discovered from uploads
let overtimeRecords = JSON.parse(localStorage.getItem('hr3_overtime')||'[]'); // [{id,name,dept,code,date,hours,reason}]
let otValueNumFmt = localStorage.getItem('hr3_ot_value_fmt') || null; // Excel number-format code (SheetJS 'z') detected from the Overtime column, e.g. "#,##0.00" or "$#,##0"
let expanded  = JSON.parse(localStorage.getItem('hr3_expanded')||'{}');
let alertsSent= JSON.parse(localStorage.getItem('hr3_alerts_sent')||'{}');
/* ─── BULK SELECT STATE (Contract Renewals) — selection is always on: checkboxes
   are shown on every row so multiple employees can be ticked and then a single
   click on "Renewed" (or "Remove Selected") applies to all of them at once ─── */
let ctSelected=new Set();
/* ─── CONTRACT RENEWALS: SEARCH / DATE FILTER STATE ─── */
let ctSearchQuery='', ctFilterDate='';
let hSelectMode=false, hSelected=new Set();
let obSelectMode=false, obSelected=new Set(); let obColSelectMode=false, obColSelected=new Set();
let curView   = 'monthly';
let curHire   = null;

function save()  { localStorage.setItem('hr3_hires',JSON.stringify(hires)); sbPush('hr3_hires',hires); }
function saveC() { localStorage.setItem('hr3_contracts',JSON.stringify(contracts)); localStorage.setItem('hr3_alerts_sent',JSON.stringify(alertsSent)); sbPush('hr3_contracts',contracts); sbPush('hr3_alerts_sent',alertsSent); }
function saveOb(){ localStorage.setItem('hr3_offboards',JSON.stringify(offboards)); sbPush('hr3_offboards',offboards); }
function saveObExtraCols(){ localStorage.setItem('hr3_ob_extra_cols',JSON.stringify(obExtraCols)); sbPush('hr3_ob_extra_cols',obExtraCols); }
function saveOt(){ localStorage.setItem('hr3_overtime',JSON.stringify(overtimeRecords)); sbPush('hr3_overtime',overtimeRecords); }
function migrateResignations(){
  // One-time migration: fold any legacy Resignation Analysis data into Offboarding records
  const raw=localStorage.getItem('hr3_resignations');
  if(!raw) return;
  try{
    const legacy=JSON.parse(raw);
    let migrated=0;
    (legacy||[]).forEach(r=>{
      if(!r.name||!r.lastDay) return;
      if(offboards.find(o=>o.name===r.name&&o.lastDay===r.lastDay)) return;
      offboards.push({
        id:'o'+Date.now()+Math.random(),
        name:r.name, lastDay:r.lastDay, dept:r.dept||'', code:r.code||'', reason:r.reason||'Resignation',
        tasks:obDefaultTasks(),
      });
      migrated++;
    });
    if(migrated>0){ saveOb(); toast(`✓ Migrated ${migrated} record${migrated!==1?'s':''} from the old Resignation Analysis page into Offboarding`,'purple',10000); }
  }catch(e){ /* ignore malformed legacy data */ }
  localStorage.removeItem('hr3_resignations');
}
function saveEx(){ localStorage.setItem('hr3_expanded',JSON.stringify(expanded)); } // per-device UI state only — not synced
function saveUsers(){ localStorage.setItem('hr3_users',JSON.stringify(users)); sbPush('hr3_users',users); }

/* ═══════════════ SHARED EXCEL DATABASE (OneDrive) ═══════════════
   Keeps the same JSON data sets the app already holds in localStorage inside ONE
   shared .xlsx workbook that lives in a OneDrive-synced folder. Every user points
   the app at that same file (File System Access API — Chrome / Edge); OneDrive
   moves the file between PCs. Writes are read-merge-write by record id, so two
   people editing different records at the same time do not overwrite each other.
   Workbook layout:  _AppData (hidden, source of truth)  +  README / Hires /
   Contracts / Offboarding / Overtime (read-only reports for humans).
   UI-only state (expanded panels, theme) intentionally stays per device. */
const SB_SYNC_KEYS=['hr3_users','hr3_hires','hr3_contracts','hr3_offboards','hr3_ob_extra_cols','hr3_overtime','hr3_alerts_sent','hr3_settings'];
const SB_EMPTY_VALUE={hr3_users:[],hr3_hires:[],hr3_contracts:[],hr3_offboards:[],hr3_ob_extra_cols:[],hr3_overtime:[],hr3_alerts_sent:{},hr3_settings:{}};
const SB_SETTING_KEYS=['hr3_ot_value_fmt','hr3_ot_std_hours','hr3_ot_mult_normal','hr3_ot_mult_sat','hr3_ot_mult_official'];
const SB_SHEET='_AppData', SB_CHUNK=30000, SB_POLL_MS=5000;
let sbConnected=false, sbApplyingRemote=false, sbPollTimer=null, sbHandle=null, sbPendingHandle=null;
let sbNeedsPerm=false, sbFileMod=0, sbSnapshot={}, sbDirty=new Set(), sbChain=Promise.resolve(), sbFlushTimer=null;
let sbLastOk=null, sbLastError=null, sbFileName='';

/* ── per-device handle + last-synced snapshot, kept in IndexedDB ── */
function sbIdb(){ return new Promise((res,rej)=>{ const r=indexedDB.open('hr3_shared',1); r.onupgradeneeded=()=>r.result.createObjectStore('kv'); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
async function sbIdbGet(k){ const db=await sbIdb(); return new Promise((res,rej)=>{ const q=db.transaction('kv').objectStore('kv').get(k); q.onsuccess=()=>res(q.result); q.onerror=()=>rej(q.error); }); }
async function sbIdbSet(k,v){ const db=await sbIdb(); return new Promise((res,rej)=>{ const t=db.transaction('kv','readwrite'); t.objectStore('kv').put(v,k); t.oncomplete=()=>res(); t.onerror=()=>rej(t.error); }); }
async function sbIdbDel(k){ const db=await sbIdb(); return new Promise((res,rej)=>{ const t=db.transaction('kv','readwrite'); t.objectStore('kv').delete(k); t.oncomplete=()=>res(); t.onerror=()=>rej(t.error); }); }
function sbSaveSnapshot(){ return sbIdbSet('snapshot',JSON.parse(JSON.stringify(sbSnapshot))).catch(()=>{}); }

/* ── local state access ── */
function sbGetLocal(key){
  if(key==='hr3_users') return users;
  if(key==='hr3_hires') return hires;
  if(key==='hr3_contracts') return contracts;
  if(key==='hr3_offboards') return offboards;
  if(key==='hr3_ob_extra_cols') return obExtraCols;
  if(key==='hr3_overtime') return overtimeRecords;
  if(key==='hr3_alerts_sent') return alertsSent;
  if(key==='hr3_settings'){ const o={}; SB_SETTING_KEYS.forEach(k=>{ const v=localStorage.getItem(k); if(v!==null) o[k]=v; }); return o; }
  return SB_EMPTY_VALUE[key];
}
function saveSettings(){ sbPush('hr3_settings',sbGetLocal('hr3_settings')); }
function sbApplyRemoteKey(key,value){
  if(value===null||value===undefined) return;
  sbApplyingRemote=true;
  try{
    if(key==='hr3_settings'){
      SB_SETTING_KEYS.forEach(k=>{ if(k in value) localStorage.setItem(k,value[k]); else localStorage.removeItem(k); });
      otValueNumFmt=localStorage.getItem('hr3_ot_value_fmt')||null;
    } else {
      localStorage.setItem(key,JSON.stringify(value));
      if(key==='hr3_users') users=value;
      else if(key==='hr3_hires') hires=value;
      else if(key==='hr3_contracts') contracts=value;
      else if(key==='hr3_offboards') offboards=value;
      else if(key==='hr3_ob_extra_cols') obExtraCols=value;
      else if(key==='hr3_overtime') overtimeRecords=value;
      else if(key==='hr3_alerts_sent') alertsSent=value;
    }
  } finally { sbApplyingRemote=false; }
}

/* ── 3-way merge by record id: base = last synced copy, local = this device, remote = file now ── */
function sbMerge(base,local,remote){
  const j=JSON.stringify;
  const idOf=o=>(o&&typeof o==='object')?(o.id!==undefined?o.id:o.key):undefined;
  function mergeMaps(b,l,r){
    const out=new Map();
    r.forEach((rv,id)=>{
      if(l.has(id)){ const lv=l.get(id); out.set(id,(!b.has(id)||j(b.get(id))!==j(lv))?lv:rv); }
      else if(!b.has(id)) out.set(id,rv);           // added by someone else
      // else: deleted on this device → stays deleted
    });
    l.forEach((lv,id)=>{
      if(!r.has(id) && (!b.has(id)||j(b.get(id))!==j(lv))) out.set(id,lv); // new here (or edited here after remote delete)
    });
    return out;
  }
  if(Array.isArray(local)&&Array.isArray(remote)){
    const all=[...local,...remote];
    if(!all.every(o=>idOf(o)!==undefined)) return local;
    const toMap=a=>new Map((a||[]).map(o=>[idOf(o),o]));
    return [...mergeMaps(toMap(Array.isArray(base)?base:[]),toMap(local),toMap(remote)).values()];
  }
  if(local&&remote&&typeof local==='object'&&typeof remote==='object'&&!Array.isArray(local)&&!Array.isArray(remote)){
    const toMap=o=>new Map(Object.entries(o||{}));
    return Object.fromEntries(mergeMaps(toMap(base&&typeof base==='object'&&!Array.isArray(base)?base:{}),toMap(local),toMap(remote)));
  }
  return local;
}

/* ── workbook read / write ── */
async function sbReadAll(){
  const f=await sbHandle.getFile();
  sbFileMod=f.lastModified;
  const out={};
  if(!f.size) return out;
  const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});
  const ws=wb.Sheets[SB_SHEET];
  if(!ws) return out;
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const parts={}, upd={};
  rows.slice(1).forEach(r=>{ const k=String(r[0]||''); if(!k) return; (parts[k]=parts[k]||[])[Number(r[1])||0]=String(r[2]===undefined?'':r[2]); upd[k]=String(r[3]||''); });
  Object.keys(parts).forEach(k=>{ try{ out[k]={value:JSON.parse(parts[k].join('')),updated_at:upd[k]}; }catch(e){ console.warn('Shared Excel: could not parse',k,e); } });
  return out;
}
function sbView(arr,extra){
  const rows=(arr||[]).map(rec=>{
    const o={};
    Object.keys(rec||{}).forEach(k=>{ const v=rec[k]; if(k!=='id' && k!=='sourceHireId' && (v===null||['string','number','boolean'].includes(typeof v))) o[k]=v; });
    if(extra) Object.assign(o,extra(rec));
    return o;
  });
  return rows.length?XLSX.utils.json_to_sheet(rows):XLSX.utils.aoa_to_sheet([['(no records yet)']]);
}
function sbBuildWorkbook(map){
  const val=k=>(map[k]&&map[k].value)||SB_EMPTY_VALUE[k];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
    ['PersonnelOps — shared database'],[''],
    ['This workbook is read and written by the PersonnelOps app. Please do not rename it, and do not edit the hidden "_AppData" sheet.'],
    ['The Hires / Contracts / Offboarding / Overtime sheets are read-only reports refreshed on every save — edits made here are NOT read back by the app.'],
    ['Last saved by the app: '+new Date().toLocaleString()]
  ]),'README');
  XLSX.utils.book_append_sheet(wb,sbView(val('hr3_hires'),h=>{ try{ return {progressPct:pct(h)}; }catch(e){ return {}; } }),'Hires');
  XLSX.utils.book_append_sheet(wb,sbView(val('hr3_contracts')),'Contracts');
  XLSX.utils.book_append_sheet(wb,sbView(val('hr3_offboards')),'Offboarding');
  XLSX.utils.book_append_sheet(wb,sbView(val('hr3_overtime')),'Overtime');
  const aoa=[['key','part','json','updated_at']];
  Object.keys(map).forEach(k=>{
    const s=JSON.stringify(map[k].value);
    for(let i=0,p=0;i<Math.max(s.length,1);i+=SB_CHUNK,p++) aoa.push([k,p,s.slice(i,i+SB_CHUNK),map[k].updated_at||'']); // Excel caps a cell at 32,767 chars
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),SB_SHEET);
  wb.Workbook={Sheets:wb.SheetNames.map(n=>({Hidden:n===SB_SHEET?1:0}))};
  return wb;
}
async function sbWriteAll(map){
  const buf=XLSX.write(sbBuildWorkbook(map),{type:'array',bookType:'xlsx'});
  const w=await sbHandle.createWritable();
  try{ await w.write(buf); await w.close(); }
  catch(e){ try{ await w.abort(); }catch(_){} throw e; }
  sbFileMod=(await sbHandle.getFile()).lastModified;
}

/* ── queue, push, pull ── */
function sbQueue(fn){ const p=sbChain.then(fn); sbChain=p.catch(()=>{}); return p; }
function sbPush(key,value){
  if(!sbConnected||sbApplyingRemote) return;
  sbDirty.add(key);
  clearTimeout(sbFlushTimer);
  sbFlushTimer=setTimeout(sbFlushDirty,400);
}
async function sbFlushDirty(){
  if(!sbConnected||!sbDirty.size) return;
  const keys=[...sbDirty]; sbDirty.clear();
  try{
    await sbQueue(async()=>{
      const remote=await sbReadAll();
      let changed=false;
      for(const k of keys){
        const local=sbGetLocal(k), rv=remote[k]?remote[k].value:undefined;
        const merged=rv===undefined?local:sbMerge(sbSnapshot[k],local,rv);
        if(JSON.stringify(merged)!==JSON.stringify(rv)){ remote[k]={value:merged,updated_at:new Date().toISOString()}; changed=true; }
        if(JSON.stringify(merged)!==JSON.stringify(local)) sbApplyRemoteKey(k,merged);
        sbSnapshot[k]=merged;
      }
      if(changed) await sbWriteAll(remote);
      await sbSaveSnapshot();
    });
    sbLastOk=new Date(); sbLastError=null;
  }catch(e){
    keys.forEach(k=>sbDirty.add(k));
    sbLastError=e; console.warn('Shared Excel: save failed, will retry',e);
    setTimeout(sbFlushDirty,8000);
  }
  renderSyncStatus();
}
let sbRefreshTimer=null;
function refreshAllViewsSoon(){ clearTimeout(sbRefreshTimer); sbRefreshTimer=setTimeout(()=>{ if(typeof appStarted!=='undefined'&&appStarted&&!sbDirty.size) refreshAllViews(); },50); }
async function sbPullMerge(){
  const remote=await sbReadAll();
  let touched=false;
  for(const k of SB_SYNC_KEYS){
    if(!(k in remote)) continue;
    const rv=remote[k].value, local=sbGetLocal(k);
    const merged=sbSnapshot[k]===undefined?sbMerge(undefined,local,rv):sbMerge(sbSnapshot[k],local,rv);
    if(JSON.stringify(merged)!==JSON.stringify(local)){ sbApplyRemoteKey(k,merged); touched=true; }
    if(JSON.stringify(merged)!==JSON.stringify(rv)) sbDirty.add(k);
    sbSnapshot[k]=rv;
  }
  await sbSaveSnapshot();
  return touched;
}
async function sbPoll(){
  if(!sbConnected) return;
  try{
    const f=await sbHandle.getFile();
    if(f.lastModified===sbFileMod) return;
    const touched=await sbQueue(sbPullMerge);
    sbLastOk=new Date(); sbLastError=null;
    if(touched && typeof appStarted!=='undefined' && appStarted) refreshAllViews();
    if(sbDirty.size) sbFlushDirty();
    renderSyncStatus();
  }catch(e){ sbLastError=e; console.warn('Shared Excel: poll failed —',e); renderSyncStatus(); }
}
function sbStartPolling(){ sbStopPolling(); sbPollTimer=setInterval(sbPoll,SB_POLL_MS); }
function sbStopPolling(){ if(sbPollTimer){ clearInterval(sbPollTimer); sbPollTimer=null; } }
document.addEventListener('visibilitychange',()=>{ if(!document.hidden) sbPoll(); });
window.addEventListener('beforeunload',e=>{ if(sbDirty.size){ e.preventDefault(); e.returnValue=''; } });

/* ── connect / disconnect ── */
function sbHasAPI(){ return 'showOpenFilePicker' in window && 'showSaveFilePicker' in window; }
const SB_PICKER_TYPES=[{description:'Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}];
async function sbCreateFile(){
  if(!sbHasAPI()){ toast('This browser cannot open shared files. Please use Chrome or Edge on a PC.','danger',0); return; }
  try{ const h=await showSaveFilePicker({suggestedName:'HR_Data.xlsx',types:SB_PICKER_TYPES}); await sbAttach(h,true); }
  catch(e){ if(e&&e.name!=='AbortError') toast('Could not create the file: '+(e.message||e),'danger',0); }
}
async function sbPickFile(){
  if(!sbHasAPI()){ toast('This browser cannot open shared files. Please use Chrome or Edge on a PC.','danger',0); return; }
  try{ const [h]=await showOpenFilePicker({types:SB_PICKER_TYPES,multiple:false}); await sbAttach(h,true); }
  catch(e){ if(e&&e.name!=='AbortError') toast('Could not open the file: '+(e.message||e),'danger',0); }
}
async function sbAttach(handle,fresh,silent){
  try{
    if(handle.requestPermission && (await handle.queryPermission({mode:'readwrite'}))!=='granted'){
      if((await handle.requestPermission({mode:'readwrite'}))!=='granted') throw new Error('Permission to edit the file was not granted');
    }
    sbHandle=handle; sbFileName=handle.name;
    const remote=await sbReadAll();
    if(fresh){
      sbSnapshot={};
      const hasRemote=SB_SYNC_KEYS.some(k=>k in remote);
      const localHas=SB_SYNC_KEYS.some(k=>{ const v=sbGetLocal(k); return Array.isArray(v)?v.length>0:Object.keys(v||{}).length>0; });
      if(hasRemote && localHas && !confirm('"'+handle.name+'" already contains shared data.\n\nConnecting will REPLACE the data currently stored in this browser with the data from that file.\n\nContinue?')){ sbHandle=null; return false; }
      SB_SYNC_KEYS.forEach(k=>{ if(k in remote) sbApplyRemoteKey(k,remote[k].value); });
    } else {
      sbSnapshot=(await sbIdbGet('snapshot'))||{};
    }
    SB_SYNC_KEYS.forEach(k=>{
      const local=sbGetLocal(k);
      if(k in remote){
        const rv=remote[k].value;
        const merged=fresh?rv:sbMerge(sbSnapshot[k],local,rv);
        if(JSON.stringify(merged)!==JSON.stringify(local)) sbApplyRemoteKey(k,merged);
        if(JSON.stringify(merged)!==JSON.stringify(rv)) sbDirty.add(k);
        sbSnapshot[k]=rv;
      } else { sbDirty.add(k); delete sbSnapshot[k]; }   // not in the file yet → seed it from this device
    });
    sbConnected=true; sbNeedsPerm=false; sbPendingHandle=null; sbLastError=null;
    await sbIdbSet('handle',handle);
    await sbSaveSnapshot();
    await sbFlushDirty();
    sbStartPolling();
    sbLastOk=new Date();
    renderSyncStatus();
    if(typeof appStarted!=='undefined' && appStarted) refreshAllViews();
    if(!silent) toast('✓ Connected — data is now shared through "'+handle.name+'"');
    return true;
  }catch(e){
    sbConnected=false; sbHandle=null; sbLastError=e;
    renderSyncStatus();
    if(!silent) toast('Could not connect: '+(e.message||e),'danger',0); else console.warn('Shared Excel: auto-connect failed —',e);
    return false;
  }
}
async function sbReconnect(){
  if(!sbPendingHandle){ openSyncModal(); return; }
  await sbAttach(sbPendingHandle,false,false);
}
async function sbSyncNow(){
  try{ await sbQueue(sbPullMerge); await sbFlushDirty(); refreshAllViews(); toast('✓ Synced'); }
  catch(e){ toast('Sync failed: '+(e.message||e),'danger'); }
}
async function sbDisconnect(){
  sbStopPolling();
  sbConnected=false; sbHandle=null; sbPendingHandle=null; sbNeedsPerm=false; sbSnapshot={}; sbDirty.clear(); sbFileName='';
  await sbIdbDel('handle'); await sbIdbDel('snapshot');
  renderSyncStatus();
  toast('Disconnected — this device now uses local data only');
}

/* ── UI ── */
function openSyncModal(){ document.getElementById('sync-modal-bg').classList.add('open'); renderSyncStatus(); }
function closeSyncModal(){ document.getElementById('sync-modal-bg').classList.remove('open'); }
function renderSyncStatus(){
  const box=document.getElementById('sync-status-box');
  const disBtn=document.getElementById('sync-disconnect-btn');
  const nowBtn=document.getElementById('sync-now-btn');
  const badge=document.getElementById('sync-nav-badge');
  const banner=document.getElementById('sync-banner');
  if(banner) banner.style.display=sbNeedsPerm?'flex':'none';
  if(!box) return;
  const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
  if(sbConnected){
    const err=sbLastError?`<div class="ab-row" style="color:var(--red-text)">⚠ Last problem: ${esc(sbLastError.message||sbLastError)} — retrying automatically (is the file open in Excel?)</div>`:'';
    box.innerHTML=`<div class="ab" style="background:var(--green-bg);border:1px solid #A7F3D0">
      <div class="ab-head" style="color:var(--green-text)">✅ Connected — sharing data via Excel</div>
      <div class="ab-row" style="color:var(--green-text)">File: <b>${esc(sbFileName)}</b></div>
      <div class="ab-row" style="color:var(--green-text)">Changes from other users appear within ~${SB_POLL_MS/1000} seconds. Last synced: ${sbLastOk?sbLastOk.toLocaleTimeString():'—'}</div>${err}
    </div>`;
    if(badge){ badge.style.display='inline-flex'; badge.style.background='var(--green-bg)'; badge.style.color='var(--green-text)'; }
  } else if(sbNeedsPerm){
    box.innerHTML=`<div class="ab ab-warn"><div class="ab-head">Disconnected — permission needed</div>
      <div class="ab-row">This device is linked to the shared file, but the browser needs your OK again for this session.</div>
      <div style="margin-top:8px"><button class="btn btn-primary btn-sm" onclick="sbReconnect()">Reconnect</button></div></div>`;
    if(badge){ badge.style.display='inline-flex'; badge.style.background='var(--amber-bg)'; badge.style.color='var(--amber-text)'; }
  } else {
    const warn=sbHasAPI()?'':'<div class="ab-row" style="color:var(--red-text)">This browser does not support shared files — please use Chrome or Edge.</div>';
    const err=sbLastError?`<div class="ab-row" style="color:var(--red-text)">⚠ ${esc(sbLastError.message||sbLastError)}</div>`:'';
    box.innerHTML=`<div class="ab ab-warn"><div class="ab-head">Not connected</div>
      <div class="ab-row">This device is only using data stored in its own browser. Connect below to share data with other users.</div>${warn}${err}</div>`;
    if(badge) badge.style.display='none';
  }
  if(disBtn) disBtn.style.display=(sbConnected||sbNeedsPerm)?'inline-flex':'none';
  if(nowBtn) nowBtn.style.display=sbConnected?'inline-flex':'none';
}
function refreshAllViews(){
  renderSidebar(); renderMonthly(); renderDashboard(); renderContracts(); renderOffboarding(); renderOvertimeAnalysis();
  if(document.getElementById('pg-resignation')?.classList.contains('active')) renderResignationAnalysis();
}

/* ─── APP ENTRY ─── (login/signup removed — app opens directly) */
function enterApp(){
  initApp();
}

/* ─── UTILS ─── */
const ini = n => n.split(' ').slice(0,2).map(w=>w[0]?.toUpperCase()||'').join('');
const fmt = ds => ds ? new Date(ds+'T00:00:00').toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}) : '—';
const daysFrom = ds => { const d=new Date(ds+'T00:00:00'),n=new Date(); n.setHours(0,0,0,0); return Math.round((d-n)/86400000); };
const addDaysISO = (ds, days) => {
  if(!ds) return ds;
  const d = new Date(ds+'T00:00:00');
  d.setDate(d.getDate()+days);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
};
/* Contract Renewals: the effective renewal date used everywhere for display/calculations is the
   stored renewal date + 1 day. */
const ctRenewalDate = c => addDaysISO(c.renewalDate, 1);
/* Renewal dates uploaded in a spreadsheet's original "mmm-yy" cell format
   (e.g. "Jan-26") are shown exactly as uploaded; every other contract still
   displays its full computed date as before. */
const ctRenewalDisplay = c => c.renewalDateDisplay || fmt(ctRenewalDate(c));
/* Contract renewal deadline = 1 week (7 days) before the (effective) renewal date */
function contractDeadline(c){
  const rd=new Date(ctRenewalDate(c)+'T00:00:00');
  const dl=new Date(rd);
  dl.setDate(dl.getDate()-7);
  return `${dl.getFullYear()}-${String(dl.getMonth()+1).padStart(2,'0')}-${String(dl.getDate()).padStart(2,'0')}`;
}
const ctDeadlineDays = c => daysFrom(contractDeadline(c));
const today = () => new Date().toISOString().split('T')[0];
const weekKey = () => { const d=new Date(); return `${d.getFullYear()}-${d.getMonth()}-w${Math.ceil(d.getDate()/7)}`; };
const addYearsISO = (ds, years) => {
  if(!ds) return '';
  const d=new Date(ds+'T00:00:00');
  d.setFullYear(d.getFullYear()+years);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
};

/* ─── ONBOARDING → CONTRACT RENEWALS AUTO-LINK ───
   Every new hire with a joining date automatically gets a matching
   Contract Renewal record (Renewal Date = Joining Date + 1 year).
   Records are linked via sourceHireId so edits update in place and
   no duplicates are ever created for the same hire. */
function syncContractFromHire(h){
  if(!h || !h.startDate) return;
  const renewalDate=addYearsISO(h.startDate,1);
  let c=contracts.find(c=>c.sourceHireId===h.id);
  if(c){
    c.name=h.name; c.dept=h.dept||''; c.code=h.code||''; c.joiningDate=h.startDate; c.renewalDate=renewalDate;
  } else {
    contracts.push({
      id:'c'+Date.now()+Math.random(),
      name:h.name, dept:h.dept||'', code:h.code||'',
      joiningDate:h.startDate, renewalDate,
      renewed:false, renewedOn:null,
      sourceHireId:h.id
    });
  }
  saveC();
}
// One-time catch-up: link any pre-existing hires that don't have a contract record yet
hires.forEach(h=>{ if(h.startDate && !contracts.find(c=>c.sourceHireId===h.id)) syncContractFromHire(h); });

function taskDone(hire,t) {
  const v=hire.tasks&&hire.tasks[t.id];
  if(t.type==='yesno') return v==='yes'||v==='no';
  if(t.type==='eligibility'){
    const elig=hire.tasks&&hire.tasks[t.id+'_elig'];
    if(elig==='no') return true;                 // legacy: imported as Not Eligible from Excel
    return v==='received'||v==='notyet';
  }
  return !!v;
}
function pct(hire) {
  return Math.round((COUNTED_TASKS.filter(t=>taskDone(hire,t)).length/COUNTED_TASKS.length)*100);
}
function phasePct(hire,phase) {
  const ts=COUNTED_TASKS.filter(t=>t.phase===phase);
  if(!ts.length) return {done:0,total:0,p:0};
  const done=ts.filter(t=>taskDone(hire,t)).length;
  return {done,total:ts.length,p:Math.round(done/ts.length*100)};
}
function pillCls(p){ return p===100?'pl-green':p>=50?'pl-amber':p>0?'pl-red':'pl-gray'; }
function pillTxt(p){ return p===100?'✓ Complete':p>=50?'In progress':p>0?'Just started':'Not started'; }
function pill(p){ return `<span class="pill ${pillCls(p)}">${pillTxt(p)}</span>`; }
function barCls(p){ return p>=70?'hi':p>=30?'mid':'lo'; }

function getAlerts(hire) {
  if(!hire.startDate) return [];
  const days=-daysFrom(hire.startDate);
  const out=[];
  if(days>=2){
    const missing=COUNTED_TASKS.filter(t=>!taskDone(hire,t));
    if(missing.length) out.push({type:'danger',label:`${missing.length} task${missing.length>1?'s':''} unchecked — ${days}d since hiring`,detail:missing.map(t=>t.label)});
  }
  return out;
}
function allAlerts(){ return hires.flatMap(h=>getAlerts(h).map(a=>({...a,hire:h.name,hid:h.id}))); }

/* ─── NAVIGATION ─── */
function nav(v, hireId) {
  curView=v;
  if(hireId) curHire=hireId;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('pg-'+v).classList.add('active');
  ['dash','monthly','contracts','offboarding','resignation','overtime'].forEach(k=>{
    document.getElementById('nb-'+k)?.classList.remove('active');
  });
  if(v==='dashboard'||v==='detail') document.getElementById('nb-dash')?.classList.add('active');
  if(v==='monthly') document.getElementById('nb-monthly')?.classList.add('active');
  if(v==='contracts') document.getElementById('nb-contracts')?.classList.add('active');
  if(v==='offboarding') document.getElementById('nb-offboarding')?.classList.add('active');
  if(v==='resignation') document.getElementById('nb-resignation')?.classList.add('active');
  if(v==='overtime') document.getElementById('nb-overtime')?.classList.add('active');
  // Render the target page
  if(v==='dashboard') renderDashboard();
  else if(v==='detail') renderDetail();
  else if(v==='monthly') renderMonthly();
  else if(v==='contracts') renderContracts();
  else if(v==='offboarding') renderOffboarding();
  else if(v==='resignation') renderResignationAnalysis();
  else if(v==='overtime') renderOvertimeAnalysis();
  renderSidebar();
}

/* ─── TOAST ─── */
function toast(msg, type='', dur=6000) {
  const w=document.getElementById('toast-wrap');
  const id='t'+Date.now()+Math.random();
  const cls=type==='warn'?'tw':type==='danger'?'td':type==='purple'?'tp':'';
  const el=document.createElement('div');
  el.className=`toast ${cls}`;
  el.id=id;
  el.innerHTML=`<span style="flex:1">${msg}</span><button class="toast-x" onclick="dismissToast('${id}')">×</button>`;
  w.appendChild(el);
  if(dur>0) setTimeout(()=>dismissToast(id),dur);
}
function dismissToast(id){
  const el=document.getElementById(id);
  if(el){el.style.animation='tOut .22s ease forwards';setTimeout(()=>el.remove(),230);}
}

/* ─── NOTIFICATION ─── */
function sendNotif(title,body,tag){
  if('Notification'in window&&Notification.permission==='granted') new Notification(title,{body,tag});
}
function reqNotif(){ if('Notification'in window&&Notification.permission==='default') Notification.requestPermission(); }

/* ─── ANIMATED COUNTER ─── */
function setVal(id, val) {
  const el=document.getElementById(id);
  if(!el||el.textContent===String(val)) return;
  el.textContent=val;
  el.classList.remove('updated');
  void el.offsetWidth;
  el.classList.add('updated');
}

/* ─── SIDEBAR ─── */
function renderSidebar() {
  const alerts=allAlerts();
  const badge=document.getElementById('dash-badge');
  if(badge){ badge.textContent=alerts.length; badge.style.display=alerts.length?'inline':'none'; }

  const cs=ctStats();
  const cb=document.getElementById('ct-badge');
  if(cb){ cb.textContent=cs.pending; cb.style.display=cs.pending?'inline':'none'; }

  const os=obStats();
  const ob=document.getElementById('ob-badge');
  if(ob){ ob.textContent=os.overdue; ob.style.display=os.overdue?'inline':'none'; }

  const list=document.getElementById('s-list');
  if(!hires.length){ list.innerHTML='<div style="font-size:12px;color:var(--text3);padding:4px 8px">No hires yet</div>'; return; }
  list.innerHTML=hires.map(h=>{
    const p=pct(h), ha=getAlerts(h);
    const dot=ha.length?'background:var(--red)':p===100?'background:var(--green)':'background:var(--amber)';
    return `<div class="s-item${curHire===h.id&&curView==='detail'?' active':''}" onclick="nav('detail','${h.id}')">
      <div class="s-av">${ini(h.name)}</div>
      <div class="s-nm">${h.name}</div>
      ${ha.length?`<div class="s-dot pulse" style="${dot}"></div>`:''}
      <div class="s-pct">${p}%</div>
    </div>`;
  }).join('');
}

/* ─── DASHBOARD ─── */
/* ─── SHARED PERIOD FILTER (All · Monthly · Quarterly · Yearly) ───────────────
   Used by the Onboarding Tracker (filters on hire date) and Offboarding (filters on last working day).
   Same control as the Overtime page, but with CALENDAR quarters:
   Q1 = Jan–Mar · Q2 = Apr–Jun · Q3 = Jul–Sep · Q4 = Oct–Dec.  (Overtime keeps its own Dec–Feb quarters.) */
const PF_STATE={
  hire:{period:'all',month:null,quarter:null,year:null},
  ob:{period:'all',month:null,quarter:null,year:null},
};
const PF_RERENDER={ hire:()=>renderDashboard(), ob:()=>renderOffboarding() };
const PF_Q_MONTHS=['Jan–Mar','Apr–Jun','Jul–Sep','Oct–Dec'];
function pfQuarterKey(iso){ const [y,m]=iso.split('-'); return `${y}-Q${Math.ceil(Number(m)/3)}`; }
function pfMonthLabel(k){ const [y,m]=k.split('-'); return new Date(Number(y),Number(m)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'}); }
function pfQuarterLabel(k){ const [y,q]=k.split('-Q'); return `Q${q} ${y} (${PF_Q_MONTHS[Number(q)-1]})`; }
function pfMatch(scope,iso){
  const st=PF_STATE[scope];
  if(st.period==='all') return true;
  if(!iso) return false;
  iso=String(iso).slice(0,10);
  if(st.period==='monthly') return !!st.month && iso.slice(0,7)===st.month;
  if(st.period==='quarterly') return !!st.quarter && pfQuarterKey(iso)===st.quarter;
  return !!st.year && iso.slice(0,4)===st.year;
}
function pfSetPeriod(scope,period){
  PF_STATE[scope].period=period;
  pfClearSelection(scope);
  PF_RERENDER[scope]();
}
function pfSetSel(scope,kind,val){
  PF_STATE[scope][kind]=val;
  pfClearSelection(scope);
  PF_RERENDER[scope]();
}
function pfClearSelection(scope){ // a bulk selection must never silently include rows the filter has hidden
  if(scope==='hire' && typeof hSelected!=='undefined') hSelected.clear();
  if(scope==='ob' && typeof obSelected!=='undefined') obSelected.clear();
}
function pfRenderBar(scope,hostId,dates,shown,total){
  const host=document.getElementById(hostId); if(!host) return;
  const st=PF_STATE[scope];
  const months=new Set(), quarters=new Set(), years=new Set();
  dates.forEach(d=>{ if(!d) return; d=String(d).slice(0,10); months.add(d.slice(0,7)); quarters.add(pfQuarterKey(d)); years.add(d.slice(0,4)); });
  const M=[...months].sort().reverse(), Q=[...quarters].sort().reverse(), Y=[...years].sort().reverse();
  if(!M.includes(st.month)) st.month=M[0]||null;
  if(!Q.includes(st.quarter)) st.quarter=Q[0]||null;
  if(!Y.includes(st.year)) st.year=Y[0]||null;
  const opts=(arr,sel,lab)=>arr.length?arr.map(k=>`<option value="${k}" ${k===sel?'selected':''}>${lab(k)}</option>`).join(''):'<option value="">No data yet</option>';
  const seg=[['all','All'],['monthly','Monthly'],['quarterly','Quarterly'],['yearly','Yearly']];
  host.innerHTML=`<div class="ra-filterbar" style="margin-bottom:16px">
    <div class="ra-seg">${seg.map(([k,l])=>`<button class="ra-seg-btn${st.period===k?' active':''}" onclick="pfSetPeriod('${scope}','${k}')">${l}</button>`).join('')}</div>
    <select class="ra-select" style="${st.period==='monthly'?'':'display:none'}" onchange="pfSetSel('${scope}','month',this.value)">${opts(M,st.month,pfMonthLabel)}</select>
    <select class="ra-select" style="${st.period==='quarterly'?'':'display:none'}" onchange="pfSetSel('${scope}','quarter',this.value)">${opts(Q,st.quarter,pfQuarterLabel)}</select>
    <select class="ra-select" style="${st.period==='yearly'?'':'display:none'}" onchange="pfSetSel('${scope}','year',this.value)">${opts(Y,st.year,k=>k)}</select>
    <span class="ra-note">${st.period==='all'?`Showing all ${total}`:`Showing ${shown} of ${total}`}</span>
  </div>`;
}

function renderDashboard() {
  // Period filter (hire date). `shown` is what every part of this page works from.
  const shown=hires.filter(h=>pfMatch('hire',h.startDate));
  pfRenderBar('hire','h-period-filter',hires.map(h=>h.startDate),shown.length,hires.length);
  document.getElementById('d-sub').textContent=shown.length===hires.length
    ? `${hires.length} employee${hires.length!==1?'s':''} being onboarded`
    : `${shown.length} of ${hires.length} employees · filtered by hire date`;

  // Bulk select/remove toolbar
  const hToolbar=document.getElementById('h-toolbar');
  const hSelBtnLbl=document.getElementById('h-select-btn-lbl');
  const hSelBtn=document.getElementById('h-select-btn');
  if(hSelBtnLbl) hSelBtnLbl.textContent = hSelectMode ? 'Cancel selecting' : 'Select';
  if(hSelBtn){ hSelBtn.classList.toggle('btn-primary', hSelectMode); hSelBtn.classList.toggle('btn-outline', !hSelectMode); }
  if(hToolbar){
    hToolbar.innerHTML = hSelectMode ? `<div class="ab ab-purple" style="margin-bottom:14px;flex-direction:row;align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-weight:600;font-size:13px">${hSelected.size} selected</span>
      <button class="btn btn-outline btn-sm" onclick="hSelectAllVisible(${shown.length>0 && shown.every(h=>hSelected.has(h.id)) ? 'false':'true'})">${shown.length>0 && shown.every(h=>hSelected.has(h.id))?'Deselect all':'Select all'}</button>
      <button class="btn btn-danger btn-sm" ${hSelected.size?'':'disabled'} onclick="hRemoveSelected()">Remove Selected</button>
      <button class="btn btn-outline btn-sm" onclick="hToggleSelectMode()">Done</button>
      <span style="font-size:11.5px;color:var(--text3);margin-left:auto">Tick the employees you want to remove, then confirm.</span>
    </div>` : '';
  }

  // Alerts
  const shownIds=new Set(shown.map(h=>h.id));
  const alerts=allAlerts().filter(a=>shownIds.has(a.hid));
  const aw=document.getElementById('d-alerts');
  aw.innerHTML='';
  const danger=alerts.filter(a=>a.type==='danger');
  const warn=alerts.filter(a=>a.type==='warn');
  if(danger.length){
    const div=document.createElement('div');
    div.className='ab ab-danger';
    div.innerHTML=`<div class="ab-head"><svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" width="13" height="13"><path d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>Overdue tasks (2-day alert)</div>`+
      danger.map(a=>`<div class="ab-row">→ <strong>${a.hire}</strong>: ${a.label}</div>`).join('');
    aw.appendChild(div);
  }
  if(warn.length){
    const div=document.createElement('div');
    div.className='ab ab-warn';
    div.innerHTML=`<div class="ab-head"><svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" width="13" height="13"><path d="M15 17h5l-1.405-1.405A2.032 2.032 0 0118 14.158V11a6.002 6.002 0 00-4-5.659V5a2 2 0 10-4 0v.341C7.67 6.165 6 8.388 6 11v3.159c0 .538-.214 1.055-.595 1.436L4 17h5m6 0v1a3 3 0 11-6 0v-1m6 0H9"/></svg>Medical mail reminders (5-day alert)</div>`+
      warn.map(a=>`<div class="ab-row">→ <strong>${a.hire}</strong>: ${a.label}</div>`).join('');
    aw.appendChild(div);
  }

  // Hire list — update in-place
  const grid=document.getElementById('d-list');
  if(!hires.length){
    grid.innerHTML=`<div class="empty"><svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1.5"><path d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z"/></svg><div class="empty-title">No hires yet</div><div>Click "Add new hire" to get started</div></div>`;
    return;
  }
  if(!shown.length){
    grid.innerHTML=`<div class="empty"><div class="empty-title">No hires in this period</div><div>Try a different month, quarter or year — or choose "All".</div></div>`;
    return;
  }
  grid.querySelectorAll('.empty').forEach(el=>el.remove());
  shown.forEach(h => {
    const p=pct(h), ha=getAlerts(h);
    const existing=document.getElementById('hcard-'+h.id);
    const checked=hSelected.has(h.id);
    const html=`<div class="hc-row">
      ${hSelectMode?`<input type="checkbox" ${checked?'checked':''} onclick="event.stopPropagation()" onchange="hToggleRow('${h.id}')" style="margin-right:2px;width:16px;height:16px;flex-shrink:0">`:''}
      <div class="hc-av">${ini(h.name)}</div>
      <div style="flex:1;min-width:0"><div class="hc-name">${h.name}</div><div class="hc-meta">${h.title||'—'} · ${h.dept||'—'} · Hired ${fmt(h.startDate)}</div></div>
      ${ha.length?`<span class="pill pl-red">⚠ ${ha.length} alert${ha.length>1?'s':''}</span>`:pill(p)}
    </div>
    <div class="prog-bar"><div class="prog-fill ${barCls(p)}" style="width:${p}%"></div></div>
    <div class="prog-meta"><span>${COUNTED_TASKS.filter(t=>taskDone(h,t)).length}/${COUNTED_TASKS.length} tasks done</span><span>${p}%</span></div>`;
    let card=existing;
    if(!card){
      card=document.createElement('div');
      card.className='hire-card';
      card.id='hcard-'+h.id;
      grid.appendChild(card);
    }
    card.innerHTML=html;
    card.onclick = hSelectMode ? ()=>hToggleRow(h.id) : ()=>nav('detail',h.id);
    card.style.outline = (hSelectMode && checked) ? '2px solid var(--accent)' : '';
    card.style.cursor = 'pointer';
  });
  // Remove deleted cards
  [...grid.querySelectorAll('[id^="hcard-"]')].forEach(el=>{
    if(!shown.find(h=>'hcard-'+h.id===el.id)) el.remove();
  });
}

/* ─── DETAIL ─── */
function renderDetail() {
  const h=hires.find(h=>h.id===curHire); if(!h) return nav('dashboard');
  const p=pct(h);

  document.getElementById('det-av').textContent=ini(h.name);
  document.getElementById('det-name').textContent=h.name;
  document.getElementById('det-sub').textContent=`${h.title||'—'} · ${h.dept||'—'}`;
  document.getElementById('det-del').onclick=()=>deleteHire(h.id);

  // Info chips
  document.getElementById('det-chips').innerHTML=[
    ['Employee code',h.code||'—'],
    ['Hiring date',fmt(h.startDate),true],
    ['Days since hire',h.startDate?Math.max(0,-daysFrom(h.startDate)):'—'],
    ['Department',h.dept||'—'],
    ['HR person',h.hr||'—'],
  ].map(([l,v,editable])=>`<div class="chip"${editable?` style="cursor:pointer" title="Click to edit joining date" onclick="editJoiningDate('${h.id}')"`:''}><div class="chip-lbl">${l}${editable?' ✎':''}</div><div class="chip-val">${v}</div></div>`).join('');

  // Progress
  const done=COUNTED_TASKS.filter(t=>taskDone(h,t)).length;
  document.getElementById('det-pct-txt').textContent=`${p}% · ${done}/${COUNTED_TASKS.length} tasks`;
  const fill=document.getElementById('det-pct-bar');
  fill.style.width=p+'%';
  fill.style.background=p>=70?'var(--green)':p>=30?'var(--amber)':'var(--red)';

  // Alerts
  const alerts=getAlerts(h);
  const aw=document.getElementById('det-alerts');
  aw.innerHTML='';
  alerts.forEach(a=>{
    const div=document.createElement('div');
    div.className='ab ab-'+(a.type==='danger'?'danger':'warn');
    div.innerHTML=`<div class="ab-head">${a.label}</div>`+(a.detail||[]).map(d=>`<div class="ab-row" style="font-size:12px;padding-left:10px">• ${d}</div>`).join('');
    aw.appendChild(div);
  });

  // Phases — DOM-based, no full replace
  const container=document.getElementById('det-phases');
  // Remove any leftover phase sections from previously-viewed hires so the
  // checklist doesn't appear duplicated when switching between employees
  Array.from(container.children).forEach(child=>{
    if(!child.id.startsWith(`phase-${h.id}-`)) child.remove();
  });
  PHASES.forEach(phase=>{
    const ph_id=`phase-${h.id}-${phase.replace(/\W+/g,'_')}`;
    let sec=document.getElementById(ph_id);
    const expKey=h.id+phase;
    const isOpen=expanded[expKey]!==false;

    if(!sec){
      sec=document.createElement('div');
      sec.className='phase-sec';
      sec.id=ph_id;
      container.appendChild(sec);
    }

    const pp=phasePct(h,phase);
    const tasks=TASKS.filter(t=>t.phase===phase);

    // Header — always re-render counts
    let hdr=sec.querySelector('.phase-hdr');
    if(!hdr){
      hdr=document.createElement('div');
      hdr.className='phase-hdr';
      hdr.onclick=()=>togglePhase(h.id,phase);
      sec.appendChild(hdr);
    }
    hdr.innerHTML=`<div class="phase-title">${PICONS[phase]||''}${phase}</div>
      <div class="phase-right"><span class="phase-cnt">${pp.done}/${pp.total}</span>${pill(pp.p)}<svg class="chevron${isOpen?' open':''}" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M19 9l-7 7-7-7"/></svg></div>`;

    // Body
    let body=sec.querySelector('.phase-body');
    if(!body){ body=document.createElement('div'); body.className='phase-body'; sec.appendChild(body); }
    body.classList.toggle('open',isOpen);

    if(isOpen){
      tasks.forEach(t=>{
        const row_id=`tr-${h.id}-${t.id}`;
        let row=document.getElementById(row_id);
        const done=taskDone(h,t);
        if(!row){
          row=document.createElement('div');
          row.className='task-row';
          row.id=row_id;
          body.appendChild(row);
        }
        if(phase==='Preparations'){
          if(t.type==='yesno'){
            const v=h.tasks&&h.tasks[t.id];
            if(row.dataset.val!==String(v)){
              row.dataset.val=String(v);
              row.style.padding='10px 0';
              row.innerHTML=`<div style="font-size:13px">
                <div style="margin-bottom:5px">${t.label}</div>
                <div style="display:flex;gap:6px">
                  <button type="button" class="btn btn-sm ${v==='yes'?'btn-green':'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setTaskYesNo('${h.id}','${t.id}','yes')">Yes</button>
                  <button type="button" class="btn btn-sm ${v==='no'?'btn-danger':'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setTaskYesNo('${h.id}','${t.id}','no')">No</button>
                </div>
              </div>`;
            }
            return;
          }
          if(t.type==='eligibility'){
            const elig=h.tasks&&h.tasks[t.id+'_elig'];
            const v=h.tasks&&h.tasks[t.id];
            const stamp=`${elig}|${v}`;
            if(row.dataset.val!==stamp){
              row.dataset.val=stamp;
              row.style.padding='10px 0';
              row.innerHTML=`<div style="font-size:13px">
                <div style="margin-bottom:5px">${t.label}</div>
                <div style="display:flex;gap:6px">
                  <button type="button" class="btn btn-sm ${v==='received'?'btn-green':'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setTaskYesNo('${h.id}','${t.id}','received')">Received</button>
                  <button type="button" class="btn btn-sm ${v==='notyet'?'btn-danger':'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setTaskYesNo('${h.id}','${t.id}','notyet')">Not Yet</button>
                </div>
                ${elig==='no'?`<div style="font-size:11px;color:var(--text2);margin-top:4px">Marked Not Eligible in the Excel sheet</div>`:''}
                            </div>`;
            }
            return;
          }
          // Offboarding-style plain checkbox row, checkbox on the right
          const wasChecked=row.querySelector('input[type=checkbox]')?.checked;
          if(wasChecked!==done||!row.querySelector('input[type=checkbox]')){
            row.style.padding='10px 0';
            row.innerHTML=`<label style="display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:13px;cursor:pointer">
              <span>${t.label}${t.noCount?' <span style="font-size:11px;color:var(--text2)">(not counted in %)</span>':''}</span>
              <input type="checkbox" ${done?'checked':''} onchange="toggleTask('${h.id}','${t.id}')">
            </label>`;
          }
          return;
        }
        const note=h.notes&&h.notes[t.id]||'';
        // Only update DOM if state changed
        const wasChecked=row.querySelector('.chk')?.classList.contains('on');
        if(wasChecked!==done||!row.querySelector('.chk')){
          row.innerHTML=`<div class="chk${done?' on':''}" onclick="toggleTask('${h.id}','${t.id}')">
            ${done?'<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="3"><path d="M5 13l4 4L19 7"/></svg>':''}
          </div>
          <div class="task-body">
            <div class="task-lbl${done?' done':''}">${t.label}</div>
            ${t.tag&&t.tag!=='medical-alert'?`<span class="task-tag">${t.tag}</span>`:''}
            <textarea class="note-area" rows="1" placeholder="Add a note…" onchange="saveNote('${h.id}','${t.id}',this.value)">${note}</textarea>
          </div>`;
        }
      });
    }
  });
}

function togglePhase(hireId, phase) {
  const key=hireId+phase;
  expanded[key]=expanded[key]===false?true:false;
  saveEx();
  renderDetail();
}
function toggleTask(hireId, taskId) {
  const h=hires.find(h=>h.id===hireId); if(!h) return;
  if(!h.tasks) h.tasks={};
  h.tasks[taskId]=!h.tasks[taskId];
  save();
  renderDetail();
  renderSidebar();
  if(curView==='dashboard') renderDashboard();
}
function saveNote(hireId, taskId, val) {
  const h=hires.find(h=>h.id===hireId); if(!h) return;
  if(!h.notes) h.notes={};
  h.notes[taskId]=val; save();
}
function setTaskEligibility(hireId, taskId, val) {
  const h=hires.find(h=>h.id===hireId); if(!h) return;
  if(!h.tasks) h.tasks={};
  const key=taskId+'_elig';
  h.tasks[key]=(h.tasks[key]===val)?null:val;
  if(h.tasks[key]!=='yes') h.tasks[taskId]=null; // clear Received/Not Yet answer if not eligible
  save();
  renderDetail();
  renderSidebar();
  if(curView==='dashboard') renderDashboard();
}
function setTaskYesNo(hireId, taskId, val) {
  const h=hires.find(h=>h.id===hireId); if(!h) return;
  if(!h.tasks) h.tasks={};
  h.tasks[taskId]=(h.tasks[taskId]===val)?null:val;
  if((val==='received'||val==='notyet') && h.tasks[taskId]) h.tasks[taskId+'_elig']='yes';
  save();
  renderDetail();
  renderSidebar();
  if(curView==='dashboard') renderDashboard();
}

/* ─── MONTHLY ─── */
function getMonthly() {
  const map={};
  hires.forEach(h=>{
    if(!h.startDate) return;
    const d=new Date(h.startDate+'T00:00:00');
    const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const lbl=d.toLocaleDateString('en-GB',{month:'short',year:'numeric'});
    if(!map[key]) map[key]={key,lbl,count:0,total:0};
    map[key].count++;
    map[key].total+=pct(h);
  });
  return Object.values(map).sort((a,b)=>a.key.localeCompare(b.key)).map(m=>({...m,avg:m.count?Math.round(m.total/m.count):0}));
}

/* ─── QUARTERLY CONTRACT STATS ─── */
/* Standard calendar-quarter definitions (Q1 Jan–Mar, Q2 Apr–Jun, Q3 Jul–Sep, Q4 Oct–Dec) — used everywhere in the system */
const CAL_Q_DEFS = [
  { q:1, label:'Q1', range:'Jan – Mar', months:[1,2,3] },
  { q:2, label:'Q2', range:'Apr – Jun', months:[4,5,6] },
  { q:3, label:'Q3', range:'Jul – Sep', months:[7,8,9] },
  { q:4, label:'Q4', range:'Oct – Dec', months:[10,11,12] },
];
function calendarQuarter(dateStr) {
  // Returns { q, label, range, year, key } for a given YYYY-MM-DD, using standard calendar quarters
  const d = new Date(dateStr + 'T00:00:00');
  const m = d.getMonth() + 1;
  const y = d.getFullYear();
  const def = CAL_Q_DEFS.find(qd => qd.months.includes(m));
  return { ...def, year: y, key: `${y}-Q${def.q}` };
}
/* Custom fiscal-quarter definition used ONLY for Contract Renewals ("Breakdown by Quarter"):
   Q1 = Mar/Apr/May, Q2 = Jun/Jul/Aug, Q3 = Sep/Oct/Nov, Q4 = Dec/Jan/Feb.
   Q4 spans a calendar-year boundary, so Jan & Feb are attributed back to the fiscal
   year that started the previous December (e.g. Dec 2026, Jan 2027, Feb 2027 are all "FY2026 Q4"). */
const CUSTOM_CQ_DEFS = [
  { q:1, label:'Q1', range:'Mar – May', months:[3,4,5] },
  { q:2, label:'Q2', range:'Jun – Aug', months:[6,7,8] },
  { q:3, label:'Q3', range:'Sep – Nov', months:[9,10,11] },
  { q:4, label:'Q4', range:'Dec – Feb', months:[12,1,2] },
];
function contractQuarter(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const m = d.getMonth() + 1;
  const y = d.getFullYear();
  const def = CUSTOM_CQ_DEFS.find(qd => qd.months.includes(m));
  // Jan & Feb belong to the Q4 that began the previous December, so roll them back a fiscal year
  const fiscalYear = (m === 1 || m === 2) ? y - 1 : y;
  return { ...def, year: fiscalYear, key: `${fiscalYear}-Q${def.q}` };
}

function getQuarterlyContracts() {
  const map = {};
  contracts.forEach(c => {
    if(!c.renewalDate) return;
    const qInfo = contractQuarter(ctRenewalDate(c));
    if(!map[qInfo.key]) map[qInfo.key] = { ...qInfo, total:0, renewed:0, pending:0 };
    map[qInfo.key].total++;
    if(c.renewed) map[qInfo.key].renewed++;
    else map[qInfo.key].pending++;
  });
  // Sort by year then Q number
  return Object.values(map).sort((a,b) => a.year!==b.year ? a.year-b.year : a.q-b.q)
    .map(q => ({ ...q, rate: q.total ? Math.round(q.renewed/q.total*100) : 0 }));
}
function getYearlyContracts() {
  const map = {};
  contracts.forEach(c => {
    if(!c.renewalDate) return;
    const y = new Date(ctRenewalDate(c)+'T00:00:00').getFullYear();
    if(!map[y]) map[y] = { year:y, total:0, renewed:0, pending:0 };
    map[y].total++;
    if(c.renewed) map[y].renewed++;
    else map[y].pending++;
  });
  return Object.values(map).sort((a,b) => a.year-b.year)
    .map(y => ({ ...y, rate: y.total ? Math.round(y.renewed/y.total*100) : 0 }));
}

function getQuarterlyHires() {
  // Calendar-quarter grouping (Jan–Mar / Apr–Jun / Jul–Sep / Oct–Dec), applied to hire start dates
  const map = {};
  hires.forEach(h => {
    if(!h.startDate) return;
    const qInfo = calendarQuarter(h.startDate);
    if(!map[qInfo.key]) map[qInfo.key] = { ...qInfo, count:0, total:0 };
    map[qInfo.key].count++;
    map[qInfo.key].total += pct(h);
  });
  return Object.values(map).sort((a,b) => a.year!==b.year ? a.year-b.year : a.q-b.q)
    .map(q => ({ ...q, avg: q.count ? Math.round(q.total/q.count) : 0 }));
}

function getOffboardingMonthly() {
  const map={};
  offboards.forEach(o=>{
    if(!o.lastDay) return;
    const d=new Date(o.lastDay+'T00:00:00');
    const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const lbl=d.toLocaleDateString('en-GB',{month:'short',year:'numeric'});
    if(!map[key]) map[key]={key,lbl,total:0,resignations:0,overdue:0,taskDone:0,taskTotal:0};
    map[key].total++;
    if((o.reason||'').toLowerCase().includes('resign')) map[key].resignations++;
    if(obIsOverdue(o)) map[key].overdue++;
    OB_TASKS.forEach(t=>{ map[key].taskTotal++; if(obTaskDone(o,t)) map[key].taskDone++; });
  });
  return Object.values(map).sort((a,b)=>a.key.localeCompare(b.key))
    .map(m=>({...m,avg:m.taskTotal?Math.round(m.taskDone/m.taskTotal*100):0}));
}
function getQuarterlyOffboarding() {
  // Calendar-quarter grouping (Jan–Mar / Apr–Jun / Jul–Sep / Oct–Dec), same as Hires per Quarter
  const map = {};
  offboards.forEach(o => {
    if(!o.lastDay) return;
    const qInfo = calendarQuarter(o.lastDay);
    if(!map[qInfo.key]) map[qInfo.key] = { ...qInfo, total:0, resignations:0 };
    map[qInfo.key].total++;
    if((o.reason||'').toLowerCase().includes('resign')) map[qInfo.key].resignations++;
  });
  return Object.values(map).sort((a,b) => a.year!==b.year ? a.year-b.year : a.q-b.q);
}

function renderMonthly() {
  const stats=getMonthly();
  const totalHires=hires.length;
  const complete=hires.filter(h=>pct(h)===100).length;
  const avg=totalHires?Math.round(hires.reduce((s,h)=>s+pct(h),0)/totalHires):0;

  setVal('mm-total', totalHires);
  setVal('mm-done', complete);
  setVal('mm-rate', totalHires?Math.round(complete/totalHires*100)+'%':'0%');
  setVal('mm-avg', avg+'%');

  // Bar chart — now aggregated per quarter (Breakdown by Month grid below stays monthly, unchanged)
  const qHires=getQuarterlyHires();
  const maxCount=qHires.length?Math.max(...qHires.map(s=>s.count),1):1;
  const bars=document.getElementById('m-bars');
  bars.innerHTML=qHires.length?qHires.map(s=>{
    const h=Math.round((s.count/maxCount)*80);
    return `<div class="bc-col">
      <div class="bc-n">${s.count}</div>
      <div class="bc-b" style="height:${h}px"></div>
      <div class="bc-lbl">${s.label} ${s.year}</div>
    </div>`;
  }).join(''):'<div style="font-size:13px;color:var(--text3);padding:20px">No data yet — add hires with hiring dates</div>';

  // Month cards
  const grid=document.getElementById('m-grid');
  if(!stats.length){
    grid.innerHTML=`<div class="empty" style="grid-column:1/-1"><div class="empty-title">No hire data yet</div><div>Add hires with hiring dates to see monthly stats</div></div>`;
  } else {
    stats.forEach(s=>{
      const id='mc-'+s.key;
      let card=document.getElementById(id);
      if(!card){ card=document.createElement('div'); card.className='mcrd'; card.id=id; grid.appendChild(card); }
      const maxC=stats.length?Math.max(...stats.map(x=>x.count),1):1;
      card.innerHTML=`<div class="mc-lbl">${s.lbl}</div>
        <div class="mc-n c-blue">${s.count}</div>
        <div class="mc-sub">employee${s.count!==1?'s':''} hired</div>
        <div class="mc-bar-wrap"><div class="mc-bar" style="width:${Math.round(s.count/maxC*100)}%"></div></div>
        <div style="margin-top:7px;font-size:12px;color:var(--text2)">Avg completion: <strong style="color:${s.avg>=70?'var(--green)':s.avg>=30?'var(--amber)':'var(--red)'}">${s.avg}%</strong></div>`;
    });
    [...grid.querySelectorAll('[id^="mc-"]')].forEach(el=>{
      if(!stats.find(s=>'mc-'+s.key===el.id)) el.remove();
    });
  }

  /* ── QUARTERLY CONTRACTS SECTION ── */
  const qStats = getQuarterlyContracts();
  const cqTotal = contracts.length;
  const cqRenewed = contracts.filter(c=>c.renewed).length;
  const cqPending = contracts.filter(c=>!c.renewed).length;
  const cqRate = cqTotal ? Math.round(cqRenewed/cqTotal*100) : 0;

  setVal('cq-total', cqTotal);
  setVal('cq-renewed', cqRenewed);
  setVal('cq-pending', cqPending);
  setVal('cq-rate', cqRate+'%');

  // Year bar chart
  const yStats = getYearlyContracts();
  const cqBars = document.getElementById('cq-bars');
  if(!yStats.length){
    cqBars.innerHTML='<div style="font-size:13px;color:var(--text3);padding:20px">No contract data yet — add contracts or upload a CSV</div>';
  } else {
    const maxY = Math.max(...yStats.map(y=>y.total), 1);
    cqBars.innerHTML = yStats.map(y => {
      const hTotal = Math.round((y.total/maxY)*85);
      return `<div class="bc-col" style="position:relative">
        <div class="bc-n">${y.total}</div>
        <div style="position:relative;width:100%;height:${hTotal}px;background:var(--surface3);border-radius:5px 5px 0 0;overflow:hidden">
          <div style="position:absolute;bottom:0;left:0;right:0;height:${y.total?Math.round(y.renewed/y.total*100):0}%;background:linear-gradient(180deg,var(--green),#34D399);border-radius:5px 5px 0 0;transition:height .5s"></div>
        </div>
        <div class="bc-lbl">${y.year}</div>
      </div>`;
    }).join('');
  }

  // Quarter cards (same mcrd style as month cards)
  const cqGrid = document.getElementById('cq-grid');
  if(!qStats.length){
    cqGrid.innerHTML=`<div class="empty" style="grid-column:1/-1"><div class="empty-title">No contract data yet</div><div>Add contracts with renewal dates to see quarterly stats</div></div>`;
  } else {
    // Clear and re-render (quarter cards are few and static enough for full re-render)
    cqGrid.innerHTML = '';
    qStats.forEach(q => {
      const rateColor = q.rate===100?'var(--green)':q.rate>=50?'var(--amber)':'var(--red)';
      const pillCl = q.rate===100?'pl-green':q.rate>=50?'pl-amber':q.rate>0?'pl-red':'pl-gray';
      const pillTx = q.rate===100?'✓ All done':q.rate>=50?'In progress':q.rate>0?'Just started':'Not started';
      const card = document.createElement('div');
      card.className = 'mcrd';
      card.innerHTML = `
        <div class="mc-lbl">${q.label} ${q.year} · ${q.range}</div>
        <div class="mc-n" style="color:var(--accent)">${q.total}</div>
        <div class="mc-sub">contract${q.total!==1?'s':''} due</div>
        <div class="mc-bar-wrap" style="margin-top:8px;height:4px">
          <div class="mc-bar" style="width:${q.rate}%;background:${rateColor}"></div>
        </div>
        <div style="margin-top:8px;display:flex;flex-direction:column;gap:3px">
          <div style="font-size:12px;color:var(--text2)">Renewed: <strong style="color:var(--green)">${q.renewed}</strong> &nbsp;·&nbsp; Pending: <strong style="color:${q.pending>0?'var(--amber)':'var(--text3)'}">${q.pending}</strong></div>
          <div style="display:flex;align-items:center;justify-content:space-between;margin-top:2px">
            <span style="font-size:12px;color:var(--text2)">Completion</span>
            <span style="font-size:13px;font-weight:700;font-family:'JetBrains Mono',monospace;color:${rateColor}">${q.rate}%</span>
          </div>
          <span class="pill ${pillCl}" style="margin-top:2px;align-self:flex-start">${pillTx}</span>
        </div>`;
      cqGrid.appendChild(card);
    });
  }

  /* ── OFFBOARDING BY MONTH SECTION ── */
  const obStatsAll = obStats();
  const obMonthly = getOffboardingMonthly();
  const obPerfAll = obChecklistPerformance();

  setVal('ob-mm-total', obStatsAll.total);
  setVal('ob-mm-resign', obStatsAll.completed);
  setVal('ob-mm-overdue', obStatsAll.overdue);
  setVal('ob-mm-perf', obPerfAll+'%');

  const obBars = document.getElementById('ob-m-bars');
  const obQuarterly = getQuarterlyOffboarding();
  if(!obQuarterly.length){
    obBars.innerHTML='<div style="font-size:13px;color:var(--text3);padding:20px">No offboarding data yet — add offboarding records with a last working day</div>';
  } else {
    const maxObQ = Math.max(...obQuarterly.map(q=>q.total), 1);
    obBars.innerHTML = obQuarterly.map(q=>{
      const h=Math.round((q.total/maxObQ)*80);
      return `<div class="bc-col">
        <div class="bc-n">${q.total}</div>
        <div class="bc-b" style="height:${h}px"></div>
        <div class="bc-lbl">${q.label} ${q.year}</div>
      </div>`;
    }).join('');
  }

  const obGrid = document.getElementById('ob-m-grid');
  if(!obMonthly.length){
    obGrid.innerHTML=`<div class="empty" style="grid-column:1/-1"><div class="empty-title">No offboarding data yet</div><div>Add offboarding records with a last working day to see monthly stats</div></div>`;
  } else {
    obGrid.innerHTML = '';
    obMonthly.forEach(m=>{
      const rateColor = m.avg>=70?'var(--green)':m.avg>=30?'var(--amber)':'var(--red)';
      const card = document.createElement('div');
      card.className = 'mcrd';
      card.innerHTML = `
        <div class="mc-lbl">${m.lbl}</div>
        <div class="mc-n c-blue">${m.total}</div>
        <div class="mc-sub">offboarding record${m.total!==1?'s':''}</div>
        <div class="mc-bar-wrap"><div class="mc-bar" style="width:${m.avg}%;background:${rateColor}"></div></div>
        <div style="margin-top:8px;display:flex;flex-direction:column;gap:3px">
          <div style="font-size:12px;color:var(--text2)">Resignations: <strong style="color:var(--amber)">${m.resignations}</strong> &nbsp;·&nbsp; Overdue: <strong style="color:${m.overdue>0?'var(--red)':'var(--text3)'}">${m.overdue}</strong></div>
          <div style="display:flex;align-items:center;justify-content:space-between;margin-top:2px">
            <span style="font-size:12px;color:var(--text2)">Checklist performance</span>
            <span style="font-size:13px;font-weight:700;font-family:'JetBrains Mono',monospace;color:${rateColor}">${m.avg}%</span>
          </div>
        </div>`;
      obGrid.appendChild(card);
    });
  }
}

/* ─── CONTRACTS ─── */
function ctStats(){
  return{
    total:contracts.length,
    pending:contracts.filter(c=>!c.renewed).length,
    renewed:contracts.filter(c=>c.renewed).length,
    overdue:contracts.filter(c=>!c.renewed&&ctDeadlineDays(c)<0).length,
  };
}
function ctStatus(c){
  if(c.renewed) return 'renewed';
  const d=ctDeadlineDays(c);
  if(d<0) return 'overdue';
  if(d<=14) return 'urgent';
  if(d<=30) return 'soon';
  return 'ok';
}
function ctStatusPill(c){
  const s=ctStatus(c);
  const m={renewed:'pl-green',overdue:'pl-red',urgent:'pl-amber',soon:'pl-blue',ok:'pl-gray'};
  const t={renewed:'✓ Renewed',overdue:'Overdue',urgent:'Due soon',soon:'Upcoming',ok:'OK'};
  return `<span class="pill ${m[s]}">${t[s]}</span>`;
}

/* ─── SEARCH (by Employee Name) + MONTH FILTER (by Contract Renewal Date) ───
   Both filters are independent and can be combined; they only affect the
   Pending/Renewed listing tables below (search & selection), not the
   quarter alert, weekly preview, or any dashboard/quarterly calculations,
   which continue to reflect the full contract dataset. */
function getFilteredContracts(){
  return contracts.filter(c=>{
    const nameOk = !ctSearchQuery || (c.name||'').toLowerCase().includes(ctSearchQuery.trim().toLowerCase());
    const dateOk = !ctFilterDate || ctRenewalDate(c).slice(0,7)===ctFilterDate;
    return nameOk && dateOk;
  });
}
function ctSetSearch(v){ ctSearchQuery=v; renderContracts(); }
function ctSetDateFilter(v){ ctFilterDate=v; renderContracts(); }
function ctClearFilters(){
  ctSearchQuery=''; ctFilterDate='';
  const si=document.getElementById('ct-search-name'); if(si) si.value='';
  const di=document.getElementById('ct-filter-date'); if(di) di.value='';
  renderContracts();
}

function renderContracts(){
  // Search / date filter bar (inputs live outside this function's re-rendered regions so focus is preserved)
  const filtered=getFilteredContracts();
  const filtersActive = !!(ctSearchQuery || ctFilterDate);
  const clearBtn=document.getElementById('ct-clear-filters-btn');
  if(clearBtn) clearBtn.style.display = filtersActive ? '' : 'none';
  const filterSummary=document.getElementById('ct-filter-summary');
  if(filterSummary) filterSummary.textContent = filtersActive ? `${filtered.length} of ${contracts.length} contract${contracts.length!==1?'s':''} match` : '';

  // Bulk select/remove/renew toolbar — always visible; tick rows below, then act on all selected at once
  const ctToolbar=document.getElementById('ct-toolbar');
  if(ctToolbar){
    ctToolbar.innerHTML = contracts.length ? `<div class="ab ab-purple" style="margin-bottom:14px;flex-direction:row;align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-weight:600;font-size:13px">${ctSelected.size} employee${ctSelected.size!==1?'s':''} selected</span>
      <button class="btn btn-green btn-sm" ${ctSelected.size?'':'disabled'} onclick="ctBulkMarkRenewed()">Renewed</button>
      <button class="btn btn-danger btn-sm" ${ctSelected.size?'':'disabled'} onclick="ctRemoveSelected()">Remove Selected</button>
      <span style="font-size:11.5px;color:var(--text3);margin-left:auto">Tick one or more employees below, then click Renewed to mark them all at once.</span>
    </div>` : '';
  }
  // Quarter alert banner
  const qWrap=document.getElementById('c-quarter-alert');
  const qPending=getQuarterPendingContracts();
  if(qPending.length){
    const qLabel=getQuarter(new Date()).label;
    const qEnd=getQuarterEnd(new Date());
    const daysLeft=Math.round((qEnd-new Date())/86400000);
    qWrap.innerHTML=`<div class="ab ab-warn" style="margin-bottom:14px">
      <div class="ab-head">
        <svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" width="13" height="13"><path d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"/></svg>
        Quarter Pending Alert — ${qLabel}
        <button class="btn btn-outline btn-sm" style="margin-left:auto;font-size:11px;padding:3px 10px" onclick="sendQuarterAlert()">Send reminder</button>
      </div>
      <div class="ab-row" style="font-size:13px;font-weight:500">⚠ ${qPending.length} contract${qPending.length>1?'s are':' is'} past the renewal date but still pending — the quarter has <strong>${daysLeft} day${daysLeft!==1?'s':''}</strong> remaining.</div>
      ${qPending.slice(0,5).map(c=>`<div class="ab-row" style="font-size:12px;padding-left:8px">→ <strong>${c.name}</strong>${c.dept?' · '+c.dept:''} — due ${ctRenewalDisplay(c)} (${Math.abs(daysFrom(ctRenewalDate(c)))}d overdue)</div>`).join('')}
      ${qPending.length>5?`<div class="ab-row" style="font-size:12px;padding-left:8px;opacity:.7">…and ${qPending.length-5} more</div>`:''}
    </div>`;
  } else {
    qWrap.innerHTML='';
  }
  const wpWrap=document.getElementById('c-weekly-prev');
  const pending=contracts.filter(c=>!c.renewed).sort((a,b)=>daysFrom(ctRenewalDate(a))-daysFrom(ctRenewalDate(b)));
  if(pending.length){
    wpWrap.innerHTML=`<div class="weekly-prev">
      <div class="wp-head">
        <div class="wp-title">
          <svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" width="13" height="13"><path d="M15 17h5l-1.405-1.405A2.032 2.032 0 0118 14.158V11a6.002 6.002 0 00-4-5.659V5a2 2 0 10-4 0v.341C7.67 6.165 6 8.388 6 11v3.159c0 .538-.214 1.055-.595 1.436L4 17h5m6 0v1a3 3 0 11-6 0v-1m6 0H9"/></svg>
          Weekly Alert Preview — ${pending.length} pending renewal
        </div>
        <button class="btn btn-outline btn-sm" onclick="sendWeeklyAlert()">Send now</button>
      </div>
      ${pending.slice(0,6).map(c=>{
        const d=daysFrom(ctRenewalDate(c));
        const dl=d<0?`${Math.abs(d)}d overdue`:d===0?'Today!':d<=30?`${d}d left`:'';
        return `<div class="wp-row">
          <div class="s-av">${ini(c.name)}</div>
          <div class="wp-name">${c.name}</div>
          <div class="wp-dept">${c.dept||'—'}</div>
          <div class="wp-days">${dl}</div>
        </div>`;
      }).join('')}
      ${pending.length>6?`<div style="font-size:12px;color:var(--text3);padding:5px 0">…and ${pending.length-6} more</div>`:''}
    </div>`;
  } else {
    wpWrap.innerHTML='';
  }

  // Pending table — in-place (reflects Name search + Date filter)
  const pWrap=document.getElementById('c-pending-wrap');
  const pendingFiltered=filtered.filter(c=>!c.renewed).sort((a,b)=>daysFrom(ctRenewalDate(a))-daysFrom(ctRenewalDate(b)));
  if(pendingFiltered.length){
    let tbl=document.getElementById('ct-pending-tbl');
    if(!tbl){
      pWrap.innerHTML=`<div class="sec-title">Pending renewal (${pendingFiltered.length})</div><div class="ct-table"><table><thead><tr><th style="width:30px"><input type="checkbox" onchange="ctSelectAllVisible(this.checked,'pending')" ${pendingFiltered.length&&pendingFiltered.every(c=>ctSelected.has(c.id))?'checked':''}></th><th>Name</th><th>Code</th><th>Dept</th><th>Renewal Date</th><th>Deadline</th><th>Status</th><th>Actions</th></tr></thead><tbody id="ct-pending-tbody"></tbody></table></div>`;
      tbl=document.getElementById('ct-pending-tbody');
    } else {
      pWrap.querySelector('.sec-title').textContent=`Pending renewal (${pendingFiltered.length})`;
    }
    const tbody=document.getElementById('ct-pending-tbody');
    tbody.innerHTML=pendingFiltered.map(c=>{
      const d=ctDeadlineDays(c);
      const dl=d<0?`${Math.abs(d)}d overdue`:d===0?'Deadline today!':d===1?'1d left':`${d}d left`;
      const dc=d<0?'var(--red-text)':d<=14?'var(--amber-text)':'var(--text2)';
      return `<tr id="ctr-${c.id}">
        <td><input type="checkbox" ${ctSelected.has(c.id)?'checked':''} onchange="ctToggleRow('${c.id}')"></td>
        <td><strong>${c.name}</strong></td><td>${c.code||'—'}</td><td>${c.dept||'—'}</td>
        <td>${ctRenewalDisplay(c)}</td>
        <td><span style="color:${dc};font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:600">${dl}</span><div style="font-size:10.5px;color:var(--text3);margin-top:1px">deadline ${fmt(contractDeadline(c))}</div></td>
        <td>${ctStatusPill(c)}</td>
        <td style="white-space:nowrap">
          <div style="display:flex;gap:4px;margin-bottom:5px">
            <button class="btn btn-sm ${c.renewed?'btn-green':'btn-outline'}" onclick="setCtStatus('${c.id}','renewed')">Renewed</button>
            <button class="btn btn-sm ${!c.renewed?'btn-pending-active':'btn-outline'}" onclick="setCtStatus('${c.id}','pending')">Pending</button>
            <button class="btn btn-danger btn-sm" onclick="removeCt('${c.id}')">Remove</button>
          </div>
          ${!c.renewed?`<input type="text" placeholder="Reason for pending…" value="${(c.pendingReason||'').replace(/"/g,'&quot;')}" style="width:100%;min-width:160px;padding:4px 8px;font-size:11px;border:1px solid var(--border-strong);border-radius:6px" onchange="setCtReason('${c.id}',this.value)">`:''}
        </td>
      </tr>`;
    }).join('');
  } else if(filtersActive){
    pWrap.innerHTML=`<div class="empty" style="margin-bottom:18px"><div class="empty-title">No matching pending renewals</div><div>Try a different name or clear the month filter</div></div>`;
  } else {
    pWrap.innerHTML=`<div class="empty" style="margin-bottom:18px"><div class="empty-title">No pending renewals</div><div>Add contracts above or upload a CSV file</div></div>`;
  }

  // Renewed table (reflects Name search + Date filter)
  const rWrap=document.getElementById('c-renewed-wrap');
  const renewedFiltered=filtered.filter(c=>c.renewed);
  if(renewedFiltered.length){
    rWrap.innerHTML=`<div class="sec-title">Already renewed (${renewedFiltered.length})</div>
    <div class="ct-table"><table><thead><tr><th style="width:30px"><input type="checkbox" onchange="ctSelectAllVisible(this.checked,'renewed')" ${renewedFiltered.length&&renewedFiltered.every(c=>ctSelected.has(c.id))?'checked':''}></th><th>Name</th><th>Code</th><th>Dept</th><th>Renewal Date</th><th>Renewed On</th><th>Status</th><th>Actions</th></tr></thead>
    <tbody>${renewedFiltered.map(c=>`<tr class="dim">
      <td><input type="checkbox" ${ctSelected.has(c.id)?'checked':''} onchange="ctToggleRow('${c.id}')"></td>
      <td><strong>${c.name}</strong></td><td>${c.code||'—'}</td><td>${c.dept||'—'}</td>
      <td>${ctRenewalDisplay(c)}</td><td>${fmt(c.renewedOn)}</td>
      <td>${ctStatusPill(c)}</td>
      <td style="white-space:nowrap">
        <div style="display:flex;gap:4px">
          <button class="btn btn-sm ${c.renewed?'btn-green':'btn-outline'}" onclick="setCtStatus('${c.id}','renewed')">Renewed</button>
          <button class="btn btn-sm ${!c.renewed?'btn-pending-active':'btn-outline'}" onclick="setCtStatus('${c.id}','pending')">Pending</button>
          <button class="btn btn-danger btn-sm" onclick="removeCt('${c.id}')">Remove</button>
        </div>
      </td>
    </tr>`).join('')}</tbody></table></div>`;
  } else if(filtersActive && contracts.some(c=>c.renewed)){
    rWrap.innerHTML=`<div class="empty" style="margin-bottom:18px"><div class="empty-title">No matching renewed contracts</div><div>Try a different name or clear the month filter</div></div>`;
  } else {
    rWrap.innerHTML='';
  }
}

function setCtStatus(id,status){
  const c=contracts.find(c=>c.id===id); if(!c) return;
  if(status==='renewed'){
    c.renewed=true; c.renewedOn=today();
    toast(`✓ ${c.name}'s contract marked renewed`);
  } else {
    c.renewed=false; c.renewedOn=null;
  }
  saveC(); renderContracts(); renderSidebar(); if(document.getElementById("pg-monthly").classList.contains("active")) renderMonthly();
}
function setCtReason(id,val){
  const c=contracts.find(c=>c.id===id); if(!c) return;
  c.pendingReason=val; saveC();
}
function removeCt(id){
  if(!confirm('Remove this contract entry?')) return;
  contracts=contracts.filter(c=>c.id!==id); saveC(); renderContracts(); renderSidebar(); if(document.getElementById("pg-monthly").classList.contains("active")) renderMonthly();
}
function ctToggleRow(id){
  if(ctSelected.has(id)) ctSelected.delete(id); else ctSelected.add(id);
  renderContracts();
}
function ctSelectAllVisible(checked, group){
  // Operates only on the currently visible (search/date-filtered) rows for that group
  const ids=getFilteredContracts().filter(c=> group==='pending' ? !c.renewed : c.renewed).map(c=>c.id);
  ids.forEach(id=> checked ? ctSelected.add(id) : ctSelected.delete(id));
  renderContracts();
}
function ctRemoveSelected(){
  const n=ctSelected.size;
  if(!n){ toast('No employees selected'); return; }
  if(!confirm(`Remove ${n} selected contract${n!==1?'s':''}? This will delete them from the currently loaded dataset and cannot be undone.`)) return;
  contracts=contracts.filter(c=>!ctSelected.has(c.id));
  ctSelected.clear();
  saveC(); renderContracts(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();
  toast(`✓ Removed ${n} contract${n!==1?'s':''}`);
}
function ctBulkMarkRenewed(){
  const n=ctSelected.size;
  if(!n){ toast('No employees selected'); return; }
  if(!confirm(`Mark ${n} selected employee${n!==1?'s':''} as Renewed?\n\nAre you sure you want to mark the selected contracts as Renewed?`)) return;
  const stamp=today();
  contracts.forEach(c=>{
    if(ctSelected.has(c.id)){ c.renewed=true; c.renewedOn=stamp; }
  });
  ctSelected.clear();
  saveC(); renderContracts(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();
  toast(`✓ Successfully marked ${n} contract${n!==1?'s':''} as Renewed`);
}
function ctResetAll(){
  if(!contracts.length){ toast('No contract data to reset'); return; }
  if(!confirm('Reset ALL contract renewal data? This permanently removes every uploaded/linked contract record and cannot be undone.')) return;
  contracts=[];
  ctSelected.clear();
  saveC(); renderContracts(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();
  toast('✓ Contract Renewals data has been reset');
}

/* ─── OFFBOARDING ─── */
const OB_TASKS=[
  {key:'oracle',type:'checkbox',label:'Terminate the employee on Oracle'},
  {key:'documents',type:'checkbox',label:'Signed the required documents'},
  {key:'cards',type:'checkbox',label:'Provide the access & medical cards'},
  {key:'incentives',type:'yesno',label:'Eligible for Main Product / Achievement or Sales incentives?'},
  {key:'clearance',type:'clearance',label:'Clearance Status'},
  {key:'medicalDeletion',type:'checkbox',label:'Sent the medical deletion'},
];
function obDefaultTasks(){
  const t={};
  OB_TASKS.forEach(task=>{
    if(task.type==='checkbox') t[task.key]=false;
    else if(task.type==='amount') t[task.key]='';
    else if(task.type==='clearance'){
      t.clearanceStatus=null;      // 'clear' | 'not_cleared'
      t.clearAmount='';            // Amount to be Transferred
      t.clearDocsReceived=null;    // 'yes' | 'no'
      t.clearExpLetter=null;       // 'yes' | 'no'
      t.clearBankTransfer=null;    // 'yes' | 'no'
      t.dueAmountStatus=null;      // 'paid' | 'not_yet'
      t.clearanceReason='';        // free text
    }
    else { t[task.key]=null; if(task.type==='yesno_reason') t[task.key+'Reason']=''; }
  });
  return t;
}
offboards.forEach(o=>{
  o.tasks = Object.assign(obDefaultTasks(), o.tasks||{});
});
function obTaskDone(o,t){
  const v=o.tasks[t.key];
  if(t.type==='checkbox') return !!v;
  if(t.type==='yesno') return v==='yes'||v==='no';
  if(t.type==='amount') return v!==undefined&&v!==null&&String(v).trim()!=='';
  if(t.type==='yesno_reason'){
    if(v==='yes') return true;
    if(v==='no') return !!(o.tasks[t.key+'Reason']&&o.tasks[t.key+'Reason'].trim());
    return false;
  }
  if(t.type==='clearance'){
    const status=o.tasks.clearanceStatus;
    if(status==='clear'){
      return !!(o.tasks.clearAmount && String(o.tasks.clearAmount).trim()!=='' &&
        (o.tasks.clearDocsReceived==='yes'||o.tasks.clearDocsReceived==='no') &&
        (o.tasks.clearExpLetter==='yes'||o.tasks.clearExpLetter==='no') &&
        (o.tasks.clearBankTransfer==='yes'||o.tasks.clearBankTransfer==='no'));
    }
    if(status==='not_cleared'){
      return !!(o.tasks.dueAmountStatus && o.tasks.clearanceReason && o.tasks.clearanceReason.trim()!=='');
    }
    return false;
  }
  return false;
}
function obProgress(o){
  const done=OB_TASKS.filter(t=>obTaskDone(o,t)).length;
  return Math.round(done/OB_TASKS.length*100);
}
function obDeadline(o){
  if(!o.lastDay) return null;
  const last=new Date(o.lastDay+'T00:00:00');
  const plus14=new Date(last); plus14.setDate(plus14.getDate()+14);
  const today=new Date(); today.setHours(0,0,0,0);
  if(today<plus14) return null; // "Not yet eligible"
  return new Date(plus14.getFullYear(), plus14.getMonth()+1, 0); // EOMONTH(A2+14,0)
}
function obDeadlineISO(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function obDeadlineLabel(o){
  const d=obDeadline(o);
  return d?fmt(obDeadlineISO(d)):'Not yet eligible';
}
function obIsOverdue(o){
  const d=obDeadline(o);
  if(!d) return false;
  const today=new Date(); today.setHours(0,0,0,0);
  return today>d;
}
function obStatus(o){
  const p=obProgress(o);
  if(p===100) return 'completed';
  if(obIsOverdue(o)) return 'overdue';
  return 'progress';
}
function obStatusPill(o){
  const s=obStatus(o);
  const m={completed:'pl-green',overdue:'pl-red',progress:'pl-amber'};
  const t={completed:'✓ Completed',overdue:'Overdue',progress:'In progress'};
  return `<span class="pill ${m[s]}">${t[s]}</span>`;
}
function obStats(){
  return{
    total:offboards.length,
    completed:offboards.filter(o=>obProgress(o)===100).length,
    progress:offboards.filter(o=>obProgress(o)<100).length,
    overdue:offboards.filter(o=>obProgress(o)<100&&obIsOverdue(o)).length,
  };
}
/* ─── TENURE (Joining Date → Last Working Date) ─── */
function obTenureYears(o){
  if(!o.joinDate||!o.lastDay) return null;
  const j=new Date(o.joinDate+'T00:00:00'), l=new Date(o.lastDay+'T00:00:00');
  if(isNaN(j)||isNaN(l)||l<j) return null;
  return (l-j)/(1000*60*60*24*365.25);
}
function obTenureBucket(o){
  const y=obTenureYears(o);
  if(y===null) return 'Not specified';
  if(y<1) return '<1 year';
  if(y<2) return '1–2 years';
  if(y<5) return '2–5 years';
  if(y<10) return '5–10 years';
  return '10+ years';
}
function obTenureLabel(o){
  const y=obTenureYears(o);
  if(y===null) return '—';
  if(y<1) return `${Math.round(y*12)} mo`;
  return `${y.toFixed(1)} yrs`;
}
/* ─── DRILL-DOWN MODAL (chart/KPI click-through to employee records) ─── */
let drillRecords=[];
function openDrillModal(title, recs, subtitle){
  drillRecords=recs||[];
  document.getElementById('drill-title').textContent=title;
  document.getElementById('drill-sub').textContent=subtitle||`${drillRecords.length} employee${drillRecords.length!==1?'s':''}`;
  const wrap=document.getElementById('drill-table-wrap');
  if(!drillRecords.length){
    wrap.innerHTML=`<div class="empty"><div class="empty-title">No matching records</div><div>No offboarding records match this selection.</div></div>`;
  } else {
    wrap.innerHTML=`<div class="ct-table" style="max-height:52vh"><table><thead><tr>
      <th>Employee ID</th><th>Name</th><th>Gender</th><th>Department</th><th>Job Title</th><th>Grade</th><th>Rating</th>
      <th>Joining Date</th><th>Resignation Date</th><th>Last Working Day</th><th>Tenure</th><th>Reason for Leaving</th><th>Employment Type</th><th>Nationality</th>
    </tr></thead><tbody>${drillRecords.map(o=>`<tr>
        <td>${o.code||'—'}</td><td><strong>${o.name}</strong></td><td>${o.gender||'—'}</td><td>${o.dept||'—'}</td>
        <td>${o.jobTitle||'—'}</td><td>${o.grade||'—'}</td><td>${o.rating||'—'}</td>
        <td>${o.joinDate?fmt(o.joinDate):'—'}</td><td>${o.resignDate?fmt(o.resignDate):'—'}</td><td>${fmt(o.lastDay)}</td>
        <td>${obTenureLabel(o)}</td><td>${o.reason||'—'}</td><td>${o.employmentType||'—'}</td><td>${o.nationality||'—'}</td>
      </tr>`).join('')}</tbody></table></div>`;
  }
  document.getElementById('drill-bg').classList.add('open');
}
function closeDrillModal(){
  document.getElementById('drill-bg').classList.remove('open');
}
function exportDrillToExcel(){
  if(!drillRecords.length){ toast('No records to export','danger'); return; }
  const rows=[
    ['Employee ID','Name','Gender','Department','Job Title','Grade','Rating','Joining Date','Resignation Date','Last Working Day','Tenure','Reason for Leaving','Employment Type','Nationality'],
    ...drillRecords.map(o=>[o.code||'', o.name||'', o.gender||'', o.dept||'', o.jobTitle||'', o.grade||'', o.rating||'', o.joinDate||'', o.resignDate||'', o.lastDay||'', obTenureLabel(o), o.reason||'', o.employmentType||'', o.nationality||'']),
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=rows[0].map(()=>({wch:17}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Employees');
  XLSX.writeFile(wb, `resignation_analysis_export_${Date.now()}.xlsx`);
  toast(`✓ Exported ${drillRecords.length} record${drillRecords.length!==1?'s':''} to Excel`);
}
/* ─── DRILL-DOWN MODAL (Overtime Analysis) ─── */
let otDrillRecords=[];
function openOtDrillModal(title, recs, subtitle){
  otDrillRecords=recs||[];
  document.getElementById('ot-drill-title').textContent=title;
  document.getElementById('ot-drill-sub').textContent=subtitle||`${otDrillRecords.length} record${otDrillRecords.length!==1?'s':''}`;
  const wrap=document.getElementById('ot-drill-table-wrap');
  if(!otDrillRecords.length){
    wrap.innerHTML=`<div class="empty"><div class="empty-title">No matching records</div><div>No overtime records match this selection.</div></div>`;
  } else {
    const sorted=[...otDrillRecords].sort((a,b)=>(a.date<b.date?1:-1));
    wrap.innerHTML=`<div class="ct-table" style="max-height:52vh"><table><thead><tr>
      <th>Name</th><th>Job Title</th><th>Division</th><th>Department</th><th>Employee ID</th><th>Date</th><th>Actual Hrs</th><th>Normal Hrs</th><th>Saturday Hrs</th><th>Official Days Hrs</th><th>OT Hours</th><th>Monthly Salary</th><th>OT Rate</th><th>OT Amount</th><th>OT %</th><th>Status</th><th>Reason</th>
    </tr></thead><tbody>${sorted.map(o=>{
        const pct=otPercentOf(o);
        return `<tr>
        <td><strong>${o.name}</strong></td><td>${o.jobTitle||'—'}</td><td>${otDivisionOf(o)}</td><td>${otOrgOf(o)}</td><td>${o.code||'—'}</td>
        <td>${fmt(o.date)}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.actualHours>0?o.actualHours:'N/A'}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.normalHours||0}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.saturdayHours||0}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.officialHours||0}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.hours}</td>
        <td style="font-family:'JetBrains Mono',monospace">${o.monthlySalary?o.monthlySalary.toLocaleString():'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace">${otHourlyRate(o)?otHourlyRate(o).toFixed(2):'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace">${otAmountOf(o).toLocaleString()}</td>
        <td style="font-family:'JetBrains Mono',monospace">${pct!=null?pct.toFixed(1)+'%':'—'}</td>
        <td>${o.status||'—'}</td><td>${o.reason||'—'}</td>
      </tr>`;}).join('')}</tbody></table></div>`;
  }
  document.getElementById('ot-drill-bg').classList.add('open');
}
function closeOtDrillModal(){
  document.getElementById('ot-drill-bg').classList.remove('open');
}
function exportOtDrillToExcel(){
  if(!otDrillRecords.length){ toast('No records to export','danger'); return; }
  const rows=[
    ['Name','Job Title','Division','Department','Employee ID','Date','Actual Hrs','Normal Hrs','Saturday Hrs','Official Days Hrs','OT Hours','Monthly Salary','OT Rate','OT Amount','OT %','Status','Reason'],
    ...otDrillRecords.map(o=>[o.name||'', o.jobTitle||'', otDivisionOf(o), otOrgOf(o), o.code||'', o.date||'', o.actualHours||'', o.normalHours||0, o.saturdayHours||0, o.officialHours||0, o.hours||'', o.monthlySalary||'', otHourlyRate(o)?+otHourlyRate(o).toFixed(2):'', otAmountOf(o), otPercentOf(o)!=null?+otPercentOf(o).toFixed(1):'', o.status||'', o.reason||'']),
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=rows[0].map(()=>({wch:17}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Overtime');
  XLSX.writeFile(wb, `overtime_analysis_export_${Date.now()}.xlsx`);
  toast(`✓ Exported ${otDrillRecords.length} record${otDrillRecords.length!==1?'s':''} to Excel`);
}
/* ═══════════════ SMART IMPORT (any spreadsheet) ═══════════════
   One front door for all four uploaders (Onboarding, Contracts, Offboarding, Overtime).
   - Opens .xlsx .xlsm .xlsb .xls .xltx .ods .csv .tsv .txt
   - Finds the real header row (skips title / blank rows above it)
   - Picks the right sheet; asks only when several sheets qualify or none does
   - If a required column is not recognised, shows a column-mapping screen
   The chosen sheet is trimmed to start at the header row and any mapped headers are renamed to
   the names the existing importers already understand, so their logic is untouched. */
const IMP_EXTS=['xlsx','xlsm','xlsb','xls','xltx','xltm','ods','fods','csv','tsv','txt'];
const IMP_TEXT_EXTS=['csv','tsv','txt'];
function impClassifyContract(h){
  if(/^(status|renewal status|renewed|renewed \/ pending|renewed\/pending|contract status)$/.test(h)) return 'status';
  if(h.includes('pending reason')||h==='reason'||h.includes('reason for pending')||h==='notes'||h==='note') return 'reason';
  if(h.includes('name')) return 'name';
  if(h.includes('date')||h.includes('renewal')||h.includes('expir')) return 'date';
  if(h.includes('dept')||h.includes('depart')) return 'dept';
  if(h.includes('code')||h==='id'||h.includes('emp')) return 'code';
  return null;
}
const IMP_SPECS={
  hire:{label:'Onboarding (new hires)', readOpts:{cellDates:true},
    classify:h=>classifyHireHeader(h), skipSheet:n=>normKey(n)==='checklist',
    fields:[{key:'name',label:'Employee name',canon:'Full Name'},{key:'startDate',label:'Hiring date',canon:'Hiring Date'}]},
  contract:{label:'Contract renewals', readOpts:{},
    classify:impClassifyContract,
    fields:[{key:'name',label:'Employee name',canon:'Name'},{key:'date',label:'Renewal date',canon:'Renewal Date'}]},
  offboard:{label:'Offboarding / resignations', readOpts:{cellDates:true},
    classify:h=>{ const tk=classifyObChecklistHeader(h); return tk?null:classifyObHeader(h); },
    fields:[{key:'name',label:'Employee name',canon:'Name'},{key:'lastday',label:'Last working date',canon:'Last Working Date'}]},
  overtime:{label:'Overtime', readOpts:{cellNF:true},
    classify:h=>classifyOtHeader(h),
    fields:[{key:'name',label:'Employee name',canon:'Name'},{key:'date',label:'Date / month',canon:'Date'},
            {key:'hours',label:'Overtime hours',canon:'OT Hours',alias:['hours','normalhours','saturdayhours','officialhours']}]}
};
function impCellText(ws,r,c){ const x=ws[XLSX.utils.encode_cell({r,c})]; if(!x) return ''; return String(x.w!==undefined?x.w:(x.v!==undefined&&x.v!==null?x.v:'')).replace(/\s+/g,' ').trim(); }
function impRange(ws){ return ws&&ws['!ref']?XLSX.utils.decode_range(ws['!ref']):null; }
function impFieldKeys(f){ return f.alias||[f.key]; }
/* detect which column holds each required field, for a given header row (relative to the sheet's first row) */
function impColsFor(ws,spec,rowRel){
  const rg=impRange(ws); const found={}; const keysSeen=new Set();
  if(!rg) return {found,keysSeen,missing:spec.fields.map(f=>f.key)};
  const r=rg.s.r+rowRel;
  for(let c=rg.s.c;c<=rg.e.c;c++){
    const t=impCellText(ws,r,c); if(!t||/^[\d.,\-\s%\/:]+$/.test(t)) continue;
    let k=null; try{ k=spec.classify(t.toLowerCase()); }catch(e){ k=null; }
    if(!k) continue; keysSeen.add(k);
    spec.fields.forEach(f=>{ if(impFieldKeys(f).includes(k) && !(f.key in found)) found[f.key]=c; });
  }
  return {found,keysSeen,missing:spec.fields.filter(f=>!(f.key in found)).map(f=>f.key)};
}
/* scan the first rows for the most header-like row */
function impDetectHeader(ws,spec){
  const rg=impRange(ws); if(!rg) return {rowRel:0,found:{},missing:spec.fields.map(f=>f.key),score:-1};
  let best=null; const lim=Math.min(rg.e.r-rg.s.r,40);
  for(let i=0;i<=lim;i++){
    const res=impColsFor(ws,spec,i);
    const score=(spec.fields.length-res.missing.length)*100+res.keysSeen.size;
    if(!best||score>best.score) best={rowRel:i,found:res.found,missing:res.missing,score};
  }
  return best;
}
/* trim sheet to start at the header row and apply column choices (key → absolute column) */
function impApply(ws,spec,rowRel,choice){
  const rg=impRange(ws); const r0=rg.s.r+rowRel;
  const auto=impColsFor(ws,spec,rowRel).found;
  Object.entries(choice).forEach(([key,col])=>{
    if(col===''||col==null||auto[key]===col) return;
    const f=spec.fields.find(x=>x.key===key); const keys=impFieldKeys(f);
    for(let c=rg.s.c;c<=rg.e.c;c++){           // stop another column from out-ranking the chosen one
      if(c===col) continue;
      const t=impCellText(ws,r0,c); let k=null; try{ k=t?spec.classify(t.toLowerCase()):null; }catch(e){}
      if(k&&keys.includes(k)) ws[XLSX.utils.encode_cell({r:r0,c})]={t:'s',v:'Unmapped '+(c+1),w:'Unmapped '+(c+1)};
    }
    ws[XLSX.utils.encode_cell({r:r0,c:col})]={t:'s',v:f.canon,w:f.canon};
  });
  ws['!ref']=XLSX.utils.encode_range({s:{r:r0,c:rg.s.c},e:rg.e});
  return ws;
}
async function impReadWorkbook(file,spec){
  const ext=(file.name.split('.').pop()||'').toLowerCase();
  if(!IMP_EXTS.includes(ext)) throw new Error('Unsupported file type ".'+ext+'". Supported: '+IMP_EXTS.map(e=>'.'+e).join(' '));
  if(IMP_TEXT_EXTS.includes(ext)){
    let text=await file.text(); text=text.replace(/^\uFEFF/,'');
    return {wb:XLSX.read(text,{type:'string',raw:true}),isText:true};   // raw:true keeps dates exactly as typed
  }
  return {wb:XLSX.read(await file.arrayBuffer(),Object.assign({type:'array'},spec.readOpts||{})),isText:false};
}
/* main entry → resolves {wb,ws,sheetName,isText} (or null if cancelled / failed) */
async function smartImport(file,kind){
  const spec=IMP_SPECS[kind];
  let rd;
  try{ rd=await impReadWorkbook(file,spec); }
  catch(e){ toast('Could not read that file: '+(e.message||e)+'. If it is password-protected, remove the password and try again.','danger',0); return null; }
  const wb=rd.wb;
  const names=wb.SheetNames.filter(n=>wb.Sheets[n]&&wb.Sheets[n]['!ref']);
  if(!names.length){ toast('That file has no data in it','danger',0); return null; }
  const pool=names.filter(n=>!(spec.skipSheet&&spec.skipSheet(n))); const cand=pool.length?pool:names;
  const det={}; cand.forEach(n=>det[n]=impDetectHeader(wb.Sheets[n],spec));
  const ok=cand.filter(n=>det[n].missing.length===0);
  let sheet,rowRel,choice={};
  if(ok.length===1){ sheet=ok[0]; rowRel=det[sheet].rowRel; }
  else {
    const pre=ok.length?ok[0]:cand.slice().sort((a,b)=>det[b].score-det[a].score)[0];
    const res=await impModal(file,wb,spec,names,pre,det);
    if(!res) return null;
    sheet=res.sheet; rowRel=res.rowRel; choice=res.choice;
  }
  const ws=impApply(wb.Sheets[sheet],spec,rowRel,choice);
  return {wb,ws,sheetName:sheet,isText:rd.isText};
}
function impModal(file,wb,spec,names,preSheet,det){
  return new Promise(resolve=>{
    let bg=document.getElementById('imp-modal-bg');
    if(!bg){ bg=document.createElement('div'); bg.id='imp-modal-bg'; bg.className='modal-bg'; document.body.appendChild(bg); }
    const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
    const colLetter=c=>XLSX.utils.encode_col(c);
    const st={sheet:preSheet,rowRel:det[preSheet]?det[preSheet].rowRel:0,choice:{}};
    function reset(keepRow){
      const ws=wb.Sheets[st.sheet];
      if(!keepRow){ const d=det[st.sheet]||impDetectHeader(ws,spec); det[st.sheet]=d; st.rowRel=d.rowRel; }
      st.choice={}; const f=impColsFor(ws,spec,st.rowRel).found; spec.fields.forEach(x=>{ st.choice[x.key]=(x.key in f)?f[x.key]:''; });
    }
    reset(false);
    function done(val){ bg.classList.remove('open'); bg.innerHTML=''; resolve(val); }
    function render(){
      const ws=wb.Sheets[st.sheet], rg=impRange(ws), r0=rg.s.r+st.rowRel;
      const maxRow=rg.e.r-rg.s.r;
      const colOpts=['<option value="">— choose column —</option>'];
      for(let c=rg.s.c;c<=rg.e.c;c++){ const t=impCellText(ws,r0,c); if(t) colOpts.push(`<option value="${c}">${colLetter(c)} — ${esc(t.slice(0,40))}</option>`); }
      const cols=[]; for(let c=rg.s.c;c<=Math.min(rg.e.c,rg.s.c+11);c++) cols.push(c);
      const prevRows=[]; for(let r=r0;r<=Math.min(rg.e.r,r0+5);r++) prevRows.push(r);
      const allOk=spec.fields.every(f=>st.choice[f.key]!==''&&st.choice[f.key]!=null);
      const sel='background:var(--surface);color:var(--text);border:1px solid var(--border-strong);border-radius:8px;padding:7px 9px;font-size:12.5px;width:100%';
      bg.innerHTML=`<div class="modal" style="width:720px;max-width:96vw">
        <div class="modal-title">Import — ${esc(spec.label)}</div>
        <div style="font-size:12.5px;color:var(--text2);margin-bottom:12px;line-height:1.6">File: <b>${esc(file.name)}</b>. ${allOk?'Check the preview, then import.':'The app could not recognise every required column automatically — please match them below.'}</div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px">
          ${names.length>1?`<div style="flex:2;min-width:200px"><div style="font-size:11.5px;font-weight:600;margin-bottom:4px">Sheet</div><select id="imp-sheet" style="${sel}">${names.map(n=>`<option value="${esc(n)}" ${n===st.sheet?'selected':''}>${esc(n)}</option>`).join('')}</select></div>`:''}
          <div style="flex:1;min-width:140px"><div style="font-size:11.5px;font-weight:600;margin-bottom:4px">Header is on row</div><input id="imp-row" type="number" min="1" max="${maxRow+1}" value="${st.rowRel+1}" style="${sel}"></div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;margin-bottom:12px">
          ${spec.fields.map(f=>`<div><div style="font-size:11.5px;font-weight:600;margin-bottom:4px">${esc(f.label)} <span style="color:var(--red)">*</span></div>
            <select class="imp-map" data-key="${f.key}" style="${sel}">${colOpts.join('').replace(`value="${st.choice[f.key]}"`,`value="${st.choice[f.key]}" selected`)}</select></div>`).join('')}
        </div>
        <div style="font-size:11.5px;font-weight:600;margin-bottom:4px">Preview (first rows after the header)</div>
        <div style="overflow:auto;border:1px solid var(--border);border-radius:8px;max-height:210px;margin-bottom:14px"><table style="border-collapse:collapse;font-size:11.5px;width:100%">
          ${prevRows.map((r,i)=>`<tr>${cols.map(c=>`<td style="padding:5px 8px;border-bottom:1px solid var(--border);white-space:nowrap;${i===0?'font-weight:600;background:var(--surface2)':''}">${esc(impCellText(ws,r,c).slice(0,28))}</td>`).join('')}</tr>`).join('')}
        </table></div>
        <div class="modal-actions"><button class="btn btn-outline" id="imp-cancel">Cancel</button><button class="btn btn-primary" id="imp-go" ${allOk?'':'disabled style="opacity:.5;cursor:not-allowed"'}>Import</button></div>
      </div>`;
      bg.classList.add('open');
      const sh=bg.querySelector('#imp-sheet'); if(sh) sh.onchange=()=>{ st.sheet=sh.value; reset(false); render(); };
      bg.querySelector('#imp-row').onchange=e=>{ const v=Math.max(1,Math.min(maxRow+1,parseInt(e.target.value)||1)); st.rowRel=v-1; reset(true); render(); };
      bg.querySelectorAll('.imp-map').forEach(s=>{ s.onchange=()=>{ st.choice[s.dataset.key]=s.value===''?'':Number(s.value); render(); }; });
      bg.querySelector('#imp-cancel').onclick=()=>done(null);
      bg.querySelector('#imp-go').onclick=()=>{ if(allOk) done({sheet:st.sheet,rowRel:st.rowRel,choice:Object.assign({},st.choice)}); };
    }
    render();
  });
}

/* ─── OFFBOARDING EXCEL/CSV BULK UPLOAD ─── */
function downloadObTemplate(){
  // ONE master sheet: analysis columns + checklist columns. Upload it on Offboarding OR Resignation Analysis.
  const rows=[
    ['Employee Number','Name','Company','Cluster/Supportive Function','Division','Unit/Department/Line','Title','Job Level','Hire Date','Last Working Date','Resignation Reason','Company Name','Business Partner Code','Business Partner Name','Grade','Grade Level','Bell Curve Class','Gender',
     'Terminate the employee on Oracle','Signed the required documents','Provide the access & medical cards','Eligible for Main Product / Achievement or Sales incentives?',
     'Clearance Status','Amount to be Transferred','Documents Received','Experience Letter','Bank Transfer','Due Amount Status','Clearance Reason','Sent the medical deletion'],
    ['EMP-001','Ahmed Hassan','Pharma Group','R&D','R&D Business Division','Finance','Financial Analyst','Senior','2022-04-01','2026-09-15','Resignation','Pharma Group','BP-01','Sara Mahmoud','Grade 3','G3','Meets Expectations','Male',
     'Yes','Yes','Yes','Yes','Clear','15000','Yes','Yes','Yes','','','Yes'],
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=rows[0].map(h=>({wch:Math.min(34,Math.max(12,h.length+2))}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Offboarding');
  XLSX.writeFile(wb,'offboarding_master_template.xlsx');
}
function handleObFile(file){
  if(!file) return;
  document.getElementById('ob-upload-errors').innerHTML='';
  smartImport(file,'offboard').then(res=>{
    if(!res) return;
    // header:1 + raw:true keeps each cell's native type (Date / number / string) instead of collapsing to text
    processObRows(XLSX.utils.sheet_to_json(res.ws,{header:1,raw:true,defval:''}));
  });
  document.getElementById('file-in-ob').value='';
}
/* Classifies a lower-cased header string into a known offboarding field.
   Order matters — more specific patterns are checked before generic ones
   (e.g. "date of joining" must not be swallowed by the generic 'date' fallback). */
function classifyObHeader(h){
  // Master-sheet columns that must be recognised BEFORE the generic rules below
  // (otherwise "Company Name" is read as the employee name, "Job Level" as the grade, etc.)
  if(h.includes('business partner')||h==='company name') return null;      // kept as extra columns
  if(h.includes('job level')||h.includes('grade level')) return null;      // kept as extra columns
  if(h.includes('bell curve')) return 'rating';                            // Bell Curve Class = performance rating
  if(h.includes('cluster')||h.includes('supportive function')) return 'cluster';
  if(h==='company') return 'company';
  if(h.includes('division')) return 'division';
  if(h.includes('name')) return 'name';
  if(h.includes('lastday')||h.includes('last day')||h.includes('last_day')||h.includes('last working')||h.includes('lwd')||h.includes('exit date')) return 'lastday';
  if(h.includes('resignation date')||h.includes('resign date')||h.includes('notice date')||h.includes('date of resignation')) return 'resigndate';
  if(h.includes('join')||h.includes('doj')||h.includes('hire date')||h.includes('start date')) return 'joindate';
  if(h.includes('gender')||h.includes('sex')) return 'gender';
  if(h.includes('job title')||h.includes('designation')||h.includes('position')||h.includes('title')) return 'jobtitle';
  if(h.includes('employment type')||h.includes('emp type')||h.includes('contract type')||h.includes('worker type')) return 'emptype';
  if(h.includes('nationality')||h.includes('citizenship')) return 'nationality';
  if(h.includes('grade')||h.includes('band')||h.includes('level')) return 'grade';
  if(h.includes('rating')||h.includes('performance')) return 'rating';
  if(h.includes('dept')||h.includes('department')) return 'dept';
  if(h.includes('code')||h.includes('emp')||h.includes('id')) return 'code';
  if(h.includes('reason')) return 'reason';
  if(h.includes('date')) return 'lastday'; // generic fallback
  return null;
}
function normalizeGender(v){
  const g=(v||'').trim().toLowerCase();
  if(g==='m'||g==='male') return 'Male';
  if(g==='f'||g==='female') return 'Female';
  return (v||'').trim();
}

/* ── OFFBOARDING DATE VALIDATION ──────────────────────────────────────────
   Dates are read straight from the uploaded file and checked field-by-field
   (day / month / year individually) instead of being handed to a lenient
   parser like `new Date(str)`, which silently "rolls over" impossible dates
   (e.g. 31/02/2026 → 03/03/2026) instead of rejecting them. Nothing here
   ever changes or guesses a value — a date is either accepted exactly as
   entered, or rejected with a specific, human-readable reason. */
const OB_DATE_COLS={lastday:'Last Working Date',joindate:'Joining Date',resigndate:'Resignation Date'};
const OB_MONTH_NAMES=['','January','February','March','April','May','June','July','August','September','October','November','December'];

function obEsc(s){
  return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function obIsLeapYear(y){ return (y%4===0 && y%100!==0) || y%400===0; }
function obDaysInMonth(m,y){
  const d=[31,obIsLeapYear(y)?29:28,31,30,31,30,31,31,30,31,30,31];
  return d[m-1];
}
function obFmtISO(y,m,d){
  return `${String(y).padStart(4,'0')}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}
// Converts an Excel date *serial number* (days since 1899-12-30, Excel's own
// epoch convention) to an ISO date — used only when a cell holds a raw
// numeric serial that Excel itself produced, never applied to guesswork.
function obExcelSerialToISO(serial){
  const n=Math.floor(Number(serial)); // floor: a time-of-day fraction must never push the date to the next day
  if(!isFinite(n)||n<10000||n>2958465) return null; // 10000 ≈ year 1927 — smaller numbers are not dates
  const d=new Date(Date.UTC(1899,11,30)+n*86400000);
  if(isNaN(d.getTime())) return null;
  return obFmtISO(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate());
}
// Validates an already-split (year, month, day) triple. This is the single
// place range/existence rules live, so every input path (ISO string, D/M/Y
// string, Excel serial fallback) is checked identically and strictly.
function obCheckYMD(y,mo,d,expectedFmt){
  if(!Number.isInteger(y)||y<1900||y>2100){
    return {ok:false, reason:`The year ${y} is outside the expected range (1900–2100).`, expected:expectedFmt};
  }
  if(!Number.isInteger(mo)||mo<1||mo>12){
    return {ok:false, reason:`The month ${mo} does not exist.`, expected:expectedFmt};
  }
  const dim=obDaysInMonth(mo,y);
  if(!Number.isInteger(d)||d<1||d>dim){
    const leapNote=(mo===2&&d===29)?' (this is not a leap year)':'';
    return {ok:false, reason:`The day ${d} does not exist in ${OB_MONTH_NAMES[mo]} ${y}${leapNote}.`, expected:expectedFmt};
  }
  return {ok:true, iso:obFmtISO(y,mo,d)};
}
// Reads a *text* date in almost any common layout — never auto-corrects an impossible date.
// Accepted: 2026-03-15 · 2026/03/15 · 20260315 · 15/03/2026 · 15-3-26 · 15.03.2026 ·
//           15 Mar 2026 · 15-Mar-26 · March 15, 2026 · 15th of March 2026 · Sun, 15 Mar 2026 ·
//           Mar 15 2026 · 2026 Mar 15 · any of the above followed by a time (e.g. 15/03/2026 14:30,
//           2026-03-15T00:00:00Z) · Arabic-Indic digits (١٥/٠٣/٢٠٢٦) and Arabic month names ·
//           a bare Excel serial typed as text (e.g. 46096).
// Numeric day/month dates follow DD/MM/YYYY. If one number is above 12 the order is unambiguous
// and is used as-is (03/15/2026 → 15 March); when both could be a month (e.g. 04/05/2026) it is
// read as DD/MM and flagged with assumed:true so the user is told.
const OB_DATE_EXPECTED='DD/MM/YYYY (e.g. 15/03/2026), YYYY-MM-DD, or 15 Mar 2026';
const OB_MONTH_LOOKUP=(()=>{
  const en=['january','february','march','april','may','june','july','august','september','october','november','december'];
  const ar=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
  const map={};
  en.forEach((n,i)=>{ map[n]=i+1; map[n.slice(0,3)]=i+1; });
  map.sept=9;
  ar.forEach((n,i)=>{ map[n]=i+1; });
  Object.assign(map,{'ابريل':4,'يونية':6,'يوليه':7,'اغسطس':8,'اكتوبر':10});
  return map;
})();
function obMonthFromName(t){ return OB_MONTH_LOOKUP[String(t).toLowerCase().replace(/\./g,'')]||null; }
function obFullYear(y){ return y<100 ? (y<=29?2000+y:1900+y) : y; } // 2-digit years: 00–29 → 20xx, 30–99 → 19xx
function obValidateDateString(raw){
  let s=String(raw)
    .replace(/[\u0660-\u0669]/g,c=>c.charCodeAt(0)-0x0660)   // Arabic-Indic digits
    .replace(/[\u06F0-\u06F9]/g,c=>c.charCodeAt(0)-0x06F0)   // Persian digits
    .trim();
  if(!s) return {ok:false, missing:true};
  const original=s;
  s=s.replace(/[T\s]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?\s*(?:[AaPp]\.?[Mm]\.?)?\s*(?:Z|[+-]\d{2}:?\d{2})?$/,'') // trailing time / timezone
     .replace(/\b(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\b\.?/gi,' ')                                          // weekday names
     .replace(/(\d)\s*(?:st|nd|rd|th)\b/gi,'$1')                                                           // 1st, 2nd, 15th
     .replace(/\bof\b/gi,' ').replace(/,/g,' ').replace(/\s+/g,' ').trim();
  const W='[A-Za-z\u0600-\u06FF]+', SEP='[-\\/.\\s]';
  let m;
  // YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
  if((m=s.match(/^(\d{4})[-\/.\s](\d{1,2})[-\/.\s](\d{1,2})$/))) return obCheckYMD(+m[1],+m[2],+m[3],OB_DATE_EXPECTED);
  // YYYYMMDD
  if((m=s.match(/^(\d{4})(\d{2})(\d{2})$/))) return obCheckYMD(+m[1],+m[2],+m[3],OB_DATE_EXPECTED);
  // numeric day/month/year (2- or 4-digit year)
  if((m=s.match(/^(\d{1,2})[-\/.\s](\d{1,2})[-\/.\s](\d{2}|\d{4})$/))){
    const a=+m[1], b=+m[2], y=obFullYear(+m[3]);
    if(b>12 && a<=12) return obCheckYMD(y,a,b,OB_DATE_EXPECTED);        // clearly MM/DD/YYYY
    const r=obCheckYMD(y,b,a,OB_DATE_EXPECTED);                          // DD/MM/YYYY
    if(r.ok && a<=12 && b<=12 && a!==b) r.assumed=true;                  // could have been MM/DD
    return r;
  }
  // 15 Mar 2026 · 15-Mar-26 · 15March2026
  if((m=s.match(new RegExp(`^(\\d{1,2})${SEP}*(${W})\\.?${SEP}*(\\d{2}|\\d{4})$`)))){
    const mo=obMonthFromName(m[2]); if(mo) return obCheckYMD(obFullYear(+m[3]),mo,+m[1],OB_DATE_EXPECTED);
  }
  // March 15 2026 · Mar-15-26
  if((m=s.match(new RegExp(`^(${W})\\.?${SEP}*(\\d{1,2})${SEP}*(\\d{2}|\\d{4})$`)))){
    const mo=obMonthFromName(m[1]); if(mo) return obCheckYMD(obFullYear(+m[3]),mo,+m[2],OB_DATE_EXPECTED);
  }
  // 2026 Mar 15
  if((m=s.match(new RegExp(`^(\\d{4})${SEP}*(${W})\\.?${SEP}*(\\d{1,2})$`)))){
    const mo=obMonthFromName(m[2]); if(mo) return obCheckYMD(+m[1],mo,+m[3],OB_DATE_EXPECTED);
  }
  // A bare Excel serial number typed as plain text (cell wasn't date-formatted)
  if(/^\d{5}(?:\.\d+)?$/.test(s)){
    const iso=obExcelSerialToISO(+s);
    if(iso) return {ok:true, iso};
  }
  return {ok:false, reason:`"${original}" is not in a recognized date format.`, expected:OB_DATE_EXPECTED};
}
// Resolves one Excel/CSV cell — whatever native type it came in as — into a
// uniform {provided, ok, iso, display, reason, expected} result. This is the
// single entry point every date column is validated through.
function obResolveDateCell(value){
  if(value===undefined||value===null) return {provided:false};
  if(value instanceof Date){
    if(isNaN(value.getTime())){
      return {provided:true, ok:false, display:'(unreadable Excel date cell)', reason:'Excel could not represent this date value — the cell may be corrupted.', expected:'Re-enter the date using Excel\'s date picker, or as DD/MM/YYYY text.'};
    }
    const iso=obFmtISO(value.getUTCFullYear(),value.getUTCMonth()+1,value.getUTCDate());
    return {provided:true, ok:true, iso, display:iso};
  }
  if(typeof value==='number'){
    let iso=null;
    if(Number.isInteger(value)&&value>=19000101&&value<=21001231){ // 20260315 typed as a number
      const r=obCheckYMD(Math.floor(value/10000),Math.floor(value/100)%100,value%100);
      if(r.ok) iso=r.iso;
    } else iso=obExcelSerialToISO(value);
    if(iso) return {provided:true, ok:true, iso, display:String(value)};
    return {provided:true, ok:false, display:String(value), reason:`"${value}" is not a valid date.`, expected:OB_DATE_EXPECTED};
  }
  const s=String(value).trim();
  if(!s) return {provided:false};
  const res=obValidateDateString(s);
  if(res.ok) return {provided:true, ok:true, iso:res.iso, display:s, assumed:!!res.assumed};
  return {provided:true, ok:false, display:s, reason:res.reason, expected:res.expected};
}
function obRenderDateAlert(e){
  return `<div class="ab-row" style="padding:8px 0;border-top:1px solid rgba(220,38,38,.15)">
    <div style="font-weight:600">⚠️ Invalid Date</div>
    <div>Employee: ${obEsc(e.employee)}</div>
    <div>Excel Row: ${obEsc(e.row)}</div>
    <div>Column: ${obEsc(e.column)}</div>
    <div>Value: ${obEsc(e.value)}</div>
    <div>Reason: ${obEsc(e.reason)}</div>
    <div>Expected format: ${obEsc(e.expected)}</div>
  </div>`;
}

/* rows = array of arrays (first row = header). Works identically whether the
   cells came from a CSV (always strings) or an .xlsx sheet (native Date /
   number / string types, thanks to cellDates:true + raw:true above). */
/* ── CHECKLIST COLUMNS → checklist fields ─────────────────────────────────
   Columns in the uploaded sheet whose header matches a checklist item are read
   straight into that employee's offboarding checklist (o.tasks), so the
   checklist is filled automatically. Computed columns (Completion %, Deadline,
   Status) are ignored because the app calculates them itself. */
function classifyObChecklistHeader(h){
  h=h.trim();
  if(h==='completion %'||h==='completion'||h==='deadline'||h==='status'||h==='progress') return 'ignore';
  if(h==='duration'||h==='tenure'||h==='month') return 'ignore'; // Resignation-analysis helper columns — the app calculates these itself
  if(h.includes('terminate')||h.includes('oracle')) return 'oracle';
  if(h.includes('signed')) return 'documents';
  if(h.includes('access')&&h.includes('card')) return 'cards';
  if(h.includes('incentive')) return 'incentives';
  if(h.includes('clearance status')||h==='clearance'||h==='clear status') return 'clearanceStatus';
  if(h.includes('due amount')) return 'dueAmountStatus';
  if(h.includes('amount')) return 'clearAmount';
  if(h.includes('document')&&h.includes('received')) return 'clearDocsReceived';
  if(h.includes('experience letter')||h.includes('exp letter')||h.includes('exp. letter')) return 'clearExpLetter';
  if(h.includes('bank transfer')) return 'clearBankTransfer';
  if(h.includes('clearance reason')||h.includes('unclear reason')||h.includes('not cleared reason')) return 'clearanceReason';
  if(h.includes('medical deletion')) return 'medicalDeletion';
  return null;
}
function obCellIsYes(v){ return ['yes','y','true','1','done','✓','✔','x'].includes(String(v??'').trim().toLowerCase()); }
function obCellIsNo(v){ return ['no','n','false','0'].includes(String(v??'').trim().toLowerCase()); }
function obParseTaskCell(key,raw){
  const t=String(raw??'').trim();
  if(t==='') return undefined; // blank cell → leave the checklist item untouched
  const low=t.toLowerCase();
  switch(key){
    case 'oracle': case 'documents': case 'cards': case 'medicalDeletion':
      return obCellIsYes(t)?true:(obCellIsNo(t)?false:undefined);
    case 'incentives': case 'clearDocsReceived': case 'clearExpLetter': case 'clearBankTransfer':
      return obCellIsYes(t)?'yes':(obCellIsNo(t)?'no':undefined);
    case 'clearanceStatus':
      if(/^(not|un)[\s_-]*clear/.test(low)) return 'not_cleared';
      if(low==='clear'||low==='cleared') return 'clear';
      return undefined;
    case 'dueAmountStatus':
      if(low==='paid') return 'paid';
      if(/^not[\s_-]*yet/.test(low)||low==='unpaid') return 'not_yet';
      return undefined;
    case 'clearAmount': {
      const n=String(t).replace(/[^0-9.\-]/g,'');
      return (n!==''&&!isNaN(Number(n)))?n:undefined;
    }
    case 'clearanceReason': return t;
  }
  return undefined;
}
function processObRows(rows){
  const errBox=document.getElementById('ob-upload-errors');
  if(!rows||rows.length<2){ toast('File appears empty','danger',0); return; }
  const rawHdr=rows[0].map(h=>String(h??'').trim());
  const hdr=rawHdr.map(h=>h.toLowerCase());

  const idx={};
  const taskCols={};            // checklist field → column index
  const ignoredCols=new Set(); // computed columns the app works out itself
  hdr.forEach((h,i)=>{
    const tk=classifyObChecklistHeader(h);
    if(tk==='ignore'){ ignoredCols.add(i); return; }
    if(tk){ if(!(tk in taskCols)) taskCols[tk]=i; return; }
    const cat=classifyObHeader(h);
    if(cat && !(cat in idx)) idx[cat]=i;
  });
  const hasTaskCols=Object.keys(taskCols).length>0;
  const readTasks=cols=>{
    const out={};
    Object.entries(taskCols).forEach(([k,i])=>{
      const v=obParseTaskCell(k,cols[i]);
      if(v!==undefined) out[k]=v;
    });
    return out;
  };
  const {name:ni,lastday:di,dept:dpi,code:ci,reason:ri,gender:gi,grade:gri,rating:rti,joindate:ji,resigndate:rdi,jobtitle:jti,emptype:eti,nationality:nai,division:dvi,cluster:cli,company:coi}=
    {name:-1,lastday:-1,dept:-1,code:-1,reason:-1,gender:-1,grade:-1,rating:-1,joindate:-1,resigndate:-1,jobtitle:-1,emptype:-1,nationality:-1,division:-1,cluster:-1,company:-1, ...idx};
  if(ni<0||di<0){
    toast('File must have "name" and "lastday" columns — download the template for the correct format','danger',0);
    return;
  }
  const knownIdx=new Set(Object.values(idx));
  const extraCols=rawHdr.map((label,i)=>({i,label})).filter(c=>!knownIdx.has(c.i)&&!Object.values(taskCols).includes(c.i)&&!ignoredCols.has(c.i)&&c.label.trim()!=='');
  extraCols.forEach(({label})=>{
    const key=label.trim().toLowerCase().replace(/\s+/g,'_');
    if(!obExtraCols.find(c=>c.key===key)) obExtraCols.push({key,label:label.trim()});
  });
  if(extraCols.length) saveObExtraCols();

  let added=0, updated=0;
  const rowErrors=[];   // non-date row problems (missing name, duplicates)
  const dateErrors=[];  // structured, detailed date problems — every one found, not just the first
  const assumedDates=[]; // numeric dates where day/month order was ambiguous and DD/MM was assumed

  const cellAt=(cols,i)=> i>=0 ? cols[i] : undefined;
  const asText=v=> v===undefined||v===null ? '' : String(v).trim();

  rows.slice(1).forEach((cols,idx2)=>{
    const rowNum=idx2+2; // header is row 1
    if(!cols||cols.every(c=>asText(c)==='')) return; // skip fully blank rows

    const name=asText(cellAt(cols,ni));
    const code=ci>=0?asText(cellAt(cols,ci)):'';
    const employeeLabel=name?(code?`${name} (${code})`:name):(code?`Unnamed (${code})`:`Unnamed — Row ${rowNum}`);

    if(!name){ rowErrors.push(`Row ${rowNum}: missing employee name`); return; }

    const lastdayCell=obResolveDateCell(cellAt(cols,di));
    const joinCell=ji>=0?obResolveDateCell(cellAt(cols,ji)):{provided:false};
    const resignCell=rdi>=0?obResolveDateCell(cellAt(cols,rdi)):{provided:false};

    let rowHasDateError=false;
    if(!lastdayCell.provided){
      dateErrors.push({employee:employeeLabel, row:rowNum, column:OB_DATE_COLS.lastday, value:'(empty)', reason:'This date is required, but the cell is empty.', expected:'DD/MM/YYYY (e.g. 15/03/2026) or YYYY-MM-DD'});
      rowHasDateError=true;
    } else if(!lastdayCell.ok){
      dateErrors.push({employee:employeeLabel, row:rowNum, column:OB_DATE_COLS.lastday, value:lastdayCell.display, reason:lastdayCell.reason, expected:lastdayCell.expected});
      rowHasDateError=true;
    }
    if(joinCell.provided && !joinCell.ok){
      dateErrors.push({employee:employeeLabel, row:rowNum, column:OB_DATE_COLS.joindate, value:joinCell.display, reason:joinCell.reason, expected:joinCell.expected});
      rowHasDateError=true;
    }
    if(resignCell.provided && !resignCell.ok){
      dateErrors.push({employee:employeeLabel, row:rowNum, column:OB_DATE_COLS.resigndate, value:resignCell.display, reason:resignCell.reason, expected:resignCell.expected});
      rowHasDateError=true;
    }

    // Cross-field relationship checks — only meaningful once both sides are individually valid dates.
    if(lastdayCell.ok && resignCell.ok && lastdayCell.iso<resignCell.iso){
      dateErrors.push({
        employee:employeeLabel, row:rowNum,
        column:`${OB_DATE_COLS.lastday} vs ${OB_DATE_COLS.resigndate}`,
        value:`${OB_DATE_COLS.lastday}: ${lastdayCell.iso} · ${OB_DATE_COLS.resigndate}: ${resignCell.iso}`,
        reason:'The Last Working Date cannot be earlier than the Resignation Date.',
        expected:`${OB_DATE_COLS.lastday} must be on or after ${resignCell.iso}.`
      });
      rowHasDateError=true;
    }
    if(lastdayCell.ok && joinCell.ok && lastdayCell.iso<joinCell.iso){
      dateErrors.push({
        employee:employeeLabel, row:rowNum,
        column:`${OB_DATE_COLS.lastday} vs ${OB_DATE_COLS.joindate}`,
        value:`${OB_DATE_COLS.lastday}: ${lastdayCell.iso} · ${OB_DATE_COLS.joindate}: ${joinCell.iso}`,
        reason:'The Offboarding Date cannot be earlier than the Joining Date.',
        expected:`${OB_DATE_COLS.lastday} must be on or after ${joinCell.iso}.`
      });
      rowHasDateError=true;
    }

    if(rowHasDateError) return; // never import a guessed/corrected date — the row waits for a corrected re-upload

    const existing=offboards.find(o=>o.name===name&&o.lastDay===lastdayCell.iso);
    if(existing){
      // Same employee already loaded → the master sheet refreshes BOTH the checklist and the
      // analysis fields (gender, grade, division, hire date…). Blank cells never overwrite existing data.
      const setIf=(key,val)=>{ if(val!==undefined && val!=='' && val!==null) existing[key]=val; };
      if(hasTaskCols) existing.tasks=Object.assign(obDefaultTasks(),existing.tasks||{},readTasks(cols));
      setIf('code',ci>=0?asText(cellAt(cols,ci)):'');
      setIf('reason',ri>=0?asText(cellAt(cols,ri)):'');
      setIf('gender',gi>=0?normalizeGender(asText(cellAt(cols,gi))):'');
      setIf('grade',gri>=0?asText(cellAt(cols,gri)):'');
      setIf('rating',rti>=0?asText(cellAt(cols,rti)):'');
      setIf('jobTitle',jti>=0?asText(cellAt(cols,jti)):'');
      setIf('employmentType',eti>=0?asText(cellAt(cols,eti)):'');
      setIf('nationality',nai>=0?asText(cellAt(cols,nai)):'');
      setIf('division',dvi>=0?asText(cellAt(cols,dvi)):'');
      setIf('cluster',cli>=0?asText(cellAt(cols,cli)):'');
      setIf('company',coi>=0?asText(cellAt(cols,coi)):'');
      setIf('dept',(dpi>=0?asText(cellAt(cols,dpi)):'')||(dvi>=0?asText(cellAt(cols,dvi)):''));
      if(joinCell.ok) existing.joinDate=joinCell.iso;
      if(resignCell.ok) existing.resignDate=resignCell.iso;
      existing.extra=existing.extra||{};
      extraCols.forEach(({i,label})=>{
        const v=asText(cellAt(cols,i));
        if(v!=='') existing.extra[label.trim().toLowerCase().replace(/\s+/g,'_')]=v;
      });
      updated++;
      return;
    }

    [[lastdayCell,OB_DATE_COLS.lastday],[joinCell,OB_DATE_COLS.joindate],[resignCell,OB_DATE_COLS.resigndate]].forEach(([c,col])=>{
      if(c.assumed) assumedDates.push({employee:employeeLabel,row:rowNum,column:col,value:c.display,iso:c.iso});
    });
    const division=dvi>=0?asText(cellAt(cols,dvi)):'';
    const cluster=cli>=0?asText(cellAt(cols,cli)):'';
    const company=coi>=0?asText(cellAt(cols,coi)):'';
    const dept=(dpi>=0?asText(cellAt(cols,dpi)):'')||division; // a sheet with only a Division column still fills the department chart
    const reason=ri>=0?asText(cellAt(cols,ri)):'';
    const gender=gi>=0?normalizeGender(asText(cellAt(cols,gi))):'';
    const grade=gri>=0?asText(cellAt(cols,gri)):'';
    const rating=rti>=0?asText(cellAt(cols,rti)):'';
    const jobTitle=jti>=0?asText(cellAt(cols,jti)):'';
    const employmentType=eti>=0?asText(cellAt(cols,eti)):'';
    const nationality=nai>=0?asText(cellAt(cols,nai)):'';
    const extra={};
    extraCols.forEach(({i,label})=>{
      const key=label.trim().toLowerCase().replace(/\s+/g,'_');
      extra[key]=asText(cellAt(cols,i));
    });

    offboards.push({
      id:'o'+Date.now()+Math.random(),
      name, lastDay:lastdayCell.iso, dept, code, reason, gender, grade, rating,
      joinDate: joinCell.ok?joinCell.iso:null,
      resignDate: resignCell.ok?resignCell.iso:null,
      jobTitle, employmentType, nationality, division, cluster, company, extra,
      tasks:Object.assign(obDefaultTasks(),readTasks(cols)),
    });
    added++;
  });

  saveOb(); renderOffboarding(); renderSidebar();

  let html='';
  if(dateErrors.length){
    html+=`<div class="alert-wrap" style="margin-bottom:14px">
      <div class="ab ab-danger">
        <div class="ab-head">⚠️ ${dateErrors.length} invalid date${dateErrors.length!==1?'s':''} found — nothing was auto-corrected or guessed</div>
        ${dateErrors.map(obRenderDateAlert).join('')}
        <div class="ab-row" style="margin-top:8px;font-weight:500">Please correct these dates in the Excel file and upload it again. Rows with a date problem were not imported.</div>
      </div>
    </div>`;
  } else {
    html+=`<div class="alert-wrap" style="margin-bottom:14px">
      <div class="ab" style="background:var(--green-bg);border:1px solid #A7F3D0">
        <div class="ab-head" style="color:var(--green-text)">✅ All dates have been successfully validated.</div>
      </div>
    </div>`;
  }
  if(assumedDates.length){
    html+=`<div class="alert-wrap" style="margin-bottom:14px">
      <div class="ab ab-warn">
        <div class="ab-head">ℹ️ ${assumedDates.length} date${assumedDates.length!==1?'s':''} could be day/month or month/day — read as DD/MM/YYYY</div>
        ${assumedDates.slice(0,15).map(e=>`<div class="ab-row">Row ${obEsc(e.row)} · ${obEsc(e.employee)} · ${obEsc(e.column)}: ${obEsc(e.value)} → ${obEsc(e.iso)}</div>`).join('')}
        ${assumedDates.length>15?`<div class="ab-row">…and ${assumedDates.length-15} more</div>`:''}
        <div class="ab-row" style="margin-top:6px">If any of these should be month/day, write the date unambiguously (e.g. 5 Apr 2026) and re-upload.</div>
      </div>
    </div>`;
  }
  if(rowErrors.length){
    html+=`<div class="alert-wrap" style="margin-bottom:18px">
      <div class="ab ab-warn">
        <div class="ab-head">⚠ ${rowErrors.length} row${rowErrors.length!==1?'s':''} could not be imported</div>
        ${rowErrors.map(e=>`<div class="ab-row">${obEsc(e)}</div>`).join('')}
      </div>
    </div>`;
  }
  errBox.innerHTML=html;

  // The same sheet feeds the Resignation Analysis: refresh it, or jump to Offboarding to show date problems
  const onAnalysis=document.getElementById('pg-resignation')?.classList.contains('active');
  if(onAnalysis){ if(dateErrors.length) nav('offboarding'); else renderResignationAnalysis(); }
  if(dateErrors.length){
    toast(`⚠️ ${dateErrors.length} invalid date${dateErrors.length!==1?'s':''} found — see details on the Offboarding page`,'danger',0);
  } else if(rowErrors.length){
    toast(`✓ Imported ${added} offboarding record${added!==1?'s':''}${updated?` · updated ${updated} existing record${updated!==1?"s":""}`:''} — ${rowErrors.length} row${rowErrors.length!==1?'s':''} had errors`,'warn',0);
  } else {
    toast(`✓ Imported ${added} offboarding record${added!==1?'s':''}${updated?` · updated ${updated} existing record${updated!==1?"s":""}`:''} — all dates validated`);
  }
}
function toggleObTask(id,key){
  const o=offboards.find(o=>o.id===id); if(!o) return;
  o.tasks[key]=!o.tasks[key];
  saveOb(); renderOffboarding(); renderSidebar();
}
/* ─── ONBOARDING EXCEL/CSV BULK UPLOAD ───
   Columns mirror the "Add new hire" form: Full name*, Department, Job title,
   Employee code, Hiring date*, HR person. */
function downloadHireTemplate(){
  const rows=[
    ['name','dept','title','code','startdate','hr'],
    ['Ahmed Hassan','Finance','Accountant','EMP-001','2026-08-15','Sara Mahmoud'],
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:20},{wch:16},{wch:18},{wch:12},{wch:14},{wch:18}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Onboarding');
  XLSX.writeFile(wb,'onboarding_upload_template.xlsx');
}
function handleHireFile(file){
  if(!file) return;
  document.getElementById('h-upload-errors').innerHTML='';
  smartImport(file,'hire').then(res=>{
    if(!res) return;
    parseHireCSV(XLSX.utils.sheet_to_csv(res.ws,{dateNF:'yyyy-mm-dd'}));
    // Checklist status comes from the sheet named "Checklist" in the same workbook (if there is one)
    applyChecklistSheet(res.wb, res.sheetName);
  });
  document.getElementById('file-in-h').value='';
}
/* Classifies a lower-cased header string into a known onboarding field —
   same six fields as the "Add new hire" form. */

/* ─── ONBOARDING CHECKLIST SHEET ───
   Reads the sheet named "Checklist" from the uploaded workbook. Each row is matched to an
   existing hire by Employee Code (or Full name), and each column whose header equals a
   checklist task label sets that task's status. Blank cells are left unchanged. */
function normKey(s){ return String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,''); }
function checklistCell(task, raw){
  const v=String(raw==null?'':raw).trim().toLowerCase();
  if(!v) return null;
  const yes=/^(yes|y|true|1|x|done|✓|✔|☑)$/.test(v), no=/^(no|n|false|0)$/.test(v);
  if(task.type==='yesno'){ return yes?{[task.id]:'yes'}:no?{[task.id]:'no'}:null; }
  if(task.type==='eligibility'){
    if(/^not\s*eligible$|^n\/?a$/.test(v)) return {[task.id+'_elig']:'no',[task.id]:null};
    if(/not\s*yet|not\s*received|pending/.test(v)) return {[task.id+'_elig']:'yes',[task.id]:'notyet'};
    if(/received/.test(v)) return {[task.id+'_elig']:'yes',[task.id]:'received'};
    if(/^eligible$/.test(v)) return {[task.id+'_elig']:'yes'};
    return null;
  }
  return yes?{[task.id]:true}:no?{[task.id]:false}:null;
}
function applyChecklistSheet(wb, mainName){
  const sn=wb.SheetNames.find(n=>normKey(n)==='checklist');
  if(!sn) return;
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:''});
  if(rows.length<2) return;
  const hdr=rows[0].map(normKey);
  const nameCol=hdr.findIndex(h=>h==='fullname'||h==='name'||h==='employeename');
  const codeCol=hdr.findIndex(h=>h==='employeecode'||h==='empcode'||h==='code');
  // Eligibility-type tasks (Laptop, Mobile Line) may be one column ("Laptop") or two columns:
  // "Laptop - Eligibility" (Eligible / Not Eligible) + "Laptop - Received" (Received / Not Yet)
  const taskCols=TASKS.map(t=>{
    const k=normKey(t.label);
    return {t, i:hdr.indexOf(k), ie:hdr.indexOf(k+'eligibility'), ir:hdr.indexOf(k+'received')};
  }).filter(x=>x.i>=0||(x.t.type==='eligibility'&&(x.ie>=0||x.ir>=0)));
  if(!taskCols.length){ toast('Checklist sheet found, but none of its column headers match the checklist items','warn',0); return; }
  // Fallback: the Checklist sheet mirrors the employee sheet row-for-row. If its name/code cells
  // are empty (e.g. formulas that were never calculated), read them from the same row of the
  // employee sheet instead, so the laptop / mobile answers are still matched to the right person.
  let mainRows=null, mName=-1, mCode=-1;
  if(mainName && wb.Sheets[mainName]){
    mainRows=XLSX.utils.sheet_to_json(wb.Sheets[mainName],{header:1,raw:true,defval:''});
    const mh=(mainRows[0]||[]).map(normKey);
    mName=mh.findIndex(h=>h==='fullname'||h==='name'||h==='employeename');
    mCode=mh.findIndex(h=>h==='employeecode'||h==='empcode'||h==='code');
  }
  let updated=0; const unmatched=[];
  rows.slice(1).forEach((r,ri)=>{
    let name=nameCol>=0?String(r[nameCol]||'').trim():'';
    let code=codeCol>=0?String(r[codeCol]||'').trim():'';
    if(!name&&!code&&mainRows&&mainRows[ri+1]){
      const mr=mainRows[ri+1];
      name=mName>=0?String(mr[mName]||'').trim():'';
      code=mCode>=0?String(mr[mCode]||'').trim():'';
    }
    if(!name&&!code) return;
    const h=(code&&hires.find(x=>x.code&&normKey(x.code)===normKey(code)))||
            (name&&hires.find(x=>normKey(x.name)===normKey(name)));
    if(!h){ unmatched.push(name||code); return; }
    if(!h.tasks) h.tasks={};
    let touched=false;
    taskCols.forEach(({t,i,ie,ir})=>{
      let upd=null;
      if(t.type==='eligibility' && (ie>=0||ir>=0)){
        // Step 1: eligibility column. Step 2: received column (ignored when Not Eligible)
        const e=ie>=0?checklistCell(t,r[ie]):null;
        if(e) upd=Object.assign({},e);
        const notElig=upd&&upd[t.id+'_elig']==='no';
        const rc=(ir>=0&&!notElig)?checklistCell(t,r[ir]):null;
        if(rc) upd=Object.assign(upd||{},rc);
      } else {
        upd=checklistCell(t,r[i]);
      }
      if(upd){ Object.assign(h.tasks,upd); touched=true; }
    });
    if(touched) updated++;
  });
  save(); renderDashboard(); renderSidebar();
  if(curView==='detail') renderDetail();
  toast(`✓ Checklist sheet applied to ${updated} employee${updated!==1?'s':''}`+(unmatched.length?` — ${unmatched.length} row${unmatched.length!==1?'s':''} had no matching employee: ${unmatched.slice(0,5).join(', ')}`:''), unmatched.length?'warn':undefined, unmatched.length?0:undefined);
}
function classifyHireHeader(h){
  if(h.includes('full name')||h.includes('employee name')||h.includes('name')) return 'name';
  if(h.includes('dept')||h.includes('department')||h.includes('division')) return 'dept';
  if(h.includes('job title')||h.includes('designation')||h.includes('position')||h.includes('title')) return 'title';
  if(h.includes('employee code')||h.includes('emp code')||h.includes('code')||h.includes('employee id')) return 'code';
  if(h.includes('hiring date')||h.includes('start date')||h.includes('joining date')||h.includes('date of joining')||h.includes('doj')||h.includes('join')) return 'startDate';
  if(h.includes('hr person')||h.includes('hr rep')||h.includes('hr manager')||h.includes('recruiter')||h.includes('hr')) return 'hr';
  return null;
}
function parseHireCSV(text){
  const lines=text.trim().split(/\r?\n/).filter(l=>l.trim()!=='');
  const errBox=document.getElementById('h-upload-errors');
  if(lines.length<2){ toast('File appears empty','danger',0); return; }
  const rawHdr=splitCSVLine(lines[0]).map(h=>h.trim().replace(/"/g,''));
  const hdr=rawHdr.map(h=>h.toLowerCase());
  const idx={};
  hdr.forEach((h,i)=>{
    const cat=classifyHireHeader(h);
    if(cat && !(cat in idx)) idx[cat]=i;
  });
  const {name:ni,dept:dpi,title:ti,code:ci,startDate:si,hr:hri}=
    {name:-1,dept:-1,title:-1,code:-1,startDate:-1,hr:-1, ...idx};
  if(ni<0||si<0){
    toast('File must have "name" and "hiring date" columns — download the template for the correct format','danger',0);
    return;
  }
  let added=0;
  const errors=[];
  lines.slice(1).forEach((ln,idx2)=>{
    const rowNum=idx2+2; // account for header row, 1-indexed
    const cols=splitCSVLine(ln);
    if(cols.every(c=>!c.trim())) return; // skip fully blank rows
    const name=(cols[ni]||'').trim();
    const rawDate=(cols[si]||'').trim();
    const dept=dpi>=0?(cols[dpi]||'').trim():'';
    const title=ti>=0?(cols[ti]||'').trim():'';
    const code=ci>=0?(cols[ci]||'').trim():'';
    const hr=hri>=0?(cols[hri]||'').trim():'';

    if(!name){ errors.push(`Row ${rowNum}: missing employee name`); return; }
    if(!rawDate){ errors.push(`Row ${rowNum} (${name}): missing hiring date`); return; }
    const parsed=parseDate(rawDate);
    if(!parsed){ errors.push(`Row ${rowNum} (${name}): "${rawDate}" is not a recognizable date`); return; }
    if(hires.find(h=>h.name===name&&h.startDate===parsed)){
      errors.push(`Row ${rowNum} (${name}): duplicate record already exists — skipped`);
      return;
    }
    const h={id:'h'+Date.now()+Math.random(), name, dept, title, code, startDate:parsed, hr, tasks:{}, notes:{}, createdAt:new Date().toISOString()};
    hires.push(h);
    syncContractFromHire(h); // keep the linked Contract Renewal record in sync, same as manually-added hires
    added++;
  });

  save(); renderDashboard(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();

  if(errors.length){
    errBox.innerHTML=`<div class="alert-wrap" style="margin-bottom:18px">
      <div class="ab ab-danger">
        <div class="ab-head">⚠ ${errors.length} row${errors.length!==1?'s':''} could not be imported</div>
        ${errors.map(e=>`<div class="ab-row">${e}</div>`).join('')}
      </div>
    </div>`;
    toast(`✓ Imported ${added} hire${added!==1?'s':''} — ${errors.length} row${errors.length!==1?'s':''} had errors`,'warn',0);
  } else {
    toast(`✓ Imported ${added} hire${added!==1?'s':''} from file`);
  }
}
/* ─── OVERTIME EXCEL/CSV BULK UPLOAD ─── */
function downloadOtTemplate(){
  const rows=[
    ['name','jobtitle','division','organization','code','date','actualhours','normalhours','saturdayhours','officialdayshours','hours','monthlysalary','rate','amount','status','reason'],
    ['Ahmed Hassan','Accountant','Finance Division','Finance','EMP-001','2026-07-14','8','3.5','0','0','','20000','','','Approved','Month-end closing'],
    ['Sara Mahmoud','HR Specialist','Corporate Services','HR','EMP-014','2026-07-16','8','0','5','0','','16000','83.3','','Pending','Payroll cycle'],
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:20},{wch:16},{wch:18},{wch:16},{wch:12},{wch:12},{wch:11},{wch:11},{wch:13},{wch:16},{wch:8},{wch:14},{wch:8},{wch:10},{wch:12},{wch:24}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Overtime');
  XLSX.writeFile(wb,'overtime_upload_template.xlsx');
  toast('✓ Template downloaded — the "organization" column is used as the Department field throughout the dashboard. Fill in "normalhours"/"saturdayhours"/"officialdayshours" for a full category breakdown, or just "hours" for a simple total (treated as Normal).');
}
function otResetAll(){
  if(!overtimeRecords.length){ toast('No overtime data to reset'); return; }
  if(!confirm('Reset ALL Overtime Analysis data? This permanently removes every uploaded overtime record, clears every chart, table and filter on this page, and cannot be undone.')) return;
  overtimeRecords=[];
  setOtValueFormat(null);
  otPeriod='monthly'; otMonthSel=null; otQuarterSel=null; otYearSel=null; otDrillRecords=[];
  otFilterDept=''; otFilterEmp=''; otFilterTitle=''; otFilterDivision=''; otFilterYear='';
  document.getElementById('ot-upload-errors').innerHTML='';
  saveOt();
  renderOvertimeAnalysis();
  renderSidebar();
  toast('✓ Overtime Analysis data has been reset');
}
/* ── Detect the Excel number/currency format ('z' format code, e.g. "$#,##0.00", "#,##0")
   applied to the Overtime value column, so computed totals can be displayed in the same
   format the source workbook uses — never a format Claude/the app invents. ── */
function otValueColIndex(ws){
  if(!ws['!ref']) return -1;
  const range=XLSX.utils.decode_range(ws['!ref']);
  const hdrs=[];
  for(let c=range.s.c;c<=range.e.c;c++){ const cell=ws[XLSX.utils.encode_cell({r:range.s.r,c})]; hdrs.push(cell&&cell.v!=null?String(cell.v):''); }
  const i=otFindValueCol(hdrs);
  return i<0?-1:range.s.c+i;
}
/* Write the exact stored number of every cell in the Overtime Value column into its display text,
   so the CSV conversion never rounds it to the cell's display format (e.g. 1,234.56 shown as 1,235). */
function otForceRawValueColumn(ws){
  try{
    const colIdx=otValueColIndex(ws); if(colIdx<0) return;
    const range=XLSX.utils.decode_range(ws['!ref']);
    for(let r=range.s.r+1;r<=range.e.r;r++){
      const cell=ws[XLSX.utils.encode_cell({r,c:colIdx})];
      if(cell&&cell.t==='n'&&typeof cell.v==='number') cell.w=String(cell.v);
    }
  }catch(e){}
}
function detectOtValueFormat(ws){
  try{
    if(!ws['!ref']) return null;
    const range=XLSX.utils.decode_range(ws['!ref']);
    const headerRow=range.s.r;
    const colIdx=otValueColIndex(ws);
    if(colIdx<0) return null;
    for(let r=headerRow+1;r<=range.e.r;r++){
      const cell=ws[XLSX.utils.encode_cell({r,c:colIdx})];
      if(cell&&typeof cell.v==='number'&&cell.z&&cell.z!=='General') return cell.z;
    }
  }catch(e){ /* fall through to no detected format */ }
  return null;
}
function setOtValueFormat(fmt){
  otValueNumFmt=fmt||null;
  if(otValueNumFmt) localStorage.setItem('hr3_ot_value_fmt', otValueNumFmt);
  else localStorage.removeItem('hr3_ot_value_fmt');
  saveSettings();
}
function handleOtFile(file){
  if(!file) return;
  document.getElementById('ot-upload-errors').innerHTML='';
  smartImport(file,'overtime').then(res=>{
    if(!res) return;
    if(res.isText) setOtValueFormat(null); // plain text files carry no cell formatting — fall back to a plain number display
    else { setOtValueFormat(detectOtValueFormat(res.ws)); otForceRawValueColumn(res.ws); }
    parseOtCSV(XLSX.utils.sheet_to_csv(res.ws));
  });
  document.getElementById('file-in-ot').value='';
}
/* ── Overtime Value column lookup. An explicit "Overtime Value" header (or OT Value / Overtime Amount)
   always wins over looser matches such as "Salary" or "Cost", so the KPI reads exactly that column. ── */
const OT_VALUE_HEADER_RE=/^(total\s+)?(overtime|ot)\s*(value|amount)\b/;
function otNormHeader(h){ return String(h==null?'':h).replace(/\s+/g,' ').trim().toLowerCase(); }
function otFindValueCol(headers){ // headers: array of raw header strings → column index or -1
  const norm=headers.map(otNormHeader);
  let i=norm.findIndex(h=>OT_VALUE_HEADER_RE.test(h));
  if(i>=0) return i;
  return norm.findIndex(h=>classifyOtHeader(h)==='salary');
}
/* Robust number reader: handles "1,234.50", "$1,200", "EGP 1 200", "(500)", "12%" — parseFloat alone
   would read "1,234.50" as 1. */
function otParseNum(raw){
  if(raw==null) return NaN;
  let t=String(raw).trim();
  if(!t) return NaN;
  let neg=false;
  if(/^\(.*\)$/.test(t)){ neg=true; t=t.slice(1,-1); }
  t=t.replace(/[^0-9.,\-]/g,'');
  if(/^-/.test(t)){ neg=!neg; t=t.slice(1); }
  if(/,/.test(t)&&/\./.test(t)){ t=t.lastIndexOf(',')>t.lastIndexOf('.') ? t.replace(/\./g,'').replace(',','.') : t.replace(/,/g,''); }
  else if(/,/.test(t)){ t=/^\d{1,3}(,\d{3})+$/.test(t) ? t.replace(/,/g,'') : t.replace(',','.'); }
  const v=parseFloat(t);
  return isNaN(v)?NaN:(neg?-v:v);
}
function classifyOtHeader(h){
  h=otNormHeader(h);
  if(h.includes('monthly salary')||h.includes('monthly pay')||h.includes('basic salary')||h==='monthlysalary') return 'monthlysalary';
  if(h.includes('job title')||h.includes('jobtitle')||h.includes('designation')||h.includes('position')||h.includes('title')) return 'jobtitle';
  if(h.includes('status')) return 'status';
  if(h.includes('name')) return 'name';
  if(h.includes('actual hour')||h==='actualhours'||h.includes('actual hrs')||h.includes('worked hour')) return 'actualhours';
  if(h.includes('normal hour')||h==='normalhours'||h.includes('normal hrs')||h.includes('normal ot')) return 'normalhours';
  if(h.includes('saturday')||h.includes('weekend')||h.includes('rest day')||h.includes('restday')||h.includes('sat day')||h.includes('sat.days')) return 'saturdayhours';
  if(h.includes('official day')||h==='officialdayshours'||h.includes('official holiday')||h.includes('holiday hour')) return 'officialhours';
  if(h.includes('salary')||h.includes('amount')||h.includes('pay')||h.includes('cost')||h.includes('overtime value')||h.includes('ot value')||/^(total\s+)?overtime(\s+(value|amount|paid|due|total|earned))?$/.test(h.trim())) return 'salary';
  if(h.includes('rate')) return 'rate';
  if(h.includes('hour')||h.includes('ot hrs')||h.includes('overtime hrs')) return 'hours';
  if(h.includes('division')) return 'division';
  if(h.includes('org')) return 'org'; /* Organization column — this is the ONLY source for the Department field throughout the Overtime module. A literal "Department" header in the sheet is intentionally NOT captured below, so it can never be used as the Department source. */
  if(h.includes('code')||h.includes('emp')||h.includes('id')) return 'code';
  if(h.includes('reason')||h.includes('note')||h.includes('remark')) return 'reason';
  if(h.includes('date')||h.trim()==='month'||h.includes('month')) return 'date';
  return null;
}
function parseOtCSV(text){
  const lines=text.trim().split(/\r?\n/).filter(l=>l.trim()!=='');
  const errBox=document.getElementById('ot-upload-errors');
  if(lines.length<2){ toast('File appears empty','danger',0); return; }
  const rawHdr=splitCSVLine(lines[0]).map(h=>h.trim().replace(/"/g,''));
  const hdr=rawHdr.map(h=>h.toLowerCase());
  const idx={};
  hdr.forEach((h,i)=>{
    const cat=classifyOtHeader(h);
    if(cat && !(cat in idx)) idx[cat]=i;
  });
  { const vi=otFindValueCol(rawHdr); if(vi>=0) idx.salary=vi; } // Overtime Value column always wins
  const {name:ni,date:di,hours:hi,division:dvi,code:ci,reason:ri,salary:si,rate:rai,jobtitle:jti,monthlysalary:msi,status:sti,actualhours:ai,normalhours:nhi,saturdayhours:shi,officialhours:ohi,org:oi}=
    {name:-1,date:-1,hours:-1,division:-1,code:-1,reason:-1,salary:-1,rate:-1,jobtitle:-1,monthlysalary:-1,status:-1,actualhours:-1,normalhours:-1,saturdayhours:-1,officialhours:-1,org:-1, ...idx};
  if(ni<0||di<0||(hi<0&&nhi<0&&shi<0&&ohi<0)){
    toast('File must have a "name", "date" column, and either "hours" or a Normal/Saturday/Official Days hours breakdown — download the template for the correct format','danger',0);
    return;
  }
  // This sheet is treated as the single source of truth: every valid row found here becomes
  // the complete Overtime dataset, replacing whatever was loaded before (not merged/appended).
  const newRecords=[];
  const seen=new Set(); // dedupe rows within this same sheet (name+date+hours)
  const errors=[];
  lines.slice(1).forEach((ln,idx2)=>{
    const rowNum=idx2+2;
    const cols=splitCSVLine(ln);
    if(cols.every(c=>!c.trim())) return;
    const name=(cols[ni]||'').trim();
    const rawDate=(cols[di]||'').trim();
    const rawHours=hi>=0?(cols[hi]||'').trim():'';
    const division=dvi>=0?(cols[dvi]||'').trim():'';
    const org=oi>=0?(cols[oi]||'').trim():''; // Organization column — the sole source of the Department field
    const code=ci>=0?(cols[ci]||'').trim():'';
    const reason=ri>=0?(cols[ri]||'').trim():'';
    const jobTitle=jti>=0?(cols[jti]||'').trim():'';
    const status=sti>=0?(cols[sti]||'').trim():'';
    const rawSalary=si>=0?(cols[si]||'').trim():'';
    const rawRate=rai>=0?(cols[rai]||'').trim():'';
    const rawMonthlySalary=msi>=0?(cols[msi]||'').trim():'';
    const rawActual=ai>=0?(cols[ai]||'').trim():'';
    const rawNormal=nhi>=0?(cols[nhi]||'').trim():'';
    const rawSaturday=shi>=0?(cols[shi]||'').trim():'';
    const rawOfficial=ohi>=0?(cols[ohi]||'').trim():'';

    if(!name){ errors.push(`Row ${rowNum}: missing employee name`); return; }
    if(!rawDate){ errors.push(`Row ${rowNum} (${name}): missing date`); return; }
    const parsed=parseDate(rawDate);
    if(!parsed){ errors.push(`Row ${rowNum} (${name}): "${rawDate}" is not a recognizable date`); return; }

    // Overtime hours category breakdown (rule 10: never double-counted). If the sheet gives a
    // Normal/Saturday/Official Days breakdown, that is the source of truth and "hours" is derived
    // as their sum. Otherwise, a plain "hours" total is treated as Normal-category overtime.
    let normalHours=0, saturdayHours=0, officialHours=0;
    if(rawNormal){ const v=parseFloat(rawNormal); if(!isNaN(v)&&v>0) normalHours=v; }
    if(rawSaturday){ const v=parseFloat(rawSaturday); if(!isNaN(v)&&v>0) saturdayHours=v; }
    if(rawOfficial){ const v=parseFloat(rawOfficial); if(!isNaN(v)&&v>0) officialHours=v; }
    let hours;
    if(normalHours+saturdayHours+officialHours>0){
      hours=+(normalHours+saturdayHours+officialHours).toFixed(4);
    } else {
      hours=parseFloat(rawHours);
      if(!rawHours||isNaN(hours)||hours<=0){ errors.push(`Row ${rowNum} (${name}): no valid overtime hours found (need "hours" or a Normal/Saturday/Official Days value)`); return; }
      normalHours=hours; // legacy single-total sheets: treated as Normal-category overtime
    }
    // Actual (worked) hours — basis for the overtime-percentage calculations; 0/blank shows as N/A downstream.
    let actualHours=0;
    if(rawActual){ const v=parseFloat(rawActual); if(!isNaN(v)&&v>0) actualHours=v; }
    // Monthly salary (basis for OT % of salary and, when no explicit rate is given, the hourly OT rate).
    let monthlySalary=0;
    if(rawMonthlySalary){ const v=parseFloat(rawMonthlySalary); if(!isNaN(v)) monthlySalary=v; }
    // Explicit hourly OT rate, if the sheet gives one.
    let otRate=0;
    if(rawRate){ const r=parseFloat(rawRate); if(!isNaN(r)) otRate=r; }
    // Explicit OT amount overrides automatic calculation; otherwise it is always derived
    // (category hours × explicit rate × multiplier, or hours × (monthly salary ÷ standard monthly hours)).
    let salary=0;
    if(rawSalary){ const v=otParseNum(rawSalary); if(!isNaN(v)) salary=v; }
    const dupKey=`${name}|${parsed}|${hours}|${salary}`; // value included so distinct Overtime Value rows are never dropped
    if(seen.has(dupKey)){
      errors.push(`Row ${rowNum} (${name}): duplicate row within this file — skipped`);
      return;
    }
    seen.add(dupKey);
    newRecords.push({id:'ot'+Date.now()+Math.random()+idx2, name, date:parsed, hours, normalHours, saturdayHours, officialHours, actualHours, salary, otRate, monthlySalary, division, org, code, jobTitle, status, reason});
  });

  if(!newRecords.length){
    errBox.innerHTML=`<div class="alert-wrap" style="margin-bottom:18px"><div class="ab ab-danger">
      <div class="ab-head">⚠ Nothing to import</div>
      <div class="ab-row">No valid rows were found in this file — the existing dashboard data was left unchanged.</div>
      ${errors.length?errors.map(e=>`<div class="ab-row">${e}</div>`).join(''):''}
    </div></div>`;
    toast('No valid rows found — existing data unchanged','danger',0);
    return;
  }

  if(overtimeRecords.length){
    const ok=confirm(`This file has ${newRecords.length} valid record${newRecords.length!==1?'s':''} and will replace the ${overtimeRecords.length} overtime record${overtimeRecords.length!==1?'s':''} currently in the dashboard (the sheet is treated as the full, current dataset). Continue?`);
    if(!ok){ toast('Import cancelled — existing data unchanged'); return; }
  }

  overtimeRecords=newRecords;
  saveOt(); renderOvertimeAnalysis(); renderSidebar();

  if(errors.length){
    errBox.innerHTML=`<div class="alert-wrap" style="margin-bottom:18px">
      <div class="ab ab-danger">
        <div class="ab-head">⚠ ${errors.length} row${errors.length!==1?'s':''} could not be imported</div>
        ${errors.map(e=>`<div class="ab-row">${e}</div>`).join('')}
      </div>
    </div>`;
    toast(`✓ Loaded ${newRecords.length} record${newRecords.length!==1?'s':''} — now the dashboard's full dataset — ${errors.length} row${errors.length!==1?'s':''} had errors`,'warn',0);
  } else {
    toast(`✓ Loaded ${newRecords.length} record${newRecords.length!==1?'s':''} from file — dashboard fully refreshed`);
  }
}
function setObYesNo(id,key,val){
  const o=offboards.find(o=>o.id===id); if(!o) return;
  o.tasks[key]=(o.tasks[key]===val)?null:val;
  saveOb(); renderOffboarding(); renderSidebar();
}
function setObAmount(id,key,val){
  const o=offboards.find(o=>o.id===id); if(!o) return;
  o.tasks[key]=val;
  saveOb(); renderOffboarding(); renderSidebar();
}
function setObReason(id,key,val){
  const o=offboards.find(o=>o.id===id); if(!o) return;
  o.tasks[key+'Reason']=val;
  saveOb(); renderOffboarding(); renderSidebar();
}
function setObText(id,key,val){
  const o=offboards.find(o=>o.id===id); if(!o) return;
  o.tasks[key]=val;
  saveOb(); renderOffboarding(); renderSidebar();
}
function removeOb(id){
  if(!confirm('Remove this offboarding record?')) return;
  offboards=offboards.filter(o=>o.id!==id); saveOb(); renderOffboarding(); renderSidebar();
}
function obToggleSelectMode(){
  obSelectMode=!obSelectMode;
  if(!obSelectMode) obSelected.clear();
  if(obSelectMode){ obColSelectMode=false; obColSelected.clear(); }
  renderOffboarding();
}
function obToggleRow(id){
  if(obSelected.has(id)) obSelected.delete(id); else obSelected.add(id);
  renderOffboarding();
}
function obSelectAllVisible(checked){
  offboards.filter(o=>pfMatch('ob',o.lastDay)).forEach(o=> checked ? obSelected.add(o.id) : obSelected.delete(o.id));
  renderOffboarding();
}
function obRemoveSelected(){
  const n=obSelected.size;
  if(!n){ toast('No rows selected'); return; }
  if(!confirm(`Remove ${n} selected offboarding record${n!==1?'s':''}? This will delete them from the currently loaded dataset and cannot be undone.`)) return;
  offboards=offboards.filter(o=>!obSelected.has(o.id));
  obSelected.clear(); obSelectMode=false;
  saveOb(); renderOffboarding(); renderSidebar();
  toast(`✓ Removed ${n} offboarding record${n!==1?'s':''}`);
}
function obToggleColSelectMode(){
  obColSelectMode=!obColSelectMode;
  if(!obColSelectMode) obColSelected.clear();
  if(obColSelectMode){ obSelectMode=false; obSelected.clear(); }
  renderOffboarding();
}
function obToggleColSelect(key){
  if(obColSelected.has(key)) obColSelected.delete(key); else obColSelected.add(key);
  renderOffboarding();
}
function obRemoveSelectedCols(){
  const n=obColSelected.size;
  if(!n){ toast('No columns selected'); return; }
  const labels=obExtraCols.filter(c=>obColSelected.has(c.key)).map(c=>c.label).join(', ');
  if(!confirm(`Remove ${n} column${n!==1?'s':''} (${labels}) from every offboarding record? This cannot be undone.`)) return;
  offboards.forEach(o=>{ if(o.extra){ obColSelected.forEach(key=>delete o.extra[key]); } });
  obExtraCols=obExtraCols.filter(c=>!obColSelected.has(c.key));
  obColSelected.clear(); obColSelectMode=false;
  saveOb(); saveObExtraCols(); renderOffboarding(); renderSidebar();
  toast(`✓ Removed ${n} column${n!==1?'s':''}`);
}
function obResetAll(){
  if(!offboards.length){ toast('No offboarding data to reset'); return; }
  if(!confirm('Reset ALL offboarding data? This permanently removes every uploaded offboarding record and cannot be undone.')) return;
  offboards=[]; obExtraCols=[];
  obSelected.clear(); obSelectMode=false; obColSelected.clear(); obColSelectMode=false;
  saveOb(); saveObExtraCols(); renderOffboarding(); renderSidebar();
  toast('✓ Offboarding data has been reset');
}
function ynBtn(id,key,val,label,active){
  return `<button type="button" class="btn btn-sm ${active?(val==='yes'?'btn-green':'btn-danger'):'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setObYesNo('${id}','${key}','${val}')">${label}</button>`;
}
function optBtn(id,key,val,label,active,activeCls){
  return `<button type="button" class="btn btn-sm ${active?activeCls:'btn-outline'}" style="padding:3px 10px;font-size:11px" onclick="setObYesNo('${id}','${key}','${val}')">${label}</button>`;
}
function obTaskRow(o,t){
  const v=o.tasks[t.key];
  if(t.type==='checkbox'){
    return `<label style="display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer">
      <input type="checkbox" ${v?'checked':''} onchange="toggleObTask('${o.id}','${t.key}')"> ${t.label}
    </label>`;
  }
  if(t.type==='yesno'){
    return `<div style="font-size:12px">
      <div style="margin-bottom:3px">${t.label}</div>
      <div style="display:flex;gap:6px">
        ${ynBtn(o.id,t.key,'yes','Yes',v==='yes')}
        ${ynBtn(o.id,t.key,'no','No',v==='no')}
      </div>
    </div>`;
  }
  if(t.type==='amount'){
    return `<div style="font-size:12px">
      <div style="margin-bottom:3px">${t.label}</div>
      <input type="number" placeholder="Amount" value="${v||''}" style="width:120px;padding:4px 8px;font-size:12px;border:1px solid var(--border-strong);border-radius:6px" onchange="setObAmount('${o.id}','${t.key}',this.value)">
    </div>`;
  }
  if(t.type==='yesno_reason'){
    const reason=o.tasks[t.key+'Reason']||'';
    return `<div style="font-size:12px">
      <div style="margin-bottom:3px">${t.label}</div>
      <div style="display:flex;gap:6px">
        ${ynBtn(o.id,t.key,'yes','Yes',v==='yes')}
        ${ynBtn(o.id,t.key,'no','No',v==='no')}
      </div>
      ${v==='no'?`<input type="text" placeholder="Reason" value="${reason.replace(/"/g,'&quot;')}" style="margin-top:5px;width:180px;padding:4px 8px;font-size:12px;border:1px solid var(--border-strong);border-radius:6px" onchange="setObReason('${o.id}','${t.key}',this.value)">`:''}
    </div>`;
  }
  if(t.type==='clearance'){
    const status=o.tasks.clearanceStatus;
    let html=`<div style="font-size:12px">
      <div style="margin-bottom:3px">${t.label}</div>
      <div style="display:flex;gap:6px">
        ${optBtn(o.id,'clearanceStatus','clear','Clear',status==='clear','btn-green')}
        ${optBtn(o.id,'clearanceStatus','not_cleared','Unclear',status==='not_cleared','btn-danger')}
      </div>`;
    if(status==='clear'){
      const amt=o.tasks.clearAmount||'';
      html+=`<div style="margin-top:8px;display:flex;flex-direction:column;gap:8px;padding-left:2px">
        <div>
          <div style="margin-bottom:3px">Amount to be Transferred</div>
          <input type="number" placeholder="Amount" value="${amt}" style="width:140px;padding:4px 8px;font-size:12px;border:1px solid var(--border-strong);border-radius:6px" onchange="setObAmount('${o.id}','clearAmount',this.value)">
        </div>
        <div>
          <div style="margin-bottom:3px">Documents Received</div>
          <div style="display:flex;gap:6px">
            ${ynBtn(o.id,'clearDocsReceived','yes','Yes',o.tasks.clearDocsReceived==='yes')}
            ${ynBtn(o.id,'clearDocsReceived','no','No',o.tasks.clearDocsReceived==='no')}
          </div>
        </div>
        <div>
          <div style="margin-bottom:3px">Experience Letter Received</div>
          <div style="display:flex;gap:6px">
            ${ynBtn(o.id,'clearExpLetter','yes','Yes',o.tasks.clearExpLetter==='yes')}
            ${ynBtn(o.id,'clearExpLetter','no','No',o.tasks.clearExpLetter==='no')}
          </div>
        </div>
        <div>
          <div style="margin-bottom:3px">Transferred Through Bank</div>
          <div style="display:flex;gap:6px">
            ${ynBtn(o.id,'clearBankTransfer','yes','Yes',o.tasks.clearBankTransfer==='yes')}
            ${ynBtn(o.id,'clearBankTransfer','no','No',o.tasks.clearBankTransfer==='no')}
          </div>
        </div>
      </div>`;
    } else if(status==='not_cleared'){
      const reason=o.tasks.clearanceReason||'';
      html+=`<div style="margin-top:8px;display:flex;flex-direction:column;gap:8px;padding-left:2px">
        <div>
          <div style="margin-bottom:3px">Due Amount Status</div>
          <div style="display:flex;gap:6px">
            ${optBtn(o.id,'dueAmountStatus','paid','Paid',o.tasks.dueAmountStatus==='paid','btn-green')}
            ${optBtn(o.id,'dueAmountStatus','not_yet','Not Yet',o.tasks.dueAmountStatus==='not_yet','btn-pending-active')}
          </div>
        </div>
        <div>
          <div style="margin-bottom:3px">Reason</div>
          <textarea rows="2" placeholder="Explain why clearance is not completed" style="width:100%;max-width:260px;padding:6px 8px;font-size:12px;border:1px solid var(--border-strong);border-radius:6px;font-family:'Inter',sans-serif;resize:vertical" onchange="setObText('${o.id}','clearanceReason',this.value)">${reason}</textarea>
        </div>
      </div>`;
    }
    html+='</div>';
    return html;
  }
  return '';
}
function obMonthKey(dateStr){
  const d=new Date(dateStr+'T00:00:00');
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}
function obMonthLabel(key){
  const [y,m]=key.split('-');
  return new Date(Number(y),Number(m)-1,1).toLocaleDateString('en-GB',{month:'short',year:'numeric'});
}
function obChecklistPerformance(){
  let done=0,total=0;
  offboards.forEach(o=>{
    OB_TASKS.forEach(t=>{ total++; if(obTaskDone(o,t)) done++; });
  });
  return total?Math.round(done/total*100):0;
}
function renderOffboarding(){
  const wrap=document.getElementById('ob-list-wrap');
  // Period filter (last working day). `list` drives the metrics, chart, department breakdown and table.
  const list=offboards.filter(o=>pfMatch('ob',o.lastDay));
  pfRenderBar('ob','ob-period-filter',offboards.map(o=>o.lastDay),list.length,offboards.length);

  // Bulk row-select toolbar
  const obToolbar=document.getElementById('ob-toolbar');
  const obSelBtnLbl=document.getElementById('ob-select-btn-lbl');
  const obSelBtn=document.getElementById('ob-select-btn');
  if(obSelBtnLbl) obSelBtnLbl.textContent = obSelectMode ? 'Cancel selecting' : 'Select rows';
  if(obSelBtn){ obSelBtn.classList.toggle('btn-primary', obSelectMode); obSelBtn.classList.toggle('btn-outline', !obSelectMode); }
  if(obToolbar){
    obToolbar.innerHTML = obSelectMode ? `<div class="ab ab-purple" style="margin-bottom:14px;flex-direction:row;align-items:center;flex-wrap:wrap;gap:10px">
      <span style="font-weight:600;font-size:13px">${obSelected.size} row${obSelected.size!==1?'s':''} selected</span>
      <button class="btn btn-danger btn-sm" ${obSelected.size?'':'disabled'} onclick="obRemoveSelected()">Remove Selected</button>
      <button class="btn btn-outline btn-sm" onclick="obToggleSelectMode()">Done</button>
      <span style="font-size:11.5px;color:var(--text3);margin-left:auto">Tick rows in the table below, then remove — remaining data stays untouched.</span>
    </div>` : '';
  }

  // Bulk column-select toolbar (only meaningful for extra/uploaded columns)
  const obColToolbar=document.getElementById('ob-col-toolbar');
  const obColBtnLbl=document.getElementById('ob-colselect-btn-lbl');
  const obColBtn=document.getElementById('ob-colselect-btn');
  const activeExtraColsForToolbar=obExtraCols.filter(c=>offboards.some(o=>o.extra&&o.extra[c.key]));
  if(obColBtnLbl) obColBtnLbl.textContent = obColSelectMode ? 'Cancel selecting' : 'Select columns';
  if(obColBtn){ obColBtn.classList.toggle('btn-primary', obColSelectMode); obColBtn.classList.toggle('btn-outline', !obColSelectMode); }
  if(obColToolbar){
    if(obColSelectMode && !activeExtraColsForToolbar.length){
      obColToolbar.innerHTML=`<div class="ab ab-purple" style="margin-bottom:14px"><span style="font-size:12.5px">No extra uploaded columns to remove — the standard fields (Name, Dept, Last day, etc.) can't be removed as columns.</span></div>`;
    } else if(obColSelectMode){
      obColToolbar.innerHTML=`<div class="ab ab-purple" style="margin-bottom:14px;flex-direction:column;gap:8px">
        <div style="font-weight:600;font-size:13px">Choose extra columns to remove from every record</div>
        <div style="display:flex;flex-wrap:wrap;gap:8px">
          ${activeExtraColsForToolbar.map(c=>`<label style="display:flex;align-items:center;gap:5px;font-size:12.5px;background:var(--surface);border:1px solid var(--border-strong);border-radius:6px;padding:4px 9px;cursor:pointer">
            <input type="checkbox" ${obColSelected.has(c.key)?'checked':''} onchange="obToggleColSelect('${c.key}')">${c.label}
          </label>`).join('')}
        </div>
        <div style="display:flex;gap:10px;align-items:center">
          <button class="btn btn-danger btn-sm" ${obColSelected.size?'':'disabled'} onclick="obRemoveSelectedCols()">Remove Selected Column${obColSelected.size!==1?'s':''}</button>
          <button class="btn btn-outline btn-sm" onclick="obToggleColSelectMode()">Done</button>
        </div>
      </div>`;
    } else {
      obColToolbar.innerHTML='';
    }
  }

  /* ── Metrics (migrated from Resignation Analysis) ── */
  const total=list.length;
  const now=new Date();
  const thisMonthCount=list.filter(o=>{
    if(!o.lastDay) return false;
    const d=new Date(o.lastDay+'T00:00:00');
    return d.getFullYear()===now.getFullYear() && d.getMonth()===now.getMonth();
  }).length;
  const thisQuarterCount=list.filter(o=>{
    if(!o.lastDay) return false;
    const d=new Date(o.lastDay+'T00:00:00');
    return d.getFullYear()===now.getFullYear() && Math.ceil((d.getMonth()+1)/3)===Math.ceil((now.getMonth()+1)/3);
  }).length;
  const deptMap={};
  list.forEach(o=>{ const d=o.dept||'Unspecified'; deptMap[d]=(deptMap[d]||0)+1; });
  const deptEntries=Object.entries(deptMap).sort((a,b)=>b[1]-a[1]);
  const topDept=deptEntries.length?deptEntries[0][0]:'—';
  setVal('ob-total', total);
  setVal('ob-month', thisMonthCount);
  setVal('ob-quarter', thisQuarterCount);
  setVal('ob-topdept', topDept);

  /* ── Monthly bar chart ── */
  const map={};
  list.forEach(o=>{
    if(!o.lastDay) return;
    const d=new Date(o.lastDay+'T00:00:00');
    const key=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const lbl=d.toLocaleDateString('en-GB',{month:'short',year:'numeric'});
    if(!map[key]) map[key]={key,lbl,count:0};
    map[key].count++;
  });
  const stats=Object.values(map).sort((a,b)=>a.key.localeCompare(b.key));
  const maxCount=stats.length?Math.max(...stats.map(s=>s.count),1):1;
  const bars=document.getElementById('ob-bars');
  if(bars){
    bars.innerHTML=stats.length?stats.map(s=>{
      const h=Math.round((s.count/maxCount)*80);
      return `<div class="bc-col">
        <div class="bc-n">${s.count}</div>
        <div class="bc-b" style="height:${h}px"></div>
        <div class="bc-lbl">${s.lbl}</div>
      </div>`;
    }).join(''):'<div style="font-size:13px;color:var(--text3);padding:20px">No offboarding data yet — upload an Excel/CSV file</div>';
  }

  /* ── Department breakdown ── */
  const deptWrap=document.getElementById('ob-dept-wrap');
  if(deptWrap){
    if(deptEntries.length){
      const maxDept=Math.max(...deptEntries.map(([,n])=>n),1);
      deptWrap.innerHTML=`<div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--rl);padding:15px 17px;box-shadow:var(--shadow-sm)">
        <div class="sec-title">By department</div>
        <div style="display:flex;flex-direction:column;gap:8px">
          ${deptEntries.map(([dept,n])=>`
            <div>
              <div style="display:flex;justify-content:space-between;font-size:12.5px;margin-bottom:3px"><span>${dept}</span><span style="color:var(--text3);font-family:'JetBrains Mono',monospace">${n}</span></div>
              <div style="height:6px;background:var(--surface3);border-radius:99px;overflow:hidden"><div style="height:100%;width:${Math.round(n/maxDept*100)}%;background:linear-gradient(90deg,var(--accent),#7C3AED);border-radius:99px"></div></div>
            </div>`).join('')}
        </div>
      </div>`;
    } else {
      deptWrap.innerHTML='';
    }
  }

  /* ── Full table (every column from uploaded data) ── */
  if(!offboards.length){
    wrap.innerHTML=`<div class="empty"><div class="empty-title">No offboarding records yet</div><div>Click "Upload Excel" to import offboarding records from a spreadsheet</div></div>`;
    renderResignationAnalysis();
    return;
  }
  if(!list.length){
    wrap.innerHTML=`<div class="empty"><div class="empty-title">No offboarding records in this period</div><div>Try a different month, quarter or year — or choose "All".</div></div>`;
    renderResignationAnalysis();
    return;
  }
  const sorted=[...list].sort((a,b)=>{
    const ap=obProgress(a)===100, bp=obProgress(b)===100;
    if(ap!==bp) return ap?1:-1;
    return daysFrom(a.lastDay)-daysFrom(b.lastDay);
  });
  // Only show extra columns that at least one record actually has a value for
  const activeExtraCols=obExtraCols.filter(c=>offboards.some(o=>o.extra&&o.extra[c.key]));
  wrap.innerHTML=`<div class="sec-title" style="justify-content:space-between">
    <span>${list.length===offboards.length?'All offboarding records':'Offboarding records'} (${list.length}${list.length===offboards.length?'':' of '+offboards.length})</span>
    ${activeExtraCols.length?'<span style="font-weight:400;font-size:11px;color:var(--text3)">↔ scroll table to see all columns</span>':''}
  </div>
  <div class="ct-table"><table><thead><tr>${obSelectMode?`<th style="width:30px"><input type="checkbox" onchange="obSelectAllVisible(this.checked)" ${sorted.length&&sorted.every(o=>obSelected.has(o.id))?'checked':''}></th>`:''}<th>Name</th><th>Code</th><th>Dept</th><th>Job Title</th><th>Last day</th><th>Deadline</th><th>Reason</th><th>Gender</th><th>Grade</th><th>Rating</th><th>Employment Type</th><th>Nationality</th><th>Joined</th><th>Resigned</th><th>Tenure</th>${activeExtraCols.map(c=>`<th>${c.label}</th>`).join('')}<th>Checklist</th><th>Status</th><th>Actions</th></tr></thead>
  <tbody>${sorted.map(o=>{
    const d=daysFrom(o.lastDay);
    const dl=d<0?`${Math.abs(d)}d ago`:d===0?'Today':`in ${d}d`;
    return `<tr id="obr-${o.id}"${obProgress(o)===100?' class="dim"':''}>
      ${obSelectMode?`<td><input type="checkbox" ${obSelected.has(o.id)?'checked':''} onchange="obToggleRow('${o.id}')"></td>`:''}
      <td><strong>${o.name}</strong></td><td>${o.code||'—'}</td><td>${o.dept||'—'}</td><td>${o.jobTitle||'—'}</td>
      <td>${fmt(o.lastDay)}<div style="font-size:11px;color:var(--text3)">${dl}</div></td>
      <td${obIsOverdue(o)?' style="color:var(--red-text);font-weight:600"':''}>${obDeadlineLabel(o)}</td>
      <td>${o.reason||'—'}</td>
      <td>${o.gender||'—'}</td><td>${o.grade||'—'}</td><td>${o.rating||'—'}</td>
      <td>${o.employmentType||'—'}</td><td>${o.nationality||'—'}</td>
      <td>${o.joinDate?fmt(o.joinDate):'—'}</td><td>${o.resignDate?fmt(o.resignDate):'—'}</td><td>${obTenureLabel(o)}</td>
      ${activeExtraCols.map(c=>`<td>${(o.extra&&o.extra[c.key])||'—'}</td>`).join('')}
      <td style="min-width:280px">
        <div style="display:flex;flex-direction:column;gap:7px">
          ${OB_TASKS.map(t=>obTaskRow(o,t)).join('')}
        </div>
      </td>
      <td>${obStatusPill(o)}</td>
      <td style="white-space:nowrap"><button class="btn btn-danger btn-sm" onclick="removeOb('${o.id}')">Remove</button></td>
    </tr>`;
  }).join('')}</tbody></table></div>`;
  renderResignationAnalysis();
}
/* ─── RESIGNATION ANALYSIS DASHBOARD (100% derived from Offboarding data) ─── */
let raPeriod='monthly';       // 'monthly' | 'quarterly' | 'yearly'
let raMonthSel=null;          // 'YYYY-MM'
let raQuarterSel=null;        // 'YYYY-Q#'
let raCharts={};              // canvasId -> Chart.js instance (destroyed/recreated on each render)
const RA_COLORS=['#2563EB','#7C3AED','#D97706','#059669','#DC2626','#0EA5E9','#DB2777','#65A30D','#7C2D12','#4338CA'];
function raColor(i){ return RA_COLORS[i%RA_COLORS.length]; }

function raRenderChart(canvasId, config){
  const el=document.getElementById(canvasId);
  if(!el||typeof Chart==='undefined') return;
  if(raCharts[canvasId]){ raCharts[canvasId].destroy(); delete raCharts[canvasId]; }
  raCharts[canvasId]=new Chart(el.getContext('2d'), config);
}
function raCardShell(canvasId, title, height){
  return `<div class="ra-chart-card">
    <div class="sec-title">${title}</div>
    <div style="position:relative;height:${height||230}px"><canvas id="${canvasId}"></canvas></div>
  </div>`;
}
function setRaPeriod(period){
  raPeriod=period;
  document.querySelectorAll('#ra-period-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.p===period));
  document.getElementById('ra-month-select').style.display=period==='monthly'?'':'none';
  document.getElementById('ra-quarter-select').style.display=period==='quarterly'?'':'none';
  document.getElementById('ra-yearly-note').style.display=period==='yearly'?'':'none';
  renderResignationAnalysis();
}
function setRaSelection(val){
  if(raPeriod==='monthly') raMonthSel=val;
  else if(raPeriod==='quarterly') raQuarterSel=val;
  renderResignationAnalysis();
}
function raMonthLabel(key){ const [y,m]=key.split('-'); return new Date(Number(y),Number(m)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'}); }
function raQuarterLabel(key){ const [y,q]=key.split('-Q'); return `Q${q} ${y}`; }
function getRaPeriodKeys(){
  const months=new Set(), quarters=new Set();
  offboards.forEach(o=>{
    if(!o.lastDay) return;
    const [y,m]=o.lastDay.split('-');
    months.add(`${y}-${m}`);
    quarters.add(`${y}-Q${Math.ceil(Number(m)/3)}`);
  });
  return{ months:[...months].sort().reverse(), quarters:[...quarters].sort().reverse() };
}
function raPopulateSelects(){
  const {months,quarters}=getRaPeriodKeys();
  if(!months.includes(raMonthSel)) raMonthSel=months[0]||null;
  if(!quarters.includes(raQuarterSel)) raQuarterSel=quarters[0]||null;
  const monthSel=document.getElementById('ra-month-select');
  const quarterSel=document.getElementById('ra-quarter-select');
  monthSel.innerHTML=months.length?months.map(k=>`<option value="${k}" ${k===raMonthSel?'selected':''}>${raMonthLabel(k)}</option>`).join(''):'<option value="">No data yet</option>';
  quarterSel.innerHTML=quarters.length?quarters.map(k=>`<option value="${k}" ${k===raQuarterSel?'selected':''}>${raQuarterLabel(k)}</option>`).join(''):'<option value="">No data yet</option>';
}
function raFilteredOb(){
  if(raPeriod==='monthly'){
    if(!raMonthSel) return [];
    return offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,7)===raMonthSel);
  }
  if(raPeriod==='quarterly'){
    if(!raQuarterSel) return [];
    const [y,q]=raQuarterSel.split('-Q');
    return offboards.filter(o=>{
      if(!o.lastDay) return false;
      const [oy,om]=o.lastDay.split('-');
      return oy===y && Math.ceil(Number(om)/3)===Number(q);
    });
  }
  return offboards; // yearly: full dataset, grouped by year internally
}
function raYears(){ return [...new Set(offboards.map(o=>o.lastDay?o.lastDay.slice(0,4):null).filter(Boolean))].sort(); }
function raCurrentPeriodLabel(){
  if(raPeriod==='monthly') return raMonthSel?raMonthLabel(raMonthSel):'selected month';
  if(raPeriod==='quarterly') return raQuarterSel?raQuarterLabel(raQuarterSel):'selected quarter';
  return 'all years';
}

/* ── KPI drill-down handlers (called from onclick on the stat cards) ── */
function raKPIClickTotal(){
  const recs=raFilteredOb();
  openDrillModal('All resignations', recs, `${recs.length} employee${recs.length!==1?'s':''} · ${raCurrentPeriodLabel()}`);
}
function raKPIClickTenure(){
  const recs=raFilteredOb().filter(o=>obTenureYears(o)!==null);
  openDrillModal('Resignations with known tenure', recs, `${recs.length} employee${recs.length!==1?'s':''} · ${raCurrentPeriodLabel()}`);
}
function raKPIClickTopDept(){
  const recs=raFilteredOb();
  const counts={}; recs.forEach(o=>{ const d=o.dept||'Not specified'; counts[d]=(counts[d]||0)+1; });
  const top=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0];
  if(!top) return;
  const filtered=recs.filter(o=>(o.dept||'Not specified')===top[0]);
  openDrillModal(`Department: ${top[0]}`, filtered, `${filtered.length} employee${filtered.length!==1?'s':''} · ${raCurrentPeriodLabel()}`);
}
function raKPIClickTopReason(){
  const recs=raFilteredOb();
  const counts={}; recs.forEach(o=>{ const r=o.reason||'Not specified'; counts[r]=(counts[r]||0)+1; });
  const top=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0];
  if(!top) return;
  const filtered=recs.filter(o=>(o.reason||'Not specified')===top[0]);
  openDrillModal(`Reason: ${top[0]}`, filtered, `${filtered.length} employee${filtered.length!==1?'s':''} · ${raCurrentPeriodLabel()}`);
}
function raKPIClickYear(year){
  const recs=offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===year);
  openDrillModal(`All resignations — ${year}`, recs, `${recs.length} employee${recs.length!==1?'s':''}`);
}

/* ── KPI area: 4 stat cards for monthly/quarterly, 2 comparison charts for yearly ── */
function raRenderKPIs(){
  const el=document.getElementById('rz-metrics');
  if(raPeriod!=='yearly'){
    const recs=raFilteredOb();
    const tenureVals=recs.map(obTenureYears).filter(y=>y!==null);
    const avgTenure=tenureVals.length?(tenureVals.reduce((s,y)=>s+y,0)/tenureVals.length):null;
    const deptCounts={}; recs.forEach(o=>{ const d=o.dept||'Not specified'; deptCounts[d]=(deptCounts[d]||0)+1; });
    const topDept=Object.entries(deptCounts).sort((a,b)=>b[1]-a[1])[0];
    const reasonCounts={}; recs.forEach(o=>{ const r=o.reason||'Not specified'; reasonCounts[r]=(reasonCounts[r]||0)+1; });
    const topReason=Object.entries(reasonCounts).sort((a,b)=>b[1]-a[1])[0];
    el.innerHTML=`
      <div class="metric metric-click" onclick="raKPIClickTotal()"><div class="m-lbl">Total resignations</div><div class="m-val">${recs.length}</div></div>
      <div class="metric metric-click" onclick="raKPIClickTenure()"><div class="m-lbl">Avg. tenure</div><div class="m-val c-purple" style="font-size:20px">${avgTenure!==null?avgTenure.toFixed(1)+' yrs':'—'}</div></div>
      <div class="metric metric-click" onclick="raKPIClickTopDept()"><div class="m-lbl">Top department</div><div class="m-val c-blue" style="font-size:16px">${topDept?topDept[0]:'—'}</div></div>
      <div class="metric metric-click" onclick="raKPIClickTopReason()"><div class="m-lbl">Top reason</div><div class="m-val c-amber" style="font-size:16px">${topReason?topReason[0]:'—'}</div></div>`;
    return;
  }
  const years=raYears();
  el.innerHTML=`
    <div class="metric" style="grid-column:span 2;padding:15px 17px">
      <div class="sec-title" style="margin-bottom:8px">Total resignations by year</div>
      <div style="position:relative;height:150px"><canvas id="rz-total-by-year"></canvas></div>
    </div>
    <div class="metric" style="grid-column:span 2;padding:15px 17px">
      <div class="sec-title" style="margin-bottom:8px">Average tenure by year (yrs)</div>
      <div style="position:relative;height:150px"><canvas id="rz-tenure-by-year"></canvas></div>
    </div>`;
  const totals=years.map(y=>offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===y).length);
  raRenderChart('rz-total-by-year',{
    type:'bar',
    data:{labels:years, datasets:[{data:totals, backgroundColor:years.map((_,i)=>raColor(i)), borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:true}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,ticks:{precision:0},grid:{color:chartGridColor()}}},
      onClick:(evt,elements)=>{ if(!elements.length) return; raKPIClickYear(years[elements[0].index]); }}
  });
  const tenureByYear=years.map(y=>{
    const vals=offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===y).map(obTenureYears).filter(v=>v!==null);
    return vals.length?+(vals.reduce((s,v)=>s+v,0)/vals.length).toFixed(2):0;
  });
  raRenderChart('rz-tenure-by-year',{
    type:'bar',
    data:{labels:years, datasets:[{data:tenureByYear, backgroundColor:years.map((_,i)=>raColor(i)), borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:true}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:chartGridColor()}}},
      onClick:(evt,elements)=>{ if(!elements.length) return; raKPIClickYear(years[elements[0].index]); }}
  });
}

/* ── Trend chart: days-in-month (monthly) / months-in-quarter (quarterly) / year-over-year (yearly) ── */
function raBuildTrend(){
  const monthShort=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  if(raPeriod==='yearly'){
    const years=raYears();
    const datasets=years.map((y,i)=>({
      label:y,
      data:monthShort.map((_,mi)=>offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===y&&Number(o.lastDay.slice(5,7))===mi+1).length),
      borderColor:raColor(i), backgroundColor:raColor(i), tension:.35, fill:false, pointRadius:3, pointHoverRadius:5,
    }));
    return{type:'line', labels:monthShort, datasets, title:'Resignations by month — year-over-year comparison'};
  }
  if(raPeriod==='quarterly'){
    if(!raQuarterSel) return{type:'bar', labels:[], datasets:[], title:'Resignations by month'};
    const [y,q]=raQuarterSel.split('-Q');
    const startMonth=(Number(q)-1)*3+1;
    const monthsInQ=[startMonth,startMonth+1,startMonth+2];
    const labels=monthsInQ.map(m=>monthShort[m-1]);
    const data=monthsInQ.map(m=>offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===y&&Number(o.lastDay.slice(5,7))===m).length);
    return{type:'bar', labels, datasets:[{label:`Q${q} ${y}`, data, backgroundColor:raColor(0), borderRadius:4}], title:`Resignations by month — Q${q} ${y}`};
  }
  // monthly: day-by-day within the selected month
  if(!raMonthSel) return{type:'bar', labels:[], datasets:[], title:'Resignations by day'};
  const [y,m]=raMonthSel.split('-');
  const daysInMonth=new Date(Number(y),Number(m),0).getDate();
  const labels=Array.from({length:daysInMonth},(_,i)=>String(i+1));
  const data=labels.map(d=>offboards.filter(o=>o.lastDay===`${y}-${m}-${String(d).padStart(2,'0')}`).length);
  return{type:'bar', labels, datasets:[{label:raMonthLabel(raMonthSel), data, backgroundColor:raColor(0), borderRadius:4}], title:`Resignations by day — ${raMonthLabel(raMonthSel)}`};
}

/* ── Breakdown chart data: single-period snapshot, or grouped-by-year comparison ── */
function raBuildBreakdown(catFn, opts){
  opts=opts||{};
  if(raPeriod==='yearly'){
    const years=raYears();
    let categories=opts.fixedCategories?opts.fixedCategories.slice():null;
    if(!categories){
      const totals={};
      offboards.forEach(o=>{ const c=catFn(o)||'Not specified'; totals[c]=(totals[c]||0)+1; });
      categories=Object.keys(totals).sort((a,b)=>totals[b]-totals[a]);
    } else {
      const hasOther=offboards.some(o=>!categories.includes(catFn(o)||'Not specified'));
      if(hasOther&&!categories.includes('Not specified')) categories.push('Not specified');
    }
    const datasets=years.map((y,i)=>({
      label:y,
      data:categories.map(cat=>offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===y&&(catFn(o)||'Not specified')===cat).length),
      backgroundColor:raColor(i), borderRadius:4,
    }));
    return{type:'bar', labels:categories, datasets};
  }
  const recs=raFilteredOb();
  const counts={};
  recs.forEach(o=>{ const c=catFn(o)||'Not specified'; counts[c]=(counts[c]||0)+1; });
  let categories=opts.fixedCategories?opts.fixedCategories.slice():null;
  if(!categories){
    categories=Object.keys(counts).sort((a,b)=>counts[b]-counts[a]);
  } else {
    const hasOther=Object.keys(counts).some(c=>!categories.includes(c));
    if(hasOther&&!categories.includes('Not specified')) categories.push('Not specified');
  }
  const data=categories.map(c=>counts[c]||0);
  const type=opts.singleType||'bar';
  return{type, labels:categories, datasets:[{data, backgroundColor:categories.map((_,i)=>raColor(i)), borderRadius:type==='bar'?4:0, borderWidth:type==='bar'?0:2, borderColor:chartSurfaceColor()}]};
}
function raRenderBreakdownCard(canvasId, catFn, opts){
  opts=opts||{};
  const cfg=raBuildBreakdown(catFn, opts);
  const isPieLike=cfg.type==='pie'||cfg.type==='doughnut';
  const isHorizontal=!!opts.horizontal&&cfg.type==='bar';
  let scales={};
  if(cfg.type==='bar'){
    scales=isHorizontal?{
      x:{beginAtZero:true, ticks:{precision:0,font:{size:11}}, grid:{color:chartGridColor()}},
      y:{grid:{display:false}, ticks:{font:{size:11}}},
    }:{
      x:{grid:{display:false}, ticks:{font:{size:11}}},
      y:{beginAtZero:true, ticks:{precision:0,font:{size:11}}, grid:{color:chartGridColor()}},
    };
  }
  raRenderChart(canvasId,{
    type:cfg.type,
    data:{labels:cfg.labels, datasets:cfg.datasets},
    options:{
      responsive:true, maintainAspectRatio:false,
      indexAxis:isHorizontal?'y':'x',
      plugins:{
        legend:{display:isPieLike||raPeriod==='yearly', position:isPieLike?'right':'top', labels:{boxWidth:10,font:{size:11}}},
        tooltip:{enabled:true},
      },
      scales,
      onClick:(evt,elements,chart)=>{
        if(!elements.length) return;
        const el=elements[0];
        const label=chart.data.labels[el.index];
        let recs, subtitle;
        if(raPeriod==='yearly'){
          const years=raYears();
          const year=years[el.datasetIndex];
          recs=offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,4)===year&&(catFn(o)||'Not specified')===label);
          subtitle=`${recs.length} employee${recs.length!==1?'s':''} · ${year}`;
        } else {
          recs=raFilteredOb().filter(o=>(catFn(o)||'Not specified')===label);
          subtitle=`${recs.length} employee${recs.length!==1?'s':''} · ${raCurrentPeriodLabel()}`;
        }
        openDrillModal(`${opts.title||''}: ${label}`, recs, subtitle);
      },
    }
  });
}

function renderResignationAnalysis(){
  const trendCanvas=document.getElementById('rz-trend-canvas');
  const cardsEl=document.getElementById('rz-cards');
  if(!trendCanvas||!cardsEl||typeof Chart==='undefined') return; // page not mounted / Chart.js not loaded yet

  raPopulateSelects();
  raRenderKPIs();

  const trend=raBuildTrend();
  const titleEl=document.getElementById('rz-trend-title');
  if(titleEl) titleEl.textContent=trend.title;
  const monthShort=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  raRenderChart('rz-trend-canvas',{
    type:trend.type,
    data:{labels:trend.labels, datasets:trend.datasets},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:trend.datasets.length>1, position:'top', labels:{boxWidth:10,font:{size:11}}}, tooltip:{enabled:true} },
      scales:{ x:{grid:{display:false}, ticks:{font:{size:11}}}, y:{beginAtZero:true, ticks:{precision:0,font:{size:11}}, grid:{color:chartGridColor()}} },
      onClick:(evt,elements,chart)=>{
        if(!elements.length) return;
        const el=elements[0];
        if(raPeriod==='yearly'){
          const years=raYears();
          const year=years[el.datasetIndex];
          const mm=String(el.index+1).padStart(2,'0');
          const recs=offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,7)===`${year}-${mm}`);
          openDrillModal(`Resignations — ${monthShort[el.index]} ${year}`, recs, `${recs.length} employee${recs.length!==1?'s':''}`);
        } else if(raPeriod==='quarterly'){
          const [y,q]=raQuarterSel.split('-Q');
          const startMonth=(Number(q)-1)*3+1;
          const monthNum=startMonth+el.index;
          const mm=String(monthNum).padStart(2,'0');
          const recs=offboards.filter(o=>o.lastDay&&o.lastDay.slice(0,7)===`${y}-${mm}`);
          openDrillModal(`Resignations — ${monthShort[monthNum-1]} ${y}`, recs, `${recs.length} employee${recs.length!==1?'s':''}`);
        } else {
          const [y,m]=raMonthSel.split('-');
          const day=String(el.index+1).padStart(2,'0');
          const recs=offboards.filter(o=>o.lastDay===`${y}-${m}-${day}`);
          openDrillModal(`Resignations — ${day} ${raMonthLabel(raMonthSel)}`, recs, `${recs.length} employee${recs.length!==1?'s':''}`);
        }
      },
    }
  });

  cardsEl.innerHTML=[
    raCardShell('rz-gender-canvas','By gender'),
    raCardShell('rz-reason-canvas','By reason'),
    raCardShell('rz-dept-canvas','By department / division'),
    raCardShell('rz-grade-canvas','By employee grade'),
    raCardShell('rz-rating-canvas','By performance rating'),
    raCardShell('rz-tenure-canvas','By tenure (joining → last working day)'),
    ...(offboards.some(o=>o.division)?[raCardShell('rz-division-canvas','By division')]:[]),
    ...(offboards.some(o=>o.cluster)?[raCardShell('rz-cluster-canvas','By cluster / supportive function')]:[]),
    ...(offboards.some(o=>o.company)?[raCardShell('rz-company-canvas','By company')]:[]),
  ].join('');

  raRenderBreakdownCard('rz-gender-canvas', o=>o.gender==='Male'?'Male':o.gender==='Female'?'Female':'Not specified', {fixedCategories:['Male','Female'], singleType:'doughnut', title:'Gender'});
  raRenderBreakdownCard('rz-reason-canvas', o=>o.reason?o.reason.trim():'Not specified', {singleType:'pie', title:'Reason'});
  raRenderBreakdownCard('rz-dept-canvas', o=>o.dept?o.dept.trim():'Not specified', {singleType:'bar', horizontal:true, title:'Department'});
  raRenderBreakdownCard('rz-grade-canvas', o=>o.grade?o.grade.trim():'Not specified', {singleType:'bar', title:'Grade'});
  raRenderBreakdownCard('rz-rating-canvas', o=>o.rating?o.rating.trim():'Not specified', {singleType:'doughnut', title:'Rating'});
  if(offboards.some(o=>o.division)) raRenderBreakdownCard('rz-division-canvas', o=>o.division?o.division.trim():'Not specified', {singleType:'bar', horizontal:true, title:'Division'});
  if(offboards.some(o=>o.cluster)) raRenderBreakdownCard('rz-cluster-canvas', o=>o.cluster?o.cluster.trim():'Not specified', {singleType:'bar', horizontal:true, title:'Cluster'});
  if(offboards.some(o=>o.company)) raRenderBreakdownCard('rz-company-canvas', o=>o.company?o.company.trim():'Not specified', {singleType:'doughnut', title:'Company'});
  raRenderBreakdownCard('rz-tenure-canvas', obTenureBucket, {fixedCategories:['<1 year','1–2 years','2–5 years','5–10 years','10+ years','Not specified'], singleType:'bar', title:'Tenure'});

  // Reflect active segmented-control state (in case this render was triggered by data changing elsewhere, not a click)
  document.querySelectorAll('#ra-period-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.p===raPeriod));
}
/* ─── OVERTIME ANALYSIS DASHBOARD (100% derived from uploaded Overtime data) ─── */
let otPeriod='monthly';       // 'monthly' | 'quarterly' | 'yearly'
let otMonthSel=null;          // 'YYYY-MM'
let otQuarterSel=null;        // 'YYYY-Q#'
let otYearSel=null;           // 'YYYY' — selected year for the Yearly KPI period
let otMetric='hours';         // 'hours' | 'cost' — drives the trend & breakdown charts
let otFilterDept='';
let otFilterEmp='';
let otFilterTitle='';
let otFilterDivision='';
let otFilterYear='';
let otDivMonthMetric='hours';

/* ── Standard monthly working hours: used to derive an hourly OT rate from monthly salary
   whenever a record doesn't already carry its own explicit rate or amount. Persisted so it
   stays consistent across sessions; recalculating with a new value updates every OT amount
   and OT % live, since nothing is stored pre-computed. ── */
function otStdHours(){ const v=Number(localStorage.getItem('hr3_ot_std_hours')); return v>0?v:240; }
function setOtStdHours(val){
  const v=Number(val);
  if(!v||v<=0){ toast('Enter a positive number of hours','danger'); document.getElementById('ot-std-hours').value=otStdHours(); return; }
  localStorage.setItem('hr3_ot_std_hours', v); saveSettings();
  renderOvertimeAnalysis();
  toast(`✓ Standard monthly hours set to ${v} — all OT rates and amounts recalculated`);
}

/* ── Optional per-category OT rate multipliers (e.g. 1.35× for Normal, 2.0× for Saturday/Official
   under Egyptian labour law). Default to 1 so costs match the plain hourly rate unless management
   explicitly configures a multiplier — this keeps rule 10's "existing calculation logic" intact. ── */
function otMultNormal(){ const v=Number(localStorage.getItem('hr3_ot_mult_normal')); return v>0?v:1; }
function otMultSaturday(){ const v=Number(localStorage.getItem('hr3_ot_mult_sat')); return v>0?v:1; }
function otMultOfficial(){ const v=Number(localStorage.getItem('hr3_ot_mult_official')); return v>0?v:1; }
function setOtMult(kind,val){
  const v=Number(val);
  if(!v||v<=0){ toast('Enter a positive multiplier','danger'); renderOvertimeAnalysis(); return; }
  const key=kind==='normal'?'hr3_ot_mult_normal':kind==='sat'?'hr3_ot_mult_sat':'hr3_ot_mult_official';
  localStorage.setItem(key, v); saveSettings();
  renderOvertimeAnalysis();
  toast('✓ Overtime rate multiplier updated — all costs recalculated');
}

/* ── Core automatic calculations (rule 10): never stored pre-computed, always derived live ── */
function otHourlyRate(o){
  if(o.otRate>0) return o.otRate;
  if(o.monthlySalary>0) return o.monthlySalary/otStdHours();
  return null;
}
function otDivisionOf(o){ return (o.division||'').trim() || 'Not specified'; }
/* Department = Organization, exclusively. Every Department-labeled field, filter, KPI, table, chart,
   and ranking in the Overtime module reads this function — never a separate "Department" column. */
function otOrgOf(o){ return (o.org||'').trim() || 'Not specified'; }
function otCatCost(o,hrs,mult){
  hrs=Number(hrs)||0;
  if(hrs<=0) return 0;
  const r=otHourlyRate(o);
  return r ? +(hrs*r*mult).toFixed(2) : 0;
}
function otNormalCost(o){ return otCatCost(o,o.normalHours,otMultNormal()); }
function otSaturdayCost(o){ return otCatCost(o,o.saturdayHours,otMultSaturday()); }
function otOfficialCost(o){ return otCatCost(o,o.officialHours,otMultOfficial()); }
function otAmountOf(o){
  if(o.salary>0) return o.salary; // explicit OT amount from the sheet overrides the formula
  return +(otNormalCost(o)+otSaturdayCost(o)+otOfficialCost(o)).toFixed(2);
}
function otPercentOf(o){
  if(!(o.monthlySalary>0)) return null;
  return otAmountOf(o)/o.monthlySalary*100;
}

function setOtPeriod(period){
  otPeriod=period;
  document.querySelectorAll('#ot-period-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.p===period));
  document.getElementById('ot-month-select').style.display=period==='monthly'?'':'none';
  document.getElementById('ot-quarter-select').style.display=period==='quarterly'?'':'none';
  document.getElementById('ot-year-select').style.display=period==='yearly'?'':'none';
  document.getElementById('ot-yearly-note').style.display=period==='yearly'?'':'none';
  renderOvertimeAnalysis();
}
function setOtSelection(val){
  if(otPeriod==='monthly') otMonthSel=val;
  else if(otPeriod==='quarterly') otQuarterSel=val;
  else if(otPeriod==='yearly') otYearSel=val;
  renderOvertimeAnalysis();
}
function setOtMetric(m){
  otMetric=m;
  document.querySelectorAll('#ot-metric-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.m===m));
  renderOvertimeAnalysis();
}
function setOtFilter(kind,val){
  if(kind==='dept') otFilterDept=val;
  else if(kind==='emp') otFilterEmp=val;
  else if(kind==='title') otFilterTitle=val;
  else if(kind==='division') otFilterDivision=val;
  else if(kind==='year') otFilterYear=val;
  renderOvertimeAnalysis();
}
function resetOtFilters(){
  otFilterDept=''; otFilterEmp=''; otFilterTitle=''; otFilterDivision=''; otFilterYear='';
  renderOvertimeAnalysis();
}
function setOtDivMonthMetric(m){
  otDivMonthMetric=m;
  document.querySelectorAll('#ot-divmonth-metric-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.dm===m));
  renderOvertimeAnalysis();
}
function otMonthLabel(key){ const [y,m]=key.split('-'); return new Date(Number(y),Number(m)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'}); }
function otQuarterLabel(key){ const [y,q]=key.split('-Q'); return `Q${q} ${y}`; }
/* ── Custom quarters: Q1 = Dec+Jan+Feb, Q2 = Mar+Apr+May, Q3 = Jun+Jul+Aug, Q4 = Sep+Oct+Nov.
   December belongs to Q1 of the FOLLOWING year (Dec 2025 + Jan 2026 + Feb 2026 = Q1 2026). ── */
function otQuarterOfMonth(y,m){ y=Number(y); m=Number(m); return m===12 ? {year:y+1,q:1} : {year:y,q:Math.floor(m/3)+1}; }
function otQuarterKeyOf(dateStr){ const [y,m]=dateStr.split('-'); const r=otQuarterOfMonth(y,m); return `${r.year}-Q${r.q}`; }
function otQuarterMonths(y,q){ y=Number(y); q=Number(q); return q===1 ? [{y:y-1,m:12},{y:y,m:1},{y:y,m:2}] : [0,1,2].map(i=>({y:y,m:(q-1)*3+i})); }
function otQuarterRangeLabel(y,q){
  const ms=otQuarterMonths(y,q), n=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${n[ms[0].m-1]} ${ms[0].y} – ${n[ms[2].m-1]} ${ms[2].y}`;
}
function getOtPeriodKeys(){
  const months=new Set(), quarters=new Set();
  overtimeRecords.forEach(o=>{
    if(!o.date) return;
    const [y,m]=o.date.split('-');
    months.add(`${y}-${m}`);
    quarters.add(otQuarterKeyOf(o.date));
  });
  return{ months:[...months].sort().reverse(), quarters:[...quarters].sort().reverse() };
}
function otPopulateSelects(){
  const {months,quarters}=getOtPeriodKeys();
  const years=[...otYears()].reverse(); // most recent first
  if(!months.includes(otMonthSel)) otMonthSel=months[0]||null;
  if(!quarters.includes(otQuarterSel)) otQuarterSel=quarters[0]||null;
  if(!years.includes(otYearSel)) otYearSel=years[0]||null;
  const monthSel=document.getElementById('ot-month-select');
  const quarterSel=document.getElementById('ot-quarter-select');
  const yearSel=document.getElementById('ot-year-select');
  monthSel.innerHTML=months.length?months.map(k=>`<option value="${k}" ${k===otMonthSel?'selected':''}>${otMonthLabel(k)}</option>`).join(''):'<option value="">No data yet</option>';
  quarterSel.innerHTML=quarters.length?quarters.map(k=>`<option value="${k}" ${k===otQuarterSel?'selected':''}>${otQuarterLabel(k)}</option>`).join(''):'<option value="">No data yet</option>';
  if(yearSel) yearSel.innerHTML=years.length?years.map(k=>`<option value="${k}" ${k===otYearSel?'selected':''}>${k}</option>`).join(''):'<option value="">No data yet</option>';
}
function otPopulateFilterSelects(){
  const depts=[...new Set(overtimeRecords.map(o=>otOrgOf(o)).filter(d=>d&&d!=='Not specified'))].sort();
  const titles=[...new Set(overtimeRecords.map(o=>o.jobTitle?o.jobTitle.trim():'').filter(Boolean))].sort();
  const emps=[...new Set(overtimeRecords.map(o=>o.name).filter(Boolean))].sort();
  const divisions=[...new Set(overtimeRecords.map(o=>otDivisionOf(o)))].sort();
  const years=otYears();
  const deptSel=document.getElementById('ot-filter-dept'), titleSel=document.getElementById('ot-filter-title'), empSel=document.getElementById('ot-filter-emp');
  const divSel=document.getElementById('ot-filter-division'), yearSel=document.getElementById('ot-filter-year');
  if(deptSel) deptSel.innerHTML=`<option value="">All departments</option>`+depts.map(d=>`<option value="${d}" ${d===otFilterDept?'selected':''}>${d}</option>`).join('');
  if(titleSel) titleSel.innerHTML=`<option value="">All job titles</option>`+titles.map(t=>`<option value="${t}" ${t===otFilterTitle?'selected':''}>${t}</option>`).join('');
  if(empSel) empSel.innerHTML=`<option value="">All employees</option>`+emps.map(e=>`<option value="${e}" ${e===otFilterEmp?'selected':''}>${e}</option>`).join('');
  if(divSel) divSel.innerHTML=`<option value="">All divisions</option>`+divisions.map(d=>`<option value="${d}" ${d===otFilterDivision?'selected':''}>${d}</option>`).join('');
  if(yearSel) yearSel.innerHTML=`<option value="">All years</option>`+years.map(y=>`<option value="${y}" ${y===otFilterYear?'selected':''}>${y}</option>`).join('');
}
function otApplyExtraFilters(recs){
  return recs.filter(o=>
    (!otFilterDept || otOrgOf(o)===otFilterDept) &&
    (!otFilterTitle || (o.jobTitle||'').trim()===otFilterTitle) &&
    (!otFilterEmp || o.name===otFilterEmp) &&
    (!otFilterDivision || otDivisionOf(o)===otFilterDivision) &&
    (!otFilterYear || (o.date&&o.date.slice(0,4)===otFilterYear))
  );
}
function otFilteredRecords(){
  let recs;
  if(otPeriod==='monthly'){
    recs=otMonthSel?overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===otMonthSel):[];
  } else if(otPeriod==='quarterly'){
    if(!otQuarterSel){ recs=[]; }
    else{
      recs=overtimeRecords.filter(o=>o.date && otQuarterKeyOf(o.date)===otQuarterSel);
    }
  } else {
    recs=otYearSel?overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===otYearSel):[]; // yearly: records for the selected year
  }
  return otApplyExtraFilters(recs);
}
function otYears(){ return [...new Set(overtimeRecords.map(o=>o.date?o.date.slice(0,4):null).filter(Boolean))].sort(); }
function otCurrentPeriodLabel(){
  if(otPeriod==='monthly') return otMonthSel?otMonthLabel(otMonthSel):'selected month';
  if(otPeriod==='quarterly') return otQuarterSel?otQuarterLabel(otQuarterSel):'selected quarter';
  return otYearSel||'selected year';
}
function otSumHours(recs){ return recs.reduce((s,o)=>s+(Number(o.hours)||0),0); }
function otSumAmount(recs){ return recs.reduce((s,o)=>s+otAmountOf(o),0); }
function otSumActual(recs){ return recs.reduce((s,o)=>s+(Number(o.actualHours)||0),0); }
function otSumNormalHrs(recs){ return recs.reduce((s,o)=>s+(Number(o.normalHours)||0),0); }
function otSumSaturdayHrs(recs){ return recs.reduce((s,o)=>s+(Number(o.saturdayHours)||0),0); }
function otSumOfficialHrs(recs){ return recs.reduce((s,o)=>s+(Number(o.officialHours)||0),0); }
function otSumNormalCost(recs){ return recs.reduce((s,o)=>s+otNormalCost(o),0); }
function otSumSaturdayCost(recs){ return recs.reduce((s,o)=>s+otSaturdayCost(o),0); }
function otSumOfficialCost(recs){ return recs.reduce((s,o)=>s+otOfficialCost(o),0); }
/* ── Division-level rollup (rule 10: purely derived, never duplicated/stored). "N/A" (null) is
   returned for percentages whenever Actual Hours is 0/missing for that division & period. ── */
function otBuildDivisionStats(recs){
  const groups={};
  recs.forEach(o=>{ const d=otDivisionOf(o); (groups[d]=groups[d]||[]).push(o); });
  return Object.keys(groups).map(d=>{
    const g=groups[d];
    const actual=otSumActual(g), normal=otSumNormalHrs(g), sat=otSumSaturdayHrs(g), official=otSumOfficialHrs(g);
    const otHrs=otSumHours(g), otCost=otSumAmount(g);
    const normalCost=otSumNormalCost(g), satCost=otSumSaturdayCost(g), officialCost=otSumOfficialCost(g);
    const basis=actual>0?actual:null;
    return {
      division:d, recs:g, empCount:otEmployeeCount(g),
      actual, normal, sat, official, otHrs, otCost, normalCost, satCost, officialCost,
      normalPct: basis?normal/basis*100:null,
      satPct: basis?sat/basis*100:null,
      officialPct: basis?official/basis*100:null,
    };
  });
}
function otAvgPercent(recs){
  const withSalary=recs.filter(o=>o.monthlySalary>0);
  if(!withSalary.length) return null;
  return withSalary.reduce((s,o)=>s+otPercentOf(o),0)/withSalary.length;
}
function otEmployeeCount(recs){ return new Set(recs.map(o=>o.name)).size; }
function otMonthsSorted(){ return [...new Set(overtimeRecords.map(o=>o.date?o.date.slice(0,7):null).filter(Boolean))].sort(); }
function otFmtCurrency(n){ return (Number(n)||0).toLocaleString(undefined,{maximumFractionDigits:0}); }
/* ── Overtime Value KPI: strictly the raw value from the uploaded Overtime column — never the
   hourly-rate-derived estimate used elsewhere in this dashboard, and never rounded/averaged. ── */
function otRawValue(o){ return Number(o.salary)||0; }
function otSumRawValue(recs){ return recs.reduce((s,o)=>s+otRawValue(o),0); }
/* ── Displays a number using the exact number/currency format detected from the uploaded Excel
   file's Overtime column (via SheetJS SSF, the same engine Excel itself uses to render a format
   code). Falls back to a plain thousands-separated number when no format was detected (e.g. a
   .csv upload, which carries no cell formatting). ── */
function otFmtValue(n){
  n=Number(n)||0;
  if(otValueNumFmt && typeof XLSX!=='undefined' && XLSX.SSF && XLSX.SSF.format){
    try{ return XLSX.SSF.format(otValueNumFmt, n); }catch(e){ /* fall through */ }
  }
  return otFmtCurrency(n);
}
function otMetricValue(recs,metric){
  metric=metric||otMetric;
  if(metric==='cost') return otSumAmount(recs);
  return otSumHours(recs);
}
function otMetricLabel(metric){
  metric=metric||otMetric;
  return metric==='cost'?'OT cost':'OT hours';
}
function otMetricDisplay(v,metric){
  metric=metric||otMetric;
  if(metric==='cost') return otFmtCurrency(v);
  return v.toFixed(1)+' hrs';
}

/* ── Month-over-month trend cards: always compares the two most recent months present in the data,
   so the comparison advances automatically as new monthly data is uploaded. ── */
function otMomCard(label, curKey, prevKey, cur, prev, unit){
  // For OT hours/cost, a decrease is the desirable outcome, so color is inverted vs. a
  // typical "up=good" metric: decrease → green, increase → red. Arrow direction (▲/▼)
  // still always reflects the actual direction of change.
  let changeHtml;
  if(prevKey===null){
    changeHtml=`<div class="mom-change flat">No prior month to compare yet</div>`;
  } else if(prev===0){
    changeHtml=cur>0
      ? `<div class="mom-change down">▲ New this month</div>`
      : `<div class="mom-change flat">No change</div>`;
  } else {
    const pct=((cur-prev)/prev)*100;
    const isIncrease=pct>=0;
    const colorClass=isIncrease?'down':'up'; // inverted: increase=red(down), decrease=green(up)
    changeHtml=`<div class="mom-change ${colorClass}">${isIncrease?'▲':'▼'} ${isIncrease?'+':''}${pct.toFixed(1)}%</div>`;
  }
  let curDisplay, prevDisplay;
  if(unit==='value'){ curDisplay=otFmtValue(cur); prevDisplay=prevKey?otFmtValue(prev):'—'; }
  else if(unit==='currency'){ curDisplay=otFmtCurrency(cur); prevDisplay=prevKey?otFmtCurrency(prev):'—'; }
  else if(unit==='percent'){ curDisplay=cur.toFixed(1)+'%'; prevDisplay=prevKey?prev.toFixed(1)+'%':'—'; }
  else if(unit==='count'){ curDisplay=Math.round(cur); prevDisplay=prevKey?Math.round(prev):'—'; }
  else { curDisplay=cur.toFixed(1)+' hrs'; prevDisplay=prevKey?prev.toFixed(1)+' hrs':'—'; }
  return `<div class="mom-card">
    <div class="mom-lbl">${label} — ${otMonthLabel(curKey)}</div>
    <div class="mom-val">${curDisplay}</div>
    ${changeHtml}
    <div class="mom-sub">vs ${prevKey?otMonthLabel(prevKey):'previous month'}: ${prevDisplay}</div>
  </div>`;
}
function otRenderMoM(){
  const el=document.getElementById('ot-mom-grid');
  if(!el) return;
  const months=otMonthsSorted();
  if(!months.length){ el.innerHTML=''; return; }
  const curKey=months[months.length-1];
  const prevKey=months.length>1?months[months.length-2]:null;
  const curRecs=overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===curKey);
  const prevRecs=prevKey?overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===prevKey):[];
  el.innerHTML=[
    otMomCard('Total OT Hours', curKey, prevKey, otSumHours(curRecs), otSumHours(prevRecs), 'hours'),
    otMomCard('Overtime Value', curKey, prevKey, otSumRawValue(curRecs), otSumRawValue(prevRecs), 'value'),
  ].join('');
}

/* ── KPI drill-down handlers ── */
function otKPIClickTotal(){
  const recs=otFilteredRecords();
  openOtDrillModal('All overtime records', recs, `${recs.length} record${recs.length!==1?'s':''} · ${otCurrentPeriodLabel()}`);
}
function otKPIClickTopEmployee(){
  const recs=otFilteredRecords();
  const totals={}; recs.forEach(o=>{ totals[o.name]=(totals[o.name]||0)+otRawValue(o); });
  const top=Object.entries(totals).sort((a,b)=>b[1]-a[1])[0];
  if(!top) return;
  const filtered=recs.filter(o=>o.name===top[0]);
  openOtDrillModal(`Employee: ${top[0]}`, filtered, `${filtered.length} record${filtered.length!==1?'s':''} · ${otFmtValue(top[1])} · ${otCurrentPeriodLabel()}`);
}
function otKPIClickEmployee(el){
  const name=el.dataset.emp, recs=otFilteredRecords().filter(o=>o.name===name);
  const tot=recs.reduce((a,o)=>a+otRawValue(o),0);
  openOtDrillModal(`Employee: ${name}`, recs, `${recs.length} record${recs.length!==1?'s':''} · ${otFmtValue(tot)} · ${otCurrentPeriodLabel()}`);
}
function otKPIClickTopDept(){
  const recs=otFilteredRecords();
  const totals={}; recs.forEach(o=>{ const d=otOrgOf(o); totals[d]=(totals[d]||0)+otRawValue(o); });
  const top=Object.entries(totals).sort((a,b)=>b[1]-a[1])[0];
  if(!top) return;
  const filtered=recs.filter(o=>otOrgOf(o)===top[0]);
  openOtDrillModal(`Organization: ${top[0]}`, filtered, `${filtered.length} record${filtered.length!==1?'s':''} · ${otFmtValue(top[1])} · ${otCurrentPeriodLabel()}`);
}
function otKPIClickYear(year){
  const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===year));
  openOtDrillModal(`All overtime records — ${year}`, recs, `${recs.length} record${recs.length!==1?'s':''}`);
}

/* ── Dashboard summary cards (KPI area). Required KPIs, always computed live from the
   uploaded Overtime / Organization / employee-name columns, and always in sync with
   the Month / Quarter / Year selector:
     1. Overtime Value  = SUM of the raw Overtime column (no estimate, average, or %)
     2. Highest Department = Organization with the highest SUM(Overtime)
     3. Highest Employee   = employee with the highest individual Overtime value
   Overtime Value and Highest Employee are displayed using the exact number/currency format
   detected from the uploaded Excel file's Overtime column.
   For the Yearly period this is scoped to the single selected Year (via otYearSel), not every
   year in the sheet — a year-over-year comparison chart is shown underneath for context. ── */
function otRenderKPIs(){
  const el=document.getElementById('ot-metrics');
  const recs=otFilteredRecords();
  const totalHours=otSumHours(recs);
  const totalValue=otSumRawValue(recs); // Overtime Value = SUM(Overtime) — raw column values only

  const empTotals={}; recs.forEach(o=>{ empTotals[o.name]=(empTotals[o.name]||0)+otRawValue(o); });
  const top10=Object.entries(empTotals).sort((a,b)=>b[1]-a[1]).slice(0,10); // Top 10 Employees by Overtime value

  const orgTotals={}; recs.forEach(o=>{ const d=otOrgOf(o); orgTotals[d]=(orgTotals[d]||0)+otRawValue(o); });
  const top5Orgs=Object.entries(orgTotals).sort((a,b)=>b[1]-a[1]).slice(0,5); // Top 5 Departments (by Organization)

  el.innerHTML=`
    <div class="metric metric-click" onclick="otKPIClickTotal()"><div class="m-lbl">Total OT hours</div><div class="m-val">${totalHours.toFixed(1)}</div><div class="m-sub">${recs.length} record${recs.length!==1?'s':''}</div></div>
    <div class="metric metric-click" onclick="otKPIClickTotal()"><div class="m-lbl">Overtime Value</div><div class="m-val c-red">${otFmtValue(totalValue)}</div><div class="m-sub">${otCurrentPeriodLabel()}</div></div>
    <div class="metric"><div class="m-lbl">Top 5 Departments</div>${top5Orgs.length?`<div class="top10-list">${top5Orgs.map((e,i)=>`<div class="top10-row" data-dept="${otEscAttr(e[0])}" onclick="otDeptTableClick(this)" title="Click to see this department's records"><span class="top10-rk">${i+1}</span><span class="top10-nm">${e[0]}</span><span class="top10-v" style="color:var(--amber)">${otFmtValue(e[1])}</span></div>`).join('')}</div>`:`<div class="m-val c-amber" style="font-size:16px">—</div>`}</div>
    <div class="metric"><div class="m-lbl">Top 10 Employees</div>${top10.length?`<div class="top10-list">${top10.map((e,i)=>`<div class="top10-row" data-emp="${otEscAttr(e[0])}" onclick="otKPIClickEmployee(this)" title="Click to see this employee's records"><span class="top10-rk">${i+1}</span><span class="top10-nm">${e[0]}</span><span class="top10-v">${otFmtValue(e[1])}</span></div>`).join('')}</div>`:`<div class="m-val c-green" style="font-size:16px">—</div>`}</div>`;

  if(otPeriod!=='yearly') return;

  // Yearly extra: year-over-year comparison charts (context only — the KPI cards above already
  // reflect the single selected year).
  const years=otYears();
  if(!years.length) return;
  el.innerHTML+=`
    <div class="metric" style="grid-column:span 2;padding:15px 17px">
      <div class="sec-title" style="margin-bottom:8px">Total OT hours by year</div>
      <div style="position:relative;height:150px"><canvas id="ot-total-by-year"></canvas></div>
    </div>
    <div class="metric" style="grid-column:span 2;padding:15px 17px">
      <div class="sec-title" style="margin-bottom:8px">Overtime Value by year</div>
      <div style="position:relative;height:150px"><canvas id="ot-avg-by-year"></canvas></div>
    </div>`;
  const totals=years.map(y=>otSumHours(otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===y))));
  raRenderChart('ot-total-by-year',{
    type:'bar',
    data:{labels:years, datasets:[{data:totals, backgroundColor:years.map((_,i)=>raColor(i)), borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:true}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:chartGridColor()}}},
      onClick:(evt,elements)=>{ if(!elements.length) return; otKPIClickYear(years[elements[0].index]); }}
  });
  const costByYear=years.map(y=>otSumRawValue(otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===y))));
  raRenderChart('ot-avg-by-year',{
    type:'bar',
    data:{labels:years, datasets:[{data:costByYear, backgroundColor:years.map((_,i)=>raColor(i)), borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:true,callbacks:{label:ctx=>otFmtValue(ctx.parsed.y)}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:chartGridColor()}}},
      onClick:(evt,elements)=>{ if(!elements.length) return; otKPIClickYear(years[elements[0].index]); }}
  });
}

/* ── Trend chart: days-in-month (monthly) / months-in-quarter (quarterly) / year-over-year (yearly) ── */
function otBuildTrend(){
  const monthShort=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const base=recs=>otMetricValue(otApplyExtraFilters(recs));
  const titleWord=otMetric==='cost'?'cost':'hours';
  if(otPeriod==='yearly'){
    const years=otYears();
    const datasets=years.map((y,i)=>({
      label:y,
      data:monthShort.map((_,mi)=>base(overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===y&&Number(o.date.slice(5,7))===mi+1))),
      borderColor:raColor(i), backgroundColor:raColor(i), tension:.35, fill:false, pointRadius:3, pointHoverRadius:5,
    }));
    return{type:'line', labels:monthShort, datasets, title:`Overtime ${titleWord} by month — year-over-year comparison`};
  }
  if(otPeriod==='quarterly'){
    if(!otQuarterSel) return{type:'bar', labels:[], datasets:[], title:`Overtime ${titleWord} by month`};
    const [y,q]=otQuarterSel.split('-Q');
    const monthsInQ=otQuarterMonths(y,q);
    const labels=monthsInQ.map(x=>monthShort[x.m-1]);
    const data=monthsInQ.map(x=>base(overtimeRecords.filter(o=>o.date&&Number(o.date.slice(0,4))===x.y&&Number(o.date.slice(5,7))===x.m)));
    return{type:'bar', labels, datasets:[{label:`Q${q} ${y}`, data, backgroundColor:raColor(0), borderRadius:4}], title:`Overtime ${titleWord} by month — Q${q} ${y}`};
  }
  // monthly: day-by-day within the selected month
  if(!otMonthSel) return{type:'bar', labels:[], datasets:[], title:`Overtime ${titleWord} by day`};
  const [y,m]=otMonthSel.split('-');
  const daysInMonth=new Date(Number(y),Number(m),0).getDate();
  const labels=Array.from({length:daysInMonth},(_,i)=>String(i+1));
  const data=labels.map(d=>base(overtimeRecords.filter(o=>o.date===`${y}-${m}-${String(d).padStart(2,'0')}`)));
  return{type:'bar', labels, datasets:[{label:otMonthLabel(otMonthSel), data, backgroundColor:raColor(0), borderRadius:4}], title:`Overtime ${titleWord} by day — ${otMonthLabel(otMonthSel)}`};
}

/* ── Breakdown: totals grouped by a category function (employee / department), top N shown as horizontal bars.
   Uses the currently selected metric (hours / cost / avg OT %) so it always matches the trend chart above it. ── */
function otRenderBreakdownCard(canvasId, catFn, opts){
  opts=opts||{};
  const recs=otFilteredRecords();
  const groups={};
  recs.forEach(o=>{ const c=catFn(o)||'Not specified'; (groups[c]=groups[c]||[]).push(o); });
  let categories=Object.keys(groups).sort((a,b)=>otMetricValue(groups[b])-otMetricValue(groups[a]));
  if(opts.topN) categories=categories.slice(0,opts.topN);
  const data=categories.map(c=>+otMetricValue(groups[c]).toFixed(otMetric==='hours'?1:0));
  raRenderChart(canvasId,{
    type:'bar',
    data:{labels:categories, datasets:[{data, backgroundColor:categories.map((_,i)=>raColor(i)), borderRadius:4}]},
    options:{
      responsive:true, maintainAspectRatio:false,
      indexAxis:'y',
      plugins:{ legend:{display:false}, tooltip:{enabled:true, callbacks:{label:ctx=>otMetricDisplay(ctx.parsed.x,otMetric)}} },
      scales:{ x:{beginAtZero:true, ticks:{font:{size:11}}, grid:{color:chartGridColor()}}, y:{grid:{display:false}, ticks:{font:{size:11}}} },
      onClick:(evt,elements,chart)=>{
        if(!elements.length) return;
        const el=elements[0];
        const label=chart.data.labels[el.index];
        const filtered=groups[label]||[];
        openOtDrillModal(`${opts.title||''}: ${label}`, filtered, `${filtered.length} record${filtered.length!==1?'s':''} · ${otMetricDisplay(otMetricValue(filtered),otMetric)} · ${otCurrentPeriodLabel()}`);
      },
    }
  });
}

/* ── Monthly overtime summary table (requirement 3): every month found in the data, with
   MoM change and an "employees with OT" count, so months can be compared at a glance. ── */
function otEscAttr(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function otMonthTableClick(mKey){
  const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===mKey));
  openOtDrillModal(`Overtime — ${otMonthLabel(mKey)}`, recs, `${recs.length} record${recs.length!==1?'s':''}`);
}
function otDeptTableClick(el){
  const dept=el.dataset.dept;
  const recs=otFilteredRecords().filter(o=>otOrgOf(o)===dept);
  openOtDrillModal(`Department: ${dept}`, recs, `${recs.length} record${recs.length!==1?'s':''}`);
}
let otSummaryDrill=[];
function otSummaryClick(i){
  const d=otSummaryDrill[i]; if(!d) return;
  openOtDrillModal(d.title, d.recs, `${d.recs.length} record${d.recs.length!==1?'s':''}`);
}
function otChangePill(cost,prevCost){
  if(prevCost>0){
    const chg=((cost-prevCost)/prevCost)*100;
    return `<span class="pill ${chg>=0?'pl-red':'pl-green'}">${chg>=0?'▲':'▼'} ${chg>=0?'+':''}${chg.toFixed(1)}%</span>`;
  }
  if(cost>0) return `<span class="pill pl-red">▲ New</span>`;
  return `<span class="pill pl-gray">No change</span>`;
}
/* Period-aware summary:
   Monthly   → one row per month, MoM % change vs the previous month (unchanged behaviour).
   Quarterly → 4 rows (Q1..Q4) of the selected year using the custom quarters; QoQ % vs previous quarter, Q1 = N/A.
   Yearly    → a single total row for the selected year (sum of all 12 months), no breakdown. */
function otRenderMonthlySummary(){
  const wrap=document.getElementById('ot-monthly-table-wrap');
  const titleEl=document.getElementById('ot-summary-title');
  if(!wrap) return;
  const mono='font-family:\'JetBrains Mono\',monospace';
  const cells=recs=>{
    const hours=otSumHours(recs), cost=otSumAmount(recs);
    const actual=otSumActual(recs), normal=otSumNormalHrs(recs), sat=otSumSaturdayHrs(recs), official=otSumOfficialHrs(recs);
    const empCount=otEmployeeCount(recs);
    return `<td style="${mono}">${actual>0?actual.toFixed(1):'N/A'}</td>
      <td style="${mono}">${normal.toFixed(1)}</td>
      <td style="${mono}">${sat.toFixed(1)}</td>
      <td style="${mono}">${official.toFixed(1)}</td>
      <td style="${mono}">${hours.toFixed(1)}</td>
      <td style="${mono}">${otFmtCurrency(cost)}</td>
      <td style="${mono}">${otFmtCurrency(empCount?cost/empCount:0)}</td>`;
  };
  const head=(first,last)=>`<thead><tr><th>${first}</th><th>Actual Hrs</th><th>Normal Hrs</th><th>Saturday Hrs</th><th>Official Days Hrs</th><th>Total OT Hrs</th><th>Total OT Cost</th><th>Avg OT / Employee</th>${last?`<th>${last}</th>`:''}</tr></thead>`;
  otSummaryDrill=[];

  if(otPeriod==='yearly'){
    if(titleEl) titleEl.textContent='Yearly overtime summary';
    if(!otYearSel){ wrap.innerHTML=`<div class="empty"><div class="empty-title">No yearly data yet</div></div>`; return; }
    const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,4)===otYearSel)); // sum of all 12 months
    otSummaryDrill.push({title:`Overtime — ${otYearSel}`, recs});
    wrap.innerHTML=`<div class="ct-table"><table>${head('Year',null)}<tbody>
      <tr style="cursor:pointer" onclick="otSummaryClick(0)"><td><strong>Total Overtime – ${otYearSel}</strong></td>${cells(recs)}</tr></tbody></table></div>`;
    return;
  }

  if(otPeriod==='quarterly'){
    if(!otQuarterSel){ if(titleEl) titleEl.textContent='Quarterly overtime summary'; wrap.innerHTML=`<div class="empty"><div class="empty-title">No quarterly data yet</div></div>`; return; }
    const year=otQuarterSel.split('-Q')[0];
    if(titleEl) titleEl.textContent=`Quarterly overtime summary — ${year}`;
    const qRecs=[1,2,3,4].map(q=>otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&otQuarterKeyOf(o.date)===`${year}-Q${q}`)));
    // Sum of each quarter = Overtime Value column from the sheet (falls back to calculated cost if the sheet has no value column)
    const useRaw=qRecs.some(r=>otSumRawValue(r)>0);
    const qVal=qRecs.map(r=>useRaw?otSumRawValue(r):otSumAmount(r));
    const rows=qRecs.map((recs,i)=>{
      const q=i+1, cur=qVal[i], prev=i>0?qVal[i-1]:null;
      let chg='<span style="color:var(--text3)">N/A</span>', sub='';
      if(q>1){
        if(prev>0){
          const pct=((cur-prev)/prev)*100;                       // QoQ % = (current − previous) / previous × 100
          const up=pct>=0;
          chg=`<span class="pill ${up?'pl-red':'pl-green'}">${up?'▲ Increased':'▼ Decreased'} ${up?'+':''}${pct.toFixed(1)}%</span>`;
        } else if(cur>0){ chg='<span class="pill pl-red">▲ New (no Q'+(q-1)+' value)</span>'; }
        else { chg='<span class="pill pl-gray">No change</span>'; }
        sub=`<span class="ot-sum-sub">vs Q${q-1}: ${otFmtValue(prev)}</span>`;
      }
      otSummaryDrill.push({title:`Overtime — Q${q} ${year}`, recs});
      const hours=otSumHours(recs), emp=otEmployeeCount(recs);
      return `<tr style="cursor:pointer" onclick="otSummaryClick(${i})">
        <td><strong>Total Overtime – Q${q}</strong><span class="ot-sum-sub">${otQuarterRangeLabel(year,q)}</span></td>
        <td style="${mono};font-weight:700">${otFmtValue(cur)}</td>
        <td>${chg}${sub}</td>
        <td style="${mono}">${hours.toFixed(1)}</td>
        <td style="${mono}">${emp}</td></tr>`;
    }).join('');
    const total=qVal.reduce((a,b)=>a+b,0), totalHrs=qRecs.reduce((a,r)=>a+otSumHours(r),0);
    wrap.innerHTML=`<div class="ct-table"><table><thead><tr><th>Quarter</th><th>${useRaw?'Overtime Value (sum)':'Total OT Cost (sum)'}</th><th>QoQ Change %</th><th>Total OT Hrs</th><th>Employees</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr><td><strong>Total ${year} (Q1–Q4)</strong></td><td style="${mono};font-weight:700">${otFmtValue(total)}</td><td></td><td style="${mono}">${totalHrs.toFixed(1)}</td><td></td></tr></tfoot></table></div>`;
    return;
  }

  // Monthly (existing behaviour)
  if(titleEl) titleEl.textContent='Monthly overtime summary';
  const months=otMonthsSorted().reverse(); // most recent first
  if(!months.length){ wrap.innerHTML=`<div class="empty"><div class="empty-title">No monthly data yet</div></div>`; return; }
  const rows=months.map((mKey,i)=>{
    const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===mKey));
    const prevKey=months[i+1]; // previous chronologically = next in this reversed array
    const prevRecs=prevKey?otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===prevKey)):null;
    const momHtml=prevRecs ? otChangePill(otSumAmount(recs), otSumAmount(prevRecs)) : '<span style="color:var(--text3)">—</span>';
    otSummaryDrill.push({title:`Overtime — ${otMonthLabel(mKey)}`, recs});
    return `<tr style="cursor:pointer" onclick="otSummaryClick(${i})"><td><strong>${otMonthLabel(mKey)}</strong></td>${cells(recs)}<td>${momHtml}</td></tr>`;
  }).join('');
  wrap.innerHTML=`<div class="ct-table"><table>${head('Month','MoM Change (Cost)')}<tbody>${rows}</tbody></table></div>`;
}

/* ── Department analysis table: Department = Organization, exclusively. Sorted highest → lowest OT cost ── */
function otRenderDeptTable(){
  const wrap=document.getElementById('ot-dept-table-wrap');
  if(!wrap) return;
  const recs=otFilteredRecords();
  const groups={};
  recs.forEach(o=>{ const d=otOrgOf(o); (groups[d]=groups[d]||[]).push(o); });
  const depts=Object.keys(groups).sort((a,b)=>otSumAmount(groups[b])-otSumAmount(groups[a]));
  if(!depts.length){ wrap.innerHTML=`<div class="empty"><div class="empty-title">No department data for this period</div></div>`; return; }
  const rows=depts.map(d=>{
    const g=groups[d];
    const hours=otSumHours(g), cost=otSumAmount(g), empCount=otEmployeeCount(g);
    const avgPerEmp=empCount?cost/empCount:0;
    return `<tr style="cursor:pointer" data-dept="${otEscAttr(d)}" onclick="otDeptTableClick(this)">
      <td><strong>${d}</strong></td>
      <td style="font-family:'JetBrains Mono',monospace">${hours.toFixed(1)}</td>
      <td style="font-family:'JetBrains Mono',monospace">${otFmtCurrency(cost)}</td>
      <td style="font-family:'JetBrains Mono',monospace">${empCount}</td>
      <td style="font-family:'JetBrains Mono',monospace">${otFmtCurrency(avgPerEmp)}</td>
    </tr>`;
  }).join('');
  wrap.innerHTML=`<div class="ct-table"><table><thead><tr>
    <th>Department</th><th>Total OT Hours</th><th>Total OT Amount</th><th># Employees w/ OT</th><th>Avg OT / Employee</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

/* ── A. KPI Summary — hours-category breakdown row (requirement 1 & 11.A) ── */
function otRenderHoursKPIs(){
  const el=document.getElementById('ot-hours-metrics');
  if(!el) return;
  const recs=otFilteredRecords();
  const actual=otSumActual(recs), normal=otSumNormalHrs(recs), sat=otSumSaturdayHrs(recs), official=otSumOfficialHrs(recs);
  el.innerHTML=`
    <div class="metric"><div class="m-lbl">Actual Hours</div><div class="m-val">${actual>0?actual.toFixed(1):'N/A'}</div></div>
    <div class="metric"><div class="m-lbl">Normal Hours</div><div class="m-val">${normal.toFixed(1)}</div></div>
    <div class="metric"><div class="m-lbl">(Sat Days)</div><div class="m-val">${sat.toFixed(1)}</div></div>
    <div class="metric"><div class="m-lbl">Official Days</div><div class="m-val">${official.toFixed(1)}</div></div>`;
}

/* ── B/C/E. Division comparison, cost comparison & rankings (requirements 1–4, 7, 11.B/C/E) ── */
function otDivisionTableClick(el){
  const div=el.dataset.division;
  const recs=otFilteredRecords().filter(o=>otDivisionOf(o)===div);
  openOtDrillModal(`Division: ${div}`, recs, `${recs.length} record${recs.length!==1?'s':''} · ${otCurrentPeriodLabel()}`);
}
function otRenderDivisionTable(){
  const wrap=document.getElementById('ot-division-table-wrap');
  if(!wrap) return;
  const stats=otBuildDivisionStats(otFilteredRecords()).sort((a,b)=>b.otHrs-a.otHrs);
  if(!stats.length){ wrap.innerHTML=`<div class="empty"><div class="empty-title">No division data for this period</div></div>`; return; }
  const rows=stats.map(s=>`<tr style="cursor:pointer" data-division="${otEscAttr(s.division)}" onclick="otDivisionTableClick(this)">
    <td><strong>${s.division}</strong></td>
    <td style="font-family:'JetBrains Mono',monospace">${s.actual>0?s.actual.toFixed(1):'N/A'}</td>
    <td style="font-family:'JetBrains Mono',monospace">${s.normal.toFixed(1)}</td>
    <td style="font-family:'JetBrains Mono',monospace">${s.sat.toFixed(1)}</td>
    <td style="font-family:'JetBrains Mono',monospace">${s.official.toFixed(1)}</td>
    <td style="font-family:'JetBrains Mono',monospace">${s.otHrs.toFixed(1)}</td>
    <td style="font-family:'JetBrains Mono',monospace">${otFmtCurrency(s.otCost)}</td>
  </tr>`).join('');
  wrap.innerHTML=`<div class="ct-table"><table><thead><tr>
    <th>Division</th><th>Actual Hrs</th><th>Normal Hrs</th><th>Saturday Hrs</th><th>Official Days Hrs</th><th>OT Hrs</th><th>OT Cost</th>
  </tr></thead><tbody>${rows}</tbody></table></div>`;
}

/* ── D-extension. Division × Month comparison matrix (requirement 6) — always spans the last
   6 months present in the (filtered) data, independent of the single-month/quarter selector,
   so multiple months can be compared side by side. ── */
function otDivMonthValue(recs){ return otDivMonthMetric==='cost'?otSumAmount(recs):otSumHours(recs); }
function otDivMonthClick(div,mKey){
  const recs=otApplyExtraFilters(overtimeRecords).filter(o=>otDivisionOf(o)===div&&o.date&&o.date.slice(0,7)===mKey);
  openOtDrillModal(`${div} — ${otMonthLabel(mKey)}`, recs, `${recs.length} record${recs.length!==1?'s':''}`);
}
function otRenderDivMonthMatrix(){
  const wrap=document.getElementById('ot-divmonth-table-wrap');
  if(!wrap) return;
  const base=otApplyExtraFilters(overtimeRecords);
  const months=[...new Set(base.map(o=>o.date?o.date.slice(0,7):null).filter(Boolean))].sort().slice(-6);
  if(!months.length){ wrap.innerHTML=`<div class="empty"><div class="empty-title">No data yet</div></div>`; return; }
  const divTotal=d=>otDivMonthValue(base.filter(o=>otDivisionOf(o)===d));
  const divisions=[...new Set(base.map(o=>otDivisionOf(o)))].sort((a,b)=>divTotal(b)-divTotal(a));
  const rows=divisions.map(div=>{
    const cells=months.map(mKey=>{
      const recs=base.filter(o=>otDivisionOf(o)===div&&o.date.slice(0,7)===mKey);
      const v=otDivMonthValue(recs);
      return `<td style="cursor:pointer;font-family:'JetBrains Mono',monospace" onclick="otDivMonthClick('${otEscAttr(div)}','${mKey}')">${otDivMonthMetric==='cost'?otFmtCurrency(v):v.toFixed(1)}</td>`;
    }).join('');
    let trendHtml='<span style="color:var(--text3)">—</span>';
    if(months.length>1){
      const lastV=otDivMonthValue(base.filter(o=>otDivisionOf(o)===div&&o.date.slice(0,7)===months[months.length-1]));
      const prevV=otDivMonthValue(base.filter(o=>otDivisionOf(o)===div&&o.date.slice(0,7)===months[months.length-2]));
      if(prevV>0){
        const chg=((lastV-prevV)/prevV)*100;
        trendHtml=`<span class="pill ${chg>=0?'pl-red':'pl-green'}">${chg>=0?'▲':'▼'} ${chg>=0?'+':''}${chg.toFixed(1)}%</span>`;
      } else if(lastV>0){ trendHtml='<span class="pill pl-red">▲ New</span>'; }
      else{ trendHtml='<span class="pill pl-gray">No change</span>'; }
    }
    return `<tr><td><strong>${div}</strong></td>${cells}<td>${trendHtml}</td></tr>`;
  }).join('');
  wrap.innerHTML=`<div class="ct-table"><table><thead><tr><th>Division</th>${months.map(m=>`<th>${otMonthLabel(m)}</th>`).join('')}<th>Trend (last 2 months)</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

/* ── F. Normal vs Saturday vs Official Days hours, stacked by division (requirement 8) ── */
function otRenderHoursBreakdownChart(){
  const recs=otFilteredRecords();
  const stats=otBuildDivisionStats(recs).sort((a,b)=>b.otHrs-a.otHrs);
  const labels=stats.map(s=>s.division);
  raRenderChart('ot-hours-breakdown-canvas',{
    type:'bar',
    data:{labels, datasets:[
      {label:'Normal', data:stats.map(s=>+s.normal.toFixed(1)), backgroundColor:raColor(0), borderRadius:4},
      {label:'Saturday', data:stats.map(s=>+s.sat.toFixed(1)), backgroundColor:raColor(1), borderRadius:4},
      {label:'Official Days', data:stats.map(s=>+s.official.toFixed(1)), backgroundColor:raColor(2), borderRadius:4},
    ]},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top', labels:{boxWidth:10,font:{size:11}}}, tooltip:{enabled:true} },
      scales:{ x:{stacked:true, grid:{display:false}, ticks:{font:{size:11}}}, y:{stacked:true, beginAtZero:true, grid:{color:chartGridColor()}} },
      onClick:(evt,elements,chart)=>{
        if(!elements.length) return;
        const div=labels[elements[0].index];
        const filtered=recs.filter(o=>otDivisionOf(o)===div);
        openOtDrillModal(`Division: ${div}`, filtered, `${filtered.length} record${filtered.length!==1?'s':''} · ${otCurrentPeriodLabel()}`);
      },
    }
  });
}

function renderOvertimeAnalysis(){
  const trendCanvas=document.getElementById('ot-trend-canvas');
  const cardsEl=document.getElementById('ot-cards');
  if(!trendCanvas||!cardsEl||typeof Chart==='undefined') return; // page not mounted / Chart.js not loaded yet

  const stdInput=document.getElementById('ot-std-hours');
  if(stdInput && document.activeElement!==stdInput) stdInput.value=otStdHours();

  const multNormalInput=document.getElementById('ot-mult-normal'), multSatInput=document.getElementById('ot-mult-sat'), multOfficialInput=document.getElementById('ot-mult-official');
  if(multNormalInput && document.activeElement!==multNormalInput) multNormalInput.value=otMultNormal();
  if(multSatInput && document.activeElement!==multSatInput) multSatInput.value=otMultSaturday();
  if(multOfficialInput && document.activeElement!==multOfficialInput) multOfficialInput.value=otMultOfficial();

  otPopulateSelects();
  otPopulateFilterSelects();
  otRenderKPIs();
  otRenderHoursKPIs();
  otRenderMoM();
  otRenderDivisionTable();
  otRenderMonthlySummary();
  otRenderDivMonthMatrix();
  otRenderDeptTable();
  document.querySelectorAll('#ot-metric-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.m===otMetric));
  document.querySelectorAll('#ot-divmonth-metric-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.dm===otDivMonthMetric));

  if(!overtimeRecords.length){
    trendCanvas.parentElement.parentElement.style.display='none';
    const breakdownCanvas=document.getElementById('ot-hours-breakdown-canvas');
    if(breakdownCanvas) breakdownCanvas.parentElement.parentElement.style.display='none';
    cardsEl.innerHTML=`<div class="empty" style="grid-column:1/-1"><div class="empty-title">No overtime records yet</div><div>Click "Upload Excel/CSV" to import overtime records from a spreadsheet</div></div>`;
    document.querySelectorAll('#ot-period-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.p===otPeriod));
    return;
  }
  trendCanvas.parentElement.parentElement.style.display='';
  const breakdownCanvas=document.getElementById('ot-hours-breakdown-canvas');
  if(breakdownCanvas) breakdownCanvas.parentElement.parentElement.style.display='';

  const trend=otBuildTrend();
  const titleEl=document.getElementById('ot-trend-title');
  if(titleEl) titleEl.textContent=trend.title;
  const monthShort=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  raRenderChart('ot-trend-canvas',{
    type:trend.type,
    data:{labels:trend.labels, datasets:trend.datasets},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:trend.datasets.length>1, position:'top', labels:{boxWidth:10,font:{size:11}}}, tooltip:{enabled:true, callbacks:{label:ctx=>`${ctx.dataset.label||''}: ${otMetricDisplay(ctx.parsed.y,otMetric)}`}} },
      scales:{ x:{grid:{display:false}, ticks:{font:{size:11}}}, y:{beginAtZero:true, ticks:{font:{size:11}}, grid:{color:chartGridColor()}} },
      onClick:(evt,elements,chart)=>{
        if(!elements.length) return;
        const el=elements[0];
        if(otPeriod==='yearly'){
          const years=otYears();
          const year=years[el.datasetIndex];
          const mm=String(el.index+1).padStart(2,'0');
          const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===`${year}-${mm}`));
          openOtDrillModal(`Overtime — ${monthShort[el.index]} ${year}`, recs, `${recs.length} record${recs.length!==1?'s':''} · ${otMetricDisplay(otMetricValue(recs),otMetric)}`);
        } else if(otPeriod==='quarterly'){
          const [y,q]=otQuarterSel.split('-Q');
          const qm=otQuarterMonths(y,q)[el.index];
          const mm=String(qm.m).padStart(2,'0');
          const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date&&o.date.slice(0,7)===`${qm.y}-${mm}`));
          openOtDrillModal(`Overtime — ${monthShort[qm.m-1]} ${qm.y}`, recs, `${recs.length} record${recs.length!==1?'s':''} · ${otMetricDisplay(otMetricValue(recs),otMetric)}`);
        } else {
          const [y,m]=otMonthSel.split('-');
          const day=String(el.index+1).padStart(2,'0');
          const recs=otApplyExtraFilters(overtimeRecords.filter(o=>o.date===`${y}-${m}-${day}`));
          openOtDrillModal(`Overtime — ${day} ${otMonthLabel(otMonthSel)}`, recs, `${recs.length} record${recs.length!==1?'s':''} · ${otMetricDisplay(otMetricValue(recs),otMetric)}`);
        }
      },
    }
  });

  cardsEl.innerHTML=[
    raCardShell('ot-division-canvas',`By division (${otMetricLabel()})`, 280),
    raCardShell('ot-employee-canvas',`By employee (top 10 — ${otMetricLabel()})`, 280),
    raCardShell('ot-dept-canvas',`By department (${otMetricLabel()})`, 280),
  ].join('');

  otRenderBreakdownCard('ot-division-canvas', o=>otDivisionOf(o), {title:'Division'});
  otRenderBreakdownCard('ot-employee-canvas', o=>o.name, {topN:10, title:'Employee'});
  otRenderBreakdownCard('ot-dept-canvas', o=>otOrgOf(o), {title:'Department'});
  otRenderHoursBreakdownChart();

  document.querySelectorAll('#ot-period-seg .ra-seg-btn').forEach(b=>b.classList.toggle('active', b.dataset.p===otPeriod));
}
/* ─── FILE IMPORT ─── */
function handleFile(file){
  if(!file) return;
  smartImport(file,'contract').then(res=>{
    if(!res) return;
    // No cellDates here on purpose: a renewal-date column formatted "mmm-yy" must arrive as the text "Jan-26"
    parseCSV(XLSX.utils.sheet_to_csv(res.ws));
  });
  document.getElementById('file-in').value='';
}
function parseCSV(text){
  const lines=text.replace(/^\uFEFF/,'').trim().split(/\r?\n/).filter(l=>l.replace(/,/g,'').trim()!=='');
  if(lines.length<2){ toast('CSV file appears empty','danger'); return; }
  const hdr=splitCSVLine(lines[0]).map(h=>h.trim().replace(/"/g,'').toLowerCase());
  // Status / reason columns are identified first so they are never mistaken for the name, date or code column
  const si=hdr.findIndex(h=>h==='status'||h==='renewal status'||h==='renewed'||h==='renewed / pending'||h==='renewed/pending'||h==='contract status');
  const pri=hdr.findIndex(h=>h.includes('pending reason')||h==='reason'||h.includes('reason for pending')||h==='notes'||h==='note');
  const taken=new Set([si,pri].filter(i=>i>=0));
  const find=fn=>hdr.findIndex((h,i)=>!taken.has(i)&&fn(h));
  const ni=find(h=>h.includes('name'));
  const di=find(h=>h.includes('date')||h.includes('renewal')||h.includes('expir'));
  const dpi=find(h=>h.includes('dept')||h.includes('depart'));   // also catches the "Departmet" typo
  const ci=find(h=>h.includes('code')||h==='id'||h.includes('emp'));
  if(ni<0||di<0){ toast('CSV must have "name" and "date"/"renewal_date" columns','danger',0); return; }
  let added=0, updated=0, markedRenewed=0, markedPending=0, unknownStatus=[];
  const stamp=today();
  const readStatus=v=>{
    const t=String(v??'').trim().toLowerCase();
    if(['renewed','yes','y','done','true','1','✓','✔'].includes(t)) return 'renewed';
    if(['pending','no','n','false','0'].includes(t)) return 'pending';
    return null; // blank / unrecognised → leave unchanged
  };
  lines.slice(1).forEach((ln,li)=>{
    const cols=splitCSVLine(ln).map(c=>c.replace(/^"|"$/g,''));
    const name=cols[ni]; const rawDate=cols[di];
    if(!name||!rawDate) return;
    const my=ctParseMonthYear(rawDate);
    const parsed=my ? my.iso : parseDate(rawDate);
    if(!parsed) return;
    const st=si>=0?readStatus(cols[si]):null;
    if(si>=0 && String(cols[si]||'').trim()!=='' && !st) unknownStatus.push(`Row ${li+2} (${name}): "${cols[si]}"`);
    const reason=pri>=0?(cols[pri]||''):'';
    const apply=c=>{
      if(st==='renewed'){ if(!c.renewed){ c.renewed=true; c.renewedOn=stamp; } markedRenewed++; }
      else if(st==='pending'){ c.renewed=false; c.renewedOn=null; if(reason) c.pendingReason=reason; markedPending++; }
    };
    const existing=contracts.find(c=>c.name===name&&c.renewalDate===parsed);
    if(existing){
      // Already loaded → a re-upload refreshes the Renewed / Pending state from the sheet
      if(st){ apply(existing); updated++; }
      return;
    }
    const rec={id:'c'+Date.now()+Math.random(),name,renewalDate:parsed,renewalDateDisplay:my?my.raw:null,dept:dpi>=0?cols[dpi]:'',code:ci>=0?cols[ci]:'',renewed:false,renewedOn:null};
    apply(rec);
    contracts.push(rec);
    added++;
  });
  saveC(); renderContracts(); renderSidebar(); if(document.getElementById("pg-monthly").classList.contains("active")) renderMonthly();
  let msg=`✓ Imported ${added} contract${added!==1?'s':''}`;
  if(updated) msg+=` · updated ${updated} existing`;
  if(si>=0) msg+=` · ${markedRenewed} Renewed, ${markedPending} Pending`;
  toast(msg);
  if(unknownStatus.length) toast(`⚠ ${unknownStatus.length} Status value${unknownStatus.length!==1?'s were':' was'} not Renewed/Pending and left unchanged: ${unknownStatus.slice(0,3).join('; ')}${unknownStatus.length>3?'…':''}`,'warn',0);
}
/* Contract Renewals: some uploaded spreadsheets store the renewal date as an
   Excel "mmm-yy" month cell (e.g. "Jan-26") rather than a full day/month/year
   date. Excel's own two-digit year rule for this format always resolves to
   20XX (never 19XX), so "Jan-26" is January 2026, "Jan-27" is January 2027,
   etc. — never 1926/1927. JavaScript's generic `new Date("Jan-26")` parser
   misreads this as day 26 of a fallback year, so it must be matched and
   converted explicitly, before any generic date parsing is attempted. */
const CT_MONTH_ABBR=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
function ctParseMonthYear(raw){
  if(!raw) return null;
  const s=String(raw).trim();
  const m=s.match(/^([A-Za-z]{3,9})[\s\-\/]?(\d{2}|\d{4})$/);
  if(!m) return null;
  const monIdx=CT_MONTH_ABBR.indexOf(m[1].slice(0,3).toLowerCase());
  if(monIdx<0) return null;
  const yearDigits=m[2];
  const year = yearDigits.length===2 ? 2000+parseInt(yearDigits,10) : parseInt(yearDigits,10);
  const mm=String(monIdx+1).padStart(2,'0');
  // Original text is preserved verbatim (e.g. "Jan-26") so the UI can keep
  // showing it exactly as uploaded; the ISO value is only used internally
  // for sorting/filtering/deadline math.
  return { iso:`${year}-${mm}-01`, raw:s };
}
function parseDate(s){
  if(!s) return null;
  s=String(s).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dmy=s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if(dmy) return `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}`;
  // Raw Excel serial date number (e.g. cell wasn't formatted as a date)
  if(/^\d{4,6}(\.\d+)?$/.test(s)){
    const serial=parseFloat(s);
    const d=new Date(Date.UTC(1899,11,30)+serial*86400000);
    if(!isNaN(d)) return d.toISOString().split('T')[0];
  }
  const my=ctParseMonthYear(s);
  if(my) return my.iso;
  const d=new Date(s);
  if(!isNaN(d)) return d.toISOString().split('T')[0];
  return null;
}
function splitCSVLine(line){
  const out=[]; let cur=''; let inQ=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(ch==='"'){
      if(inQ && line[i+1]==='"'){ cur+='"'; i++; }
      else inQ=!inQ;
    } else if(ch===',' && !inQ){
      out.push(cur); cur='';
    } else { cur+=ch; }
  }
  out.push(cur);
  return out.map(c=>c.trim());
}

/* ─── QUARTERLY CONTRACT ALERT ─── */
// Q1 Mar–May | Q2 Jun–Aug | Q3 Sep–Nov | Q4 Dec–Feb (matches contractQuarter/Breakdown by Quarter)
function getQuarter(date) {
  const m = date.getMonth() + 1; // 1-based month
  if(m>=3&&m<=5) return {q:1, label:'Q1 (Mar–May)'};
  if(m>=6&&m<=8) return {q:2, label:'Q2 (Jun–Aug)'};
  if(m>=9&&m<=11) return {q:3, label:'Q3 (Sep–Nov)'};
  return {q:4, label:'Q4 (Dec–Feb)'}; // Dec, Jan, Feb
}

function getQuarterEnd(date) {
  const m = date.getMonth() + 1;
  const y = date.getFullYear();
  // Returns the last day of the current fiscal quarter
  if(m>=3&&m<=5) return new Date(y, 4, 31);        // May 31
  if(m>=6&&m<=8) return new Date(y, 7, 31);         // Aug 31
  if(m>=9&&m<=11) return new Date(y, 10, 30);       // Nov 30
  if(m===12) return new Date(y+1, 1, 28);           // Feb 28/29 of next year
  return new Date(y, 1, 28);                        // Jan/Feb -> Feb 28/29 of this year
}

function isStillInSameQuarter(renewalDateStr) {
  // Returns true if today is still within the same fiscal quarter as the renewal date
  // meaning the contract has passed its renewal date but we're still in that quarter
  const renewal = new Date(renewalDateStr + 'T00:00:00');
  const now = new Date(); now.setHours(0,0,0,0);
  if(renewal >= now) return false; // not yet past due
  const renewalQ = getQuarter(renewal);
  const nowQ = getQuarter(now);
  // Fiscal year for comparison purposes: Jan/Feb count as part of the previous calendar year's Q4
  const fy = d => { const mo=d.getMonth()+1; return (mo===1||mo===2) ? d.getFullYear()-1 : d.getFullYear(); };
  return renewalQ.q === nowQ.q && fy(renewal) === fy(now);
}

function getQuarterPendingContracts() {
  return contracts.filter(c => !c.renewed && isStillInSameQuarter(ctRenewalDate(c)));
}

function autoQuarterAlert() {
  if(!contracts.length) return;
  const qKey = 'qt_' + (() => { const n=new Date(); return getQuarter(n).q+'_'+n.getFullYear(); })();
  if(alertsSent[qKey]) return;
  const qPending = getQuarterPendingContracts();
  if(!qPending.length) return;
  alertsSent[qKey]=1; saveC();
  setTimeout(()=>{
    const qLabel = getQuarter(new Date()).label;
    toast(`📅 ${qPending.length} contract${qPending.length>1?'s':''} past renewal date but still pending — ${qLabel} is not yet over`, 'warn', 12000);
    sendNotif('Quarterly Renewal Alert', `${qPending.length} contracts still pending within current quarter`, 'ct-quarter');
  }, 800);
}

function sendQuarterAlert() {
  const qPending = getQuarterPendingContracts();
  if(!qPending.length){ toast('No contracts pending within current quarter',''); return; }
  const qLabel = getQuarter(new Date()).label;
  toast(`📅 ${qLabel}: ${qPending.length} contract${qPending.length>1?'s are':' is'} past due date but still pending renewal`, 'warn', 12000);
  sendNotif('Quarterly Renewal Reminder', `${qPending.length} contracts still pending in ${qLabel}`, 'ct-quarter-manual');
}

/* ─── WEEKLY CONTRACT ALERT ─── */
function sendWeeklyAlert(){
  const pending=contracts.filter(c=>!c.renewed);
  if(!pending.length){ toast('No pending contracts'); return; }
  const overdue=pending.filter(c=>daysFrom(ctRenewalDate(c))<0);
  toast(`📋 Weekly reminder: ${pending.length} contract${pending.length!==1?'s':''} pending renewal${overdue.length?' — '+overdue.length+' overdue':''}`, 'purple', 10000);
  sendNotif('Weekly Contract Reminder',`${pending.length} contracts pending renewal`,'ct-weekly');
}
function autoWeeklyAlert(){
  if(!contracts.length) return;
  const wk=weekKey();
  if(alertsSent['wk_'+wk]) return;
  const pending=contracts.filter(c=>!c.renewed);
  if(!pending.length) return;
  alertsSent['wk_'+wk]=1; saveC();
  setTimeout(()=>{
    toast(`📋 Weekly: ${pending.length} contract${pending.length!==1?'s':''} still pending renewal`, 'purple', 10000);
    sendNotif('Weekly Contract Reminder',`${pending.length} contracts pending renewal`,'ct-auto');
  },2200);
}

/* ─── ONBOARDING ALERTS ─── */
function autoOnboardingAlerts(){
  const alerts=allAlerts();
  const key='ob_'+today();
  if(alertsSent[key]||!alerts.length) return;
  alertsSent[key]=1; saveC();
  const danger=alerts.filter(a=>a.type==='danger');
  const warn=alerts.filter(a=>a.type==='warn');
  if(danger.length) setTimeout(()=>{
    toast(`⚠ ${danger.length} overdue onboarding task${danger.length>1?'s':''}`, 'danger', 8000);
    sendNotif('Onboarding Alert',`${danger.length} overdue task${danger.length>1?'s':''}`, 'ob-danger');
  },600);
  if(warn.length) setTimeout(()=>{
    toast(`🔔 ${warn.length} medical mail reminder${warn.length>1?'s':''}`, 'warn', 8000);
    sendNotif('Onboarding Reminder',`${warn.length} medical mail reminders`, 'ob-warn');
  },1400);
}

/* ─── MODAL ─── */
function openModal(){
  document.getElementById('modal-bg').classList.add('open');
  document.getElementById('f-name').focus();
}
function closeModal(){
  document.getElementById('modal-bg').classList.remove('open');
  ['f-name','f-dept','f-title','f-code','f-start','f-hr'].forEach(id=>document.getElementById(id).value='');
}
function saveHire(){
  const name=document.getElementById('f-name').value.trim();
  if(!name){ document.getElementById('f-name').focus(); return; }
  const h={id:'h'+Date.now(),name,dept:document.getElementById('f-dept').value.trim(),title:document.getElementById('f-title').value.trim(),code:document.getElementById('f-code').value.trim(),startDate:document.getElementById('f-start').value,hr:document.getElementById('f-hr').value.trim(),tasks:{},notes:{},createdAt:new Date().toISOString()};
  hires.push(h); save();
  syncContractFromHire(h); // auto-create linked Contract Renewal record (Renewal Date = Joining Date + 1 year)
  closeModal();
  renderSidebar(); renderDashboard();
  if(curView==='monthly') renderMonthly();
  if(curView==='contracts') renderContracts();
  toast(`✓ ${name} added to onboarding`);
}
function editJoiningDate(id){
  const h=hires.find(h=>h.id===id); if(!h) return;
  const val=prompt('Joining date (YYYY-MM-DD):', h.startDate||'');
  if(val===null) return;
  const trimmed=val.trim();
  if(trimmed && !/^\d{4}-\d{2}-\d{2}$/.test(trimmed)){ toast('Please use YYYY-MM-DD format','danger'); return; }
  h.startDate=trimmed; save();
  syncContractFromHire(h); // keep linked Renewal Date (Joining Date + 1 year) in sync
  renderDetail(); renderSidebar(); renderDashboard();
  if(curView==='monthly') renderMonthly();
  if(curView==='contracts') renderContracts();
  toast('✓ Joining date updated — renewal date recalculated');
}
function deleteHire(id){
  if(!confirm('Remove this hire?')) return;
  hires=hires.filter(h=>h.id!==id); save();
  nav('monthly');
  toast('Hire removed');
}
function hToggleSelectMode(){
  hSelectMode=!hSelectMode;
  if(!hSelectMode) hSelected.clear();
  renderDashboard();
}
function hToggleRow(id){
  if(hSelected.has(id)) hSelected.delete(id); else hSelected.add(id);
  renderDashboard();
}
function hSelectAllVisible(checked){
  hires.filter(h=>pfMatch('hire',h.startDate)).forEach(h=> checked ? hSelected.add(h.id) : hSelected.delete(h.id));
  renderDashboard();
}
function hRemoveSelected(){
  const n=hSelected.size;
  if(!n){ toast('No employees selected'); return; }
  if(!confirm(`Remove ${n} selected hire${n!==1?'s':''}? This will delete them from the onboarding list and cannot be undone.`)) return;
  hires=hires.filter(h=>!hSelected.has(h.id));
  hSelected.clear(); hSelectMode=false;
  save(); renderDashboard(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();
  toast(`✓ Removed ${n} hire${n!==1?'s':''}`);
}
function hResetAll(){
  if(!hires.length){ toast('No onboarding data to reset'); return; }
  if(!confirm('Reset ALL Onboarding Tracker data? This permanently removes every hire record and returns this page to its initial empty state. It will not affect Contract Renewals, Offboarding, or Overtime data.')) return;
  hires=[];
  hSelected.clear(); hSelectMode=false;
  document.getElementById('h-upload-errors').innerHTML='';
  save(); renderDashboard(); renderSidebar();
  if(document.getElementById('pg-monthly').classList.contains('active')) renderMonthly();
  toast('✓ Onboarding Tracker data has been reset');
}

document.addEventListener('keydown',e=>{ if(e.key==='Escape') closeModal(); });

/* ─── BOOT ─── */
let appStarted=false;
function initApp(){
  if(appStarted) return;
  appStarted=true;
  reqNotif();
  migrateResignations();
  renderSidebar();
  renderMonthly();
  renderDashboard();
  renderContracts();
  renderOffboarding();
  renderOvertimeAnalysis();
  setTimeout(autoOnboardingAlerts, 800);
  setTimeout(autoWeeklyAlert, 1600);
  setTimeout(autoQuarterAlert, 2400);
}
(async function boot(){
  // If this device was linked to the shared Excel file before, reconnect and merge before the
  // first render. Browsers re-ask for file permission each session, so if it is not already
  // granted we show a "Reconnect" bar instead (one click) and the app keeps working locally.
  try{
    const h=await sbIdbGet('handle');
    if(h){
      sbPendingHandle=h; sbFileName=h.name;
      const p=h.queryPermission?await h.queryPermission({mode:'readwrite'}):'denied';
      if(p==='granted') await sbAttach(h,false,true);
      else sbNeedsPerm=true;
    }
  }catch(e){ console.warn('Shared Excel: boot check failed',e); }
  enterApp();
  renderSyncStatus();
})();
