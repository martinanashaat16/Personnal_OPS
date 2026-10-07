/* Loads data/config.json + data/data.json, validates them, then starts the app.
   JSON = configuration + seed data   |   app.js = logic   |   index.html = structure   |   styles.css = look */
(function(){
  const CONFIG_URL='data/config.json', DATA_URL='data/data.json';
  const REQUIRED_CONFIG={
    storageKeys:'object', sync:'object', onboarding:'object', offboarding:'object', quarters:'object',
    dates:'object', overtime:'object', charts:'object', templates:'object'
  };
  const SEED_SHAPE={users:'array',hires:'array',contracts:'array',offboards:'array',obExtraCols:'array',overtime:'array',alertsSent:'object'};

  function typeOf(v){ return Array.isArray(v)?'array':(v===null?'null':typeof v); }
  function fail(title,details){
    const box=document.createElement('div');
    box.style.cssText='font-family:Inter,system-ui,sans-serif;max-width:640px;margin:60px auto;padding:24px;border:1px solid #FCA5A5;background:#FEF2F2;color:#7F1D1D;border-radius:12px;line-height:1.6';
    box.innerHTML='<h2 style="margin:0 0 8px;font-size:18px"></h2><div style="font-size:13.5px"></div>';
    box.firstChild.textContent=title;
    box.lastChild.innerHTML=details.map(d=>'<div>• '+String(d).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))+'</div>').join('');
    document.body.innerHTML=''; document.body.appendChild(box);
  }
  async function getJSON(url){
    let res;
    try{ res=await fetch(url,{cache:'no-cache'}); }
    catch(e){ throw new Error('Could not load '+url+'. If you opened index.html directly from disk (file://), serve the folder over http instead, e.g. "python -m http.server" and open http://localhost:8000.'); }
    if(!res.ok) throw new Error('Could not load '+url+' (HTTP '+res.status+').');
    try{ return await res.json(); }
    catch(e){ throw new Error(url+' is not valid JSON: '+e.message); }
  }
  function validate(cfg,seed){
    const errs=[];
    Object.entries(REQUIRED_CONFIG).forEach(([k,t])=>{ if(typeOf(cfg[k])!==t) errs.push('config.json: "'+k+'" is missing or not an '+t); });
    if(!errs.length){
      if(!Array.isArray(cfg.onboarding.tasks)||!cfg.onboarding.tasks.length) errs.push('config.json: onboarding.tasks must be a non-empty array');
      else cfg.onboarding.tasks.forEach((t,i)=>{ if(!t||!t.id||!t.label) errs.push('config.json: onboarding.tasks['+i+'] needs "id" and "label"'); });
      if(!Array.isArray(cfg.offboarding.tasks)||!cfg.offboarding.tasks.length) errs.push('config.json: offboarding.tasks must be a non-empty array');
      else cfg.offboarding.tasks.forEach((t,i)=>{ if(!t||!t.key||!t.type) errs.push('config.json: offboarding.tasks['+i+'] needs "key" and "type"'); });
      ['calendar','contractFiscal'].forEach(k=>{ if(!Array.isArray(cfg.quarters[k])||cfg.quarters[k].length!==4) errs.push('config.json: quarters.'+k+' must have 4 quarter definitions'); });
      if(!Array.isArray(cfg.dates.monthNames)||cfg.dates.monthNames.length!==13) errs.push('config.json: dates.monthNames must have 13 entries (blank + 12 months)');
      if(!(cfg.overtime.defaultStdHours>0)) errs.push('config.json: overtime.defaultStdHours must be a positive number');
    }
    Object.entries(SEED_SHAPE).forEach(([k,t])=>{ if(typeOf(seed[k])!==t) errs.push('data.json: "'+k+'" is missing or not an '+t); });
    return errs;
  }
  function loadScript(src){
    return new Promise((ok,bad)=>{ const s=document.createElement('script'); s.src=src; s.onload=ok; s.onerror=()=>bad(new Error('Could not load '+src)); document.body.appendChild(s); });
  }
  (async function(){
    try{
      const [cfg,seed]=await Promise.all([getJSON(CONFIG_URL),getJSON(DATA_URL)]);
      const errs=validate(cfg,seed);
      if(errs.length){ fail('The JSON files have problems',errs); return; }
      window.CFG=cfg; window.SEED=seed;
      await loadScript('js/app.js');
      await loadScript('js/theme.js');
    }catch(e){ fail('The app could not start',[e.message]); }
  })();
})();
