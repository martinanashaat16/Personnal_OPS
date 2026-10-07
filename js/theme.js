function chartGridColor(){ return document.documentElement.getAttribute('data-theme')==='dark' ? '#2A3650' : '#EEF1F5'; }
function chartSurfaceColor(){ return document.documentElement.getAttribute('data-theme')==='dark' ? '#161E2E' : '#fff'; }
function applyChartTheme(){
  if(typeof Chart==='undefined') return;
  const dark=document.documentElement.getAttribute('data-theme')==='dark';
  Chart.defaults.color=dark?'#A3B1C8':'#666';
  Chart.defaults.borderColor=dark?'#2A3650':'rgba(0,0,0,0.1)';
}
function updateThemeBtn(){
  const b=document.getElementById('theme-toggle'); if(!b) return;
  const dark=document.documentElement.getAttribute('data-theme')==='dark';
  b.textContent=dark?'☀️':'🌙';
  b.title=dark?'Switch to light mode':'Switch to dark mode';
}
function toggleTheme(){
  const dark=document.documentElement.getAttribute('data-theme')==='dark';
  if(dark) document.documentElement.removeAttribute('data-theme'); else document.documentElement.setAttribute('data-theme','dark');
  try{ localStorage.setItem(CFG.storageKeys.theme, dark?'light':'dark'); }catch(e){}
  updateThemeBtn(); applyChartTheme();
  try{ refreshAllViews(); }catch(e){}
}
applyChartTheme(); updateThemeBtn();
