/* Applies the saved light/dark theme before the page paints (prevents a flash). Keep this loaded in <head>. */
try{var _t=localStorage.getItem('hr3_theme');if(_t==='dark')document.documentElement.setAttribute('data-theme','dark');}catch(e){}
