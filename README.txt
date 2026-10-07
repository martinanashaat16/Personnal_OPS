PersonnelOps — folder layout
  index.html        page markup (open this file in Chrome or Edge)
  css/styles.css    all styling
  js/theme-init.js  applies the saved light/dark theme before the page draws
  js/app.js         main application logic (data, imports, shared Excel sync, charts)
  js/theme.js       theme toggle and chart colours

Keep the folder structure exactly as it is. To share with other users, put the WHOLE
folder in the shared OneDrive location (not just index.html).
The page loads three things from the internet (Excel reader, Chart.js, Google Fonts),
so each PC needs internet access when opening the app.
