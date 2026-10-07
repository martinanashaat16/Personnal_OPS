# PersonnelOps – JSON-based structure

index.html        structure / containers
css/styles.css    all styling (unchanged from original)
js/loader.js      loads + validates data/*.json, then starts the app
js/app.js         all application logic (calculations, filters, uploads, charts, sync)
js/theme.js, theme-init.js   dark/light theme
data/config.json  configuration: onboarding & offboarding checklists, quarter definitions,
                  month names/aliases, chart palette, OT defaults, Excel templates, storage keys, sync settings
data/data.json    seed data (users, hires, contracts, offboards, obExtraCols, overtime, alertsSent).
                  Used only when the browser has no saved data for that collection.

Run: serve the folder over http (e.g. `python -m http.server`) and open http://localhost:8000
(browsers block fetch() of JSON from file://).
