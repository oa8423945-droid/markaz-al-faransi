# French Center - Developer Guide

## Stack and architecture

This is a Windows-local garage management system. The frontend is vanilla HTML, CSS, and JavaScript. The backend is Node.js. Excel is the operational data store.

| File | Purpose |
| --- | --- |
| `server.js` | HTTP API, Excel persistence, code generation, financial logic, PDF/Excel exports |
| `public/index.html` | Page and dialog markup |
| `public/app.js` | Frontend state, events, forms, API calls |
| `public/style.css` | RTL layout and visual styling |
| `data/main data 2.xlsx` | Live operating data - never use production data for development |
| `installer/FrenchCenter.iss` | Inno Setup installer definition |

## Local development

Install Node.js LTS, run `npm install`, then `npm start`, and open `http://localhost:3210`. Use `GARAGE_DATA_FILE` to point to a separate test workbook before making changes.

## Excel schema

Sheet names and their header arrays are declared at the top of `server.js`. When adding a field, update its header array, `readData()`, save logic, API validation, and the frontend form/table. Do not rename existing sheets or headers without a migration.

## Key functions

- `readData()` reads the workbook into application objects.
- `saveData()` writes changed objects to their Excel sheets.
- `nextCode()` generates automatic C, V, M, S, E, A, and T codes.
- `appendAccount()` calculates the running center balance.
- `paymentInfo()` validates split payments.
- `inventoryMovementsWorkbook()` and `inventoryMovementsPdf()` export inventory movement reports.

## Safe change process

1. Back up the workbook and work on a test copy.
2. Implement backend validation first.
3. Update HTML, app logic, and CSS separately.
4. Check syntax with `node --check server.js` and `node --check public/app.js`.
5. Test create, search, financial balance, and export flows.
6. Commit a focused change and publish only after verification.

## Troubleshooting

- Missing workbook: check `GARAGE_DATA_FILE` and the `data/` location.
- Save failure: Excel usually has the workbook open, or a required header changed.
- Route not found: the frontend URL and server endpoint do not match.
- Old PDF: stop old local server instances and restart the application after a server change.

Never put production Excel data, customer data, generated PDFs, or local Node modules in a source-only repository.
