# French Center Source Code Guide

The source-code repository contains the application UI, server, installer scripts, and documentation only. It must not contain a real customer workbook.

## Getting started

1. Create a blank workbook using the required sheet names and headers, then save it as `data/main data 2.xlsx`.
2. Run `npm install`.
3. Run `npm start`.
4. Open `http://localhost:3210`.

## Where to make changes

- UI markup: `public/index.html`
- UI behavior: `public/app.js`
- Styles and responsive RTL layout: `public/style.css`
- Data, APIs, business rules, reports: `server.js`
- Windows installer: `installer/FrenchCenter.iss`

## Repository hygiene

Keep `data/main data 2.xlsx`, generated reports, logs, and `node_modules` out of the source repository. Use `.gitignore`, test with an empty/example workbook, and keep commits small and descriptive.
