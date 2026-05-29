# AGENTS.md

## Commands

- Install with `npm install`; this repo uses `package-lock.json`, not pnpm/yarn.
- Run the CLI as `node index.js <laufname> <kategorie> [output-datei]` or `npm start -- <laufname> <kategorie> [output-datei]`.
- Quote categories containing spaces or shell metacharacters, e.g. `node index.js ostufer-fischhallen-lauf-2026 "Hauptlauf (10km)"`.
- `npm test` is a placeholder that exits 1; use `node --check index.js` for a no-network syntax check.
- Full smoke tests call the live Davengo API and write an `.xlsx` file, so only run them when network access and output creation are acceptable.

## Repo Shape

- This is a single-file CommonJS CLI; `index.js` is both the executable entrypoint and the module exporting `fetchAllResults` and `createExcelFile`.
- There is no build step, linter, formatter, typecheck, CI workflow, or test suite configured in the repo.
- `node_modules` and generated `*.xlsx` files are ignored; do not treat spreadsheet outputs as source changes.

## Runtime Details

- `fetchAllResults` POSTs to `https://www.davengo.com/event/result/${laufname}/search/list` and paginates with `data.navigation.nextOffset`.
- Excel headers come from the API `fields`; the hard-coded field list is only a fallback when the API returns no field definitions.
- Workbook sheets are `Gesamt`, optionally `Männer`/`Frauen`, and one `AK <ageGroupShort>` sheet per age group.
