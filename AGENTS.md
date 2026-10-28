# Spreadsheet Utils — Project Guidelines

## Project Overview
`spreadsheet-utils` is an ultra-lightweight, zero-dependency core, low-memory universal spreadsheet toolkit merging streaming SAX reading, multi-format plugins, and styled multi-format export for 16+ formats in Node.js and the browser.

## Architecture & Code Organization
- **Core Dispatcher & Registry** (`src/core/`):
  - `registry.js`: Dynamic format parser/writer plugin registry and diagnostics.
  - `dispatcher.js`: Multi-format router (`readSpreadsheet`, `createReader`) and streaming coordinator.
  - `detector.js`: Format sniffing via magic bytes and heuristics.
- **Reading Pipelines**:
  - `src/xlsx/`, `src/xml/`, `src/schema/`: SAX-based streaming XLSX parser (`saxen`).
  - `src/plugins/`: Text (`src/plugins/text.js`), XLS (`src/plugins/xls.js`), Markup/ODS (`src/plugins/markup.js`, `src/plugins/ods.js`), Legacy/DBF (`src/plugins/legacy.js`, `src/plugins/dbf.js`).
- **Writing Pipelines** (`src/write/`):
  - `index.js` & `browser.js`: `writeSpreadsheet` universal entrypoints supporting `xlsx`, `xls`, `ods`, `dbf`, `csv`, `tsv`, `html`, `json`.
  - Format writers: `src/write/xlsx/`, `src/write/xls.js`, `src/write/ods.js`, `src/write/dbf.js`, `src/write/formats.js`.
  - Schema mapping: `src/write/schema.js` for converting schema-based data to columns.
- **Decompression & Archives** (`src/zip/`):
  - Strictly relies on `fflate` and `unzipper-esm`. No `jszip` or heavy archive libraries allowed.

## Architectural Invariants
1. **Core Zero-Dependency**: Keep the core lightweight; only minimal required parsing/compression libraries (`fflate`, `saxen`, `unzipper-esm`). Never introduce heavy dependencies like `exceljs` or `xlsx`/`sheetjs`.
2. **Dual Node.js & Browser Support**: All export and parsing pathways must maintain browser compatibility. Avoid unconditional Node-specific built-ins (`fs`, `stream`) in universal modules without appropriate browser guards or isomorphic abstractions.
3. **Buffer & File Symmetry**: `writeSpreadsheet` must support both writing to a filepath (Node.js) or returning a `Buffer` (`{ buffer: true }` in Node, `Blob` in browser) across all exportable formats.

## Development & Verification Workflow
- **Build**: `npm run build` bundles ESM, CJS, and browser minified targets into `dist/`. Always build before running tests or finalizing changes.
- **Test Suite**: Run `npm test` to execute the full mocha test suite across unit, integration, and upstream test cases.
- **Integration Tests**: Place new format export or plugin tests in `test/integration/`.
