# Upstream Synchronization & Maintenance Guide

`spreadsheet-utils` merges and extends three core upstream open-source repositories:
1. [`read-excel-file`](https://gitlab.com/catamphetamine/read-excel-file) by catamphetamine
2. [`tabularjs`](https://github.com/jspreadsheet/tabularjs) by Jspreadsheet
3. [`write-excel-file`](https://gitlab.com/catamphetamine/write-excel-file) by catamphetamine

This document details the file mappings, architectural modifications, and procedures for pulling upstream updates.

---

## 1. Upstream Mapping & Architecture

```
spreadsheet-utils/
├── src/
│   ├── core/                      [Unified Dispatcher & Plugin Registry]
│   │   ├── registry.js            (Dynamic plugin registration & diagnostics)
│   │   ├── dispatcher.js          (Format router & scoped createReader)
│   │   └── detector.js            (Magic-byte sniffer for XLSX & fallback heuristics)
│   │
│   ├── xlsx/, xml/, schema/       [From read-excel-file @ 5.8.6]
│   │   ├── parseSpreadsheetContents.js
│   │   ├── parseXml.js            (Streaming SAX parser using saxen)
│   │   └── parseSheetData.js      (Schema mapping, typecasting & validation)
│   │
│   ├── parser/, plugins/          [From tabularjs @ 0.0.12]
│   │   ├── text.js                (textPlugin: CSV, TSV, TAB, TXT)
│   │   ├── xls.js                 (xlsPlugin: BIFF5/BIFF8 + CFB + PTG formula token decoder)
│   │   ├── markup.js              (markupPlugin: ODS, HTML tables, XML SpreadsheetML 2003)
│   │   └── legacy.js              (legacyPlugin: DBF, Lotus 1-2-3, DIF, SLK, Numbers)
│   │
│   ├── write/                     [From write-excel-file @ 4.1.1]
│   │   ├── index.js               (writeSpreadsheet universal writer + writeXlsxFile alias)
│   │   ├── browser.js             (Browser-safe writer with downloadBlob)
│   │   ├── schema.js              (Bi-directional schema-to-columns transformer)
│   │   ├── xlsx/                  (styles.xml, workbook.xml, sheet.xml generators)
│   │   └── zip/                   (fflate zip streaming and buffer packers)
│   │
│   └── zip/                       [Zero-JSZip Architecture]
│       ├── unzipFromArrayBuffer.js (uses fflate for all zip decompression)
│       └── zip-utils.js           (Zero-decompression ZIP Central Directory header reader)
```

---

## 2. Key Architectural Invariants

When merging upstream updates, ensure these core principles are preserved:

1. **Zero JSZip:**
   - TabularJS originally bundled JSZip (`~100 KB`).
   - In `spreadsheet-utils`, all ZIP operations across XLSX, ODS, Numbers, and Writer use `fflate`.
   - Never reintroduce `jszip` into dependencies.

2. **Low-Memory Streaming SAX:**
   - XLSX is parsed via `saxen` streaming XML events, avoiding large in-memory DOM document trees.

3. **Core vs Plugin Separation:**
   - The default export `spreadsheet-utils` only bundles XLSX reader and schema parsing (~44 KB minified in browser).
   - Non-XLSX formats remain in `spreadsheet-utils/plugins`.
   - The writer remains in `spreadsheet-utils/write`.
   - All formats are combined in `spreadsheet-utils/full`.

4. **Multi-Target Bundles:**
   - ESM (`dist/esm/*.mjs`)
   - CommonJS (`dist/cjs/*.cjs`)
   - Browser `<script>` tags (`dist/browser/*.min.js`)

---

## 2b. Local Additions and Edits Inside Synchronized Paths

`scripts/sync-upstream.js` treats `src/write`, `src/parser`, `src/plugins`, `src/utils`, `src/xlsx`, `src/xml`, `src/schema`, `src/export` and `src/zip` as upstream territory. These are the local additions and edits in them; a sync must keep them (re-apply after taking upstream's version of the file).

### `src/write` (write-excel-file @ 4.1.1)

| File | Change |
| :--- | :--- |
| `src/write/xlsx/features/rowBreaks.js` | **New.** The `rowBreaks` sheet option (v1.1.0): a feature transform like `stickyRowsOrColumns.js` that inserts `<rowBreaks count manualBreakCount><brk id max="16383" man="1"/></rowBreaks>` by the sheet schema order (after `<pageSetup/>`/`<headerFooter/>`, before `<drawing/>`). Sorts, de-duplicates and validates (positive integers; otherwise throws). |
| `src/write/xlsx/getFeatures.js` | **Edited.** Imports `rowBreaks` and pushes it after `stickyRowsOrColumns`. |

### `src/zip` (read-excel-file @ 5.8.6)

| File | Change |
| :--- | :--- |
| `src/zip/unzipFromArrayBuffer.js` | **Edited** (v1.1.0), two bugs. (1) `unzipFromArrayBufferUsingFunction()` called `.then` on the result even for the synchronous `unzipSync()`, so the sync path threw; it now honours `isAsync` (returns the files, throws `UnzipError` for a bad archive). (2) The async `unzipAsync()` ignored the options, so the `filter` and the `maxUncompressedSize` zip-bomb guard never applied in the browser/universal reader; the options (`{ filter }`) now go to fflate's `unzip()` / `unzipSync()`. |
| `src/zip/unzipFromArrayBufferSync.js` | **Edited** (v1.1.0). Passes `isAsync: false`. |

### `src/utils` and `src/parser` (tabularjs @ 0.0.12)

| File | Change |
| :--- | :--- |
| `src/utils/loader.js` | **Edited** (v1.1.0). `import('fs')` goes through one `loadFs()` helper with a non-literal specifier and a stand-in that answers "no such file" when `fs` can't load, so the `browser-sync` entry bundles for the browser with a plain esbuild run. Behaviour in Node and in the other browser bundles is unchanged. |
| `src/parser/odsCore.js` | **New name** (v1.1.0). The former `src/parser/ods.js`, unchanged except that `parseODS(input, options)` is now `parseODSWith(unzipArchive, input, options)` and the `zip-utils.js` import is gone, so the caller picks the unzip function (async for the main entries, synchronous for `browser-sync`). |
| `src/parser/ods.js` | **New thin file** (v1.1.0). `parseODS(input, options)` = `parseODSWith(unzipArchive, ...)` with the async `unzipArchive` of `zip-utils.js`; same export as before. |

### Outside the synchronized paths (not upstream, nothing to merge)

* `src/browser-sync.js` and `src/sync/` (`unzipSync.js`, `parseXLSXSync.js`, `writeXlsxFileSync.js`): the `spreadsheet-utils/browser-sync` entry. It uses `fflate`'s `zipSync()` / `unzipSync()` only, so no Worker is started (fflate's async `zip()` / `unzip()` start a `blob:` Worker for larger files, which a Content Security Policy without `worker-src` blocks). It imports unpackers and parsers from the synchronized paths; if an upstream sync renames `validateLeadingBytes`, `filterZipArchiveEntry`, `parseSpreadsheetContents`, `parseXLSBuffer` or `generateXlsxFileContents`, update those imports. `test/integration/browser-sync.test.js` fails on any `Worker` in the entry's bundle.
* `src/core/shapeReadResult.js`: the sheet selection / `allSheets` / `asObjects` / `schema` part of `readSpreadsheet()`, moved out of `src/core/dispatcher.js` so the `browser-sync` reader shares it.

---

## 3. Synchronization CLI Tool

`scripts/sync-upstream.js` provides automation to check and sync upstream repositories:

### Check Status
Display local mappings and tracked upstream releases:
```bash
npm run sync:check
# or: node scripts/sync-upstream.js --status
```

### Check for Upstream Releases
Queries the npm registry for new versions of the upstream packages:
```bash
node scripts/sync-upstream.js --check
```

### Staging an Upstream Clone
Clone an upstream repo into `.upstream-staging/` for diffing:
```bash
# Example: clone read-excel-file
npm run sync:pull read-excel-file

# Example: clone tabularjs
npm run sync:pull tabularjs

# Example: clone write-excel-file
npm run sync:pull write-excel-file
```

### Run Verification
After merging upstream updates, run the automated verification pipeline:
```bash
npm run sync:verify
```
This builds all `dist/` targets and executes all 465+ test cases.

---

## 4. Upstream Remote Setup (Git Alternative)

For developers working directly with Git remotes:

```bash
# Add upstream remotes
git remote add upstream-read-excel-file https://gitlab.com/catamphetamine/read-excel-file.git
git remote add upstream-tabularjs https://github.com/jspreadsheet/tabularjs.git
git remote add upstream-write-excel-file https://gitlab.com/catamphetamine/write-excel-file.git

# Fetch all remotes
git fetch --all --tags
```
