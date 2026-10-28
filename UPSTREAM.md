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
