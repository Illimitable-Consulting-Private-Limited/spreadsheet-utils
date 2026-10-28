# spreadsheet-utils

[![Build & Tests](https://img.shields.io/badge/tests-576%20passing-brightgreen.svg)](https://github.com/Illimitable-Consulting-Private-Limited/spreadsheet-utils/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Bundle Size (Core Reader)](https://img.shields.io/badge/Core%20Reader-44%20KB%20min%20%2F%2014%20KB%20gzip-success)](https://bundlephobia.com)
[![Formats Supported](https://img.shields.io/badge/Formats%20Supported-16%2B-blueviolet)](#supported-formats-matrix)

**Ultra-lightweight, low-memory universal spreadsheet toolkit for JavaScript & TypeScript** (Node.js, Browsers, and Web Workers).

Merges the streaming, low-memory SAX architecture of **[`read-excel-file`](https://www.npmjs.com/package/read-excel-file)** with the multi-format versatility of **[`tabularjs`](https://github.com/jspreadsheet/tabularjs)** and the rich styled workbook generator from **[`write-excel-file`](https://www.npmjs.com/package/write-excel-file)**.

---

## Table of Contents

- [Why spreadsheet-utils?](#why-spreadsheet-utils)
  - [Architecture Comparison](#architecture-comparison)
  - [Supported Formats & Capabilities Matrix](#supported-formats--capabilities-matrix)
- [Installation](#installation)
- [Distribution Targets & Subpaths](#distribution-targets--subpaths)
- [Quick Start](#quick-start)
  - [1. Reading XLSX (Core)](#1-reading-xlsx-core)
  - [2. Multi-Format Reading with Plugins](#2-multi-format-reading-with-plugins)
  - [3. Isolated Scoped Reader (`createReader`)](#3-isolated-scoped-reader-createreader)
  - [4. All-in-One Zero-Config Preset (`spreadsheet-utils/full`)](#4-all-in-one-zero-config-preset-spreadsheet-utilsfull)
- [Streaming Row Consumer](#streaming-row-consumer)
  - [`iterateSpreadsheet` Async Generator](#iteratespreadsheet-async-generator)
  - [`onRow` Callback](#onrow-callback)
- [Universal File Converter (`convertSpreadsheet`)](#universal-file-converter-convertspreadsheet)
- [CLI Tool (`spreadsheet-utils`)](#cli-tool-spreadsheet-utils)
- [Streaming Exporters for Massive Datasets](#streaming-exporters-for-massive-datasets)
- [Background Thread Offloading (`spreadsheet-utils/worker`)](#background-thread-offloading-spreadsheet-utilsworker)
- [Security & Malicious Input Hardening](#security--malicious-input-hardening)
- [Schema Mapping & Data Validation](#schema-mapping--data-validation)
  - [Auto-Object Reading (`asObjects: true`)](#auto-object-reading-asobjects-true)
- [Styled Writer & Multi-Format Exporter](#styled-writer--multi-format-exporter)
  - [Ready-to-Use Styling Themes](#ready-to-use-styling-themes)
  - [Auto Column Widths with `maxWidth` Safeguard](#auto-column-widths-with-maxwidth-safeguard)
  - [Writing Styled XLSX Files](#writing-styled-xlsx-files)
  - [Writing Objects with Schema](#writing-objects-with-schema)
  - [Multi-Sheet Workbooks](#multi-sheet-workbooks)
  - [Multi-Format Exporter (`writeCSV`, `writeTSV`, `writeHTML`, `writeJSON`)](#multi-format-exporter-writecsv-writetsv-writehtml-writejson)
  - [DBF File Export with Schema & Plugins](#dbf-file-export-with-schema--plugins)
  - [XLS (Excel 97–2004) File Export (`xlsPlugin`)](#xls-excel-972004-file-export-xlsplugin)
  - [ODS (OpenDocument Spreadsheet) Export (`markupPlugin` / `odsPlugin`)](#ods-opendocument-spreadsheet-export-markupplugin--odsplugin)
  - [Format Styling & Differences Matrix](#format-styling--differences-matrix)
- [Browser `<script>` Tag Usage](#browser-script-tag-usage)
- [CommonJS (`require`) Support](#commonjs-require-support)
- [Interactive Browser Playground](#interactive-browser-playground)
- [Complete API Reference](#complete-api-reference)
  - [`readSpreadsheet(input, [options])`](#readspreadsheetinput-options)
  - [`iterateSpreadsheet(input, [options])`](#iteratespreadsheetinput-options)
  - [`readInWorker(input, [options])`](#readinworkerinput-options)
  - [`createReader([plugins])`](#createreaderplugins)
  - [`registerPlugin(...plugins)`](#registerpluginplugins)
  - [`writeSpreadsheet(data, [options], [extraOptions])`](#writespreadsheetdata-options-extraoptions)
  - [`writeCSV` / `writeTSV` / `writeHTML` / `writeJSON`](#writecsv--writetsv--writehtml--writejson)
- [Upstream Synchronization & CI](#upstream-synchronization--ci)
- [License](#license)

---

## Why spreadsheet-utils?

Spreadsheet libraries in JavaScript typically force a difficult trade-off:
1. **SheetJS (`xlsx`)**: Supports many formats, but has a large bundle footprint (~800 KB minified), high memory usage because it loads full DOM structures, and non-permissive dual licensing for commercial use.
2. **`read-excel-file`**: Outstanding low-memory streaming SAX parser with declarative schema validation, but strictly limited to `.xlsx`.
3. **`tabularjs`**: Impressive 16+ format parser suite, but lacks streaming row consumption, bundles `jszip` (~100 KB), and doesn't offer an integrated styled exporter.
4. **`write-excel-file`**: Lightweight, beautiful styled XLSX generator, but separate from reading toolkits.

**`spreadsheet-utils` combines the best of all worlds:**
- **Modular & Tree-Shakable**: The core package is only **~44 KB minified (~14 KB gzipped)** containing the streaming SAX `.xlsx` reader and schema validator. Formats like CSV, XLS, ODS, Numbers, and DBF are packaged in grouped plugins so applications only load what they need.
- **Zero JSZip**: JSZip (`~100 KB`, high memory) has been completely replaced with [`fflate`](https://github.com/101arrowz/fflate) (`~8 KB`, ultra-fast streaming inflate/deflate).
- **Constant Memory Streaming**: Memory remains constant whether parsing 10 rows or 1,000,000 rows using `iterateSpreadsheet` or `options.onRow`.
- **Security Hardened**: Built-in limits for uncompressed ZIP size, file size, rows, and columns prevent ZIP bombs and CPU exhaustion.
- **Background Worker Ready**: Seamlessly offloads parsing to Web Workers in the browser or `node:worker_threads` in Node.js.
- **Roundtrip Styled XLSX & Multi-Format Exporter**: Read any format, manipulate data, and export back to styled `.xlsx`, `.csv`, `.tsv`, `.html`, or `.json`.

### Architecture Comparison

| Feature | `spreadsheet-utils` | `read-excel-file` | `tabularjs` | `xlsx` (SheetJS) |
| :--- | :---: | :---: | :---: | :---: |
| **Core Bundle Size (minified)** | **44 KB** | 42 KB | 280 KB | ~850 KB |
| **ZIP Engine** | **`fflate` (~8 KB)** | `fflate` / `unzipper` | `jszip` (~100 KB) | Custom internal |
| **XML Parsing** | **Streaming SAX (`saxen`)** | Streaming SAX | DOMParser / XMLDOM | Custom regex / SAX |
| **Supported Formats** | **16+ formats** | 1 (`.xlsx`) | 16+ formats | 15+ formats |
| **Plugin Architecture** | **Yes (Grouped plugins)** | No | Monolithic | Monolithic |
| **Streaming Row Consumer** | **Yes (`iterateSpreadsheet`)** | Partial (raw SAX) | No | No |
| **Schema Validation & Typing** | **Yes (Built-in)** | Yes | No | No |
| **Styled XLSX Writer** | **Yes (`spreadsheet-utils/write`)** | No | No | Paid Pro feature |
| **Multi-Format Exporter** | **Yes (CSV, TSV, HTML, JSON)** | No | No | Partial |
| **Worker Offloader** | **Yes (`spreadsheet-utils/worker`)** | Manual | Manual | Manual |
| **Browser Script Tag Bundle** | **Yes (Standalone IIFE)** | Yes | Yes | Yes |
| **License** | **MIT** | MIT | MIT | Apache 2.0 / Commercial |

---

### Supported Formats & Capabilities Matrix

`spreadsheet-utils` provides universal reading, writing, conversion, and streaming across 16+ spreadsheet and tabular data formats:

| Format | Extensions | Read | Write | Multi-Sheet | Streaming $O(1)$ | Schema / Objects | Themes & Styles | Auto-Width | Required Module / Plugin |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Excel Modern** | `.xlsx` | Yes (SAX) | Yes | Yes | Read (Stream) | Yes | Yes (24-bit RGB) | Yes (with `maxWidth`) | **Core built-in** |
| **Excel Binary Legacy** | `.xls` | Yes (BIFF5/8) | Yes (BIFF8) | Yes | Read (Iterate) | Yes | Yes (Palette) | Yes (`COLINFO`) | `xlsPlugin` |
| **OpenDocument** | `.ods` | Yes (OASIS) | Yes | Yes | Read (Iterate) | Yes | Yes (24-bit RGB) | Yes (`table-column`) | `markupPlugin` / `odsPlugin` |
| **Comma-Separated** | `.csv` | Yes | Yes | Single | Read & Write $O(1)$ | Yes | Text | N/A | `textPlugin` |
| **Tab-Separated** | `.tsv`, `.tab`, `.txt` | Yes | Yes | Single | Read & Write $O(1)$ | Yes | Text | N/A | `textPlugin` |
| **HTML Table** | `.html`, `.htm` | Yes (`<table>`) | Yes (`createHTML`) | Single | Read (Iterate) | Yes | Yes (Inline CSS) | Yes (`<col>`) | `markupPlugin` |
| **XML Spreadsheet** | `.xml` | Yes (SpreadsheetML) | Via plugins | Yes | Read (Iterate) | Yes | Styles Parsed | Yes | `markupPlugin` |
| **dBASE Database** | `.dbf` | Yes (III, IV, V, VFP) | Yes (III, VFP) | Single | Read (Iterate) | Yes (Typed fields) | Binary fields | Yes (Field width) | `legacyPlugin` / `dbfPlugin` |
| **JSON / JSONL** | `.json`, `.jsonl` | Yes | Yes | Yes | Write $O(1)$ (JSONL) | Yes | N/A | N/A | Core built-in |
| **Apple Numbers** | `.numbers` | Yes (iWork '09–'13) | Read-only | Yes | Read (Iterate) | Yes | Text / Values | N/A | `legacyPlugin` |
| **Lotus 1-2-3** | `.wks`, `.wk1`, `.wk3`, `.wk4`, `.123` | Yes | Read-only | Single/Multi | Read (Iterate) | Yes | Text / Values | N/A | `legacyPlugin` |
| **Data Interchange** | `.dif` | Yes | Read-only | Single | Read (Iterate) | Yes | Text / Values | N/A | `legacyPlugin` |
| **Symbolic Link** | `.slk`, `.sylk` | Yes | Read-only | Single | Read (Iterate) | Yes | Text / Values | N/A | `legacyPlugin` |

---

## Installation

```bash
npm install spreadsheet-utils
```

---

## Distribution Targets & Subpaths

`spreadsheet-utils` provides first-class support for ES Modules (`import`), CommonJS (`require`), TypeScript types, and direct browser `<script>` tag inclusion:

| Subpath / Target | Description | Browser Minified Size |
| :--- | :--- | :--- |
| **`spreadsheet-utils`** | **Core Reader**: SAX streaming XLSX reader, schema validator, plugin registry | **44 KB** (~14 KB gzip) |
| **`spreadsheet-utils/plugins`** | **Format Plugins**: `textPlugin`, `xlsPlugin`, `markupPlugin`, `legacyPlugin`, `allPlugins` | **92 KB** |
| **`spreadsheet-utils/write`** | **Styled Writer & Exporter**: `writeSpreadsheet`, `writeXlsxFile`, `writeCSV`, `writeTSV`, `writeHTML`, `writeJSON` | **73 KB** |
| **`spreadsheet-utils/full`** | **All-in-One**: Core reader + all format plugins pre-registered + styled writer | **192 KB** (~60 KB gzip) |
| **`spreadsheet-utils/worker`** | **Worker Offloader**: `readInWorker` for background thread processing | Zero-setup worker thread |

---

## Quick Start

### 1. Reading XLSX (Core)

By default, the core package parses `.xlsx` files with zero plugin registration needed:

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

// Read an XLSX file into a 2D array of rows
const rows = await readSpreadsheet('financials.xlsx');
console.log(rows);
// Output:
// [
//   ['Quarter', 'Revenue', 'Profit'],
//   ['Q1', 120000, 34000],
//   ['Q2', 145000, 42000]
// ]
```

### 2. Multi-Format Reading with Plugins

Non-XLSX formats are loaded via **grouped plugins**, keeping your production bundle lean. If you attempt to parse a format without registering its plugin, `spreadsheet-utils` throws a clear diagnostic error naming the exact plugin required:

```javascript
import { readSpreadsheet, registerPlugin } from 'spreadsheet-utils';
import { textPlugin, xlsPlugin } from 'spreadsheet-utils/plugins';

// Enable CSV, TSV, and legacy Excel (.xls)
registerPlugin(textPlugin, xlsPlugin);

// Now read CSV or legacy XLS transparently
const csvRows = await readSpreadsheet('users.csv');
const xlsRows = await readSpreadsheet('archive.xls');
```

#### Available Grouped Plugins

```javascript
import {
  textPlugin,    // .csv, .tsv, .tab, .txt
  xlsPlugin,     // .xls (BIFF5/BIFF8 binary Excel 97-2004 + CFB + PTG formula decoder)
  markupPlugin,  // .ods (OpenDocument), .html/.htm (HTML tables), .xml (SpreadsheetML 2003)
  legacyPlugin,  // .dbf (dBase), .dif, .slk, Lotus 1-2-3 (.wks/.wk1/.wk3/.wk4/.123), .numbers
  allPlugins     // Array containing all 4 plugins
} from 'spreadsheet-utils/plugins';

// Register all at once:
registerPlugin(...allPlugins);
```

### 3. Isolated Scoped Reader (`createReader`)

In multi-tenant or modular applications where you don't want to mutate the global plugin registry:

```javascript
import { createReader } from 'spreadsheet-utils';
import { textPlugin } from 'spreadsheet-utils/plugins';

// Scoped reader with only textPlugin registered
const csvReader = createReader([textPlugin]);

const rows = await csvReader('data.csv');
```

### 4. All-in-One Zero-Config Preset (`spreadsheet-utils/full`)

If bundle size is not a constraint and you want all 16+ formats and the styled writer pre-loaded:

```javascript
import { readSpreadsheet, writeSpreadsheet } from 'spreadsheet-utils/full';

// Reads any format automatically
const rows = await readSpreadsheet('data.numbers');

// Writes styled XLSX
await writeSpreadsheet(rows, 'output.xlsx');
```

---

## Streaming Row Consumer

When parsing huge spreadsheets (e.g. hundreds of thousands or millions of rows), buffering the entire 2D matrix in memory can cause out-of-memory errors. `spreadsheet-utils` provides two streaming APIs:

### `iterateSpreadsheet` Async Generator

Yields rows asynchronously as they are parsed from the underlying XML SAX stream. Memory usage remains constant throughout the entire read:

```javascript
import { iterateSpreadsheet } from 'spreadsheet-utils';

// Constant heap memory regardless of file size
for await (const row of iterateSpreadsheet('massive-dataset.xlsx')) {
  console.log('Streamed row:', row);
  
  // Early termination: breaking immediately stops XML stream parsing!
  if (row[0] === 'STOP') {
    break;
  }
}
```

### `onRow` Callback

Process each row on-the-fly without accumulating rows in memory:

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

let totalSum = 0;

await readSpreadsheet('transactions.xlsx', {
  onRow: (row, rowIndex) => {
    if (rowIndex > 0) { // Skip header
      totalSum += Number(row[2] || 0);
    }
  }
});

console.log('Total:', totalSum);
```

---

## Universal File Converter (`convertSpreadsheet`)

Convert any spreadsheet file or buffer directly into another format with zero configuration. It automatically sniffs the input format and encodes to the target extension:

```javascript
import { convertSpreadsheet } from 'spreadsheet-utils/full';

// 1. One-liner CSV to styled XLSX with theme and responsive widths
await convertSpreadsheet('data.csv', 'report.xlsx', { theme: 'ocean', autoWidth: true });

// 2. Convert legacy XLS to modern ODS
await convertSpreadsheet('archive.xls', 'modern.ods');

// 3. Export spreadsheet to standalone styled HTML document
await convertSpreadsheet('sales.xlsx', 'report.html', { theme: 'forest', fullDocument: true });

// 4. Memory buffer to buffer conversion (Node.js & Browser)
const xlsxBuffer = await convertSpreadsheet(csvBuffer, {
  from: 'csv',
  format: 'xlsx',
  buffer: true,
  theme: 'slate'
});
```

---

## CLI Tool (`spreadsheet-utils`)

`spreadsheet-utils` provides a command-line interface for inspecting, converting, and viewing spreadsheets directly from your terminal:

```bash
# Convert between any formats (xlsx, ods, xls, csv, tsv, dbf, html, json)
spreadsheet-utils convert data.csv report.xlsx --theme ocean --auto-width

# Inspect spreadsheet format, sheet count, sheet names, and dimensions
spreadsheet-utils info workbook.xlsx

# Pretty-print spreadsheet to terminal as a formatted table
spreadsheet-utils cat data.xlsx --limit 25

# Dump as JSON objects using first row as headers
spreadsheet-utils cat data.xlsx --objects

# Dump as CSV
spreadsheet-utils cat data.xlsx --csv
```

---

## Streaming Exporters for Massive Datasets

Stream millions of rows from async iterators, database queries, or generators directly into CSV, TSV, or JSON Lines with **constant O(1) memory**:

```javascript
import { streamCSV, streamTSV, streamJSONL, iterateSpreadsheet } from 'spreadsheet-utils/full';

// 1. Stream 500,000+ rows directly from XLSX to CSV with 0 memory bloat:
await streamCSV(
  iterateSpreadsheet('huge-dataset.xlsx'),
  'output.csv'
);

// 2. Stream an async database cursor / generator to JSON Lines (JSONL):
async function* fetchDatabaseRows() {
  for await (const row of db.users.find()) {
    yield { id: row._id, name: row.name, email: row.email };
  }
}

await streamJSONL(fetchDatabaseRows(), 'users.jsonl');

// 3. Stream to any Node.js Writable stream (HTTP response, process.stdout):
await streamCSV(fetchDatabaseRows(), res, { delimiter: ',' });
```

---

## Background Thread Offloading (`spreadsheet-utils/worker`)

Offload heavy spreadsheet parsing off the main thread:
- **In Node.js**: Uses `node:worker_threads`.
- **In the Browser**: Uses the Web Worker API.

```javascript
import { readInWorker } from 'spreadsheet-utils/worker';

// Parsing runs in a background thread; UI / event-loop never freezes
const rows = await readInWorker('heavy-report.xlsx');
console.log('Parsed rows:', rows.length);
```

---

## Security & Malicious Input Hardening

Spreadsheet parsers are common vectors for Denial-of-Service attacks (ZIP bombs, 10-million column allocations, or infinite loops). `spreadsheet-utils` provides configurable guardrails:

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

const rows = await readSpreadsheet(userUploadedBuffer, {
  // 1. Max physical file size in bytes (e.g., 20 MB)
  maxFileSize: 20 * 1024 * 1024,

  // 2. Max decompressed size for ZIP-based formats (.xlsx, .ods, .numbers)
  // Stops decompression bombs before memory is exhausted!
  maxUncompressedSize: 50 * 1024 * 1024,

  // 3. Dimensional limits
  maxRows: 10000,
  maxColumns: 100,

  // 4. Behavior when limit exceeded: true throws error; false truncates cleanly
  throwOnLimit: true
});
```

---

## Schema Mapping & Data Validation

Extract typed JavaScript objects directly from spreadsheet rows with full validation:

```javascript
import { readSpreadsheet, registerPlugin, Integer, Email, URL } from 'spreadsheet-utils';
import { textPlugin } from 'spreadsheet-utils/plugins';

registerPlugin(textPlugin);

const schema = {
  'USER_ID': {
    prop: 'id',
    type: Integer,
    required: true
  },
  'FULL_NAME': {
    prop: 'name',
    type: String,
    required: true
  },
  'EMAIL_ADDRESS': {
    prop: 'email',
    type: Email,
    required: false
  },
  'JOIN_DATE': {
    prop: 'joinedAt',
    type: Date,
    required: false
  },
  'STATUS': {
    prop: 'status',
    type: String,
    oneOf: ['active', 'inactive', 'pending']
  }
};

const { rows, errors } = await readSpreadsheet('users.csv', { schema });

if (errors.length > 0) {
  console.warn('Validation errors:', errors);
}

console.log(rows);
// [
//   { id: 101, name: 'Alice Smith', email: 'alice@example.com', joinedAt: 2024-01-15T00:00:00.000Z, status: 'active' },
//   { id: 102, name: 'Bob Jones', email: 'bob@example.com', joinedAt: 2024-02-01T00:00:00.000Z, status: 'pending' }
// ]
```

### Schema Property Options

| Property | Type | Description |
| :--- | :--- | :--- |
| `prop` | `string` | Target object property name |
| `type` | `Function` | Type constructor (`String`, `Number`, `Boolean`, `Date`, `Integer`, `Email`, `URL`) or custom parsing function |
| `required` | `boolean \| (row) => boolean` | Whether the cell is required. Supports conditional requirements based on row data |
| `validate` | `(value) => void` | Custom validator function. Throw an error with a custom message if invalid |
| `parse` | `(value) => any` | Custom value parser / transformer |
| `oneOf` | `any[]` | Array of allowed values |

### Auto-Object Reading (`asObjects: true`)

If you don't need a formal schema and simply want an array of JavaScript objects mapped using the first row as property names:

```javascript
// Automatically maps the first row to object keys:
const users = await readSpreadsheet('users.csv', { asObjects: true });
// -> [ { USER_ID: '1', FULL_NAME: 'Alice', EMAIL_ADDRESS: 'alice@example.com' }, ... ]

// Also supported when streaming row-by-row with iterateSpreadsheet:
for await (const user of iterateSpreadsheet('large-file.xlsx', { asObjects: true })) {
  console.log(user.FULL_NAME);
}
```

---

## Styled Writer & Multi-Format Exporter

### Ready-to-Use Styling Themes

Produce professionally designed spreadsheets with zero styling boilerplate. Themes render with **100% visual parity across Microsoft Excel, LibreOffice Calc, and HTML**:

| Theme | Header Palette | Alternate Row (Zebra) | Best For |
|---|---|---|---|
| `'ocean'` | Navy `#1B365D`, white bold text | `#F0F4F8` | Corporate, finance, executive |
| `'forest'` | Emerald `#1B4332`, white bold text | `#F2F9F4` | Environmental, growth, sales |
| `'slate'` | Charcoal `#334155`, white bold text | `#F8FAFC` | Modern SaaS, technical, analytics |
| `'sunset'` | Terracotta `#9C4123`, white bold text | `#FFF7ED` | Marketing, operations |
| `'striped'` | Steel `#475569`, white bold text | `#F1F5F9` | Clean subtle alternating rows |

```javascript
await writeSpreadsheet(data, {
  filePath: 'sales.xlsx',
  theme: 'ocean',
  autoWidth: true
});
```

### Auto Column Widths with `maxWidth` Safeguard

Automatically computes responsive column widths based on maximum content length across all rows. Includes a default safeguard `maxWidth: 40` to guarantee that long descriptions or multi-sentence paragraphs never stretch column widths or break sheet layouts:

```javascript
// Automatically size columns with safe bounds:
await writeSpreadsheet(data, {
  filePath: 'report.xlsx',
  autoWidth: true // minWidth: 8, maxWidth: 40, padding: 2
});

// Or customize safeguards:
await writeSpreadsheet(data, {
  filePath: 'report.xlsx',
  autoWidth: { minWidth: 10, maxWidth: 30, padding: 3 }
});
```

Export styled spreadsheets with custom colors, fonts, borders, column widths, and format masks:

```javascript
import { writeSpreadsheet } from 'spreadsheet-utils/write';

const data = [
  [
    { value: 'Product', fontWeight: 'bold', backgroundColor: '#e2e8f0', align: 'left' },
    { value: 'Units Sold', fontWeight: 'bold', backgroundColor: '#e2e8f0', align: 'right' },
    { value: 'Price', fontWeight: 'bold', backgroundColor: '#e2e8f0', align: 'right' }
  ],
  [
    { value: 'Wireless Headphones', type: String },
    { value: 1250, type: Number, format: '#,##0' },
    { value: 89.99, type: Number, format: '$#,##0.00' }
  ],
  [
    { value: 'Mechanical Keyboard', type: String },
    { value: 640, type: Number, format: '#,##0' },
    { value: 129.50, type: Number, format: '$#,##0.00' }
  ]
];

// Write directly to file in Node.js
await writeSpreadsheet(data, 'sales.xlsx');

// Or return a Buffer (in Node.js) or Blob (in Browser)
const buffer = await writeSpreadsheet(data, { buffer: true });
```

### Cell Styling Properties

| Property | Values / Examples | Description |
| :--- | :--- | :--- |
| `fontWeight` | `'bold'`, `'normal'` | Font weight |
| `fontStyle` | `'italic'`, `'normal'` | Font style |
| `fontSize` | `12`, `14`, `18` | Font size in points |
| `fontFamily` | `'Arial'`, `'Calibri'` | Font family name |
| `textColor` | `'#ffffff'`, `'#2563eb'` | Hex text color |
| `backgroundColor` | `'#f1f5f9'`, `'#dbeafe'` | Hex fill background color |
| `borderColor` | `'#cccccc'`, `'#000000'` | Border color |
| `borderStyle` | `'thin'`, `'medium'`, `'thick'`, `'dashed'`, `'dotted'`, `'double'` | Border line style |
| `align` | `'left'`, `'center'`, `'right'` | Horizontal alignment |
| `verticalAlign` | `'top'`, `'center'`, `'bottom'` | Vertical alignment |
| `wrap` | `true`, `false` | Text wrapping inside cell |
| `format` | `'#,##0'`, `'$#,##0.00'`, `'yyyy-mm-dd'` | Excel number or date format mask |
| `columnSpan` | `2`, `3` | Number of columns to merge |
| `rowSpan` | `2`, `3` | Number of rows to merge |

### Writing Objects with Schema

```javascript
const users = [
  { id: 1, name: 'Alice Smith', balance: 1450.75 },
  { id: 2, name: 'Bob Jones', balance: 920.00 }
];

const schema = {
  id: { column: 'User ID', type: Number, width: 12 },
  name: { column: 'Customer Name', type: String, width: 25 },
  balance: { column: 'Outstanding Balance', type: Number, format: '$#,##0.00', width: 20 }
};

await writeSpreadsheet(users, { schema, filePath: 'customers.xlsx' });
```

### Multi-Sheet Workbooks

```javascript
const sheets = [
  {
    sheet: 'Q1 Sales',
    data: [
      [{ value: 'Month' }, { value: 'Revenue' }],
      [{ value: 'January' }, { value: 10000 }]
    ]
  },
  {
    sheet: 'Q2 Sales',
    data: [
      [{ value: 'Month' }, { value: 'Revenue' }],
      [{ value: 'April' }, { value: 15000 }]
    ]
  }
];

await writeSpreadsheet(sheets, 'quarterly-report.xlsx');
```

### Multi-Format Exporter (`writeCSV`, `writeTSV`, `writeHTML`, `writeJSON`)

In addition to `.xlsx`, export data to standard text and web formats:

```javascript
import { writeCSV, writeTSV, writeHTML, writeJSON, writeSpreadsheet } from 'spreadsheet-utils/write';

const data = [
  ['ID', 'Name', 'Role'],
  [1, 'Alice', 'Admin'],
  [2, 'Bob', 'Member']
];

// RFC 4180 compliant CSV
await writeCSV(data, 'export.csv');

// Tab-separated TSV
await writeTSV(data, 'export.tsv');

// HTML <table> with inline styles or complete standalone document
await writeHTML(data, { filePath: 'table.html', fullDocument: true });

// JSON array of objects or 2D matrix
await writeJSON(data, { filePath: 'data.json', asObjects: true });

// Or let writeSpreadsheet auto-route by file extension:
await writeSpreadsheet(data, 'report.csv');
await writeSpreadsheet(data, 'report.html');
await writeSpreadsheet(data, 'report.json');
```

### DBF File Export with Schema & Plugins

Export data to standard dBASE database files (`.dbf` - dBASE III, Visual FoxPro) with declarative schema definitions and full roundtrip compatibility:

```javascript
import { writeSpreadsheet, writeDBF } from 'spreadsheet-utils/write';
import { registerPlugin, dbfPlugin, legacyPlugin } from 'spreadsheet-utils/plugins';

// Enable DBF export by registering either dbfPlugin or legacyPlugin
registerPlugin(dbfPlugin);

const users = [
  { id: 101, name: 'Alice', balance: 1250.50, active: true, joined: new Date('2024-01-15') },
  { id: 102, name: 'Bob', balance: 450.00, active: false, joined: new Date('2024-06-20') }
];

// 1. Extended schema-based write:
const schema = {
  id: { column: 'ID', type: Number, size: 6 },
  name: { column: 'NAME', type: String, size: 30 },
  balance: { column: 'BALANCE', type: Number, size: 10, decimalPlaces: 2 },
  active: { column: 'ACTIVE', type: Boolean },
  joined: { column: 'JOINED', type: Date }
};

await writeSpreadsheet(users, { schema, filePath: 'users.dbf' });

// 2. Direct writeDBF function:
await writeDBF(users, { schema, filePath: 'users.dbf' });

// 3. Or with DBFFile (yortus/DBFFile) fieldDescriptors:
const fieldDescriptors = [
  { name: 'ID', type: 'I', size: 4 },
  { name: 'TITLE', type: 'C', size: 40 },
  { name: 'PRICE', type: 'N', size: 10, decimalPlaces: 2 }
];

const books = [
  { ID: 1, TITLE: 'Clean Code', PRICE: 29.99 }
];

await writeDBF(books, { fieldDescriptors, filePath: 'books.dbf', fileVersion: 0x30 });
```

#### DBF Export Options

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `schema` | `Object \| Array` | `undefined` | Extended schema mapping columns, types (`String`, `Number`, `Boolean`, `Date`), byte `size`, `decimalPlaces`, and `value` extractors |
| `fieldDescriptors` | `Array` | `undefined` | yortus/`DBFFile` style field descriptor array (`[{ name, type, size, decimalPlaces }]`) |
| `fileVersion` | `number` | `0x03` | `0x03` (dBASE III), `0x30` (Visual FoxPro) |
| `encoding` | `string` | `'utf-8'` | Character encoding: `'utf-8'`, `'latin1'` / `'iso-8859-1'`, `'ascii'` |
| `lastUpdate` | `Date` | `new Date()` | Custom last-update date in the DBF header |
| `deletedRows` | `number[]` | `undefined` | Array of row indices to mark as deleted (`0x2A` flag) |
| `isRowDeleted` | `Function` | `undefined` | Predicate `(row, index) => boolean` to mark records deleted |
| `filePath` | `string` | `undefined` | Output file path (Node.js) |
| `buffer` | `boolean` | `true` (Node) | Return Node.js `Buffer` |
| `blob` | `boolean` | `false` | Return browser `Blob` |
| `fileName` | `string` | `undefined` | File name for browser download |

### XLS (Excel 97–2004) File Export (`xlsPlugin`)

Export spreadsheets to legacy binary Microsoft Excel format (`.xls` BIFF8 with Compound File Binary / OLE2 packaging). Supports 2D grids, schema mapping, multi-sheet workbooks, and cell styling:

```javascript
import { writeSpreadsheet, writeXLS, createXLS } from 'spreadsheet-utils/write';
import { registerPlugin, xlsPlugin } from 'spreadsheet-utils/plugins';

// Enable XLS writing by registering xlsPlugin
registerPlugin(xlsPlugin);

const data = [
  ['Product', 'Q1 Sales', 'In Stock'],
  ['Widget A', 15400, true],
  ['Widget B', 23100, false]
];

// 1. Export via writeSpreadsheet (auto-detects .xls extension)
await writeSpreadsheet(data, 'sales.xls');

// 2. Direct writeXLS function
await writeXLS(data, { filePath: 'sales.xls' });

// 3. In-memory binary generation
const xlsUint8 = createXLS(data);

// 4. Multi-sheet workbooks
const sheets = [
  { sheet: 'North America', data: [['Region', 'Sales'], ['USA', 1500000]] },
  { sheet: 'Europe', data: [['Region', 'Sales'], ['UK', 800000]] }
];
await writeSpreadsheet(sheets, 'global-sales.xls');
```

### ODS (OpenDocument Spreadsheet) Export (`markupPlugin` / `odsPlugin`)

Export spreadsheets to OASIS OpenDocument Spreadsheet format (`.ods`), compatible with LibreOffice Calc, Apache OpenOffice, Google Sheets, and Microsoft Excel:

```javascript
import { writeSpreadsheet, writeODS, createODS } from 'spreadsheet-utils/write';
import { registerPlugin, markupPlugin, odsPlugin } from 'spreadsheet-utils/plugins';

// Enable ODS writing by registering markupPlugin or odsPlugin
registerPlugin(markupPlugin);

const employees = [
  { id: 101, name: 'Alice Smith', salary: 75000, hireDate: new Date('2021-03-15') },
  { id: 102, name: 'Bob Johnson', salary: 62000, hireDate: new Date('2022-07-01') }
];

const schema = [
  { column: 'Employee ID', type: Number, value: row => row.id, fontWeight: 'bold' },
  { column: 'Full Name', type: String, value: row => row.name },
  { column: 'Annual Salary', type: Number, value: row => row.salary },
  { column: 'Hire Date', type: Date, value: row => row.hireDate }
];

// 1. Export via writeSpreadsheet with schema
await writeSpreadsheet(employees, { filePath: 'employees.ods', schema });

// 2. Direct writeODS function
await writeODS(employees, { filePath: 'employees.ods', schema });

// 3. In-memory binary generation
const odsUint8 = createODS(employees, { schema });
```

---

## Universal Schema Styling & Cell Overrides

`spreadsheet-utils` provides 100% API compatibility with `write-excel-file` across **XLSX**, **XLS**, **ODS**, and **HTML**. The exact same dataset and schema definition work across all formats.

### 1. Header Styling (Object Syntax & `headerStyle`)

Headers can be styled per-column using an object definition or globally using the `headerStyle` option:

```javascript
import { writeSpreadsheet, createHTML } from 'spreadsheet-utils/write';

const employees = [
  { id: 101, name: 'Alice Smith', dept: 'Engineering', salary: 120000, bonus: 15000 },
  { id: 102, name: 'Bob Johnson', dept: 'Marketing', salary: 75000, bonus: -2000 }
];

const schema = [
  // A. Header defined as an object with value and custom styles:
  {
    header: {
      value: 'Employee ID',
      fontWeight: 'bold',
      backgroundColor: '#1F4E79', // Navy blue
      color: '#FFFFFF',           // White text
      align: 'center'
    },
    prop: 'id',
    type: Number,
    width: 15
  },

  // B. String header with headerStyle shorthand:
  {
    column: 'Full Name',
    prop: 'name',
    headerStyle: {
      fontWeight: 'bold',
      backgroundColor: '#2E75B6',
      color: '#FFFFFF'
    },
    width: 25
  },

  // C. Dynamic cell overrides via cell(row, index) callback:
  {
    header: 'Performance Bonus',
    cell: (row) => ({
      value: row.bonus,
      type: Number,
      format: '$#,##0.00',
      align: 'right',
      // Conditional styling:
      fontWeight: row.bonus > 5000 ? 'bold' : 'normal',
      backgroundColor: row.bonus > 5000 ? '#D9EAD3' : (row.bonus < 0 ? '#FFC7CE' : undefined),
      color: row.bonus > 5000 ? '#274E13' : (row.bonus < 0 ? '#9C0006' : undefined)
    }),
    width: 20
  }
];

// Export to any format using the exact same schema:
await writeSpreadsheet(employees, { schema, filePath: 'report.xlsx' }); // XLSX
await writeSpreadsheet(employees, { schema, filePath: 'report.xls' });  // XLS (BIFF8)
await writeSpreadsheet(employees, { schema, filePath: 'report.ods' });  // ODS
await writeSpreadsheet(employees, { schema, filePath: 'report.html' }); // HTML table
```

### 2. Direct HTML String Generation (`createHTML`)

For rendering HTML tables directly to strings (e.g., in server-side templates or web apps):

```javascript
import { createHTML } from 'spreadsheet-utils/write';

// Generates an HTML <table> string with <thead>, <tbody>, inline CSS, and column widths:
const htmlTable = createHTML(employees, { schema });

// Wrap in complete standalone HTML5 document:
const htmlDoc = createHTML(employees, { schema, fullDocument: true, title: 'Employee Report' });
```

### Format Styling & Differences Matrix

Every export format in `spreadsheet-utils` uses the identical API definition: `writeSpreadsheet(data, { schema, theme, autoWidth, ... })`. The table below outlines exact styling and layout feature support across the export targets:

| Feature / Styling Property | XLSX (`.xlsx`) | ODS (`.ods`) | XLS (`.xls`) | HTML Table (`.html`) | CSV / TSV / JSON / DBF |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Built-in Themes (`theme: 'ocean' \| 'forest' \| ...`)** | Yes | Yes | Yes | Yes | Plain text / data values |
| **Alternating Zebra Backgrounds** | Yes | Yes | Yes | Yes | N/A |
| **Auto Column Widths (`autoWidth: true`)** | Yes | Yes | Yes (`BIFF.COLINFO`) | Yes (`<col style="...">`) | DBF (field length) |
| **`maxWidth` Layout Safeguard** | Yes | Yes | Yes | Yes | N/A |
| **Header Styling (`headerStyle` / rich header)** | Yes | Yes | Yes | Yes (`<th>`) | Header row names |
| **Dynamic Row/Cell Overrides (`cell: (row) => ({ ... })`)** | Yes | Yes | Yes | Yes | Cell values |
| **Rich Value Return (`value: (row) => ({ value, ... })`)** | Yes | Yes | Yes | Yes | Cell values |
| **Font Styles (`fontWeight`, `fontStyle`)** | Yes | Yes | Yes | Yes (`font-weight`, `font-style`) | N/A |
| **Font Family & Size (`fontFamily`, `fontSize`)** | Yes | Yes | Yes | Yes (`font-family`, `font-size`) | N/A |
| **Text Color (`color` / `textColor`)** | Yes (24-bit RGB) | Yes (24-bit RGB) | Yes (16-color palette) | Yes (Hex/CSS) | N/A |
| **Cell Background Color (`backgroundColor`)** | Yes (24-bit RGB) | Yes (24-bit RGB) | Yes (16-color palette) | Yes (Hex/CSS) | N/A |
| **Borders (`borderColor`, `borderStyle`, per-side)** | Yes | Yes | Yes | Yes (`border: 1px solid`) | N/A |
| **Cell Alignment (`align`, `verticalAlign`)** | Yes | Yes | Yes | Yes (`text-align`, `vertical-align`)| N/A |
| **Text Wrap (`wrap: true`)** | Yes | Yes | Yes | Yes (`white-space: normal`) | N/A |
| **Merged Cells (`columnSpan`, `rowSpan`)** | Yes | Yes | Yes (`MERGECELLS`) | Yes (`colspan`, `rowspan`) | N/A |
| **Formulas** | Yes (OpenXML) | Yes (OpenFormula) | Read-only | String | String |
| **Multi-Sheet Support** | Yes | Yes | Yes | Single (or multi-table) | Single (JSON: array) |
| **Synchronous String Generation** | N/A | N/A | N/A | Yes (`createHTML(data)`) | Yes (`createCSV`, etc.) |
| **In-Memory Binary Buffer (`{ buffer: true }`)** | Yes (`Buffer`) | Yes (`Buffer`) | Yes (`Buffer` / `createXLS`) | Yes (`Buffer`) | Yes (`createDBF`) |
| **Browser Blob Download (`{ fileName: '...' }`)** | Yes (`Blob`) | Yes (`Blob`) | Yes (`Blob`) | Yes (`Blob`) | Yes (`Blob`) |
| **Constant $O(1)$ Memory Streaming** | Read (SAX) | Read (SAX) | Read (Iterate) | Read (Iterate) | **Read & Write $O(1)$** (`streamCSV`, `streamJSONL`) |
| **Embedded Images** | Yes (OpenXML) | No | No | No | No |
| **OpenXML Conditional Formatting Rules** | Yes (`<dxfs>`) | Use `cell()` callback | Use `cell()` callback | Use `cell()` callback | N/A |

---

## Browser `<script>` Tag Usage

Standalone browser bundles are available in `dist/browser/`. Include them directly via `<script>` tag:

### Core Reader

```html
<script src="dist/browser/spreadsheet-utils.min.js"></script>
<script>
  // Exposes window.spreadsheetUtils
  const fileInput = document.getElementById('myFile');
  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    const rows = await window.spreadsheetUtils.readSpreadsheet(file);
    console.log(rows);
  });
</script>
```

### All-in-One Full Bundle (Reader + All Formats + Writer)

```html
<script src="dist/browser/spreadsheet-utils-full.min.js"></script>
<script>
  // Exposes window.spreadsheetUtilsFull
  async function processSpreadsheet(file) {
    // Reads any format: XLSX, XLS, ODS, CSV, Numbers, DBF, etc.
    const rows = await window.spreadsheetUtilsFull.readSpreadsheet(file);

    // Write back and automatically trigger a browser download
    await window.spreadsheetUtilsFull.writeSpreadsheet(rows, 'converted.xlsx');
  }
</script>
```

---

## CommonJS (`require`) Support

```javascript
const { readSpreadsheet, registerPlugin } = require('spreadsheet-utils');
const { textPlugin, xlsPlugin } = require('spreadsheet-utils/plugins');
const { writeSpreadsheet } = require('spreadsheet-utils/write');

registerPlugin(textPlugin, xlsPlugin);

async function run() {
  const rows = await readSpreadsheet('data.csv');
  await writeSpreadsheet(rows, 'output.xlsx');
}
run();
```

---

## Interactive Browser Playground

An interactive demo is included in [`examples/index.html`](file:///home/illimitable/Documents/vibe-projects/spreadsheet-utils/examples/index.html):
- **Drag-and-Drop Parsing**: Drop any `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv`, `.numbers`, `.dbf`, or `.html` file.
- **Data Preview Table**: Sticky headers, column inspection, row counters, and format detection badge.
- **Multi-Format Instant Export**: Download parsed data as `.xlsx`, `.csv`, `.html`, or `.json`.
- **Sample Data Generator**: Test the toolkit in one click.

To run it locally:
```bash
npx serve .
# Open http://localhost:3000/examples/
```

---

## Complete API Reference

### `readSpreadsheet(input, [options])`

Parses a spreadsheet from a file path, Buffer, Uint8Array, ArrayBuffer, ReadableStream, or browser File/Blob.

```typescript
function readSpreadsheet(input: SpreadsheetInput, options?: CommonOptions): Promise<Row[]>;
function readSpreadsheet<T>(input: SpreadsheetInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
function readSpreadsheet(input: SpreadsheetInput, options: { format: 'tabular' }): Promise<TabularSpreadsheet>;
function readSpreadsheet(input: SpreadsheetInput, options: { allSheets: true }): Promise<SheetEntry[]>;
```

#### Options:

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `sheet` | `string \| number` | `1` | Sheet name or 1-based sheet index to parse |
| `allSheets` | `boolean` | `false` | When `true`, returns an array of `{ name, data }` for all sheets |
| `format` | `'rows' \| 'tabular'` | `'rows'` | When `'tabular'`, returns TabularJS workbook structure with styles/merges |
| `schema` | `Schema` | `undefined` | Schema object for mapping rows to typed objects |
| `onRow` | `(row: Row, index: number) => void` | `undefined` | Callback invoked as each row is parsed |
| `delimiter` | `string` | auto | Delimiter character for CSV/TSV |
| `tableIndex` | `number` | `0` | 0-based index of `<table>` when parsing HTML |
| `dateFormat` | `string` | auto | Date format mask |
| `trim` | `boolean` | `false` | Trim whitespace from string cell values |
| `maxRows` | `number` | `undefined` | Maximum number of rows to parse |
| `maxColumns` | `number` | `undefined` | Maximum number of columns per row |
| `maxFileSize` | `number` | `undefined` | Maximum input file size in bytes |
| `maxUncompressedSize` | `number` | `undefined` | Maximum uncompressed ZIP extraction size (ZIP bomb guard) |
| `throwOnLimit` | `boolean` | `true` | Throw an error when limit is exceeded (if `false`, truncates cleanly) |

---

### `iterateSpreadsheet(input, [options])`

Returns an asynchronous generator that yields rows one-by-one with constant heap memory usage.

```typescript
function iterateSpreadsheet(input: SpreadsheetInput, options?: CommonOptions): AsyncGenerator<Row, void, unknown>;
```

---

### `readInWorker(input, [options])`

Offloads spreadsheet parsing to a background thread (`node:worker_threads` in Node.js, Web Worker in browser).

```typescript
function readInWorker(input: SpreadsheetInput, options?: CommonOptions): Promise<Row[]>;
function readInWorker<T>(input: SpreadsheetInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
```

---

### `createReader([plugins])`

Creates a scoped, isolated reader instance with its own plugin registry.

```typescript
function createReader(plugins?: SpreadsheetPlugin | SpreadsheetPlugin[]): ScopedReader;
```

---

### `registerPlugin(...plugins)`

Registers one or more plugins into the global registry.

```typescript
function registerPlugin(...plugins: SpreadsheetPlugin[]): void;
function unregisterPlugin(target?: string | SpreadsheetPlugin): void;
function isPluginRegistered(name: string): boolean;
```

---

### `writeSpreadsheet(data, [options], [extraOptions])`

Generates a spreadsheet. Automatically routes between `.xlsx`, `.csv`, `.tsv`, `.html`, and `.json` based on file extension or `options.format`.

```typescript
function writeSpreadsheet(
  data: WriteGrid | Record<string, any>[] | SheetSpec[] | { worksheets: any[] },
  options?: WriteSheetOptions | string,
  extraOptions?: any
): Promise<Uint8Array | Buffer | Blob | string | void>;
```

---

### `writeCSV` / `writeTSV` / `writeHTML` / `createHTML` / `writeJSON`

Dedicated format serializes:

```typescript
function writeCSV(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
function writeTSV(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
function writeHTML(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
function createHTML(data: any, options?: WriteSheetOptions | string): string;
function writeJSON(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
```

---

## Upstream Synchronization & CI

`spreadsheet-utils` maintains compatibility with its upstream base projects (`read-excel-file`, `tabularjs`, and `write-excel-file`):

- **Automated GitHub Actions ([`.github/workflows/upstream-sync.yml`](file:///home/illimitable/Documents/vibe-projects/spreadsheet-utils/.github/workflows/upstream-sync.yml))**:
  - Automatically runs every Monday at 00:00 UTC and on manual dispatch.
  - Queries npm/GitLab registries for newer upstream releases.
  - Runs full matrix test verification across Node 18, 20, and 22.

- **Maintenance CLI Scripts**:
```bash
# Check upstream releases
npm run sync:check

# Clone upstream repo for diff comparison
npm run sync:pull read-excel-file
npm run sync:pull tabularjs
npm run sync:pull write-excel-file

# Rebuild bundles and run verification
npm run sync:verify
```

See [`UPSTREAM.md`](file:///home/illimitable/Documents/vibe-projects/spreadsheet-utils/UPSTREAM.md) for full maintenance guidelines.

---

## License

[MIT](file:///home/illimitable/Documents/vibe-projects/spreadsheet-utils/LICENSE)
