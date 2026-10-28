# spreadsheet-utils

Ultra-lightweight, low-memory universal spreadsheet toolkit for JavaScript (Node.js, Browsers, and Web Workers).

Merges the streaming, low-memory architecture of **[`read-excel-file`](https://www.npmjs.com/package/read-excel-file)** with the extensive format support of **[`tabularjs`](https://github.com/jspreadsheet/tabularjs)** and the styled XLSX generator from **[`write-excel-file`](https://www.npmjs.com/package/write-excel-file)**.

---

## Highlights

- **Lightweight Core & Modular Plugin Architecture**:
  - Core package bundles **only** the streaming SAX XLSX reader and schema validator (**~44 KB minified / ~14 KB gzipped**).
  - Non-XLSX formats are packaged into **grouped plugins** (`textPlugin`, `xlsPlugin`, `markupPlugin`, `legacyPlugin`), keeping client bundles small and tree-shaken.
  - An all-in-one preset (`spreadsheet-utils/full`) is available for zero-config multi-format support.
- **16+ File Formats Supported**:
  - **Excel**: `.xlsx` (Core streaming SAX), `.xls` (BIFF5/BIFF8 binary via `xlsPlugin`)
  - **OpenDocument**: `.ods` (via `markupPlugin`)
  - **Delimited Text**: `.csv`, `.tsv`, `.tab`, `.txt` (via `textPlugin`)
  - **Databases**: `.dbf` (dBase III, IV, V, Visual FoxPro via `legacyPlugin`)
  - **Legacy Spreadsheets**: Lotus 1-2-3 (`.wks`, `.wk1`, `.wk3`, `.wk4`, `.123`), `.dif` (Data Interchange Format), `.slk` / `.sylk` (Symbolic Link) (via `legacyPlugin`)
  - **Markup & Web**: `.html`, `.htm` (HTML tables), `.xml` (Excel 2003 XML / SpreadsheetML) (via `markupPlugin`)
  - **Apple**: `.numbers` (Apple Numbers documents via `legacyPlugin`)
- **Styled XLSX Export & Roundtrip**:
  - Integrated `writeSpreadsheet` and `writeXlsxFile` for generating styled `.xlsx` workbooks with custom fonts, colors, borders, alignments, column spans, and formatting.
  - Supports converting raw rows, objects with schemas, multi-sheet workbooks, and TabularJS worksheets directly to XLSX.
- **Low Memory & Zero JSZip**:
  - Replaced JSZip (`~100 KB`, high memory) with `fflate` (`~8 KB`) across all ZIP operations (`.xlsx`, `.ods`, `.numbers`, and writer).
  - Streaming SAX XML parsing (`saxen`) for `.xlsx`—never constructs giant in-memory XML DOM trees.
- **Multi-Target Distribution**:
  - **ESM** (`import` in Node.js 18+ and modern bundlers)
  - **CommonJS** (`require` in Node.js)
  - **Browser `<script>` tags** (Standalone IIFE bundles exposing global variables)
- **TypeScript Support**:
  - Comprehensive TypeScript declarations included (`index.d.ts`, `plugins.d.ts`, `write.d.ts`, `full.d.ts`).
- **Upstream Synchronization**:
  - Built-in CLI tool (`scripts/sync-upstream.js`) and tracking guide (`UPSTREAM.md`) to regularly inspect and sync updates from upstream repositories.

---

## Installation

```bash
npm install spreadsheet-utils
```

---

## Package Targets & Entry Points

| Subpath / Target | Description | Bundle Size (Browser Minified) |
| :--- | :--- | :--- |
| `spreadsheet-utils` | **Core Reader**: Streaming XLSX reader, schema validator, plugin registry | **44 KB** (~14 KB gzip) |
| `spreadsheet-utils/plugins` | **Format Plugins**: `textPlugin`, `xlsPlugin`, `markupPlugin`, `legacyPlugin`, `allPlugins` | **92 KB** |
| `spreadsheet-utils/write` | **Styled Writer & Multi-Exporter**: `writeSpreadsheet`, `writeXlsxFile`, `writeCSV`, `writeTSV`, `writeHTML`, `writeJSON` | **73 KB** |
| `spreadsheet-utils/full` | **All-in-One**: Core reader + all format plugins pre-registered + full writer/exporter | **192 KB** |
| `spreadsheet-utils/worker` | **Worker Offloader**: `readInWorker` (Node `worker_threads` & Browser Web Workers) | Zero-setup background thread |

---

## Quick Start

### 1. Core XLSX Reading (read-excel-file style)

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

// Reads XLSX directly into a 2D array of rows
const rows = await readSpreadsheet('data.xlsx');
console.log(rows);
// [ ['Name', 'Age'], ['Alice', 30], ['Bob', 25] ]
```

### 2. Using Grouped Plugins for Other Formats

If you attempt to read a non-XLSX format without its plugin registered, `spreadsheet-utils` provides a clear diagnostic message indicating the required plugin:

```javascript
import { readSpreadsheet, registerPlugin } from 'spreadsheet-utils';
import { textPlugin } from 'spreadsheet-utils/plugins';

// Register CSV, TSV, and TXT support
registerPlugin(textPlugin);

const rows = await readSpreadsheet('report.csv');
console.log(rows);
```

#### Available Grouped Plugins

```javascript
import {
  textPlugin,    // .csv, .tsv, .tab, .txt
  xlsPlugin,     // .xls (Excel 97-2004 BIFF5/BIFF8 + CFB + PTG decoder)
  markupPlugin,  // .ods (OpenDocument), .html/.htm (HTML tables), .xml (SpreadsheetML 2003)
  legacyPlugin,  // .dbf (dBase), .dif, .slk, Lotus 1-2-3 (.wks/.wk1/.wk3/.wk4/.123), .numbers
  allPlugins     // Array of all 4 plugins
} from 'spreadsheet-utils/plugins';
```

### 3. Isolated Scoped Reader (`createReader`)

To avoid mutating the global registry when different components require different formats:

```javascript
import { createReader } from 'spreadsheet-utils';
import { textPlugin, xlsPlugin } from 'spreadsheet-utils/plugins';

const reader = createReader([textPlugin, xlsPlugin]);
const rows = await reader('financials.xls');
```

### 4. Zero-Config All-in-One Preset (`spreadsheet-utils/full`)

If bundle size is not an issue and you want all 16+ formats and the writer ready out of the box:

```javascript
import { readSpreadsheet, writeSpreadsheet } from 'spreadsheet-utils/full';

// Reads any supported format automatically
const rows = await readSpreadsheet('archive.ods');

// Writes styled XLSX
await writeSpreadsheet(rows, 'output.xlsx');
```

---

## Streaming Row Consumer (`iterateSpreadsheet` & `onRow`)

For massive spreadsheets (hundreds of thousands of rows), holding the entire 2D matrix in RAM is unnecessary. `spreadsheet-utils` provides two ultra-low memory streaming APIs:

### 1. `iterateSpreadsheet` Async Generator
Yields rows one-by-one with constant heap memory usage:

```javascript
import { iterateSpreadsheet } from 'spreadsheet-utils';

// Heap memory remains constant whether reading 100 rows or 1,000,000 rows
for await (const row of iterateSpreadsheet('massive-dataset.xlsx')) {
  console.log('Streamed row:', row);
}
```

### 2. `onRow` Callback
Process each row as it is parsed without buffering:

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

let totalRevenue = 0;

await readSpreadsheet('sales.xlsx', {
  onRow: (row, rowIndex) => {
    if (rowIndex > 0) { // Skip header
      totalRevenue += Number(row[2] || 0);
    }
  }
});

console.log('Total Revenue:', totalRevenue);
```

---

## Security & Input Limits Hardening

Defend against **ZIP bombs**, deeply nested XML attacks, or accidental out-of-memory crashes from hostile user uploads:

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

const rows = await readSpreadsheet(userUploadBuffer, {
  // Enforce max physical file size (e.g., 20 MB)
  maxFileSize: 20 * 1024 * 1024,

  // Guard against ZIP decompression bombs (e.g., max 50 MB uncompressed)
  maxUncompressedSize: 50 * 1024 * 1024,

  // Limit parsed dimensions
  maxRows: 10000,
  maxColumns: 100,

  // Throw an error if bounds are exceeded (defaults to true)
  throwOnLimit: true
});
```

---

## Schema Mapping & Validation

Extract typed JavaScript objects with schema validation across **all** formats:

```javascript
import { readSpreadsheet, registerPlugin, Integer, Email } from 'spreadsheet-utils';
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
  }
};

const { rows, errors } = await readSpreadsheet('users.csv', { schema });

if (errors.length > 0) {
  console.warn('Validation issues:', errors);
}

console.log(rows);
// [ { id: 1, name: 'Alice', email: 'alice@example.com', joinedAt: 2024-01-15T... } ]
```

---

## Styled XLSX Writer (`spreadsheet-utils/write`)

Export spreadsheets with cell styling, custom column widths, and schema mapping:

```javascript
import { writeSpreadsheet } from 'spreadsheet-utils/write';

// 1. Grid of styled cells
const data = [
  [
    { value: 'Product', fontWeight: 'bold', backgroundColor: '#e2e8f0' },
    { value: 'Price', fontWeight: 'bold', backgroundColor: '#e2e8f0' }
  ],
  [
    { value: 'Widget', type: String },
    { value: 29.99, type: Number, format: '$#,##0.00' }
  ]
];

// Write to file (Node.js)
await writeSpreadsheet(data, 'pricing.xlsx');

// Or return Buffer in Node / Blob in Browser
const buffer = await writeSpreadsheet(data, { buffer: true });
```

### Writing Objects with Schema

```javascript
const users = [
  { id: 1, name: 'Alice', balance: 1450.75 },
  { id: 2, name: 'Bob', balance: 920.00 }
];

const schema = {
  id: { column: 'ID', type: Number, width: 10 },
  name: { column: 'Name', type: String, width: 25 },
  balance: { column: 'Balance', type: Number, format: '$#,##0.00', width: 15 }
};

await writeSpreadsheet(users, { schema, filePath: 'users.xlsx' });
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

In addition to styled `.xlsx`, export spreadsheet data directly to standard web and text formats:

```javascript
import { writeCSV, writeTSV, writeHTML, writeJSON, writeSpreadsheet } from 'spreadsheet-utils/write';

const data = [
  ['Product', 'Category', 'Price'],
  ['Widget A', 'Hardware', 29.99],
  ['Widget B', 'Hardware', 49.99]
];

// RFC 4180 compliant CSV
await writeCSV(data, 'catalog.csv');

// Tab-separated TSV
await writeTSV(data, 'catalog.tsv');

// HTML <table> or full HTML document with inline styles
await writeHTML(data, { filePath: 'catalog.html', fullDocument: true });

// JSON array of objects or 2D grid
await writeJSON(data, { filePath: 'catalog.json', asObjects: true });

// Or let writeSpreadsheet auto-route by file extension:
await writeSpreadsheet(data, 'export.csv');
await writeSpreadsheet(data, 'export.html');
```

---

## Background Worker Offloader (`spreadsheet-utils/worker`)

Offload CPU-intensive spreadsheet parsing to background threads to prevent UI freezes in browsers or event-loop stalls in Node.js servers:

```javascript
import { readInWorker } from 'spreadsheet-utils/worker';

// In Node.js: Runs inside node:worker_threads
// In Browser: Runs inside a Web Worker
const rows = await readInWorker('large-report.xlsx');
console.log('Parsed off-thread:', rows.length, 'rows');
```

---

## Interactive Browser Playground Demo

An interactive zero-setup demo is included in `examples/index.html`. You can open it directly in any browser:

- **Drag-and-Drop Parsing**: Drop any `.xlsx`, `.xls`, `.ods`, `.csv`, `.numbers`, `.dbf`, or `.html` file.
- **Dynamic Table Preview**: Inspect data with formatting, headers, row counters, and file metrics.
- **Multi-Format Export**: Convert uploaded data into `.xlsx`, `.csv`, `.html`, or `.json` on the fly.
- **Sample Data Generator**: Test the toolkit with a single click.

To launch locally:
```bash
npx serve .
# Open http://localhost:3000/examples/
```

---

## Browser `<script>` Tag Usage

Ready-to-use standalone bundles are located in `dist/browser/`:

### Core Reader

```html
<script src="dist/browser/spreadsheet-utils.min.js"></script>
<script>
  // Exposes window.spreadsheetUtils
  const input = document.querySelector('input[type="file"]');
  input.addEventListener('change', async (e) => {
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
  async function handleFile(file) {
    // Reads any format (XLSX, CSV, ODS, Numbers, etc.)
    const rows = await window.spreadsheetUtilsFull.readSpreadsheet(file);

    // Write & trigger browser download
    await window.spreadsheetUtilsFull.writeSpreadsheet(rows, 'exported.xlsx');
  }
</script>
```

---

## CommonJS (`require`) Support

```javascript
const { readSpreadsheet, registerPlugin } = require('spreadsheet-utils');
const { textPlugin } = require('spreadsheet-utils/plugins');
const { writeSpreadsheet } = require('spreadsheet-utils/write');

registerPlugin(textPlugin);

async function run() {
  const rows = await readSpreadsheet('data.csv');
  await writeSpreadsheet(rows, 'output.xlsx');
}
run();
```

---

## Upstream Synchronization

`spreadsheet-utils` includes automated tooling and CI workflows to track and merge updates from its base repositories (`read-excel-file`, `tabularjs`, and `write-excel-file`):

- **Automated GitHub Actions (`.github/workflows/upstream-sync.yml`)**:
  - Runs automatically every Monday at 00:00 UTC (and via manual `workflow_dispatch`).
  - Checks npm registry for newer upstream versions.
  - Automatically creates a GitHub issue with details if an upstream library publishes a new release.

- **Manual CLI Tooling**:
```bash
# Check status and file mappings
npm run sync:check

# Clone an upstream repository to .upstream-staging/ for diff inspection
npm run sync:pull read-excel-file
npm run sync:pull tabularjs
npm run sync:pull write-excel-file

# Verify build and test suite
npm run sync:verify
```

See [`UPSTREAM.md`](./UPSTREAM.md) for full maintenance guidelines.

---

## License

MIT
