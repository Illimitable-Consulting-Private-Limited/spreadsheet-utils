# spreadsheet-utils

Ultra-lightweight, low-memory universal spreadsheet parser for JavaScript (Node.js, Browser, and Web Workers).

Merges the streaming, low-memory architecture of **[read-excel-file](https://www.npmjs.com/package/read-excel-file)** with the extensive multi-format support of **[tabularjs](https://github.com/jspreadsheet/tabularjs)**.

---

## Highlights

- **16+ File Formats Supported**:
  - **Excel**: `.xlsx`, `.xls` (BIFF5, BIFF8 legacy binary)
  - **OpenDocument**: `.ods`
  - **Delimited Text**: `.csv`, `.tsv`, `.tab`, `.txt`
  - **Databases**: `.dbf` (dBase III, IV, V, Visual FoxPro)
  - **Legacy Spreadsheets**: Lotus 1-2-3 (`.wks`, `.wk1`, `.wk3`, `.wk4`, `.123`), `.dif` (Data Interchange Format), `.slk` / `.sylk` (Symbolic Link)
  - **Markup & Web**: `.html`, `.htm` (HTML tables), `.xml` (Excel 2003 XML / SpreadsheetML)
  - **Apple**: `.numbers` (Apple Numbers documents)
- **Low Memory & Small Footprint**:
  - Preserves `read-excel-file`'s **streaming SAX XML parsing** (`saxen`) for `.xlsx`—never loads the entire XML DOM tree into memory.
  - **Zero JSZip dependency**: replaced JSZip (~100 KB minified, high memory footprint) with `fflate` (~8 KB gzipped) for all archive formats (`.xlsx`, `.ods`, `.numbers`).
  - Selective ZIP extraction: only decompresses relevant XML components (`workbook.xml`, `sharedStrings.xml`, target `sheet.xml`), preventing out-of-memory errors on large spreadsheets.
- **Dual API Compatibility**:
  - **`read-excel-file` API**: returns 2D row arrays (`any[][]`), supports `{ schema }` object mapping & validation, `{ sheet }` selection, and `{ allSheets: true }`.
  - **`tabularjs` / Jspreadsheet API**: returns `{ worksheets: [ { data, columns, worksheetName, ... } ] }`.
  - Extends `read-excel-file` features (like `schema` parsing) to **all** formats, including `.csv`, `.xls`, `.ods`, `.html`, etc.!
- **Universal Runtime**:
  - Works out of the box in **Node.js** (file paths, Buffers, Streams), **modern browsers** (Files, Blobs, ArrayBuffers), and **Web Workers**.
  - 100% pure JavaScript with zero native compilation dependencies.

---

## Installation

```bash
npm install spreadsheet-utils
```

### Dependencies
- `fflate`: Ultra-fast, tiny (~8 KB) compression utility.
- `saxen`: Tiny (~4 KB) streaming XML SAX parser.
- `unzipper-esm`: Lightweight streaming unzipper for Node.js pipelines.

---

## Quick Start

### 1. Simple Row Parsing (read-excel-file style)

Reads any spreadsheet format directly into a 2D array of rows:

```javascript
import readSpreadsheet from 'spreadsheet-utils';

// Node.js: pass a file path
const rows = await readSpreadsheet('./data.xlsx');
// rows: [
//   ['Name', 'Age', 'Active'],
//   ['Alice', 30, true],
//   ['Bob', 25, false]
// ]

// Works identically with CSV, XLS, ODS, HTML, etc.!
const csvRows = await readSpreadsheet('./data.csv');
const odsRows = await readSpreadsheet('./data.ods');
const xlsRows = await readSpreadsheet('./legacy.xls');
```

In the browser:
```javascript
const fileInput = document.getElementById('uploader');
fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    const rows = await readSpreadsheet(file);
    console.log(rows);
});
```

---

### 2. Schema Mapping & Type Validation

Map tabular rows directly into typed JavaScript objects with validation:

```javascript
import { readSpreadsheet, Integer } from 'spreadsheet-utils';

const schema = {
    id: {
        column: 'ID',
        type: Integer,
        required: true
    },
    name: {
        column: 'Name',
        type: String,
        required: true
    },
    birthDate: {
        column: 'Date of Birth',
        type: Date
    },
    income: {
        column: 'Income',
        type: Number
    }
};

// Works on .xlsx, .csv, .xls, .ods, or any supported file!
const { rows, errors } = await readSpreadsheet('./employees.xlsx', { schema });

if (errors.length > 0) {
    console.error('Validation errors:', errors);
} else {
    console.log('Parsed objects:', rows);
    // [
    //   { id: 1, name: 'Alice', birthDate: 1994-03-12T..., income: 95000 },
    //   { id: 2, name: 'Bob', birthDate: 1999-07-25T..., income: 72000 }
    // ]
}
```

---

### 3. Multiple Sheets

```javascript
import { readSpreadsheet } from 'spreadsheet-utils';

// Read all sheets
const sheets = await readSpreadsheet('./workbook.xlsx', { allSheets: true });
// [
//   { name: 'Summary', data: [['Q1', 100], ['Q2', 200]] },
//   { name: 'Details', data: [['Item', 'Price'], ['Widget', 25]] }
// ]

// Read specific sheet by 1-based index or sheet name
const sheet2 = await readSpreadsheet('./workbook.xlsx', { sheet: 2 });
const summary = await readSpreadsheet('./workbook.xlsx', { sheet: 'Summary' });
```

---

### 4. TabularJS / Jspreadsheet Compatibility

Need output structured for spreadsheet grids like Jspreadsheet? Use `tabularjs` or `format: 'tabular'`:

```javascript
import { tabularjs } from 'spreadsheet-utils';

const result = await tabularjs('./data.xlsx');
// {
//   worksheets: [
//     {
//       worksheetName: 'Sheet1',
//       data: [['A1', 'B1'], ['A2', 'B2']],
//       columns: [{ title: 'A', width: 100 }, { title: 'B', width: 100 }],
//       mergeCells: { ... },
//       style: { ... }
//     }
//   ]
// }

// Render directly in Jspreadsheet
jspreadsheet(document.getElementById('spreadsheet'), result);
```

Or via `readSpreadsheet`:
```javascript
const result = await readSpreadsheet(file, { format: 'tabular' });
```

---

## Supported Formats

| Format | Extension | Engine / Parser | Memory Profile |
| :--- | :--- | :--- | :--- |
| **Excel OpenXML** | `.xlsx` | `read-excel-file` streaming SAX + `fflate` | Streamed, chunked SAX, minimal RAM |
| **Excel 97-2004** | `.xls` | Pure JS OLE2/CFB + BIFF parser + PTG formula decoder | Direct typed-array reading |
| **OpenDocument** | `.ods` | `fflate` selective unpack + lightweight XML parser | Only decompresses `content.xml`/`styles.xml` |
| **Comma-Separated** | `.csv` | Fast line & quote tokenizer | Linear memory |
| **Tab-Separated** | `.tsv`, `.tab`, `.txt` | Fast tab-delimited tokenizer | Linear memory |
| **dBase** | `.dbf` | Binary table reader | Direct buffer reading |
| **Data Interchange** | `.dif` | Text DIF vector reader | Minimal RAM |
| **Lotus 1-2-3** | `.wks`, `.wk1`, `.wk3`, `.wk4`, `.123` | Lotus BOF record binary parser | Direct buffer reading |
| **Symbolic Link** | `.slk`, `.sylk` | Pure JS SYLK parser | Linear memory |
| **HTML Tables** | `.html`, `.htm` | AST scanner for `<table>` elements | Minimal RAM |
| **XML Spreadsheet** | `.xml` | SpreadsheetML parser | Minimal RAM |
| **Apple Numbers** | `.numbers` | `fflate` unpack + IWA protobuf extraction | Selective unpack |

---

## Automatic Format Detection

When file extension is missing or uninformative, `spreadsheet-utils` sniffs the file content using magic byte signatures:

```javascript
import { detectFormat } from 'spreadsheet-utils';

const buffer = fs.readFileSync('./unknown_file');
const format = await detectFormat(buffer);
// 'xlsx', 'xls', 'ods', 'csv', 'tsv', 'dbf', 'wk1', 'slk', 'dif', 'html', 'xml', or 'numbers'
```

---

## API Reference

### `readSpreadsheet(input, [options])`
The primary universal parsing function.

- `input`: File path string (Node.js), `Buffer`, `Uint8Array`, `ArrayBuffer`, `File` (Browser), `Blob`, or `Stream`.
- `options`:
  - `extension` (`string`): Force a format extension (e.g. `'xlsx'`, `'csv'`, `'ods'`).
  - `format` (`'rows' | 'tabular'`): Set `'tabular'` to return TabularJS worksheets structure. Default is `'rows'`.
  - `sheet` (`string | number`): Select sheet by name or 1-based index.
  - `allSheets` (`boolean`): When `true`, returns an array of `{ name, data }`.
  - `schema` (`object`): Schema definition for transforming rows into typed objects.
  - `delimiter` (`string`): Delimiter for CSV/delimited files (default: auto-detected or `','`).
  - `firstRowAsHeader` (`boolean`): For HTML tables / CSV (default: `true`).
  - `tableIndex` (`number`): Index of HTML table to parse if document contains multiple tables.
  - `trim` (`boolean`): Whether to trim string cell values (default: `true`).
  - `dateFormat` (`string`): Custom date format pattern.
  - `parseNumber` (`function`): Custom number parsing function (e.g. for BigInt or lossless decimal precision).

### Aliases & Subpath Imports
- `readXlsxFile`: Alias to `readSpreadsheet` for 100% drop-in compatibility with `read-excel-file`.
- `readSheet`: Drop-in compatibility with `read-excel-file`'s `readSheet(input, [sheet], [options])`.
- `tabularjs`: Alias to `parseTabular` for 100% drop-in compatibility with `tabularjs`.
- `spreadsheet-utils/node`: Node-optimized entry point.
- `spreadsheet-utils/browser`: Browser-optimized entry point.
- `spreadsheet-utils/web-worker`: Web Worker-optimized entry point.

---

## Architecture & Memory Comparison

| Feature | `tabularjs` | `read-excel-file` | `spreadsheet-utils` (this library) |
| :--- | :--- | :--- | :--- |
| **XLSX XML Parsing** | DOM tree (high RAM) | Streaming SAX (`saxen`) | **Streaming SAX (`saxen`)** |
| **Zip Decompressor** | `jszip` (~100 KB minified) | `fflate` (~8 KB gzipped) | **`fflate` (0 KB `jszip` overhead)** |
| **Total Dependencies** | `jszip`, `@lemonadejs/html-to-json` | `fflate`, `saxen`, `unzipper-esm` | **`fflate`, `saxen`, `unzipper-esm`** |
| **XLSX Memory Usage** | High (loads entire DOM & files) | Minimal (streamed SAX) | **Minimal (streamed SAX)** |
| **Formats Supported** | 16+ formats | Only `.xlsx` | **All 16+ formats** |
| **Schema Object Parsing**| No | Yes | **Yes (across all 16+ formats!)** |

---

## License

MIT
