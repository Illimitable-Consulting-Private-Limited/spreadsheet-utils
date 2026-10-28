#!/usr/bin/env node

/**
 * spreadsheet-utils CLI
 * Universal command-line interface for reading, converting, and inspecting spreadsheets.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const distFull = path.resolve(__dirname, '../dist/esm/full.mjs');
const srcFull = path.resolve(__dirname, '../src/full.js');
const modulePath = fs.existsSync(distFull) ? distFull : srcFull;

const {
    readSpreadsheet,
    convertSpreadsheet,
    parseTabular,
    detectFormat
} = await import(modulePath);

const pkg = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../package.json'), 'utf-8'));

function printHelp() {
    console.log(`
spreadsheet-utils v${pkg.version}
Ultra-lightweight universal spreadsheet toolkit for 16+ formats.

USAGE:
  spreadsheet-utils <command> [arguments] [options]

COMMANDS:
  convert <input> <output>       Convert spreadsheet between any formats (xlsx, ods, xls, csv, tsv, dbf, html, json)
  cat <input>                    Print spreadsheet contents to terminal
  info <input>                   Display spreadsheet format, sheets, and dimensions

OPTIONS:
  --sheet <name|index>           Select specific sheet (1-based index or sheet name)
  --theme <name>                 Apply styling theme: ocean | forest | slate | sunset | striped
  --auto-width                   Automatically compute responsive column widths
  --csv                          Output as CSV (cat command)
  --json                         Output as JSON array (cat command)
  --objects, --as-objects        Output data as objects using first row as headers
  --limit <n>                    Limit number of displayed rows (cat command, default: 50)
  -v, --version                  Display version
  -h, --help                     Display this help message

EXAMPLES:
  # Convert CSV to styled XLSX
  spreadsheet-utils convert data.csv report.xlsx --theme ocean --auto-width

  # Convert legacy XLS to modern ODS
  spreadsheet-utils convert old.xls modern.ods

  # Export XLSX sheet to CSV
  spreadsheet-utils convert workbook.xlsx sheet.csv --sheet "Sales"

  # Inspect spreadsheet structure
  spreadsheet-utils info workbook.xlsx

  # View spreadsheet in terminal
  spreadsheet-utils cat data.xlsx --limit 20
`);
}

function parseArgs(args) {
    const positional = [];
    const options = {};

    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (arg === '--help' || arg === '-h') {
            options.help = true;
        } else if (arg === '--version' || arg === '-v') {
            options.version = true;
        } else if (arg === '--auto-width') {
            options.autoWidth = true;
        } else if (arg === '--csv') {
            options.csv = true;
        } else if (arg === '--json') {
            options.json = true;
        } else if (arg === '--objects' || arg === '--as-objects') {
            options.asObjects = true;
        } else if (arg === '--theme') {
            options.theme = args[++i];
        } else if (arg === '--sheet') {
            const val = args[++i];
            options.sheet = /^\d+$/.test(val) ? parseInt(val, 10) : val;
        } else if (arg === '--limit') {
            options.limit = parseInt(args[++i], 10);
        } else if (arg.startsWith('--')) {
            options[arg.slice(2)] = true;
        } else {
            positional.push(arg);
        }
    }

    return { positional, options };
}

function renderAsciiTable(rows, maxCols = 8) {
    if (!rows || rows.length === 0) return '(empty sheet)';
    const displayRows = rows.map(r => (r || []).slice(0, maxCols).map(c => {
        if (c == null) return '';
        if (typeof c === 'object' && 'value' in c) return String(c.value ?? '');
        return String(c);
    }));

    const colWidths = [];
    for (const r of displayRows) {
        for (let c = 0; c < r.length; c++) {
            colWidths[c] = Math.max(colWidths[c] || 0, Math.min(25, (r[c] || '').length));
        }
    }

    const formatRow = (r) => '| ' + r.map((c, i) => {
        const str = (c || '').slice(0, 25);
        return str.padEnd(colWidths[i] || 0);
    }).join(' | ') + ' |';

    const sep = '+-' + colWidths.map(w => '-'.repeat(w)).join('-+-') + '-+';

    const lines = [];
    if (displayRows.length > 0) {
        lines.push(sep);
        lines.push(formatRow(displayRows[0]));
        lines.push(sep);
        for (let i = 1; i < displayRows.length; i++) {
            lines.push(formatRow(displayRows[i]));
        }
        lines.push(sep);
    }
    return lines.join('\n');
}

async function handleInfo(inputPath, options) {
    if (!fs.existsSync(inputPath)) {
        console.error(`Error: File not found: ${inputPath}`);
        process.exit(1);
    }

    const buf = await fs.promises.readFile(inputPath);
    let format = await detectFormat(buf);
    if (!format && inputPath.includes('.')) {
        format = inputPath.split('.').pop().toLowerCase();
    }
    const tabular = await parseTabular(buf, { ...options, filePath: inputPath });
    const sheets = tabular.worksheets || [];

    console.log(`\nSpreadsheet Info: ${path.basename(inputPath)}`);
    console.log(`  Path:            ${path.resolve(inputPath)}`);
    console.log(`  Detected Format: ${format || 'unknown'}`);
    console.log(`  Worksheets:      ${sheets.length}\n`);

    sheets.forEach((ws, idx) => {
        const name = ws.worksheetName || ws.sheet || `Sheet${idx + 1}`;
        const rowCount = (ws.data || []).length;
        let colCount = 0;
        if (rowCount > 0 && Array.isArray(ws.data[0])) {
            colCount = ws.data[0].length;
        }
        console.log(`  [Sheet ${idx + 1}] "${name}" — ${rowCount} rows, ${colCount} columns`);
    });
    console.log('');
}

async function handleCat(inputPath, options) {
    if (!fs.existsSync(inputPath)) {
        console.error(`Error: File not found: ${inputPath}`);
        process.exit(1);
    }

    const readOpts = { ...options };
    if (options.limit && !options.asObjects) {
        readOpts.maxRows = options.limit;
    }

    const data = await readSpreadsheet(inputPath, readOpts);

    if (options.json || options.asObjects) {
        const limit = options.limit || 50;
        const slice = Array.isArray(data) ? data.slice(0, limit) : data;
        console.log(JSON.stringify(slice, null, 2));
        return;
    }

    if (options.csv) {
        const limit = options.limit || 50;
        const rows = (Array.isArray(data) ? data.slice(0, limit) : []).map(row => {
            return (row || []).map(cell => {
                const val = (cell && typeof cell === 'object' && 'value' in cell) ? cell.value : cell;
                const str = val == null ? '' : String(val);
                return str.includes(',') || str.includes('"') || str.includes('\n')
                    ? `"${str.replace(/"/g, '""')}"`
                    : str;
            }).join(',');
        });
        console.log(rows.join('\n'));
        return;
    }

    // Default: ASCII table
    const limit = options.limit || 50;
    const rows = Array.isArray(data) ? data.slice(0, limit) : [];
    console.log(renderAsciiTable(rows));
    if (Array.isArray(data) && data.length > limit) {
        console.log(`... and ${data.length - limit} more rows (use --limit to see more)`);
    }
}

async function handleConvert(inputPath, outputPath, options) {
    if (!fs.existsSync(inputPath)) {
        console.error(`Error: File not found: ${inputPath}`);
        process.exit(1);
    }

    const start = Date.now();
    await convertSpreadsheet(inputPath, outputPath, options);
    const duration = Date.now() - start;

    const stats = fs.statSync(outputPath);
    const sizeKB = (stats.size / 1024).toFixed(1);
    console.log(`✔ Converted ${inputPath} -> ${outputPath} (${sizeKB} KB, ${duration}ms)`);
}

async function main() {
    const { positional, options } = parseArgs(process.argv.slice(2));

    if (options.version) {
        console.log(`spreadsheet-utils v${pkg.version}`);
        process.exit(0);
    }

    if (options.help || positional.length === 0) {
        printHelp();
        process.exit(0);
    }

    const command = positional[0].toLowerCase();

    try {
        if (command === 'info') {
            const inputPath = positional[1];
            if (!inputPath) {
                console.error('Error: Missing input file path for "info" command.');
                process.exit(1);
            }
            await handleInfo(inputPath, options);
        } else if (command === 'cat') {
            const inputPath = positional[1];
            if (!inputPath) {
                console.error('Error: Missing input file path for "cat" command.');
                process.exit(1);
            }
            await handleCat(inputPath, options);
        } else if (command === 'convert') {
            const inputPath = positional[1];
            const outputPath = positional[2];
            if (!inputPath || !outputPath) {
                console.error('Error: "convert" command requires both <input> and <output> paths.');
                process.exit(1);
            }
            await handleConvert(inputPath, outputPath, options);
        } else {
            console.error(`Error: Unknown command "${command}". Run "spreadsheet-utils --help" for available commands.`);
            process.exit(1);
        }
    } catch (err) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
    }
}

main();
