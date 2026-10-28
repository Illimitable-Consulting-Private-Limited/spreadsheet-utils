import fs from 'node:fs';
import { normalizeSchemaToColumns, calculateAutoWidths } from './schema.js';
import { applyTheme } from './themes.js';
import getSheetData from './getSheetData/getSheetData.js';

/**
 * Normalize any input format (2D grid, TabularJS worksheets, or object array) to a 2D array of rows.
 * @param {any} input
 * @returns {any[][]}
 */
export function normalizeTo2DArray(input) {
    if (!input) return [];

    // TabularJS format
    if (input.worksheets && Array.isArray(input.worksheets) && input.worksheets[0]) {
        return input.worksheets[0].data || [];
    }

    if (!Array.isArray(input)) return [];
    if (input.length === 0) return [];

    // Already 2D array
    if (Array.isArray(input[0])) {
        return input.map(row => row.map(cell => (cell && typeof cell === 'object' && 'value' in cell) ? cell.value : cell));
    }

    // Array of objects
    if (typeof input[0] === 'object' && input[0] !== null) {
        const keys = Object.keys(input[0]);
        const headerRow = [...keys];
        const rows = input.map(obj => keys.map(k => obj[k]));
        return [headerRow, ...rows];
    }

    return [];
}

/**
 * Format a single CSV/TSV cell value according to RFC 4180.
 * @param {any} value
 * @param {string} delimiter
 * @returns {string}
 */
export function escapeCSVCell(value, delimiter = ',') {
    if (value === null || value === undefined) {
        return '';
    }
    if (value instanceof Date) {
        return value.toISOString();
    }
    const str = String(value);
    if (str.includes(delimiter) || str.includes('"') || str.includes('\n') || str.includes('\r')) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
}

/**
 * Serialize spreadsheet data to CSV format.
 *
 * @param {any} data - 2D array or object array
 * @param {Object} [options={}]
 * @param {string} [options.delimiter=','] - Delimiter character
 * @param {string} [options.filePath] - Output file path (Node.js)
 * @param {boolean} [options.buffer=false] - Return Buffer instead of string
 * @returns {Promise<string|Buffer>}
 */
export async function writeCSV(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }
    const delimiter = options.delimiter || ',';
    const grid = normalizeTo2DArray(data);
    const lines = grid.map(row => row.map(cell => escapeCSVCell(cell, delimiter)).join(delimiter));
    const csvContent = lines.join('\n');

    if (options.filePath) {
        await fs.promises.writeFile(options.filePath, csvContent, 'utf-8');
        return options.filePath;
    }

    if (options.buffer) {
        return Buffer.from(csvContent, 'utf-8');
    }

    return csvContent;
}

/**
 * Serialize spreadsheet data to TSV format.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Promise<string|Buffer>}
 */
export async function writeTSV(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }
    return writeCSV(data, { ...options, delimiter: '\t' });
}

function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function formatCellStyles(cell) {
    if (!cell || typeof cell !== 'object' || cell instanceof Date) return '';
    const styles = [];
    if (cell.fontWeight) styles.push(`font-weight:${cell.fontWeight}`);
    if (cell.fontStyle) styles.push(`font-style:${cell.fontStyle}`);
    if (cell.fontSize) {
        const sz = typeof cell.fontSize === 'number' ? `${cell.fontSize}pt` : cell.fontSize;
        styles.push(`font-size:${sz}`);
    }
    if (cell.fontFamily) styles.push(`font-family:${cell.fontFamily}`);
    if (cell.backgroundColor) styles.push(`background-color:${cell.backgroundColor}`);
    if (cell.textColor || cell.color) styles.push(`color:${cell.textColor || cell.color}`);
    if (cell.align) styles.push(`text-align:${cell.align}`);
    if (cell.verticalAlign) {
        const va = cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign;
        styles.push(`vertical-align:${va}`);
    }
    if (cell.wrap) styles.push('white-space:pre-wrap;word-break:break-word');
    if (cell.border) styles.push(`border:${cell.border}`);
    if (cell.borderStyle || cell.borderColor) {
        styles.push(`border:1px ${cell.borderStyle || 'solid'} ${cell.borderColor || '#ccc'}`);
    }
    if (cell.leftBorderStyle || cell.leftBorderColor) {
        styles.push(`border-left:1px ${cell.leftBorderStyle || 'solid'} ${cell.leftBorderColor || '#ccc'}`);
    }
    if (cell.rightBorderStyle || cell.rightBorderColor) {
        styles.push(`border-right:1px ${cell.rightBorderStyle || 'solid'} ${cell.rightBorderColor || '#ccc'}`);
    }
    if (cell.topBorderStyle || cell.topBorderColor) {
        styles.push(`border-top:1px ${cell.topBorderStyle || 'solid'} ${cell.topBorderColor || '#ccc'}`);
    }
    if (cell.bottomBorderStyle || cell.bottomBorderColor) {
        styles.push(`border-bottom:1px ${cell.bottomBorderStyle || 'solid'} ${cell.bottomBorderColor || '#ccc'}`);
    }
    return styles.length > 0 ? ` style="${styles.join(';')}"` : '';
}

/**
 * Synchronously render spreadsheet data to an HTML table string.
 *
 * @param {any} data - 2D grid, object array with schema/columns, or TabularJS worksheets
 * @param {Object} [options={}]
 * @param {boolean} [options.fullDocument=false] - Wrap in a complete HTML document
 * @param {string} [options.tableClass='spreadsheet-table'] - CSS class for <table>
 * @param {string} [options.title='Spreadsheet Export'] - Title when fullDocument is true
 * @param {boolean} [options.hasHeader=true] - Whether to generate a <thead> header row
 * @param {Object} [options.headerStyle] - Baseline style applied to all header cells
 * @returns {string} HTML markup string
 */
export function createHTML(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    let rows = [];
    let columns = options.columns;

    // 1. TabularJS worksheets format
    if (data && Array.isArray(data.worksheets) && data.worksheets[0]) {
        const ws = data.worksheets[0];
        rows = ws.data || [];
        if (ws.columns && Array.isArray(ws.columns) && rows.length > 0 && !Array.isArray(rows[0])) {
            columns = normalizeSchemaToColumns(ws.columns, ws);
            rows = getSheetData(rows, columns);
        }
    }
    // 2. Object array with schema
    else if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object' && data[0] !== null && !Array.isArray(data[0])) {
        const schema = options.schema || options.columns;
        if (schema) {
            columns = normalizeSchemaToColumns(schema, options);
            rows = getSheetData(data, columns);
        } else {
            const keys = Object.keys(data[0]);
            const headerStyle = options.headerStyle || { fontWeight: 'bold' };
            const headerRow = keys.map(k => ({ value: k, ...headerStyle }));
            const dataRows = data.map(obj => keys.map(k => ({ value: obj[k] })));
            rows = [headerRow, ...dataRows];
        }
    }
    // 3. 2D array
    else if (Array.isArray(data)) {
        rows = data;
    }

    if (options.theme) {
        rows = applyTheme(rows, options.theme, options);
    }
    if (options.autoWidth) {
        columns = calculateAutoWidths(rows, options, columns);
    }

    const tableClass = options.tableClass || 'spreadsheet-table';
    let html = `<table class="${tableClass}">\n`;

    // Emit <colgroup> if column widths are provided
    if (columns && Array.isArray(columns) && columns.some(c => c && c.width)) {
        html += '  <colgroup>\n';
        columns.forEach(c => {
            if (c && c.width) {
                html += `    <col style="width:${c.width}ch;">\n`;
            } else {
                html += `    <col>\n`;
            }
        });
        html += '  </colgroup>\n';
    }

    const coveredCells = new Set();
    const hasHeader = options.hasHeader !== false && rows.length > 0;
    const headerRowIndex = hasHeader ? 0 : -1;

    let inThead = false;
    let inTbody = false;

    rows.forEach((row, rowIndex) => {
        if (!Array.isArray(row)) return;

        const isHeader = rowIndex === headerRowIndex;
        if (isHeader) {
            html += '  <thead>\n';
            inThead = true;
        } else if (!inTbody) {
            if (inThead) {
                html += '  </thead>\n';
                inThead = false;
            }
            html += '  <tbody>\n';
            inTbody = true;
        }

        const tag = isHeader ? 'th' : 'td';
        html += '    <tr>\n';

        row.forEach((cell, colIndex) => {
            const cellKey = `${rowIndex},${colIndex}`;
            if (coveredCells.has(cellKey)) {
                return;
            }

            let val = cell;
            let spanAttrs = '';
            let style = '';

            if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                val = cell.value !== undefined ? cell.value : '';
                const colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
                const rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);

                if (colSpan > 1) spanAttrs += ` colspan="${colSpan}"`;
                if (rowSpan > 1) spanAttrs += ` rowspan="${rowSpan}"`;

                if (colSpan > 1 || rowSpan > 1) {
                    for (let dr = 0; dr < rowSpan; dr++) {
                        for (let dc = 0; dc < colSpan; dc++) {
                            if (dr > 0 || dc > 0) {
                                coveredCells.add(`${rowIndex + dr},${colIndex + dc}`);
                            }
                        }
                    }
                }

                style = formatCellStyles(cell);
            } else if (cell instanceof Date) {
                val = !isNaN(cell.getTime()) ? cell.toISOString().slice(0, 10) : '';
            }

            if (val === null || val === undefined) val = '';
            const safeVal = escapeHtml(val);
            html += `      <${tag}${spanAttrs}${style}>${safeVal}</${tag}>\n`;
        });

        html += '    </tr>\n';
    });

    if (inThead) {
        html += '  </thead>\n';
    }
    if (inTbody) {
        html += '  </tbody>\n';
    }

    html += '</table>';

    if (options.fullDocument) {
        const title = escapeHtml(options.title || 'Spreadsheet Export');
        html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>${title}</title>\n` +
            `<style>\ntable.${tableClass} { border-collapse: collapse; width: 100%; font-family: sans-serif; }\n` +
            `table.${tableClass} th, table.${tableClass} td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n` +
            `table.${tableClass} th { background-color: #f4f4f4; font-weight: bold; }\n</style>\n</head>\n<body>\n${html}\n</body>\n</html>`;
    }

    return html;
}

/**
 * Serialize spreadsheet data to HTML table format.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @param {boolean} [options.fullDocument=false] - Wrap in complete HTML document
 * @param {string} [options.tableClass='spreadsheet-table'] - CSS class for <table>
 * @returns {Promise<string|Buffer>}
 */
export async function writeHTML(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const html = createHTML(data, options);

    if (options.filePath) {
        await fs.promises.writeFile(options.filePath, html, 'utf-8');
        return options.filePath;
    }

    if (options.buffer) {
        return Buffer.from(html, 'utf-8');
    }

    return html;
}

/**
 * Serialize spreadsheet data to JSON format.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Promise<string|Buffer>}
 */
export async function writeJSON(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    let result = data;
    if (Array.isArray(data) && Array.isArray(data[0])) {
        // 2D grid: if options.asObjects, convert using row 0 as keys
        if (options.asObjects && data.length > 1) {
            const keys = data[0].map(k => String((k && typeof k === 'object') ? k.value : k));
            result = data.slice(1).map(row => {
                const obj = {};
                keys.forEach((key, i) => {
                    const cell = row[i];
                    obj[key] = (cell && typeof cell === 'object' && 'value' in cell) ? cell.value : cell;
                });
                return obj;
            });
        } else {
            result = data.map(r => r.map(c => (c && typeof c === 'object' && 'value' in c) ? c.value : c));
        }
    }

    const spaces = options.spaces !== undefined ? options.spaces : 2;
    const jsonString = JSON.stringify(result, null, spaces);

    if (options.filePath) {
        await fs.promises.writeFile(options.filePath, jsonString, 'utf-8');
        return options.filePath;
    }

    if (options.buffer) {
        return Buffer.from(jsonString, 'utf-8');
    }

    return jsonString;
}

export { writeDBF, createDBF, normalizeSchemaToDbfFields } from './dbf.js';
export { writeODS, createODS } from './ods.js';
export { writeXLS, createXLS } from './xls.js';
