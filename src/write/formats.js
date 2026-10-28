import fs from 'node:fs';

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

    let rows = [];
    if (Array.isArray(data) && Array.isArray(data[0])) {
        rows = data;
    } else {
        rows = normalizeTo2DArray(data);
    }

    const tableClass = options.tableClass || 'spreadsheet-table';
    let html = `<table class="${tableClass}">\n`;

    rows.forEach((row, rowIndex) => {
        const isHeader = rowIndex === 0 && options.hasHeader !== false;
        const tag = isHeader ? 'th' : 'td';
        html += '  <tr>\n';
        row.forEach(cell => {
            let val = cell;
            let style = '';
            if (cell && typeof cell === 'object') {
                val = cell.value !== undefined ? cell.value : '';
                const styles = [];
                if (cell.fontWeight) styles.push(`font-weight:${cell.fontWeight}`);
                if (cell.backgroundColor) styles.push(`background-color:${cell.backgroundColor}`);
                if (cell.textColor || cell.color) styles.push(`color:${cell.textColor || cell.color}`);
                if (cell.align) styles.push(`text-align:${cell.align}`);
                if (cell.borderStyle || cell.borderColor) {
                    styles.push(`border:1px ${cell.borderStyle || 'solid'} ${cell.borderColor || '#ccc'}`);
                }
                if (styles.length > 0) {
                    style = ` style="${styles.join(';')}"`;
                }
            }
            if (val === null || val === undefined) val = '';
            const safeVal = String(val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            html += `    <${tag}${style}>${safeVal}</${tag}>\n`;
        });
        html += '  </tr>\n';
    });

    html += '</table>';

    if (options.fullDocument) {
        html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<title>Spreadsheet Export</title>\n` +
            `<style>\ntable.${tableClass} { border-collapse: collapse; width: 100%; font-family: sans-serif; }\n` +
            `table.${tableClass} th, table.${tableClass} td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n` +
            `table.${tableClass} th { background-color: #f4f4f4; font-weight: bold; }\n</style>\n</head>\n<body>\n${html}\n</body>\n</html>`;
    }

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
