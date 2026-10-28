import writeXlsxFileBrowser from './export/writeXlsxFileBrowser.js';
import getSheetData from './getSheetData/getSheetData.js';
import { normalizeSchemaToColumns } from './schema.js';
import { normalizeTo2DArray, escapeCSVCell } from './formats.js';

function downloadFile(content, fileName, mimeType) {
    if (typeof document !== 'undefined' && typeof Blob !== 'undefined') {
        const blob = new Blob([content], { type: mimeType });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        link.click();
        setTimeout(() => URL.revokeObjectURL(link.href), 100);
        return blob;
    }
}

export async function writeCSV(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }
    const delimiter = options.delimiter || ',';
    const grid = normalizeTo2DArray(data);
    const lines = grid.map(row => row.map(cell => escapeCSVCell(cell, delimiter)).join(delimiter));
    const csvContent = lines.join('\n');

    if (options.fileName) {
        downloadFile(csvContent, options.fileName, 'text/csv;charset=utf-8;');
        return;
    }

    if (options.blob) {
        return new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    }

    return csvContent;
}

export async function writeTSV(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }
    return writeCSV(data, { ...options, delimiter: '\t' });
}

export async function writeHTML(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
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

    if (options.fileName) {
        downloadFile(html, options.fileName, 'text/html;charset=utf-8;');
        return;
    }

    if (options.blob) {
        return new Blob([html], { type: 'text/html;charset=utf-8;' });
    }

    return html;
}

export async function writeJSON(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }

    let result = data;
    if (Array.isArray(data) && Array.isArray(data[0])) {
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

    if (options.fileName) {
        downloadFile(jsonString, options.fileName, 'application/json;charset=utf-8;');
        return;
    }

    if (options.blob) {
        return new Blob([jsonString], { type: 'application/json;charset=utf-8;' });
    }

    return jsonString;
}

export async function writeSpreadsheet(data, options = {}, extraOptions = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }

    const format = options.format
        ? String(options.format).toLowerCase().replace(/^\./, '')
        : (options.fileName ? options.fileName.split('.').pop().toLowerCase() : 'xlsx');

    if (format === 'csv') {
        return writeCSV(data, options);
    }
    if (format === 'tsv') {
        return writeTSV(data, options);
    }
    if (format === 'html' || format === 'htm') {
        return writeHTML(data, options);
    }
    if (format === 'json') {
        return writeJSON(data, options);
    }

    let inputData = data;
    let sheetOptions = { ...options };

    if (inputData && Array.isArray(inputData.worksheets)) {
        inputData = inputData.worksheets.map((ws, i) => ({
            data: ws.data || [],
            sheet: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
            columns: ws.columns
        }));
    }

    if (Array.isArray(inputData) && inputData.length > 0 && !Array.isArray(inputData[0])) {
        if (sheetOptions.schema) {
            const columns = normalizeSchemaToColumns(sheetOptions.schema);
            inputData = getSheetData(inputData, columns);
            sheetOptions.columns = columns;
            delete sheetOptions.schema;
        } else if (sheetOptions.columns && typeof sheetOptions.columns[0]?.cell !== 'function') {
            const columns = normalizeSchemaToColumns(sheetOptions.columns);
            inputData = getSheetData(inputData, columns);
            sheetOptions.columns = columns;
        }
    }

    const writer = writeXlsxFileBrowser(inputData, sheetOptions, extraOptions);

    if (sheetOptions.fileName) {
        await writer.toFile(sheetOptions.fileName);
        return;
    }

    return await writer.toBlob();
}

export const writeXlsxFile = writeXlsxFileBrowser;
export {
    writeCSV as writeCSVBrowser,
    getSheetData,
    normalizeSchemaToColumns
};
export default writeSpreadsheet;
