import writeXlsxFileBrowser from './export/writeXlsxFileBrowser.js';
import getSheetData from './getSheetData/getSheetData.js';
import { normalizeSchemaToColumns, normalizeSchemaToDbfFields } from './schema.js';
import { normalizeTo2DArray, escapeCSVCell, createHTML } from './formats.js';
import { createDBF } from './dbf.js';
import { createODS } from './ods.js';
import { createXLS } from './xls.js';
import { getWriterForExtension, KNOWN_PLUGIN_MAPPINGS, createMissingPluginError } from '../core/registry.js';

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

    const html = createHTML(data, options);

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

async function writeDBF(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }
    const uint8 = createDBF(data, options);

    if (options.fileName) {
        downloadFile(uint8, options.fileName, 'application/x-dbf');
        return;
    }

    if (options.blob) {
        return new Blob([uint8], { type: 'application/x-dbf' });
    }

    return uint8;
}

async function writeODS(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }
    const uint8 = createODS(data, options);

    if (options.fileName) {
        downloadFile(uint8, options.fileName, 'application/vnd.oasis.opendocument.spreadsheet');
        return;
    }

    if (options.blob) {
        return new Blob([uint8], { type: 'application/vnd.oasis.opendocument.spreadsheet' });
    }

    return uint8;
}

async function writeXLS(data, options = {}) {
    if (typeof options === 'string') {
        options = { fileName: options };
    }
    const uint8 = createXLS(data, options);

    if (options.fileName) {
        downloadFile(uint8, options.fileName, 'application/vnd.ms-excel');
        return;
    }

    if (options.blob) {
        return new Blob([uint8], { type: 'application/vnd.ms-excel' });
    }

    return uint8;
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

    if (format === 'dbf') {
        const writer = getWriterForExtension('dbf');
        if (writer) {
            return writer('dbf', data, options, extraOptions);
        }
        if (KNOWN_PLUGIN_MAPPINGS['dbf']) {
            throw createMissingPluginError('dbf', options.fileName);
        }
        return writeDBF(data, options);
    }

    if (format === 'xls') {
        const writer = getWriterForExtension('xls');
        if (writer) {
            return writer('xls', data, options, extraOptions);
        }
        if (KNOWN_PLUGIN_MAPPINGS['xls']) {
            throw createMissingPluginError('xls', options.fileName);
        }
        return writeXLS(data, options);
    }

    if (format === 'ods') {
        const writer = getWriterForExtension('ods');
        if (writer) {
            return writer('ods', data, options, extraOptions);
        }
        if (KNOWN_PLUGIN_MAPPINGS['ods']) {
            throw createMissingPluginError('ods', options.fileName);
        }
        return writeODS(data, options);
    }

    const customWriter = getWriterForExtension(format);
    if (customWriter) {
        return customWriter(format, data, options, extraOptions);
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
            const columns = normalizeSchemaToColumns(sheetOptions.schema, sheetOptions);
            inputData = getSheetData(inputData, columns);
            sheetOptions.columns = columns;
            delete sheetOptions.schema;
        } else if (sheetOptions.columns && typeof sheetOptions.columns[0]?.cell !== 'function') {
            const columns = normalizeSchemaToColumns(sheetOptions.columns, sheetOptions);
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
    createHTML,
    getSheetData,
    normalizeSchemaToColumns,
    normalizeSchemaToDbfFields,
    writeDBF,
    createDBF,
    writeODS,
    createODS,
    writeXLS,
    createXLS
};
export default writeSpreadsheet;
