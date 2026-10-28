import writeXlsxFileNode from './export/writeXlsxFileNode.js';
import getSheetData from './getSheetData/getSheetData.js';
import { normalizeSchemaToColumns, normalizeSchemaToDbfFields } from './schema.js';
import { writeCSV, writeTSV, writeHTML, createHTML, writeJSON } from './formats.js';
import { writeDBF, createDBF } from './dbf.js';
import { writeODS, createODS } from './ods.js';
import { writeXLS, createXLS } from './xls.js';
import { getWriterForExtension, KNOWN_PLUGIN_MAPPINGS, createMissingPluginError } from '../core/registry.js';

export {
    normalizeSchemaToColumns,
    normalizeSchemaToDbfFields,
    writeCSV,
    writeTSV,
    writeHTML,
    createHTML,
    writeJSON,
    writeDBF,
    createDBF,
    writeODS,
    createODS,
    writeXLS,
    createXLS
};

/**
 * Universal styled spreadsheet writer.
 *
 * Supports multi-format export:
 * - XLSX (.xlsx) with styles, colors, fonts, borders, alignments (default)
 * - CSV (.csv) via writeCSV
 * - TSV (.tsv) via writeTSV
 * - HTML (.html) table markup via writeHTML
 * - JSON (.json) via writeJSON
 *
 * Accepts:
 * - 2D grid: [ [ { value: 'Cell', fontWeight: 'bold' } ], [ { value: 123 } ] ] or [ [ 'Cell' ], [ 123 ] ]
 * - Object array + schema/columns: [ { id: 1, name: 'Alice' } ], { schema }
 * - TabularJS worksheets: { worksheets: [ { worksheetName: 'Sheet1', data: [...] } ] }
 * - Multi-sheet array: [ { data: [...], sheet: 'Sheet1' }, { data: [...], sheet: 'Sheet2' } ]
 */
export async function writeSpreadsheet(data, options = {}, extraOptions = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const format = options.format
        ? String(options.format).toLowerCase().replace(/^\./, '')
        : (options.filePath ? options.filePath.split('.').pop().toLowerCase() : 'xlsx');

    // Route to non-XLSX formats if requested
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
            throw createMissingPluginError('dbf', options.filePath);
        }
        return writeDBF(data, options);
    }

    if (format === 'xls') {
        const writer = getWriterForExtension('xls');
        if (writer) {
            return writer('xls', data, options, extraOptions);
        }
        if (KNOWN_PLUGIN_MAPPINGS['xls']) {
            throw createMissingPluginError('xls', options.filePath);
        }
        return writeXLS(data, options);
    }

    if (format === 'ods') {
        const writer = getWriterForExtension('ods');
        if (writer) {
            return writer('ods', data, options, extraOptions);
        }
        if (KNOWN_PLUGIN_MAPPINGS['ods']) {
            throw createMissingPluginError('ods', options.filePath);
        }
        return writeODS(data, options);
    }

    // Route to other registered format writers if available
    const customWriter = getWriterForExtension(format);
    if (customWriter) {
        return customWriter(format, data, options, extraOptions);
    }

    // Default: XLSX format
    let inputData = data;
    let sheetOptions = { ...options };

    // Case 1: TabularJS format ({ worksheets: [...] })
    if (inputData && Array.isArray(inputData.worksheets)) {
        inputData = inputData.worksheets.map((ws, i) => ({
            data: ws.data || [],
            sheet: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
            columns: ws.columns
        }));
    }

    // Case 2: Object array with schema
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

    const writer = writeXlsxFileNode(inputData, sheetOptions, extraOptions);

    if (sheetOptions.filePath) {
        if (typeof writer.toFile === 'function') {
            await writer.toFile(sheetOptions.filePath);
            return sheetOptions.filePath;
        }
        throw new Error('filePath option is only supported in Node.js environments');
    }

    if (sheetOptions.buffer) {
        return await writer.toBuffer();
    }

    if (sheetOptions.blob) {
        return await writer.toBlob();
    }

    return await writer.toBuffer();
}

export const writeXlsxFile = writeXlsxFileNode;
export { getSheetData };
export default writeSpreadsheet;
