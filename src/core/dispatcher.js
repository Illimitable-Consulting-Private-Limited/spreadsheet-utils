import { parseXLS } from '../parser/xls.js';
import { parseXLSX } from '../parser/xlsx.js';
import { parseODS } from '../parser/ods.js';
import { parseLotus } from '../parser/lotus.js';
import { parseCSVFile } from '../parser/csv.js';
import { parseXMLSpreadsheet } from '../parser/xml.js';
import { parseDIF } from '../parser/dif.js';
import { parseSLK } from '../parser/slk.js';
import { parseHTMLTable } from '../parser/table.js';
import { parseDBF } from '../parser/dbf.js';
import { parseNumbers } from '../parser/numbers.js';
import { loadAsBuffer } from '../utils/loader.js';
import { detectFormat } from './detector.js';
import parseSheetData from '../schema/parseSheetData.js';

/**
 * Route a resolved format extension to its parser
 *
 * @param {string} ext - Lowercase file extension
 * @param {any} file - Input file
 * @param {Object} options - Options
 * @returns {Promise<Object>|null}
 */
export function dispatch(ext, file, options = {}) {
    switch (ext) {
        case 'xls':
            return parseXLS(file, options);
        case 'xlsx':
            return parseXLSX(file, options);
        case 'ods':
            return parseODS(file, options);
        case 'numbers':
            return parseNumbers(file, options);
        case 'wks':
        case 'wk1':
        case 'wk3':
        case 'wk4':
        case '123':
            return parseLotus(file, options);
        case 'csv':
            return parseCSVFile(file, options);
        case 'tsv':
        case 'tab':
            return parseCSVFile(file, { ...options, delimiter: '\t' });
        case 'txt':
            return parseCSVFile(file, { ...options, delimiter: options.delimiter || '\t' });
        case 'xml':
            return parseXMLSpreadsheet(file, options);
        case 'dif':
            return parseDIF(file, options);
        case 'slk':
        case 'sylk':
            return parseSLK(file, options);
        case 'html':
        case 'htm':
            return parseHTMLTable(file, options);
        case 'dbf':
            return parseDBF(file, options);
        default:
            return null;
    }
}

/**
 * Low-level spreadsheet parser returning TabularJS worksheets structure.
 * Compatible with TabularJS.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file
 * @param {Object} [options={}]
 * @returns {Promise<{ worksheets: Array<any> }>}
 */
export async function parseTabular(file, options = {}) {
    if (!file) {
        throw new Error('Invalid file');
    }

    // Normalize raw byte inputs
    if (typeof ArrayBuffer !== 'undefined' && file instanceof ArrayBuffer) {
        file = new Uint8Array(file);
    }

    const isBinaryInput = file instanceof Uint8Array
        || (typeof Buffer !== 'undefined' && Buffer.isBuffer(file))
        || (typeof Blob !== 'undefined' && file instanceof Blob);

    let fileName = '';
    if (typeof file === 'string') {
        fileName = file;
    } else if (file.name && typeof file.name === 'string') {
        fileName = file.name;
    } else if (!isBinaryInput && !(file && typeof file.pipe === 'function')) {
        throw new Error('Invalid file: must be a file path, File, Blob, Uint8Array, ArrayBuffer, or Buffer');
    }

    // Format resolution: explicit options first, then extension
    let ext = null;
    if (options.extension) {
        ext = String(options.extension).toLowerCase().replace(/^\./, '');
    } else if (options.format && options.format !== 'tabular' && options.format !== 'rows') {
        ext = String(options.format).toLowerCase().replace(/^\./, '');
    } else {
        const dot = fileName.lastIndexOf('.');
        if (dot > -1) {
            ext = fileName.toLowerCase().slice(dot + 1);
        }
    }

    if (ext) {
        const result = dispatch(ext, file, options);
        if (result) {
            return await result;
        }
    }

    // Unknown or missing extension: sniff format from content
    let buffer = null;
    let detected = null;
    try {
        buffer = await loadAsBuffer(file);
        detected = await detectFormat(buffer);
    } catch {
        // Fall through
    }

    if (detected) {
        const result = dispatch(detected, buffer || file, options);
        if (result) {
            return await result;
        }
    }

    throw new Error(`Unsupported file type: ${ext || fileName || '(buffer)'}`);
}

/**
 * Universal spreadsheet reader supporting both read-excel-file options/output
 * and TabularJS format across all 16+ spreadsheet formats.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file
 * @param {Object} [options={}]
 * @returns {Promise<any[][] | { rows: any[], errors: any[] } | Array<{ name: string, data: any[][] }> | { worksheets: any[] }>}
 */
export async function readSpreadsheet(file, options = {}) {
    // 1. Get tabular worksheets structure
    const tabularResult = await parseTabular(file, options);
    const worksheets = tabularResult.worksheets || [];

    // If caller explicitly requested TabularJS format or worksheets
    if (options.format === 'tabular' || options.worksheets === true) {
        return tabularResult;
    }

    // If caller requested allSheets (read-excel-file style)
    if (options.allSheets === true) {
        return worksheets.map((ws, i) => ({
            name: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
            data: ws.data || []
        }));
    }

    // Single sheet selection (options.sheet can be string name or 1-based index)
    let selectedWorksheet = worksheets[0];
    if (options.sheet !== undefined) {
        if (typeof options.sheet === 'number') {
            const index = options.sheet - 1;
            if (index < 0 || index >= worksheets.length) {
                throw new Error(`Sheet not found: ${options.sheet}`);
            }
            selectedWorksheet = worksheets[index];
        } else if (typeof options.sheet === 'string') {
            const found = worksheets.find(
                (ws) => (ws.worksheetName || ws.sheet) === options.sheet
            );
            if (!found) {
                throw new Error(`Sheet not found: "${options.sheet}"`);
            }
            selectedWorksheet = found;
        }
    }

    const data = (selectedWorksheet && selectedWorksheet.data) ? selectedWorksheet.data : [];

    // If caller provided schema for object transformation / validation
    if (options.schema) {
        return parseSheetData(data, options.schema, options);
    }

    // Default: return 2D array of rows
    return data;
}
