import { parseXLSX } from '../parser/xlsx.js';
import { loadAsBuffer } from '../utils/loader.js';
import { detectFormat } from './detector.js';
import { getParserForExtension, createMissingPluginError, KNOWN_PLUGIN_MAPPINGS } from './registry.js';
import parseSheetData from '../schema/parseSheetData.js';

/**
 * Route a resolved format extension to its parser.
 * Handles core XLSX directly, looks up plugins in local or global registry,
 * or throws a helpful missing-plugin error.
 *
 * @param {string} ext - Lowercase file extension
 * @param {any} file - Input file
 * @param {Object} [options={}] - Parser options
 * @param {Object} [context={}] - Execution context with optional localParsers
 * @returns {Promise<Object>}
 */
export function dispatch(ext, file, options = {}, context = {}) {
    const key = (ext || '').toLowerCase().replace(/^\./, '');

    // Core format: XLSX is built-in
    if (key === 'xlsx') {
        return parseXLSX(file, options);
    }

    // Check plugin registry (local scoped registry first, then global)
    const parser = getParserForExtension(key, context.localParsers);
    if (parser) {
        return parser(key, file, options);
    }

    // If format is a known plugin format but not registered, give helpful diagnostic
    if (KNOWN_PLUGIN_MAPPINGS[key]) {
        throw createMissingPluginError(key);
    }

    return null;
}

/**
 * Low-level spreadsheet parser returning TabularJS worksheets structure.
 * Compatible with TabularJS.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file
 * @param {Object} [options={}]
 * @param {Object} [context={}]
 * @returns {Promise<{ worksheets: Array<any> }>}
 */
export async function parseTabular(file, options = {}, context = {}) {
    if (!file) {
        throw new Error('Invalid file');
    }

    // Security check: maxFileSize if configured
    if (options.maxFileSize) {
        if (typeof file === 'object' && typeof file.size === 'number' && file.size > options.maxFileSize) {
            throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
        }
        if (file instanceof Uint8Array && file.length > options.maxFileSize) {
            throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
        }
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
        const result = dispatch(ext, file, options, context);
        if (result) {
            return await result;
        }
        if (!KNOWN_PLUGIN_MAPPINGS[ext]) {
            throw new Error(`Unsupported file type: ${ext || fileName}`);
        }
    }

    if (!ext && typeof file === 'string' && !file.includes('\n') && !file.includes('<')) {
        throw new Error(`Unsupported file type: ${fileName || '(none)'}`);
    }

    // Unknown or missing extension: sniff format from content
    let buffer = null;
    let detected = null;
    try {
        buffer = await loadAsBuffer(file);
        if (options.maxFileSize && buffer.length > options.maxFileSize) {
            throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
        }
        detected = await detectFormat(buffer, context.localSniffers);
    } catch (err) {
        if (err && err.message && err.message.includes('File size limit')) {
            throw err;
        }
        // Fall through
    }

    if (detected) {
        const result = dispatch(detected, buffer || file, options, context);
        if (result) {
            return await result;
        }
    }

    if (ext && KNOWN_PLUGIN_MAPPINGS[ext]) {
        throw createMissingPluginError(ext, fileName);
    }
    if (detected && KNOWN_PLUGIN_MAPPINGS[detected]) {
        throw createMissingPluginError(detected, fileName);
    }

    throw new Error(`Unsupported file type: ${ext || detected || fileName || '(buffer)'}`);
}

/**
 * Converts a 2D array of rows into an array of objects, using the first row as headers.
 * Handles duplicate or empty headers gracefully.
 *
 * @param {any[][]} rows
 * @returns {Array<Record<string, any>>}
 */
export function convertRowsToObjects(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const seen = new Map();
    const headers = (rows[0] || []).map((h, i) => {
        let base = h != null ? String(h).trim() : '';
        if (!base) base = `column${i + 1}`;
        const count = seen.get(base) || 0;
        seen.set(base, count + 1);
        return count > 0 ? `${base}_${count + 1}` : base;
    });

    const result = [];
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const obj = {};
        for (let c = 0; c < headers.length; c++) {
            obj[headers[c]] = row[c] !== undefined ? row[c] : null;
        }
        result.push(obj);
    }
    return result;
}

/**
 * Universal spreadsheet reader supporting both read-excel-file options/output
 * and TabularJS format across XLSX and registered plugin formats.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file
 * @param {Object} [options={}]
 * @param {Object} [context={}]
 * @returns {Promise<any[][] | { rows: any[], errors: any[] } | Array<{ name: string, data: any[][] | Array<Record<string, any>> }> | { worksheets: any[] } | Array<Record<string, any>>>}
 */
export async function readSpreadsheet(file, options = {}, context = {}) {
    // 1. Get tabular worksheets structure
    const tabularResult = await parseTabular(file, options, context);
    const worksheets = tabularResult.worksheets || [];

    // If caller explicitly requested TabularJS format or worksheets
    if (options.format === 'tabular' || options.worksheets === true) {
        return tabularResult;
    }

    const asObjects = options.asObjects === true || options.header === true;

    // If caller requested allSheets (read-excel-file style)
    if (options.allSheets === true) {
        return worksheets.map((ws, i) => {
            const sheetData = ws.data || [];
            return {
                name: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
                data: asObjects ? convertRowsToObjects(sheetData) : sheetData
            };
        });
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

    let data = (selectedWorksheet && selectedWorksheet.data) ? selectedWorksheet.data : [];

    // Enforce maxRows / maxColumns bounds if requested
    if (options.maxColumns && data.length > 0) {
        data = data.map(r => r.length > options.maxColumns ? r.slice(0, options.maxColumns) : r);
    }
    if (options.maxRows && data.length > options.maxRows) {
        data = data.slice(0, options.maxRows);
    }

    if (typeof options.onRow === 'function') {
        data.forEach((row, i) => options.onRow(row, i));
    }

    // If caller provided schema for object transformation / validation
    if (options.schema) {
        return parseSheetData(data, options.schema, options);
    }

    if (asObjects) {
        return convertRowsToObjects(data);
    }

    // Default: return 2D array of rows
    return data;
}

/**
 * Streaming row iterator yielding rows one-by-one via an async generator.
 * Provides constant O(1) memory consumption for large datasets.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} file
 * @param {Object} [options={}]
 * @param {Object} [context={}]
 * @returns {AsyncGenerator<any[] | Record<string, any>, void, unknown>}
 */
export async function* iterateSpreadsheet(file, options = {}, context = {}) {
    const rowQueue = [];
    let resolveNext = null;
    let finished = false;
    let parseError = null;

    const asObjects = options.asObjects === true || options.header === true;
    let headers = null;

    const streamingOptions = {
        ...options,
        // Override asObjects in the inner read so onRow receives raw rows
        asObjects: false,
        header: false,
        onRow: (row, rowIndex) => {
            if (asObjects) {
                if (headers === null) {
                    const seen = new Map();
                    headers = (row || []).map((h, i) => {
                        let base = h != null ? String(h).trim() : '';
                        if (!base) base = `column${i + 1}`;
                        const count = seen.get(base) || 0;
                        seen.set(base, count + 1);
                        return count > 0 ? `${base}_${count + 1}` : base;
                    });
                    return;
                }
                const obj = {};
                for (let c = 0; c < headers.length; c++) {
                    obj[headers[c]] = row[c] !== undefined ? row[c] : null;
                }
                rowQueue.push(obj);
            } else {
                rowQueue.push(row);
            }
            if (resolveNext) {
                const cb = resolveNext;
                resolveNext = null;
                cb();
            }
        }
    };

    const parsePromise = readSpreadsheet(file, streamingOptions, context)
        .then(() => {
            finished = true;
            if (resolveNext) {
                const cb = resolveNext;
                resolveNext = null;
                cb();
            }
        })
        .catch((err) => {
            parseError = err;
            finished = true;
            if (resolveNext) {
                const cb = resolveNext;
                resolveNext = null;
                cb();
            }
        });

    while (true) {
        if (rowQueue.length > 0) {
            yield rowQueue.shift();
        } else if (finished) {
            if (parseError) {
                throw parseError;
            }
            break;
        } else {
            await new Promise((resolve) => {
                resolveNext = resolve;
            });
        }
    }

    await parsePromise;
}

/**
 * Factory to create an isolated reader instance with dedicated plugins.
 * Useful when multiple parts of an app require different format sets without mutating global state.
 *
 * @param {Array<Object>|Object} plugins
 * @returns {Function & { tabular: Function, register: Function, iterate: Function }}
 */
export function createReader(plugins = []) {
    const localParsers = new Map();
    const localSniffers = new Set();
    const pluginList = Array.isArray(plugins) ? plugins : [plugins];

    function addPlugins(list) {
        for (const plugin of list) {
            if (!plugin || typeof plugin !== 'object') continue;
            if (Array.isArray(plugin.extensions) && typeof plugin.parse === 'function') {
                for (const ext of plugin.extensions) {
                    localParsers.set(ext.toLowerCase(), plugin.parse);
                }
            }
            if (typeof plugin.detect === 'function') {
                localSniffers.add(plugin.detect);
            }
        }
    }

    addPlugins(pluginList);

    const context = { localParsers, localSniffers };

    const reader = (file, options = {}) => readSpreadsheet(file, options, context);
    reader.tabular = (file, options = {}) => parseTabular(file, options, context);
    reader.iterate = (file, options = {}) => iterateSpreadsheet(file, options, context);
    reader.register = (...extra) => {
        addPlugins(extra);
        return reader;
    };

    return reader;
}
