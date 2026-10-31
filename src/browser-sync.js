/**
 * spreadsheet-utils/browser-sync
 *
 * The browser entry for pages served under a Content Security Policy that blocks `blob:` Workers (no `worker-src`
 * allowing it). fflate's async `zip()` / `unzip()` start a Worker for larger files (an archive over about 160 KB when
 * writing, over about 512 KB when reading), so every other entry fails on a big file there. This entry uses
 * `zipSync()` / `unzipSync()` only: nothing starts a Worker, at the price of doing the work on the calling thread.
 *
 * Reads `.xlsx`, `.xls` and `.ods`; writes `.xlsx` (with `rowBreaks`, `stickyRowsCount`, ... as in `/write`).
 * Other formats are in the other entries.
 */
import { shapeReadResult, convertRowsToObjects } from './core/shapeReadResult.js';
import { unzipArchiveSyncAsPromise, zipEntryNamesSync } from './sync/unzipSync.js';
import { parseXLSXSync } from './sync/parseXLSXSync.js';
import { writeXlsxFileSync } from './sync/writeXlsxFileSync.js';
import { parseXLSBuffer } from './parser/xls.js';
import { parseODSWith } from './parser/odsCore.js';
import InvalidInputError from './xlsx/file/InvalidInputError.js';
import InvalidSpreadsheetError from './xlsx/InvalidSpreadsheetError.js';
import SheetNotFoundError from './xlsx/SheetNotFoundError.js';

const SUPPORTED = 'xlsx, xls and ods';

async function toBytes(input) {
    if (input instanceof Uint8Array) {
        return input;
    }
    if (typeof ArrayBuffer !== 'undefined' && input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }
    if (typeof Blob !== 'undefined' && input instanceof Blob) {
        return new Uint8Array(await input.arrayBuffer());
    }
    throw new Error('Invalid input: expected a File, Blob, Uint8Array or ArrayBuffer');
}

function detect(bytes) {
    if (bytes.length < 4) {
        return null;
    }
    // ZIP container: xlsx or ods
    if (bytes[0] === 0x50 && bytes[1] === 0x4B && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)) {
        // An unreadable archive has no names and goes to the xlsx reader, which reports it.
        const names = zipEntryNamesSync(bytes);
        if (names.some((n) => n === '[Content_Types].xml' || n.startsWith('xl/'))) {
            return 'xlsx';
        }
        if (names.includes('mimetype') || names.includes('content.xml')) {
            return 'ods';
        }
        return 'xlsx';
    }
    // OLE2 compound file: xls
    if (bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) {
        return 'xls';
    }
    return null;
}

/**
 * Low-level reader returning TabularJS worksheets (`{ worksheets: [{ worksheetName, data, columns }] }`).
 *
 * @param {File|Blob|Uint8Array|ArrayBuffer} file
 * @param {Object} [options={}] — `extension` or `format` names the type; else the file name's extension, else sniffed.
 * @returns {Promise<{ worksheets: Array<any> }>}
 */
export async function parseTabular(file, options = {}) {
    if (!file) {
        throw new Error('Invalid file');
    }
    const bytes = await toBytes(file);

    if (options.maxFileSize && bytes.length > options.maxFileSize) {
        throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
    }

    let ext = null;
    if (options.extension) {
        ext = String(options.extension).toLowerCase().replace(/^\./, '');
    } else if (options.format && options.format !== 'tabular' && options.format !== 'rows') {
        ext = String(options.format).toLowerCase().replace(/^\./, '');
    } else if (typeof file.name === 'string' && file.name.lastIndexOf('.') > -1) {
        ext = file.name.toLowerCase().slice(file.name.lastIndexOf('.') + 1);
    }
    if (!ext || !['xlsx', 'xls', 'ods'].includes(ext)) {
        const detected = detect(bytes);
        if (!detected && ext) {
            throw new Error(`Unsupported file type: ${ext} (spreadsheet-utils/browser-sync reads ${SUPPORTED})`);
        }
        if (!detected) {
            throw new Error(`Unsupported file type (spreadsheet-utils/browser-sync reads ${SUPPORTED})`);
        }
        ext = detected;
    }

    if (ext === 'xlsx') {
        return parseXLSXSync(bytes, options);
    }
    if (ext === 'xls') {
        const result = parseXLSBuffer(bytes);
        if (options.onload) {
            options.onload(result);
        }
        return result;
    }
    return parseODSWith(unzipArchiveSyncAsPromise, bytes, options);
}

/**
 * Reads a spreadsheet to rows, by the same options as `readSpreadsheet()` of `spreadsheet-utils`.
 *
 * @param {File|Blob|Uint8Array|ArrayBuffer} file
 * @param {Object} [options={}]
 * @returns {Promise<any>}
 */
export async function readSpreadsheet(file, options = {}) {
    return shapeReadResult(await parseTabular(file, options), options);
}

// read-excel-file backwards compatibility
export function readSheet(input, sheet, options) {
    if (!options && sheet && typeof sheet !== 'number' && typeof sheet !== 'string') {
        options = sheet;
        sheet = undefined;
    }
    return readSpreadsheet(input, { ...options, sheet });
}

const readXlsxFile = readSpreadsheet;

/**
 * Creates an `*.xlsx` file; the same arguments as `writeXlsxFile()` of `spreadsheet-utils/write`.
 * `toUint8Array()` is synchronous; `toBlob()` and `toFile(fileName)` return Promises.
 */
const writeXlsxFile = writeXlsxFileSync;

export {
    readSpreadsheet as default,
    readXlsxFile,
    writeXlsxFile,
    convertRowsToObjects,
    InvalidInputError,
    InvalidSpreadsheetError,
    SheetNotFoundError
};
