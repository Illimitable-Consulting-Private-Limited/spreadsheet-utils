import { getColumnName } from '../utils/helpers.js';
import parseSpreadsheetContents from '../xlsx/parseSpreadsheetContents.js';
import parseXml from '../xml/parseXml.js';
import UnzipError from '../zip/UnzipError.js';
import InvalidInputError from '../xlsx/file/InvalidInputError.js';
import validateLeadingBytes from '../xlsx/file/validateLeadingBytes.js';
import filterZipArchiveEntry from '../export/filterZipArchiveEntry.js';
import { unzipBytesSync } from './unzipSync.js';

/**
 * Same as `parseXLSX()` of `parser/xlsx.js`, but unpacks the archive with `unzipSync()` (no Worker).
 *
 * @param {Uint8Array} bytes
 * @param {object} [options]
 * @returns {Promise<{ worksheets: Array<{ worksheetName: string, data: any[][], columns: any[] }> }>}
 */
export async function parseXLSXSync(bytes, options = {}) {
    validateLeadingBytes(bytes);

    let contents;
    try {
        contents = unzipBytesSync(bytes, {
            filter: (file) => filterZipArchiveEntry({ path: file.name }),
            maxBytes: options.maxUncompressedSize || options.maxBytes
        });
    } catch (error) {
        if (error instanceof UnzipError) {
            throw new InvalidInputError('INVALID_ZIP', error.cause);
        }
        throw error;
    }

    const sheets = await parseSpreadsheetContents(undefined, parseXml, contents, options);

    const worksheets = (sheets || []).map((s) => {
        const data = s.data || [];
        let maxCols = 0;
        for (const row of data) {
            if (row && row.length > maxCols) {
                maxCols = row.length;
            }
        }
        const columns = [];
        for (let c = 0; c < maxCols; c++) {
            columns.push({ title: getColumnName(c), width: 100 });
        }
        return { worksheetName: s.sheet || s.name || 'Sheet1', data, columns };
    });

    const result = { worksheets };
    if (options.onload) {
        options.onload(result);
    }
    return result;
}
