import { getColumnName } from '../utils/helpers.js';
import parseSpreadsheetContents from '../xlsx/parseSpreadsheetContents.js';
import parseXml from '../xml/parseXml.js';
import unpackXlsxFileUniversal from '../export/unpackXlsxFileUniversal.js';
import unpackXlsxFileNode from '../export/unpackXlsxFileNode.js';

/**
 * Low-memory, streaming XLSX parser using read-excel-file's SAX engine.
 * Completely replaces TabularJS's heavy JSZip + DOM tree implementation.
 *
 * @param {string|File|Blob|Uint8Array|ArrayBuffer|Buffer} input
 * @param {Object} [options={}]
 * @returns {Promise<{ worksheets: Array<{ worksheetName: string, data: any[][], columns: any[] }> }>}
 */
export async function parseXLSX(input, options = {}) {
    // Choose unpacker based on environment / input type
    const isNode = typeof process !== 'undefined' && process.versions && process.versions.node;
    const isNodeSpecific = typeof input === 'string' || (typeof Buffer !== 'undefined' && Buffer.isBuffer(input));

    let unpacker = unpackXlsxFileUniversal;
    if (isNode && isNodeSpecific) {
        unpacker = unpackXlsxFileNode;
    }

    const contents = await unpacker(input);
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
            columns.push({
                title: getColumnName(c),
                width: 100
            });
        }
        return {
            worksheetName: s.sheet || s.name || 'Sheet1',
            data,
            columns
        };
    });

    const result = { worksheets };

    if (options.onload) {
        options.onload(result);
    }

    return result;
}
