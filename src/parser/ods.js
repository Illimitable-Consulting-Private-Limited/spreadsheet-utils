import { unzipArchive } from '../zip/zip-utils.js';
import { parseODSWith } from './odsCore.js';

/**
 * Parse ODS file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseODS(input, options = {}) {
    return parseODSWith(unzipArchive, input, options);
}
