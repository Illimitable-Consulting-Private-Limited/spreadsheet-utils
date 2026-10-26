import { unzip, unzipSync } from 'fflate';

/**
 * Default per-entry or per-archive decompression cap.
 * Protects against zip bombs.
 */
export const DEFAULT_MAX_ENTRY_BYTES = 256 * 1024 * 1024;

/**
 * Normalize binary input to Uint8Array.
 * @param {Uint8Array|ArrayBuffer|Buffer} input
 * @returns {Uint8Array}
 */
export function toUint8Array(input) {
    if (input instanceof Uint8Array) {
        return input;
    }
    if (typeof ArrayBuffer !== 'undefined' && input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }
    if (typeof Buffer !== 'undefined' && Buffer.isBuffer(input)) {
        return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }
    throw new TypeError('Expected Uint8Array, ArrayBuffer, or Buffer');
}

/**
 * Scan zip archive entries without decompressing any file contents.
 * Very fast, near-zero memory footprint.
 *
 * @param {Uint8Array|ArrayBuffer|Buffer} buffer
 * @returns {string[]} List of entry paths in the zip archive
 */
export function getZipEntryNames(buffer) {
    const bytes = toUint8Array(buffer);
    const names = [];
    try {
        unzipSync(bytes, {
            filter: (file) => {
                names.push(file.name);
                return false; // Skips decompression entirely
            }
        });
    } catch {
        // Corrupt or truncated archive
    }
    return names;
}

/**
 * Decompress files from a zip archive with optional filtering and size protection.
 *
 * @param {Uint8Array|ArrayBuffer|Buffer} buffer - Zip archive binary data
 * @param {Object} [options={}]
 * @param {function(string): boolean} [options.filter] - Predicate returning true for entries to decompress
 * @param {number} [options.maxBytes=DEFAULT_MAX_ENTRY_BYTES] - Maximum total uncompressed bytes allowed
 * @returns {Promise<Record<string, Uint8Array>>} Map of entry path to uncompressed Uint8Array
 */
export function unzipArchive(buffer, options = {}) {
    const bytes = toUint8Array(buffer);
    const filter = options.filter;
    const maxBytes = options.maxBytes || DEFAULT_MAX_ENTRY_BYTES;

    return new Promise((resolve, reject) => {
        let totalUncompressed = 0;
        let settled = false;

        unzip(bytes, {
            filter: (file) => {
                if (filter && !filter(file.name)) {
                    return false;
                }
                if (file.originalSize) {
                    totalUncompressed += file.originalSize;
                    if (totalUncompressed > maxBytes) {
                        settled = true;
                        reject(new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`));
                        return false;
                    }
                }
                return true;
            }
        }, (err, unzipped) => {
            if (settled) return;
            if (err) return reject(err);
            resolve(unzipped || {});
        });
    });
}

/**
 * Synchronous version of unzipArchive.
 *
 * @param {Uint8Array|ArrayBuffer|Buffer} buffer
 * @param {Object} [options={}]
 * @returns {Record<string, Uint8Array>}
 */
export function unzipArchiveSync(buffer, options = {}) {
    const bytes = toUint8Array(buffer);
    const filter = options.filter;
    return unzipSync(bytes, {
        filter: filter ? (file) => filter(file.name) : undefined
    });
}

/**
 * Convert Uint8Array to specified format.
 *
 * @param {Uint8Array} bytes
 * @param {'string'|'uint8array'|'nodebuffer'|'base64'} type
 * @returns {string|Uint8Array|Buffer}
 */
export function convertEntryBytes(bytes, type = 'string') {
    if (!bytes) return null;
    switch (type) {
        case 'string':
            return new TextDecoder('utf-8').decode(bytes);
        case 'base64':
            if (typeof Buffer !== 'undefined') {
                return Buffer.from(bytes).toString('base64');
            }
            let binary = '';
            for (let i = 0; i < bytes.length; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            return btoa(binary);
        case 'nodebuffer':
            return typeof Buffer !== 'undefined' ? Buffer.from(bytes) : bytes;
        default:
            return bytes;
    }
}
