import { unzipSync } from 'fflate';

import { createUnzipError } from '../zip/UnzipError.js';

const DEFAULT_MAX_BYTES = 256 * 1024 * 1024;

// `fflate` throws errors that carry a numeric `code`.
function isFlateError(error) {
    return Boolean(error) && typeof error.code === 'number';
}

/**
 * Unpacks a `.zip` archive with fflate's `unzipSync()`, so no Worker is started however large the archive is
 * (`unzip()` starts one for archives over about 512 KB, and a Content Security Policy without `worker-src` blocks it).
 *
 * @param {Uint8Array} bytes
 * @param {object} [options]
 * @param {(entry: { name: string }) => boolean} [options.filter] — Which entries to unpack.
 * @param {number} [options.maxBytes] — Total uncompressed size limit (zip bomb guard).
 * @returns {Record<string, Uint8Array>}
 * @throws {UnzipError} When the archive is not a valid `.zip`.
 */
export function unzipBytesSync(bytes, options = {}) {
    const filter = options.filter;
    const maxBytes = options.maxBytes || DEFAULT_MAX_BYTES;
    let total = 0;
    try {
        return unzipSync(bytes, {
            filter: (file) => {
                if (filter && !filter(file)) {
                    return false;
                }
                if (file.originalSize) {
                    total += file.originalSize;
                    if (total > maxBytes) {
                        throw new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`);
                    }
                }
                return true;
            }
        });
    } catch (error) {
        throw isFlateError(error) ? createUnzipError(error) : error;
    }
}

/**
 * Same shape as `unzipArchive()` of `zip-utils.js` (the filter gets the entry name, the result is awaited), done
 * synchronously. Used for `.ods`.
 *
 * @param {Uint8Array} bytes
 * @param {object} [options]
 * @param {(name: string) => boolean} [options.filter]
 * @param {number} [options.maxBytes]
 * @returns {Promise<Record<string, Uint8Array>>}
 */
export function unzipArchiveSyncAsPromise(bytes, options = {}) {
    const nameFilter = options.filter;
    try {
        return Promise.resolve(unzipBytesSync(bytes, {
            maxBytes: options.maxBytes,
            filter: nameFilter ? (file) => nameFilter(file.name) : undefined
        }));
    } catch (error) {
        return Promise.reject(error);
    }
}

/**
 * Lists the entry names of a `.zip` archive without unpacking anything. An unreadable archive gives an empty list.
 * @param {Uint8Array} bytes
 * @returns {string[]}
 */
export function zipEntryNamesSync(bytes) {
    const names = [];
    try {
        unzipSync(bytes, {
            filter: (file) => {
                names.push(file.name);
                return false;
            }
        });
    } catch {
        // Corrupt or truncated archive
    }
    return names;
}
