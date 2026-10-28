import { parseODS } from '../parser/ods.js';
import { writeODS, createODS } from '../write/ods.js';
import { getZipEntryNames } from '../zip/zip-utils.js';

/**
 * OpenDocument Spreadsheet (.ods) format plugin.
 * Adds support for reading and writing ODS files.
 */
export const odsPlugin = {
    name: 'odsPlugin',
    extensions: ['ods'],
    writeExtensions: ['ods'],

    detect(buffer) {
        if (!buffer || buffer.length < 4) return null;
        if (buffer[0] === 0x50 && buffer[1] === 0x4B && (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07)) {
            try {
                const names = getZipEntryNames(buffer);
                if (names.includes('mimetype') || names.includes('content.xml')) {
                    return 'ods';
                }
            } catch {
                // Not ODS
            }
        }
        return null;
    },

    async parse(ext, file, options = {}) {
        return parseODS(file, options);
    },

    async write(ext, data, options = {}) {
        return writeODS(data, options);
    }
};

export { parseODS, writeODS, createODS };
export default odsPlugin;
