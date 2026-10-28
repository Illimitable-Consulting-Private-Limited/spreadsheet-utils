import { parseDBF, parseDBFBuffer } from '../parser/dbf.js';
import { writeDBF, createDBF, normalizeSchemaToDbfFields } from '../write/dbf.js';

/**
 * dBASE (.dbf) format plugin.
 * Adds support for reading and exporting DBF files with schema-based field definitions.
 */
export const dbfPlugin = {
    name: 'dbfPlugin',
    extensions: ['dbf'],
    writeExtensions: ['dbf'],

    detect(buffer) {
        if (!buffer || buffer.length < 32) return null;
        if ([0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x43, 0x63, 0x83, 0x8B, 0x8E, 0xF5, 0xFB].includes(buffer[0])) {
            const month = buffer[2];
            const day = buffer[3];
            const headerLength = buffer[8] | (buffer[9] << 8);
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && headerLength >= 32) {
                return 'dbf';
            }
        }
        return null;
    },

    async parse(ext, file, options = {}) {
        return parseDBF(file, options);
    },

    async write(ext, data, options = {}) {
        return writeDBF(data, options);
    }
};

export {
    parseDBF,
    parseDBFBuffer,
    writeDBF,
    createDBF,
    normalizeSchemaToDbfFields
};

export default dbfPlugin;
