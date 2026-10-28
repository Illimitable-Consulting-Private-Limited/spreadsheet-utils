import { parseDBF, parseDBFBuffer } from '../parser/dbf.js';
import { parseLotus, parseLotusBuffer, parseWK1, parseWK3, parseWK4 } from '../parser/lotus.js';
import { parseDIF, parseDIFString } from '../parser/dif.js';
import { parseSLK, parseSLKString } from '../parser/slk.js';
import { parseNumbers, isNumbersFile } from '../parser/numbers.js';
import { getZipEntryNames } from '../zip/zip-utils.js';

/**
 * Databases & Legacy Spreadsheets plugin.
 * Adds support for:
 * - dBase database files (.dbf - III, IV, V, FoxPro)
 * - Lotus 1-2-3 (.wks, .wk1, .wk3, .wk4, .123)
 * - Data Interchange Format (.dif)
 * - Symbolic Link (.slk, .sylk)
 * - Apple Numbers (.numbers)
 */
export const legacyPlugin = {
    name: 'legacyPlugin',
    extensions: ['dbf', 'dif', 'slk', 'sylk', 'wks', 'wk1', 'wk3', 'wk4', '123', 'numbers'],

    detect(buffer) {
        if (!buffer || buffer.length < 4) return null;

        // ZIP container check for Apple Numbers
        if (buffer[0] === 0x50 && buffer[1] === 0x4B && (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07)) {
            try {
                const names = getZipEntryNames(buffer);
                if (names.some((n) => n === 'Index.zip' || n.startsWith('Index/') || n.endsWith('.iwa'))) {
                    return 'numbers';
                }
            } catch {
                // Not Numbers
            }
        }

        // Lotus 1-2-3 BOF record
        if (buffer[0] === 0x00 && buffer[1] === 0x00 && (buffer[2] === 0x02 || buffer[2] === 0x1A) && buffer[3] === 0x00) {
            return 'wk1';
        }

        // dBase
        if ([0x02, 0x03, 0x30, 0x31, 0x83, 0x8B, 0xF5].includes(buffer[0]) && buffer.length > 32) {
            const month = buffer[2];
            const day = buffer[3];
            const headerLength = buffer[8] | (buffer[9] << 8);
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && headerLength > 32) {
                return 'dbf';
            }
        }

        // Text heuristics for SYLK and DIF
        const sample = new TextDecoder('utf-8', { fatal: false }).decode(buffer.slice(0, 4096));
        if (sample.includes('\u0000')) return null;
        const text = sample.replace(/^\uFEFF/, '').trimStart();

        if (text.startsWith('ID;')) {
            return 'slk';
        }
        if (text.startsWith('TABLE') && text.includes('VECTORS')) {
            return 'dif';
        }

        return null;
    },

    async parse(ext, file, options = {}) {
        switch (ext) {
            case 'dbf':
                return parseDBF(file, options);
            case 'dif':
                return parseDIF(file, options);
            case 'slk':
            case 'sylk':
                return parseSLK(file, options);
            case 'wks':
            case 'wk1':
            case 'wk3':
            case 'wk4':
            case '123':
                return parseLotus(file, options);
            case 'numbers':
                return parseNumbers(file, options);
            default:
                throw new Error(`Unsupported legacy format: ${ext}`);
        }
    }
};

export {
    parseDBF,
    parseDBFBuffer,
    parseLotus,
    parseLotusBuffer,
    parseWK1,
    parseWK3,
    parseWK4,
    parseDIF,
    parseDIFString,
    parseSLK,
    parseSLKString,
    parseNumbers,
    isNumbersFile
};
export default legacyPlugin;
