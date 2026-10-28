import { parseXLS, parseXLSBuffer } from '../parser/xls.js';
import { writeXLS, createXLS } from '../write/xls.js';

/**
 * Legacy Excel format plugin.
 * Adds support for Microsoft Excel 97-2004 (.xls) BIFF5/BIFF8 binary spreadsheets.
 * Includes Compound File Binary (CFB) OLE2 container reader and PTG formula token decoder.
 */
export const xlsPlugin = {
    name: 'xlsPlugin',
    extensions: ['xls'],
    writeExtensions: ['xls'],

    detect(buffer) {
        if (!buffer || buffer.length < 4) return null;
        // OLE2 compound file magic bytes: 0xD0, 0xCF, 0x11, 0xE0
        if (buffer[0] === 0xD0 && buffer[1] === 0xCF && buffer[2] === 0x11 && buffer[3] === 0xE0) {
            return 'xls';
        }
        return null;
    },

    async parse(ext, file, options = {}) {
        return parseXLS(file, options);
    },

    async write(ext, data, options = {}) {
        return writeXLS(data, options);
    }
};

export { parseXLS, parseXLSBuffer, writeXLS, createXLS };
export default xlsPlugin;
