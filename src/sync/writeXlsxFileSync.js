import { zipSync, strToU8 } from 'fflate';

import generateXlsxFileContents from '../write/xlsx/generateXlsxFileContents.js';
import downloadBlob from '../write/export/downloadBlob.js';

const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function toUint8Array(path, content) {
    if (typeof content === 'string') {
        return strToU8(content);
    }
    if (content instanceof Uint8Array) {
        return content;
    }
    if (typeof ArrayBuffer !== 'undefined' && content instanceof ArrayBuffer) {
        return new Uint8Array(content);
    }
    throw new Error(`${path}: a file inside the .xlsx given as a Blob or File can't be read synchronously; pass an ArrayBuffer or a Uint8Array`);
}

/**
 * Creates an `*.xlsx` file with `zipSync()`, so no Worker is started however large the file is.
 * Takes the same arguments as `writeXlsxFile()` of `spreadsheet-utils/write`.
 *
 * @param {object[][]|object[]} arg1 — Sheet data, or an array of `{ sheet, data, ...sheetOptions }`.
 * @param {object} [arg2] — Sheet options (single sheet) or options (multiple sheets).
 * @param {object} [arg3] — Options (single sheet).
 * @returns {{ toUint8Array: () => Uint8Array, toBlob: () => Promise<Blob>, toFile: (fileName: string) => Promise<void> }}
 */
export function writeXlsxFileSync(arg1, arg2, arg3) {
    const toUint8ArrayNow = () => {
        const files = generateXlsxFileContents(arg1, arg2, arg3);
        const entries = {};
        for (const path of Object.keys(files)) {
            entries[path] = toUint8Array(path, files[path]);
        }
        return zipSync(entries);
    };
    const toBlob = () => {
        try {
            return Promise.resolve(new Blob([toUint8ArrayNow()], { type: XLSX_MIME_TYPE }));
        } catch (error) {
            return Promise.reject(error);
        }
    };
    return {
        toUint8Array: toUint8ArrayNow,
        toBlob,
        toFile: (fileName) => toBlob().then((blob) => {
            downloadBlob(blob, fileName);
        })
    };
}
