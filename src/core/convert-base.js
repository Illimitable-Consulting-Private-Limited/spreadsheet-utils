import { readSpreadsheet, parseTabular } from './dispatcher.js';

/**
 * Creates a conversion function bound to a specific spreadsheet writer.
 *
 * @param {Function} writer
 * @returns {Function}
 */
export function createConverter(writer) {
    /**
     * Converts a spreadsheet file, buffer, or stream from one format to another.
     * Automatically sniffs input format and serializes to the destination format.
     *
     * @param {string|Buffer|Uint8Array|Blob|File} input - Source file path or buffer
     * @param {string|Object} [output] - Destination file path or output options
     * @param {Object} [options={}] - Conversion options (theme, autoWidth, sheet, format, etc.)
     * @returns {Promise<string|Buffer|Blob>}
     */
    return async function convertSpreadsheet(input, output, options = {}) {
        let destPath = null;
        let conversionOptions = {};

        if (typeof output === 'string') {
            destPath = output;
            conversionOptions = { ...options, filePath: destPath };
        } else if (typeof output === 'object' && output !== null) {
            conversionOptions = { ...options, ...output };
            destPath = conversionOptions.filePath || conversionOptions.fileName || null;
        } else {
            conversionOptions = { ...options };
        }

        let targetFormat = conversionOptions.to || conversionOptions.format || conversionOptions.outputFormat;
        if (!targetFormat && destPath) {
            const ext = destPath.split('.').pop().toLowerCase();
            if (ext) targetFormat = ext;
        }
        if (!targetFormat) {
            targetFormat = 'xlsx';
        }

        const readOptions = { ...conversionOptions };
        delete readOptions.format;
        delete readOptions.to;
        delete readOptions.outputFormat;
        delete readOptions.buffer;
        delete readOptions.blob;

        if (conversionOptions.from) {
            readOptions.format = conversionOptions.from;
        } else if (conversionOptions.inputFormat) {
            readOptions.format = conversionOptions.inputFormat;
        }

        const writeOptions = {
            ...conversionOptions,
            format: targetFormat
        };

        const multiSheetWriters = new Set(['xlsx', 'ods', 'xls', 'json']);
        const canWriteMultiSheet = multiSheetWriters.has(targetFormat.toLowerCase());

        let dataToWrite;
        if (conversionOptions.sheet !== undefined) {
            dataToWrite = await readSpreadsheet(input, { ...readOptions, sheet: conversionOptions.sheet });
        } else {
            const tabular = await parseTabular(input, readOptions);
            const worksheets = tabular.worksheets || [];
            if (worksheets.length > 1 && canWriteMultiSheet) {
                dataToWrite = tabular;
            } else if (worksheets.length > 0) {
                dataToWrite = worksheets[0].data || [];
            } else {
                dataToWrite = [];
            }
        }

        return await writer(dataToWrite, writeOptions);
    };
}
