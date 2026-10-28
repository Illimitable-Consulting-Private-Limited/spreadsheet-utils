import { parseCSVFile, parseCSVString } from '../parser/csv.js';

/**
 * Text and Delimited formats plugin.
 * Adds support for CSV, TSV, TAB, and TXT spreadsheet files.
 */
export const textPlugin = {
    name: 'textPlugin',
    extensions: ['csv', 'tsv', 'tab', 'txt'],

    detect(buffer) {
        if (!buffer || buffer.length === 0) return null;
        // Text-based formats
        const sample = new TextDecoder('utf-8', { fatal: false }).decode(buffer.slice(0, 4096));
        if (sample.includes('\u0000')) {
            return null;
        }
        const text = sample.replace(/^\uFEFF/, '').trimStart();
        const lower = text.toLowerCase();
        if (lower.startsWith('<?xml') || lower.startsWith('<!doctype html') || lower.includes('<html') || lower.includes('<table')) {
            return null; // Markup formats handled by markupPlugin
        }
        if (text.startsWith('ID;') || (text.startsWith('TABLE') && text.includes('VECTORS'))) {
            return null; // Handled by legacyPlugin
        }
        if (text.length > 0) {
            const firstLine = text.split(/\r?\n/, 1)[0];
            if (firstLine.includes('\t') && !firstLine.includes(',')) {
                return 'tsv';
            }
            return 'csv';
        }
        return null;
    },

    async parse(ext, file, options = {}) {
        switch (ext) {
            case 'tsv':
            case 'tab':
                return parseCSVFile(file, { ...options, delimiter: '\t' });
            case 'txt':
                return parseCSVFile(file, { ...options, delimiter: options.delimiter || '\t' });
            case 'csv':
            default:
                return parseCSVFile(file, options);
        }
    }
};

export { parseCSVFile, parseCSVString };
export default textPlugin;
