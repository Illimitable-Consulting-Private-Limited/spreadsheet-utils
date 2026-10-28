import { parseODS } from '../parser/ods.js';
import { parseHTMLTable, parseAllHTMLTables } from '../parser/table.js';
import { parseXMLSpreadsheet, parseXMLSpreadsheetString } from '../parser/xml.js';
import { getZipEntryNames } from '../zip/zip-utils.js';
import { writeODS, createODS } from '../write/ods.js';

/**
 * Structured Documents & Markup format plugin.
 * Adds support for:
 * - OpenDocument Spreadsheet (.ods)
 * - HTML Tables (.html, .htm)
 * - XML Spreadsheet 2003 / SpreadsheetML (.xml)
 *
 * Encapsulates the shared XML/HTML AST tree scanner.
 */
export const markupPlugin = {
    name: 'markupPlugin',
    extensions: ['ods', 'html', 'htm', 'xml'],
    writeExtensions: ['ods'],

    detect(buffer) {
        if (!buffer || buffer.length < 4) return null;

        // ZIP container check for ODS
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

        // Text heuristics
        const sample = new TextDecoder('utf-8', { fatal: false }).decode(buffer.slice(0, 4096));
        if (sample.includes('\u0000')) return null;
        const text = sample.replace(/^\uFEFF/, '').trimStart();
        const lower = text.toLowerCase();

        if (lower.startsWith('<?xml')) {
            if (lower.includes('urn:schemas-microsoft-com:office:spreadsheet') || lower.includes('<workbook')) {
                return 'xml';
            }
            if (lower.includes('<html') || lower.includes('<table')) {
                return 'html';
            }
            return 'xml';
        }

        if (lower.startsWith('<!doctype html') || lower.includes('<html') || lower.includes('<table')) {
            return 'html';
        }

        return null;
    },

    async parse(ext, file, options = {}) {
        switch (ext) {
            case 'ods':
                return parseODS(file, options);
            case 'html':
            case 'htm':
                return parseHTMLTable(file, options);
            case 'xml':
                return parseXMLSpreadsheet(file, options);
            default:
                throw new Error(`Unsupported markup format: ${ext}`);
        }
    },

    async write(ext, data, options = {}) {
        if (ext === 'ods') {
            return writeODS(data, options);
        }
        throw new Error(`Writing format .${ext} is not supported by markupPlugin`);
    }
};

export { parseODS, parseHTMLTable, parseAllHTMLTables, parseXMLSpreadsheet, parseXMLSpreadsheetString, writeODS, createODS };
export default markupPlugin;
