import { getZipEntryNames } from '../zip/zip-utils.js';

/**
 * Detect a spreadsheet format from file content (magic bytes and text heuristics).
 *
 * @param {Uint8Array} buffer - File raw bytes
 * @returns {Promise<string|null>} Detected extension (e.g. 'xlsx', 'ods', 'xls', 'csv') or null
 */
export async function detectFormat(buffer) {
    if (!buffer || buffer.length < 4) {
        return null;
    }

    // ZIP container: xlsx, ods or Apple Numbers
    if (buffer[0] === 0x50 && buffer[1] === 0x4B && (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07)) {
        try {
            const names = getZipEntryNames(buffer);
            if (names.some((n) => n === '[Content_Types].xml' || n.startsWith('xl/'))) {
                return 'xlsx';
            }
            if (names.some((n) => n === 'Index.zip' || n.startsWith('Index/') || n.endsWith('.iwa'))) {
                return 'numbers';
            }
            if (names.includes('mimetype') || names.includes('content.xml')) {
                return 'ods';
            }
        } catch {
            // Corrupt archive; fall through to xlsx as most common
        }
        return 'xlsx';
    }

    // OLE2 compound file (legacy Excel .xls)
    if (buffer[0] === 0xD0 && buffer[1] === 0xCF && buffer[2] === 0x11 && buffer[3] === 0xE0) {
        return 'xls';
    }

    // Lotus 1-2-3 BOF record (0x0002/0x001A with 16-bit length)
    if (buffer[0] === 0x00 && buffer[1] === 0x00 && (buffer[2] === 0x02 || buffer[2] === 0x1A) && buffer[3] === 0x00) {
        return 'wk1';
    }

    // dBase: known version byte + plausible header
    if ([0x02, 0x03, 0x30, 0x31, 0x83, 0x8B, 0xF5].includes(buffer[0]) && buffer.length > 32) {
        const month = buffer[2];
        const day = buffer[3];
        const headerLength = buffer[8] | (buffer[9] << 8);
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && headerLength > 32) {
            return 'dbf';
        }
    }

    // Text-based formats
    const sample = new TextDecoder('utf-8', { fatal: false }).decode(buffer.slice(0, 4096));
    if (sample.includes('\u0000')) {
        return null; // Binary content not recognized
    }
    const text = sample.replace(/^\uFEFF/, '').trimStart();
    const lower = text.toLowerCase();

    if (text.startsWith('ID;')) {
        return 'slk';
    }
    if (text.startsWith('TABLE') && text.includes('VECTORS')) {
        return 'dif';
    }
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
    if (text.length) {
        const firstLine = text.split(/\r?\n/, 1)[0];
        if (firstLine.includes('\t') && !firstLine.includes(',')) {
            return 'tsv';
        }
        return 'csv';
    }

    return null;
}
