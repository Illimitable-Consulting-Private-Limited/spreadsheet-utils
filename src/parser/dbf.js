import { loadAsBuffer, parse } from '../utils/loader.js';
import { readUInt8, readUInt16LE, readUInt32LE, readFloat64LE } from '../utils/helpers.js';

/**
 * Parse DBF (dBase) file format
 * Supports dBase III, IV, V, and FoxPro variants
 *
 * DBF is a database format, so it contains:
 * - Column definitions (name, type, width, decimals)
 * - Row data
 * - Metadata (version, record count, dates)
 *
 * DBF does NOT contain:
 * - Formulas (database stores calculated values)
 * - Fonts/styling (no visual formatting in binary format)
 * - Merged cells (database concept)
 */

// DBF version constants
const DBF_VERSION = {
    0x02: 'FoxBASE',
    0x03: 'dBASE III',
    0x04: 'dBASE IV',
    0x05: 'dBASE V',
    0x30: 'Visual FoxPro',
    0x31: 'Visual FoxPro with AutoIncrement',
    0x43: 'dBASE IV SQL table',
    0x63: 'dBASE IV SQL system',
    0x83: 'dBASE III+ with memo',
    0x8B: 'dBASE IV with memo',
    0x8E: 'dBASE IV SQL table with memo',
    0xF5: 'FoxPro 2.x with memo',
    0xFB: 'FoxPro without memo'
};

// Field type constants
const FIELD_TYPES = {
    'C': 'Character',
    'N': 'Numeric',
    'F': 'Float',
    'L': 'Logical',
    'D': 'Date',
    'M': 'Memo',
    'I': 'Integer',
    'B': 'Double',
    'G': 'General',
    'P': 'Picture',
    'T': 'DateTime',
    'Y': 'Currency',
    '@': 'Timestamp',
    'O': 'Double',
    '+': 'Autoincrement'
};

// Map DBF types to Jspreadsheet types
const DBF_TO_JS_TYPE = {
    'C': 'text',        // Character -> text
    'N': 'numeric',     // Numeric -> numeric
    'F': 'numeric',     // Float -> numeric
    'L': 'checkbox',    // Logical -> checkbox
    'D': 'calendar',    // Date -> calendar
    'M': 'text',        // Memo -> text
    'I': 'numeric',     // Integer -> numeric
    'B': 'numeric',     // Double -> numeric
    'T': 'calendar',    // DateTime -> calendar
    'Y': 'numeric',     // Currency -> numeric
    '@': 'calendar',    // Timestamp -> calendar
    'O': 'numeric',     // Double -> numeric
    '+': 'numeric'      // Autoincrement -> numeric
};

/**
 * Parse DBF header (32 bytes)
 */
function parseHeader(buffer) {
    const version = readUInt8(buffer, 0);

    // Last update date (YY MM DD)
    const year = 1900 + readUInt8(buffer, 1);
    const month = readUInt8(buffer, 2);
    const day = readUInt8(buffer, 3);
    const lastUpdate = new Date(year, month - 1, day);

    const recordCount = readUInt32LE(buffer, 4);
    const headerLength = readUInt16LE(buffer, 8);
    const recordLength = readUInt16LE(buffer, 10);

    // Reserved bytes
    const reserved = buffer.slice(12, 14);

    // Flags
    const incompleteTransaction = readUInt8(buffer, 14);
    const encryptionFlag = readUInt8(buffer, 15);

    // Multi-user processing (reserved)
    const multiUser = buffer.slice(16, 28);

    // MDX flag
    const mdxFlag = readUInt8(buffer, 28);

    // Language driver ID
    const languageDriver = readUInt8(buffer, 29);

    // Reserved
    const reserved2 = buffer.slice(30, 32);

    return {
        version,
        versionName: DBF_VERSION[version] || `Unknown (0x${version.toString(16)})`,
        lastUpdate,
        recordCount,
        headerLength,
        recordLength,
        incompleteTransaction: incompleteTransaction !== 0,
        encrypted: encryptionFlag !== 0,
        hasMDX: mdxFlag !== 0,
        languageDriver
    };
}

/**
 * Parse field descriptor (32 bytes)
 */
function parseFieldDescriptor(buffer, offset) {
    // Field name (11 bytes, null-terminated)
    let nameBytes = [];
    for (let i = 0; i < 11; i++) {
        const byte = buffer[offset + i];
        if (byte === 0) break;
        nameBytes.push(byte);
    }
    const name = String.fromCharCode(...nameBytes);

    // Field type (1 byte)
    const type = String.fromCharCode(buffer[offset + 11]);

    // Field data address (4 bytes) - in memory, not used in file
    const dataAddress = readUInt32LE(buffer, offset + 12);

    // Field length (1 byte)
    const length = readUInt8(buffer, offset + 16);

    // Decimal count (1 byte)
    const decimalCount = readUInt8(buffer, offset + 17);

    // Reserved (2 bytes)
    const reserved1 = buffer.slice(offset + 18, offset + 20);

    // Work area ID (1 byte)
    const workAreaId = readUInt8(buffer, offset + 20);

    // Reserved (10 bytes)
    const reserved2 = buffer.slice(offset + 21, offset + 31);

    // MDX flag (1 byte)
    const mdxFlag = readUInt8(buffer, offset + 31);

    return {
        name,
        type,
        typeName: FIELD_TYPES[type] || 'Unknown',
        length,
        decimalCount,
        workAreaId,
        hasMDX: mdxFlag !== 0
    };
}

/**
 * Check the read options once and return them with their defaults filled in.
 */
function resolveReadOptions(options) {
    options = options || {};
    const bad = (name, value, expected) => {
        throw new Error(`Invalid DBF option ${name}: ${JSON.stringify(value)} (expected ${expected}).`);
    };

    const encoding = options.encoding;
    if (encoding !== undefined && encoding !== 'utf-8' && encoding !== 'latin1' && encoding !== 'iso-8859-1') {
        bad('encoding', encoding, "'utf-8', 'latin1' or 'iso-8859-1'");
    }
    const trimText = options.trimText === undefined ? 'both' : options.trimText;
    if (trimText !== 'both' && trimText !== 'end') {
        bad('trimText', options.trimText, "'both' or 'end'");
    }
    const emptyText = options.emptyText === undefined ? null : options.emptyText;
    if (emptyText !== null && emptyText !== '') {
        bad('emptyText', options.emptyText, "null or ''");
    }
    const dateType = options.dateType === undefined ? 'string' : options.dateType;
    if (dateType !== 'string' && dateType !== 'date') {
        bad('dateType', options.dateType, "'string' or 'date'");
    }

    return {
        latin1: encoding === 'latin1' || encoding === 'iso-8859-1',
        trimEnd: trimText === 'end',
        emptyText,
        dateAsDate: dateType === 'date'
    };
}

/**
 * Decode field bytes to text, not trimmed. `latin1` is one byte a character; otherwise UTF-8 first, latin1 when
 * the bytes are not valid UTF-8.
 */
function decodeFieldValueBytes(bytes, latin1) {
    if (!latin1) {
        try {
            return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
        } catch {
            // not UTF-8: fall through to one byte a character
        }
    }
    let text = '';
    for (let i = 0; i < bytes.length; i += 8192) {
        text += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
    }
    return text;
}

/**
 * Parse a record value based on field type
 */
function parseFieldValue(buffer, offset, field, opts) {
    const bytes = buffer.slice(offset, offset + field.length);
    const decoded = decodeFieldValueBytes(bytes, opts.latin1);

    if (field.type === 'C' || field.type === 'M') {
        // Memo returns the memo block number, actual content in .dbt/.fpt file
        const text = opts.trimEnd ? decoded.replace(/\x20+$/, '') : decoded.trim();
        return text === '' ? opts.emptyText : text;
    }

    const rawValue = decoded.trim();

    if (!rawValue || rawValue === '') {
        return null;
    }

    switch (field.type) {
        case 'N': // Numeric
        case 'F': // Float
            const num = parseFloat(rawValue);
            return isNaN(num) ? null : num;

        case 'I': // Integer (4 bytes, little-endian)
            if (field.length === 4) {
                return readUInt32LE(bytes, 0);
            }
            const intNum = parseInt(rawValue, 10);
            return isNaN(intNum) ? null : intNum;

        case 'L': // Logical
            const logical = rawValue.toUpperCase();
            if (logical === 'T' || logical === 'Y') return true;
            if (logical === 'F' || logical === 'N') return false;
            return null;

        case 'D': // Date (YYYYMMDD)
            if (rawValue.length === 8) {
                const year = parseInt(rawValue.substring(0, 4), 10);
                const month = parseInt(rawValue.substring(4, 6), 10);
                const day = parseInt(rawValue.substring(6, 8), 10);
                if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
                    if (opts.dateAsDate) {
                        const date = new Date(Date.UTC(2000, 0, 1));
                        date.setUTCFullYear(year, month - 1, day);
                        const valid = date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
                        return valid ? date : null;
                    }
                    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                }
            }
            return null;

        case 'T': // DateTime (8 bytes: 4 for date, 4 for time)
            if (field.length === 8) {
                // Julian date (days since 4713 BC) + milliseconds since midnight
                const julianDate = readUInt32LE(bytes, 0);
                const milliseconds = readUInt32LE(bytes, 4);

                if (julianDate > 0) {
                    // Convert Julian to Gregorian (simplified)
                    const unixEpoch = 2440588; // Julian day for 1970-01-01
                    const days = julianDate - unixEpoch;
                    const date = new Date(days * 86400000 + milliseconds);
                    return date.toISOString();
                }
            }
            return null;

        case 'Y': // Currency (8 bytes, little-endian integer, 4 decimal places)
            if (field.length === 8) {
                // Read as 64-bit integer (approximate with two 32-bit)
                const low = readUInt32LE(bytes, 0);
                const high = readUInt32LE(bytes, 4);
                const value = high * 4294967296 + low;
                return value / 10000; // 4 decimal places
            }
            return parseFloat(rawValue) || null;

        case 'B': // Double (8 bytes)
        case 'O': // Double
            if (field.length === 8) {
                return readFloat64LE(bytes, 0);
            }
            return parseFloat(rawValue) || null;

        case '@': // Timestamp
            return rawValue; // Return as string for now

        case '+': // Autoincrement (4 bytes)
            if (field.length === 4) {
                return readUInt32LE(bytes, 0);
            }
            return parseInt(rawValue, 10) || null;

        case 'G': // General (OLE)
        case 'P': // Picture (binary)
            return `<binary:${rawValue.length}bytes>`;

        default:
            return rawValue;
    }
}

/**
 * Parse DBF from buffer
 * @param {Uint8Array} buffer - DBF file buffer
 * @param {Object} [options] - Read options
 * @param {'utf-8'|'latin1'|'iso-8859-1'} [options.encoding] - Text decoding; default UTF-8 with a latin1 fallback
 * @param {'both'|'end'} [options.trimText='both'] - C and M fields: trim both sides, or trailing spaces only
 * @param {null|''} [options.emptyText=null] - Value of a C / M field with nothing left after trimming
 * @param {'string'|'date'} [options.dateType='string'] - D fields as 'YYYY-MM-DD' or a Date at UTC midnight
 * @returns {object} Jspreadsheet Pro format
 */
export function parseDBFBuffer(buffer, options = {}) {
    const opts = resolveReadOptions(options);

    // Ensure buffer is Uint8Array
    if (!(buffer instanceof Uint8Array)) {
        buffer = new Uint8Array(buffer);
    }

    if (!buffer || buffer.length < 32) {
        throw new Error('Invalid DBF file: buffer too small for header');
    }

    // Parse header
    const header = parseHeader(buffer);

    // Parse field descriptors
    const fields = [];
    let offset = 32; // Start after header

    while (offset < header.headerLength - 1 && offset + 32 <= buffer.length) {
        const byte = buffer[offset];

        // 0x0D marks end of field descriptors
        if (byte === 0x0D) {
            break;
        }

        const field = parseFieldDescriptor(buffer, offset);
        fields.push(field);
        offset += 32;
    }

    // Skip to data start (should be at headerLength)
    offset = header.headerLength;

    // Parse records
    const data = [];
    const deletedRecords = [];

    const maxPossibleRecords = (header.recordLength > 0 && buffer.length > header.headerLength)
        ? Math.floor((buffer.length - header.headerLength) / header.recordLength)
        : 0;
    const count = Math.min(header.recordCount, Math.max(0, maxPossibleRecords));

    for (let i = 0; i < count; i++) {
        if (offset >= buffer.length) break;
        // First byte of record: 0x20 = valid, 0x2A = deleted
        const deletionFlag = buffer[offset];
        offset++;

        const record = [];

        for (const field of fields) {
            if (offset >= buffer.length) break;
            const value = parseFieldValue(buffer, offset, field, opts);
            record.push(value);
            offset += field.length;
        }

        if (deletionFlag === 0x2A) {
            deletedRecords.push(i);
        } else {
            data.push(record);
        }
    }

    // Create column definitions with enhanced properties
    const columns = fields.map((field, index) => {
        const column = {
            title: field.name,
            name: field.name,
            width: Math.max(field.length * 10, 100), // Convert to pixels (approximate)
        };

        // Add Jspreadsheet type
        const jsType = DBF_TO_JS_TYPE[field.type];
        if (jsType) {
            column.type = jsType;
        }

        // Add decimal places for numeric fields
        if ((field.type === 'N' || field.type === 'F') && field.decimalCount > 0) {
            column.decimal = '.'.repeat(field.decimalCount);
            column.mask = '#,##0.' + '0'.repeat(field.decimalCount);
        }

        // Add currency mask
        if (field.type === 'Y') {
            column.mask = '$#,##0.00';
        }

        // Add date/datetime format
        if (field.type === 'D') {
            column.options = { format: 'YYYY-MM-DD' };
        } else if (field.type === 'T' || field.type === '@') {
            column.options = { format: 'YYYY-MM-DD HH:mm:ss' };
        }

        return column;
    });

    // Build comprehensive metadata
    const meta = {
        // DBF header information
        dbfVersion: header.versionName,
        dbfVersionCode: `0x${header.version.toString(16)}`,
        lastModified: header.lastUpdate.toISOString(),
        totalRecords: header.recordCount,
        activeRecords: data.length,
        deletedRecords: deletedRecords.length,
        recordLength: header.recordLength,
        headerLength: header.headerLength,

        // Flags
        encrypted: header.encrypted,
        hasMDX: header.hasMDX,
        incompleteTransaction: header.incompleteTransaction,
        languageDriver: header.languageDriver,

        // Field information
        fields: fields.map(field => ({
            name: field.name,
            type: field.type,
            typeName: field.typeName,
            length: field.length,
            decimals: field.decimalCount,
            jsType: DBF_TO_JS_TYPE[field.type] || 'text'
        })),

        // Deleted record indices (if any)
        deletedRecordIndices: deletedRecords.length > 0 ? deletedRecords : undefined
    };

    const headerRow = fields.map(f => f.name);
    const sheetData = [headerRow, ...data];

    // Return in Jspreadsheet Pro format
    return {
        worksheets: [{
            worksheetName: 'DBF Data',
            data: sheetData,
            columns,
            meta
        }]
    };
}

/**
 * Parse DBF file - works in both Browser and Node.js
 * @param {string|File|Uint8Array} input - File path (Node.js), File object (Browser), or buffer
 * @param {Object} options - Parser options
 * @param {Function} options.onload - Callback when parsing completes
 * @param {Function} options.onerror - Callback when parsing fails
 * @param {string} [options.encoding] - See parseDBFBuffer
 * @param {string} [options.trimText] - See parseDBFBuffer
 * @param {null|string} [options.emptyText] - See parseDBFBuffer
 * @param {string} [options.dateType] - See parseDBFBuffer
 * @returns {Promise<object>} Jspreadsheet Pro format
 */
export async function parseDBF(input, options = {}) {
    return parse(async (inp) => {
        const buffer = await loadAsBuffer(inp);
        return parseDBFBuffer(buffer, options);
    }, input, options);
}
