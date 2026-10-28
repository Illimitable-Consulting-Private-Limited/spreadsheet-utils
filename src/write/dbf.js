import fs from 'node:fs';

/**
 * DBF (dBase) field type constants.
 */
export const DBF_FIELD_TYPES = {
    CHARACTER: 'C',
    NUMERIC: 'N',
    FLOAT: 'F',
    INTEGER: 'I',
    LOGICAL: 'L',
    DATE: 'D',
    DATETIME: 'T',
    DOUBLE: 'B',
    CURRENCY: 'Y',
    MEMO: 'M'
};

/**
 * Encode string to bytes according to specified encoding.
 * Supports UTF-8, Latin-1 (ISO-8859-1), and ASCII universally.
 *
 * @param {string} str
 * @param {string} [encoding='utf-8']
 * @returns {Uint8Array|Buffer}
 */
export function encodeString(str, encoding = 'utf-8') {
    if (str === null || str === undefined) {
        return new Uint8Array(0);
    }
    const enc = String(encoding || 'utf-8').toLowerCase().replace('-', '');
    if (typeof Buffer !== 'undefined') {
        const nodeEnc = (enc === 'latin1' || enc === 'iso88591')
            ? 'latin1'
            : (enc === 'ascii' ? 'ascii' : 'utf8');
        return Buffer.from(String(str), nodeEnc);
    }
    if (enc === 'latin1' || enc === 'iso88591' || enc === 'ascii') {
        const s = String(str);
        const bytes = new Uint8Array(s.length);
        for (let i = 0; i < s.length; i++) {
            bytes[i] = s.charCodeAt(i) & 0xFF;
        }
        return bytes;
    }
    return new TextEncoder().encode(String(str));
}

/**
 * Sanitize a field name for dBase standard compliance:
 * - Up to 10 uppercase ASCII characters (alphanumeric and underscore)
 * - Must not start with a digit
 * - Avoids collision with already registered field names
 *
 * @param {string} rawName
 * @param {number} index
 * @param {Set<string>} existingNames
 * @returns {string}
 */
export function sanitizeDbfFieldName(rawName, index = 0, existingNames = new Set()) {
    let name = String(rawName || '').trim();
    if (!name) {
        name = `FIELD_${index + 1}`;
    } else {
        name = name.replace(/[\s\-\.]+/g, '_');
        name = name.replace(/[^a-zA-Z0-9_]/g, '');
        name = name.toUpperCase();
        if (/^[0-9]/.test(name)) {
            name = `F_${name}`;
        }
        if (!name) {
            name = `FIELD_${index + 1}`;
        }
    }

    let base = name.slice(0, 10);
    let candidate = base;
    let counter = 2;

    while (existingNames.has(candidate)) {
        const suffix = `_${counter}`;
        const maxBaseLen = 10 - suffix.length;
        candidate = `${base.slice(0, Math.max(1, maxBaseLen))}${suffix}`;
        counter++;
    }

    existingNames.add(candidate);
    return candidate;
}

/**
 * Extract number of decimal places from a format pattern (e.g. '#,##0.00' -> 2).
 *
 * @param {string} format
 * @returns {number|undefined}
 */
export function extractDecimalsFromFormat(format) {
    if (!format || typeof format !== 'string') return undefined;
    const match = format.match(/\.([0#]+)/);
    if (match) {
        return match[1].length;
    }
    if (format === '0' || format === '#,##0') {
        return 0;
    }
    return undefined;
}

/**
 * Map user schema or JS type to DBF field type code.
 *
 * @param {any} type
 * @param {Object} [options={}]
 * @returns {string} Single-letter DBF type code ('C', 'N', 'F', 'I', 'L', 'D', 'T', 'B', 'Y', 'M')
 */
export function mapTypeToDbf(type, options = {}) {
    if (!type) return null;
    if (type === String || type === 'string' || type === 'text' || type === 'char' || type === 'character' || type === 'C') {
        return 'C';
    }
    if (type === Number || type === 'number' || type === 'numeric' || type === 'N') {
        return 'N';
    }
    if (type === 'integer' || type === 'int' || type === 'I') {
        return (options.fileVersion === 0x30 || type === 'I') ? 'I' : 'N';
    }
    if (type === 'float' || type === 'F') {
        return 'F';
    }
    if (type === 'double' || type === 'B' || type === 'O') {
        return 'B';
    }
    if (type === Boolean || type === 'boolean' || type === 'bool' || type === 'logical' || type === 'L') {
        return 'L';
    }
    if (type === Date || type === 'date' || type === 'D') {
        return 'D';
    }
    if (type === 'datetime' || type === 'timestamp' || type === 'T') {
        return 'T';
    }
    if (type === 'currency' || type === 'Y') {
        return 'Y';
    }
    if (type === 'memo' || type === 'M') {
        return 'M';
    }
    return 'C';
}

/**
 * Get standard default field length in bytes for a given DBF type.
 *
 * @param {string} type
 * @param {number} [decimalPlaces=0]
 * @returns {number}
 */
export function getDefaultFieldSize(type, decimalPlaces = 0) {
    switch (type) {
        case 'C': return 30;
        case 'N': return decimalPlaces > 0 ? 12 : 10;
        case 'F': return 16;
        case 'I': return 4;
        case 'L': return 1;
        case 'D': return 8;
        case 'T': return 8;
        case 'B': return 8;
        case 'Y': return 8;
        case 'M': return 10;
        default: return 30;
    }
}

/**
 * Normalize raw input data (TabularJS worksheets, 2D array, or object array)
 * into a uniform record structure with header list.
 *
 * @param {any} input
 * @returns {{ rows: any[], headers: string[], is2D: boolean }}
 */
export function extractRowsAndHeaders(input) {
    if (!input) return { rows: [], headers: [], is2D: false };

    // Case 1: TabularJS workbook / worksheets object
    if (input.worksheets && Array.isArray(input.worksheets) && input.worksheets[0]) {
        const ws = input.worksheets[0];
        const data = ws.data || [];
        if (data.length > 0) {
            const hasExplicitColumns = Array.isArray(ws.columns) && ws.columns.length > 0;
            const headers = hasExplicitColumns
                ? ws.columns.map(c => c.name || c.title || '')
                : data[0].map(c => String((c && typeof c === 'object' && 'value' in c) ? c.value : c));
            const rows = hasExplicitColumns ? data : data.slice(1);
            return { rows, headers, is2D: true };
        }
        return { rows: [], headers: [], is2D: true };
    }

    if (!Array.isArray(input)) return { rows: [], headers: [], is2D: false };
    if (input.length === 0) return { rows: [], headers: [], is2D: false };

    // Case 2: 2D Array of rows
    if (Array.isArray(input[0])) {
        const headers = input[0].map(c => String((c && typeof c === 'object' && 'value' in c) ? c.value : c));
        const rows = input.slice(1);
        return { rows, headers, is2D: true };
    }

    // Case 3: Array of objects
    if (typeof input[0] === 'object' && input[0] !== null) {
        const keySet = new Set();
        for (const obj of input) {
            if (obj && typeof obj === 'object') {
                for (const k of Object.keys(obj)) {
                    if (k !== '_deleted' && k !== '__deleted') {
                        keySet.add(k);
                    }
                }
            }
        }
        const headers = Array.from(keySet);
        return { rows: input, headers, is2D: false };
    }

    return { rows: [], headers: [], is2D: false };
}

/**
 * Infer the best matching DBF type code by scanning values.
 *
 * @param {any[]} values
 * @returns {string}
 */
function inferFieldTypeFromValues(values) {
    let hasString = false;
    let hasNumber = false;
    let hasDate = false;
    let hasBoolean = false;

    for (const raw of values) {
        const val = (raw && typeof raw === 'object' && 'value' in raw && !(raw instanceof Date)) ? raw.value : raw;
        if (val === null || val === undefined) continue;
        if (typeof val === 'boolean') {
            hasBoolean = true;
        } else if (val instanceof Date) {
            hasDate = true;
        } else if (typeof val === 'number' && !isNaN(val)) {
            hasNumber = true;
        } else if (typeof val === 'string') {
            if (/^\d{4}-\d{2}-\d{2}$/.test(val.trim())) {
                hasDate = true;
            } else {
                hasString = true;
            }
        } else {
            hasString = true;
        }
    }

    if (hasString) return 'C';
    if (hasNumber) return 'N';
    if (hasDate) return 'D';
    if (hasBoolean) return 'L';
    return 'C';
}

/**
 * Convert user schema / fieldDescriptors / columns into full DBF field descriptors.
 *
 * Supports:
 * - write-excel-file schema objects: { colName: { column, type, size, decimalPlaces, format, value } }
 * - write-excel-file schema arrays: [ { column, prop, type, size, decimalPlaces, format, value } ]
 * - yortus/DBFFile fieldDescriptors: [ { name, type, size, decimalPlaces } ]
 * - Automatic type and size detection when schema is omitted or partially defined.
 *
 * @param {Object|Array} [schema]
 * @param {any} [data=[]]
 * @param {Object} [options={}]
 * @returns {Array<{ name: string, type: string, size: number, decimalPlaces: number, value: Function, prop?: string }>}
 */
export function normalizeSchemaToDbfFields(schema, data = [], options = {}) {
    const { rows, headers, is2D } = extractRowsAndHeaders(data);
    const existingNames = new Set();
    const rawSchema = schema || options.schema || options.fieldDescriptors || options.columns;

    let fieldConfigs = [];

    if (Array.isArray(rawSchema)) {
        // Array of column definitions or DBFFile fieldDescriptors
        fieldConfigs = rawSchema.map((colDef, index) => {
            const rawHeader = colDef.name || colDef.column || colDef.header || colDef.title || colDef.prop || `FIELD_${index + 1}`;
            const prop = colDef.prop || colDef.name || colDef.column;
            const accessor = typeof colDef.value === 'function'
                ? colDef.value
                : (row) => {
                    if (is2D && Array.isArray(row)) {
                        return row[index];
                    }
                    if (row && typeof row === 'object') {
                        if (prop && prop in row) return row[prop];
                        if (rawHeader in row) return row[rawHeader];
                    }
                    return null;
                };

            return {
                rawName: rawHeader,
                prop,
                type: colDef.type,
                size: colDef.size || colDef.length || colDef.width,
                decimalPlaces: colDef.decimalPlaces !== undefined ? colDef.decimalPlaces : (colDef.decimals !== undefined ? colDef.decimals : colDef.decimalCount),
                format: colDef.format,
                value: accessor,
                index
            };
        });
    } else if (rawSchema && typeof rawSchema === 'object') {
        // Object schema: keys are property names
        fieldConfigs = Object.entries(rawSchema).map(([propKey, colDef], index) => {
            const rawHeader = colDef.column || colDef.header || colDef.name || colDef.title || propKey;
            const accessor = typeof colDef.value === 'function'
                ? colDef.value
                : (row) => {
                    if (is2D && Array.isArray(row)) {
                        return row[index];
                    }
                    if (row && typeof row === 'object') {
                        if (propKey in row) return row[propKey];
                        if (rawHeader in row) return row[rawHeader];
                    }
                    return null;
                };

            return {
                rawName: rawHeader,
                prop: propKey,
                type: colDef.type,
                size: colDef.size || colDef.length || colDef.width,
                decimalPlaces: colDef.decimalPlaces !== undefined ? colDef.decimalPlaces : (colDef.decimals !== undefined ? colDef.decimals : colDef.decimalCount),
                format: colDef.format,
                value: accessor,
                index
            };
        });
    } else {
        // Auto-generate fields from data headers
        fieldConfigs = headers.map((header, index) => {
            const accessor = (row) => {
                if (is2D && Array.isArray(row)) {
                    return row[index];
                }
                if (row && typeof row === 'object') {
                    return row[header];
                }
                return null;
            };

            return {
                rawName: header || `FIELD_${index + 1}`,
                prop: header,
                type: null,
                size: null,
                decimalPlaces: null,
                format: null,
                value: accessor,
                index
            };
        });
    }

    // Resolve and finalize each field descriptor
    return fieldConfigs.map((cfg, index) => {
        const sanitizedName = sanitizeDbfFieldName(cfg.rawName, index, existingNames);

        // Collect sample values for this field across rows
        const sampleValues = [];
        for (let r = 0; r < Math.min(rows.length, 1000); r++) {
            const rowVal = cfg.value(rows[r], r);
            const val = (rowVal && typeof rowVal === 'object' && 'value' in rowVal && !(rowVal instanceof Date))
                ? rowVal.value
                : rowVal;
            if (val !== null && val !== undefined) {
                sampleValues.push(val);
            }
        }

        // Determine DBF type
        let dbfType = mapTypeToDbf(cfg.type, options);
        if (!dbfType) {
            dbfType = inferFieldTypeFromValues(sampleValues);
        }

        // Determine decimal places
        let decPlaces = cfg.decimalPlaces;
        if (decPlaces === null || decPlaces === undefined) {
            const fmtDec = extractDecimalsFromFormat(cfg.format);
            if (fmtDec !== undefined) {
                decPlaces = fmtDec;
            } else if (dbfType === 'N' || dbfType === 'F') {
                let maxFoundDecimals = 0;
                let hasFloats = false;
                for (const val of sampleValues) {
                    if (typeof val === 'number' && !isNaN(val) && val % 1 !== 0) {
                        hasFloats = true;
                        const decStr = String(val).split('.')[1] || '';
                        if (decStr.length > maxFoundDecimals) {
                            maxFoundDecimals = Math.min(decStr.length, 15);
                        }
                    }
                }
                decPlaces = hasFloats ? Math.max(2, maxFoundDecimals) : 0;
            } else if (dbfType === 'Y') {
                decPlaces = 4;
            } else {
                decPlaces = 0;
            }
        }

        // Determine field size (in bytes)
        let fieldSize = cfg.size;
        if (!fieldSize || fieldSize <= 0) {
            if (dbfType === 'C') {
                let maxByteLen = 0;
                for (const val of sampleValues) {
                    const str = String(val);
                    const len = (typeof Buffer !== 'undefined')
                        ? Buffer.byteLength(str, 'utf-8')
                        : new TextEncoder().encode(str).length;
                    if (len > maxByteLen) maxByteLen = len;
                }
                fieldSize = Math.min(254, Math.max(10, maxByteLen || 30));
            } else if (dbfType === 'N' || dbfType === 'F') {
                let maxIntDigits = 1;
                let hasNeg = false;
                for (const val of sampleValues) {
                    if (typeof val === 'number' && !isNaN(val)) {
                        if (val < 0) hasNeg = true;
                        const intDigits = Math.floor(Math.abs(val)).toString().length;
                        if (intDigits > maxIntDigits) maxIntDigits = intDigits;
                    }
                }
                const neededSize = maxIntDigits + (decPlaces > 0 ? decPlaces + 1 : 0) + (hasNeg ? 1 : 0);
                fieldSize = Math.min(20, Math.max(decPlaces > 0 ? 12 : 10, neededSize));
            } else {
                fieldSize = getDefaultFieldSize(dbfType, decPlaces);
            }
        }

        // Enforce maximum size limits per DBF specification
        if (dbfType === 'C') {
            fieldSize = Math.min(254, Math.max(1, fieldSize));
        } else if (dbfType === 'N' || dbfType === 'F') {
            fieldSize = Math.min(20, Math.max(1, fieldSize));
        } else if (dbfType === 'I') {
            fieldSize = 4;
            decPlaces = 0;
        } else if (dbfType === 'L') {
            fieldSize = 1;
            decPlaces = 0;
        } else if (dbfType === 'D' || dbfType === 'T' || dbfType === 'B' || dbfType === 'Y') {
            fieldSize = 8;
            if (dbfType !== 'Y') decPlaces = 0;
        } else if (dbfType === 'M') {
            fieldSize = 10;
            decPlaces = 0;
        }

        return {
            name: sanitizedName,
            prop: cfg.prop,
            type: dbfType,
            size: fieldSize,
            decimalPlaces: decPlaces,
            value: cfg.value
        };
    });
}

/**
 * Format and write a single field's value into the binary record buffer.
 */
function writeFieldValue(buffer, view, offset, rawVal, field, encoding) {
    let val = rawVal;
    if (val && typeof val === 'object' && 'value' in val && !(val instanceof Date)) {
        val = val.value;
    }

    switch (field.type) {
        case 'C': {
            if (val === null || val === undefined) {
                buffer.fill(0x20, offset, offset + field.size);
            } else {
                const bytes = encodeString(String(val), encoding);
                const copyLen = Math.min(bytes.length, field.size);
                buffer.set(bytes.subarray(0, copyLen), offset);
                if (copyLen < field.size) {
                    buffer.fill(0x20, offset + copyLen, offset + field.size);
                }
            }
            break;
        }

        case 'N':
        case 'F': {
            if (val === null || val === undefined || (typeof val === 'number' && isNaN(val))) {
                buffer.fill(0x20, offset, offset + field.size);
            } else {
                const num = Number(val);
                let str = '';
                if (isNaN(num)) {
                    str = '';
                } else if (field.decimalPlaces > 0) {
                    str = num.toFixed(field.decimalPlaces);
                } else {
                    str = Math.round(num).toString();
                }

                if (str.length > field.size) {
                    // Overflow indication in dBase
                    buffer.fill(0x2A, offset, offset + field.size);
                } else {
                    const padded = str.padStart(field.size, ' ');
                    for (let i = 0; i < field.size; i++) {
                        buffer[offset + i] = padded.charCodeAt(i);
                    }
                }
            }
            break;
        }

        case 'L': {
            if (val === true || val === 'T' || val === 't' || val === 'Y' || val === 'y' || val === 1 || val === '1') {
                buffer[offset] = 0x54; // 'T'
            } else if (val === false || val === 'F' || val === 'f' || val === 'N' || val === 'n' || val === 0 || val === '0') {
                buffer[offset] = 0x46; // 'F'
            } else {
                buffer[offset] = 0x3F; // '?'
            }
            break;
        }

        case 'D': {
            let dateStr = null;
            if (val instanceof Date && !isNaN(val.getTime())) {
                const y = val.getFullYear();
                const m = String(val.getMonth() + 1).padStart(2, '0');
                const d = String(val.getDate()).padStart(2, '0');
                dateStr = `${String(y).padStart(4, '0')}${m}${d}`;
            } else if (typeof val === 'string' && val.trim().length > 0) {
                const clean = val.replace(/[^0-9]/g, '');
                if (clean.length >= 8) {
                    dateStr = clean.slice(0, 8);
                }
            } else if (typeof val === 'number' && !isNaN(val)) {
                const dt = new Date(val);
                if (!isNaN(dt.getTime())) {
                    const y = dt.getFullYear();
                    const m = String(dt.getMonth() + 1).padStart(2, '0');
                    const d = String(dt.getDate()).padStart(2, '0');
                    dateStr = `${String(y).padStart(4, '0')}${m}${d}`;
                }
            }

            if (dateStr && dateStr.length === 8) {
                for (let i = 0; i < 8; i++) {
                    buffer[offset + i] = dateStr.charCodeAt(i);
                }
            } else {
                buffer.fill(0x20, offset, offset + 8);
            }
            break;
        }

        case 'I': {
            const intVal = (val === null || val === undefined) ? 0 : Math.round(Number(val) || 0);
            view.setInt32(offset, intVal, true);
            break;
        }

        case 'B': {
            const dblVal = (val === null || val === undefined) ? 0 : (Number(val) || 0);
            view.setFloat64(offset, dblVal, true);
            break;
        }

        case 'T': {
            if (val) {
                const dt = (val instanceof Date) ? val : new Date(val);
                if (!isNaN(dt.getTime())) {
                    const epochMs = dt.getTime();
                    const julian = Math.floor(epochMs / 86400000) + 2440588;
                    const ms = ((epochMs % 86400000) + 86400000) % 86400000;
                    view.setUint32(offset, julian, true);
                    view.setUint32(offset + 4, ms, true);
                } else {
                    buffer.fill(0, offset, offset + 8);
                }
            } else {
                buffer.fill(0, offset, offset + 8);
            }
            break;
        }

        case 'Y': {
            const curNum = (val === null || val === undefined) ? 0 : (Number(val) || 0);
            const scaled = BigInt(Math.round(curNum * 10000));
            view.setBigInt64(offset, scaled, true);
            break;
        }

        case 'M': {
            const memoStr = (val === null || val === undefined) ? '' : String(val);
            const padded = memoStr.padStart(10, ' ').slice(-10);
            for (let i = 0; i < 10; i++) {
                buffer[offset + i] = padded.charCodeAt(i);
            }
            break;
        }

        default: {
            // Fallback for unknown type: write as string
            const strVal = (val === null || val === undefined) ? '' : String(val);
            const bytes = encodeString(strVal, encoding);
            const copyLen = Math.min(bytes.length, field.size);
            buffer.set(bytes.subarray(0, copyLen), offset);
            if (copyLen < field.size) {
                buffer.fill(0x20, offset + copyLen, offset + field.size);
            }
            break;
        }
    }
}

/**
 * Universal in-memory DBF binary generator.
 * Produces a complete, valid dBase / FoxPro DBF file as a Uint8Array.
 *
 * @param {any} data - 2D grid, array of objects, or TabularJS worksheets
 * @param {Object} [options={}]
 * @param {Object|Array} [options.schema] - Write-excel-file style schema definition
 * @param {Array} [options.fieldDescriptors] - yortus/DBFFile style field descriptors
 * @param {number} [options.fileVersion=0x03] - 0x03 (dBASE III), 0x30 (Visual FoxPro), etc.
 * @param {string} [options.encoding='utf-8'] - Character encoding ('utf-8', 'latin1', 'ascii')
 * @param {Date} [options.lastUpdate] - Custom last update date in header
 * @param {number} [options.languageDriver] - DBF language driver ID byte
 * @param {Function} [options.isRowDeleted] - Predicate returning true if row is marked deleted
 * @param {number[]} [options.deletedRows] - Indices of rows to mark deleted
 * @returns {Uint8Array}
 */
export function createDBF(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const { rows } = extractRowsAndHeaders(data);
    const fields = normalizeSchemaToDbfFields(
        options.schema || options.fieldDescriptors || options.columns,
        data,
        options
    );

    if (fields.length === 0) {
        throw new Error('Cannot export DBF: No columns or fields defined.');
    }

    const encoding = options.encoding || 'utf-8';
    const fileVersion = typeof options.fileVersion === 'number'
        ? options.fileVersion
        : (options.fileVersion === 'Visual FoxPro' ? 0x30 : 0x03);

    const headerLength = 32 + (fields.length * 32) + 1;
    const recordLength = 1 + fields.reduce((sum, f) => sum + f.size, 0);
    const totalLength = headerLength + (rows.length * recordLength) + 1;

    const buffer = new Uint8Array(totalLength);
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);

    // 1. Header (32 bytes)
    buffer[0] = fileVersion & 0xFF;

    const lastUpdate = options.lastUpdate instanceof Date ? options.lastUpdate : new Date();
    buffer[1] = (lastUpdate.getFullYear() - 1900) & 0xFF;
    buffer[2] = (lastUpdate.getMonth() + 1) & 0xFF;
    buffer[3] = lastUpdate.getDate() & 0xFF;

    view.setUint32(4, rows.length, true);
    view.setUint16(8, headerLength, true);
    view.setUint16(10, recordLength, true);

    const encLower = encoding.toLowerCase();
    const defaultLangDriver = (encLower.includes('latin') || encLower.includes('1252') || encLower.includes('iso'))
        ? 0x03
        : 0x00;
    buffer[29] = options.languageDriver !== undefined ? options.languageDriver : defaultLangDriver;

    // 2. Field Descriptors (32 bytes per field)
    for (let i = 0; i < fields.length; i++) {
        const field = fields[i];
        const fieldOffset = 32 + (i * 32);

        // Field name (up to 11 bytes, null padded)
        const nameBytes = encodeString(field.name, 'ascii');
        const nameLen = Math.min(nameBytes.length, 10);
        buffer.set(nameBytes.subarray(0, nameLen), fieldOffset);
        buffer.fill(0x00, fieldOffset + nameLen, fieldOffset + 11);

        // Field type
        buffer[fieldOffset + 11] = field.type.charCodeAt(0);

        // Length and decimal places
        buffer[fieldOffset + 16] = field.size & 0xFF;
        buffer[fieldOffset + 17] = (field.decimalPlaces || 0) & 0xFF;
    }

    // Header terminator
    buffer[32 + (fields.length * 32)] = 0x0D;

    // 3. Data Records
    let recOffset = headerLength;
    const deletedSet = Array.isArray(options.deletedRows) ? new Set(options.deletedRows) : null;

    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];

        const isDeleted = Boolean(row && (row._deleted === true || row.__deleted === true))
            || (typeof options.isRowDeleted === 'function' && options.isRowDeleted(row, r))
            || (deletedSet && deletedSet.has(r));

        buffer[recOffset] = isDeleted ? 0x2A : 0x20;

        let colOffset = recOffset + 1;
        for (let f = 0; f < fields.length; f++) {
            const field = fields[f];
            const rawVal = field.value(row, r);
            writeFieldValue(buffer, view, colOffset, rawVal, field, encoding);
            colOffset += field.size;
        }

        recOffset += recordLength;
    }

    // 4. File Terminator (0x1A - EOF marker)
    buffer[totalLength - 1] = 0x1A;

    return buffer;
}

/**
 * Universal styled/schema-based DBF export function.
 *
 * In Node.js:
 * - Writes to file if `options.filePath` is provided and returns the filePath.
 * - Returns Buffer by default (or Uint8Array if `options.uint8Array` is set).
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Promise<Buffer|Uint8Array|string>}
 */
export async function writeDBF(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const uint8Array = createDBF(data, options);

    if (options.filePath) {
        await fs.promises.writeFile(options.filePath, uint8Array);
        return options.filePath;
    }

    if (options.uint8Array) {
        return uint8Array;
    }

    if (typeof Buffer !== 'undefined') {
        return Buffer.from(uint8Array.buffer, uint8Array.byteOffset, uint8Array.byteLength);
    }

    return uint8Array;
}

export default writeDBF;
