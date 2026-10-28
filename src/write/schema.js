import { normalizeSchemaToDbfFields, sanitizeDbfFieldName } from './dbf.js';

export { normalizeSchemaToDbfFields, sanitizeDbfFieldName };

const STYLE_PROPERTIES = [
    'fontWeight', 'fontStyle', 'fontSize', 'fontFamily',
    'color', 'textColor', 'backgroundColor',
    'align', 'verticalAlign', 'wrap',
    'borderColor', 'borderStyle',
    'leftBorderStyle', 'rightBorderStyle', 'topBorderStyle', 'bottomBorderStyle',
    'leftBorderColor', 'rightBorderColor', 'topBorderColor', 'bottomBorderColor',
    'columnSpan', 'rowSpan', 'span'
];

/**
 * Extract style attributes from a column definition or object.
 */
function extractStyles(def) {
    if (!def || typeof def !== 'object') return {};
    const styles = {};
    for (const key of STYLE_PROPERTIES) {
        if (def[key] !== undefined) {
            styles[key] = def[key];
        }
    }
    if (def.style && typeof def.style === 'object') {
        Object.assign(styles, def.style);
    }
    return styles;
}

/**
 * Build a styled or plain header cell object.
 */
function buildHeaderCell(colDef, fallbackName, options = {}) {
    const rawHeader = colDef.header !== undefined
        ? colDef.header
        : (colDef.column !== undefined ? colDef.column : (colDef.name !== undefined ? colDef.name : fallbackName));

    const headerStyles = {
        ...(options.headerStyle || {}),
        ...(colDef.headerStyle || {})
    };

    if (rawHeader && typeof rawHeader === 'object' && !(rawHeader instanceof Date)) {
        return {
            value: rawHeader.value !== undefined ? rawHeader.value : fallbackName,
            ...headerStyles,
            ...rawHeader
        };
    }

    if (Object.keys(headerStyles).length > 0) {
        return {
            value: rawHeader !== undefined && rawHeader !== null ? rawHeader : fallbackName,
            ...headerStyles
        };
    }

    return rawHeader !== undefined && rawHeader !== null ? rawHeader : fallbackName;
}

/**
 * Convert schema definitions into columns format for write-excel-file and universal writers.
 * Supports:
 * - header as string or object { value, ...styles }
 * - headerStyle option / per-column headerStyle
 * - user-provided cell(row, index) callback
 * - rich object return from value(row, index)
 * - dynamic style(val, row, index) function
 * - top-level style properties on column definition
 *
 * @param {Array|Object} schema
 * @param {Object} [options={}]
 * @returns {Array}
 */
export function normalizeSchemaToColumns(schema, options = {}) {
    if (Array.isArray(schema)) {
        return schema.map((col) => {
            const baseStyles = extractStyles(col);
            const header = buildHeaderCell(col, '', options);

            return {
                header,
                cell: (row, index) => {
                    // 1. If user provided a custom cell function, use it
                    if (typeof col.cell === 'function') {
                        const userCell = col.cell(row, index);
                        if (userCell && typeof userCell === 'object' && !(userCell instanceof Date)) {
                            return {
                                type: col.type,
                                format: col.format,
                                ...baseStyles,
                                ...userCell
                            };
                        }
                        return {
                            value: userCell !== undefined ? userCell : null,
                            type: col.type,
                            format: col.format,
                            ...baseStyles
                        };
                    }

                    // 2. Extract value from value function, prop, column, or name
                    const val = typeof col.value === 'function'
                        ? col.value(row, index)
                        : (col.prop ? row[col.prop] : (col.column ? row[col.column] : (col.name ? row[col.name] : null)));

                    // 3. Dynamic style function support
                    const dynamicStyles = typeof col.style === 'function'
                        ? col.style(val, row, index)
                        : {};

                    // 4. If val is an object with a value property (rich cell return), merge it
                    if (val && typeof val === 'object' && !(val instanceof Date) && 'value' in val) {
                        return {
                            type: col.type,
                            format: col.format,
                            ...baseStyles,
                            ...dynamicStyles,
                            ...val
                        };
                    }

                    return {
                        value: val !== undefined ? val : null,
                        type: col.type,
                        format: col.format,
                        ...baseStyles,
                        ...dynamicStyles
                    };
                },
                width: col.width
            };
        });
    }

    if (schema && typeof schema === 'object') {
        return Object.entries(schema).map(([propKey, colDef]) => {
            const baseStyles = extractStyles(colDef);
            const header = buildHeaderCell(colDef, propKey, options);

            return {
                header,
                cell: (row, index) => {
                    // 1. If user provided a custom cell function, use it
                    if (typeof colDef.cell === 'function') {
                        const userCell = colDef.cell(row, index);
                        if (userCell && typeof userCell === 'object' && !(userCell instanceof Date)) {
                            return {
                                type: colDef.type,
                                format: colDef.format,
                                ...baseStyles,
                                ...userCell
                            };
                        }
                        return {
                            value: userCell !== undefined ? userCell : null,
                            type: colDef.type,
                            format: colDef.format,
                            ...baseStyles
                        };
                    }

                    // 2. Extract value from value function or row property
                    const val = typeof colDef.value === 'function'
                        ? colDef.value(row, index)
                        : (row ? row[propKey] : null);

                    // 3. Dynamic style function support
                    const dynamicStyles = typeof colDef.style === 'function'
                        ? colDef.style(val, row, index)
                        : {};

                    // 4. If val is an object with a value property, merge it
                    if (val && typeof val === 'object' && !(val instanceof Date) && 'value' in val) {
                        return {
                            type: colDef.type,
                            format: colDef.format,
                            ...baseStyles,
                            ...dynamicStyles,
                            ...val
                        };
                    }

                    return {
                        value: val !== undefined ? val : null,
                        type: colDef.type,
                        format: colDef.format,
                        ...baseStyles,
                        ...dynamicStyles
                    };
                },
                width: colDef.width
            };
        });
    }

    return [];
}

/**
 * Determine display character length of a cell value.
 */
function getCellLength(cell) {
    if (cell == null) return 0;
    if (typeof cell === 'object') {
        if (cell instanceof Date) {
            return 10;
        }
        if (cell.value !== undefined && cell.value !== null) {
            return getCellLength(cell.value);
        }
        return 0;
    }
    if (typeof cell === 'number') {
        return String(cell).length;
    }
    if (typeof cell === 'boolean') {
        return cell ? 4 : 5;
    }
    const str = String(cell);
    if (str.includes('\n')) {
        const lines = str.split('\n');
        let maxL = 0;
        for (const line of lines) {
            if (line.length > maxL) maxL = line.length;
        }
        return maxL;
    }
    return str.length;
}

/**
 * Calculates responsive column widths based on maximum content length in each column.
 * Enforces minWidth and maxWidth safeguards to prevent layout breakages.
 *
 * @param {Array<Array<any>>} rows - 2D grid of rows and cells
 * @param {Object} [options={}] - Options containing autoWidth, minWidth, maxWidth, padding
 * @param {Array<Object>} [existingColumns=null] - Existing column definitions to augment or preserve
 * @returns {Array<{ width: number }>}
 */
export function calculateAutoWidths(rows, options = {}, existingColumns = null) {
    if (!Array.isArray(rows) || rows.length === 0) {
        return existingColumns || [];
    }

    const autoWidthConfig = typeof options.autoWidth === 'object' && options.autoWidth !== null
        ? options.autoWidth
        : {};

    const minWidth = autoWidthConfig.minWidth ?? options.minWidth ?? 8;
    const maxWidth = autoWidthConfig.maxWidth ?? options.maxWidth ?? 40;
    const padding = autoWidthConfig.padding ?? options.padding ?? 2;
    const force = autoWidthConfig.force === true;

    // Determine total number of columns
    let numCols = 0;
    if (Array.isArray(existingColumns)) {
        numCols = existingColumns.length;
    }
    for (const row of rows) {
        if (Array.isArray(row) && row.length > numCols) {
            numCols = row.length;
        }
    }

    if (numCols === 0) return existingColumns || [];

    const colWidths = new Array(numCols).fill(0);

    for (const row of rows) {
        if (!Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
            const len = getCellLength(row[c]);
            if (len > colWidths[c]) {
                colWidths[c] = len;
            }
        }
    }

    const result = [];
    for (let c = 0; c < numCols; c++) {
        const existingCol = (existingColumns && existingColumns[c]) ? { ...existingColumns[c] } : {};
        if (!force && existingCol.width != null) {
            result.push(existingCol);
        } else {
            const calculated = Math.max(minWidth, Math.min(maxWidth, colWidths[c] + padding));
            existingCol.width = calculated;
            result.push(existingCol);
        }
    }

    return result;
}
