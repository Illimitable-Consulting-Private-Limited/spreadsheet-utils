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
