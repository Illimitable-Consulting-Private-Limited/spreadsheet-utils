import { normalizeSchemaToDbfFields, sanitizeDbfFieldName } from './dbf.js';

export { normalizeSchemaToDbfFields, sanitizeDbfFieldName };

/**
 * Convert schema definitions into columns format for write-excel-file.
 */
export function normalizeSchemaToColumns(schema) {
    if (Array.isArray(schema)) {
        return schema.map(col => ({
            header: col.column || col.header || col.name || '',
            cell: (row, index) => {
                const val = typeof col.value === 'function'
                    ? col.value(row, index)
                    : (col.prop ? row[col.prop] : (col.column ? row[col.column] : (col.name ? row[col.name] : null)));
                return {
                    value: val !== undefined ? val : null,
                    type: col.type,
                    format: col.format,
                    ...(col.style || {})
                };
            },
            width: col.width
        }));
    }
    if (schema && typeof schema === 'object') {
        return Object.entries(schema).map(([propKey, colDef]) => ({
            header: colDef.column || colDef.header || propKey,
            cell: (row, index) => {
                const val = typeof colDef.value === 'function'
                    ? colDef.value(row, index)
                    : row[propKey];
                return {
                    value: val !== undefined ? val : null,
                    type: colDef.type,
                    format: colDef.format,
                    ...(colDef.style || {})
                };
            },
            width: colDef.width
        }));
    }
    return [];
}
