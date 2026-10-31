import parseSheetData from '../schema/parseSheetData.js';

/**
 * Converts a 2D array of rows into an array of objects, using the first row as headers.
 * Handles duplicate or empty headers gracefully.
 *
 * @param {any[][]} rows
 * @returns {Array<Record<string, any>>}
 */
export function convertRowsToObjects(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const seen = new Map();
    const headers = (rows[0] || []).map((h, i) => {
        let base = h != null ? String(h).trim() : '';
        if (!base) base = `column${i + 1}`;
        const count = seen.get(base) || 0;
        seen.set(base, count + 1);
        return count > 0 ? `${base}_${count + 1}` : base;
    });

    const result = [];
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const obj = {};
        for (let c = 0; c < headers.length; c++) {
            obj[headers[c]] = row[c] !== undefined ? row[c] : null;
        }
        result.push(obj);
    }
    return result;
}

/**
 * Turns the TabularJS structure (`{ worksheets }`) a parser returned into what `readSpreadsheet()` returns, by its
 * options: the rows of one sheet, every sheet, objects, or schema-checked rows. Shared by the main reader and the
 * `browser-sync` reader.
 *
 * @param {{ worksheets: Array<any> }} tabularResult
 * @param {Object} [options={}]
 * @returns {any}
 */
export function shapeReadResult(tabularResult, options = {}) {
    const worksheets = tabularResult.worksheets || [];

    // If caller explicitly requested TabularJS format or worksheets
    if (options.format === 'tabular' || options.worksheets === true) {
        return tabularResult;
    }

    const asObjects = options.asObjects === true || options.header === true;

    // If caller requested allSheets (read-excel-file style)
    if (options.allSheets === true) {
        return worksheets.map((ws, i) => {
            const sheetData = ws.data || [];
            return {
                name: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
                data: asObjects ? convertRowsToObjects(sheetData) : sheetData
            };
        });
    }

    // Single sheet selection (options.sheet can be string name or 1-based index)
    let selectedWorksheet = worksheets[0];
    if (options.sheet !== undefined) {
        if (typeof options.sheet === 'number') {
            const index = options.sheet - 1;
            if (index < 0 || index >= worksheets.length) {
                throw new Error(`Sheet not found: ${options.sheet}`);
            }
            selectedWorksheet = worksheets[index];
        } else if (typeof options.sheet === 'string') {
            const found = worksheets.find(
                (ws) => (ws.worksheetName || ws.sheet) === options.sheet
            );
            if (!found) {
                throw new Error(`Sheet not found: "${options.sheet}"`);
            }
            selectedWorksheet = found;
        }
    }

    let data = (selectedWorksheet && selectedWorksheet.data) ? selectedWorksheet.data : [];

    // Enforce maxRows / maxColumns bounds if requested
    if (options.maxColumns && data.length > 0) {
        data = data.map(r => r.length > options.maxColumns ? r.slice(0, options.maxColumns) : r);
    }
    if (options.maxRows && data.length > options.maxRows) {
        data = data.slice(0, options.maxRows);
    }

    if (typeof options.onRow === 'function') {
        data.forEach((row, i) => options.onRow(row, i));
    }

    // If caller provided schema for object transformation / validation
    if (options.schema) {
        return parseSheetData(data, options.schema, options);
    }

    if (asObjects) {
        return convertRowsToObjects(data);
    }

    // Default: return 2D array of rows
    return data;
}
