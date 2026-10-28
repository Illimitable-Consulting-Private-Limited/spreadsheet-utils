import fs from 'node:fs';
import { zipSync, strToU8 } from 'fflate';
import { normalizeSchemaToColumns } from './schema.js';
import getSheetData from './getSheetData/getSheetData.js';

/**
 * Escape XML special characters.
 * @param {any} str
 * @returns {string}
 */
export function escapeXml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Normalize raw input data into uniform sheet structures.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Array<{ name: string, rows: any[][] }>}
 */
export function normalizeWorkbookSheets(data, options = {}) {
    let sheets = [];

    // Case 1: TabularJS format ({ worksheets: [...] })
    if (data && Array.isArray(data.worksheets)) {
        sheets = data.worksheets.map((ws, i) => {
            const sheetName = ws.worksheetName || ws.sheet || ws.name || `Sheet${i + 1}`;
            let rows = ws.data || [];
            let columns = ws.columns;
            if (ws.columns && Array.isArray(ws.columns) && rows.length > 0 && !Array.isArray(rows[0])) {
                columns = normalizeSchemaToColumns(ws.columns, { ...options, ...ws });
                rows = getSheetData(rows, columns);
            }
            return { name: sheetName, rows, columns };
        });
    }
    // Case 2: Multi-sheet array ([ { sheet: '...', data: [...] }, ... ])
    else if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object' && data[0] !== null && !Array.isArray(data[0]) && (data[0].sheet || data[0].data)) {
        sheets = data.map((s, i) => {
            const sheetName = s.sheet || s.name || `Sheet${i + 1}`;
            let rows = s.data || [];
            const schema = s.schema || s.columns || options.schema;
            let columns = s.columns;
            if (schema && rows.length > 0 && !Array.isArray(rows[0])) {
                columns = normalizeSchemaToColumns(schema, { ...options, ...s });
                rows = getSheetData(rows, columns);
            }
            return { name: sheetName, rows, columns };
        });
    }
    // Case 3: Object array with schema
    else if (Array.isArray(data) && data.length > 0 && !Array.isArray(data[0])) {
        const schema = options.schema || options.columns;
        if (schema) {
            const columns = normalizeSchemaToColumns(schema, options);
            const grid = getSheetData(data, columns);
            sheets = [{
                name: options.sheet || 'Sheet1',
                rows: grid,
                columns
            }];
        } else {
            const keys = Object.keys(data[0] || {});
            const headerStyle = options.headerStyle || { fontWeight: 'bold' };
            const headerRow = keys.map(k => ({ value: k, ...headerStyle }));
            const dataRows = data.map(obj => keys.map(k => ({ value: obj[k] })));
            sheets = [{
                name: options.sheet || 'Sheet1',
                rows: [headerRow, ...dataRows]
            }];
        }
    }
    // Case 4: 2D array
    else if (Array.isArray(data)) {
        sheets = [{
            name: options.sheet || 'Sheet1',
            rows: data,
            columns: options.columns
        }];
    }
    else {
        sheets = [{
            name: options.sheet || 'Sheet1',
            rows: []
        }];
    }

    return sheets;
}

function formatBorderWidth(b) {
    if (!b) return '0.06pt';
    const s = String(b).toLowerCase();
    if (s === 'medium') return '1.5pt';
    if (s === 'thick') return '2.5pt';
    if (s.includes('px') || s.includes('pt')) return s.split(' ')[0];
    return '0.06pt';
}

function formatBorderStyle(b) {
    if (!b) return 'solid';
    const s = String(b).toLowerCase();
    if (s.includes('dashed')) return 'dashed';
    if (s.includes('dotted')) return 'dotted';
    if (s.includes('double')) return 'double';
    return 'solid';
}

function extractCellStyle(cell, isDate) {
    if (!cell || typeof cell !== 'object' || cell instanceof Date) {
        return isDate ? { isDate: true } : null;
    }

    const style = {};
    if (isDate) style.isDate = true;
    if (cell.backgroundColor) style.backgroundColor = cell.backgroundColor;
    if (cell.textColor || cell.color) style.color = cell.textColor || cell.color;
    if (cell.fontWeight === 'bold' || cell.bold) style.bold = true;
    if (cell.fontStyle === 'italic' || cell.italic) style.italic = true;
    if (cell.fontSize) style.fontSize = cell.fontSize;
    if (cell.fontFamily) style.fontFamily = cell.fontFamily;
    if (cell.align) style.align = cell.align;

    // Borders
    if (cell.borderStyle || cell.borderColor) {
        style.border = `${formatBorderWidth(cell.borderStyle)} ${formatBorderStyle(cell.borderStyle)} ${cell.borderColor || '#000000'}`;
    }
    if (cell.leftBorderStyle || cell.leftBorderColor) {
        style.borderLeft = `${formatBorderWidth(cell.leftBorderStyle)} ${formatBorderStyle(cell.leftBorderStyle)} ${cell.leftBorderColor || '#000000'}`;
    }
    if (cell.rightBorderStyle || cell.rightBorderColor) {
        style.borderRight = `${formatBorderWidth(cell.rightBorderStyle)} ${formatBorderStyle(cell.rightBorderStyle)} ${cell.rightBorderColor || '#000000'}`;
    }
    if (cell.topBorderStyle || cell.topBorderColor) {
        style.borderTop = `${formatBorderWidth(cell.topBorderStyle)} ${formatBorderStyle(cell.topBorderStyle)} ${cell.topBorderColor || '#000000'}`;
    }
    if (cell.bottomBorderStyle || cell.bottomBorderColor) {
        style.borderBottom = `${formatBorderWidth(cell.bottomBorderStyle)} ${formatBorderStyle(cell.bottomBorderStyle)} ${cell.bottomBorderColor || '#000000'}`;
    }

    return Object.keys(style).length > 0 ? style : null;
}

function serializeOdsStyle(styleName, style) {
    let xml = `    <style:style style:name="${escapeXml(styleName)}" style:family="table-cell"`;
    if (style.isDate) {
        xml += ` style:data-style-name="N_DATE"`;
    }
    xml += `>\n`;

    const cellProps = [];
    if (style.backgroundColor) {
        cellProps.push(`fo:background-color="${escapeXml(style.backgroundColor)}"`);
    }
    if (style.border) {
        cellProps.push(`fo:border="${escapeXml(style.border)}"`);
    }
    if (style.borderLeft) {
        cellProps.push(`fo:border-left="${escapeXml(style.borderLeft)}"`);
    }
    if (style.borderRight) {
        cellProps.push(`fo:border-right="${escapeXml(style.borderRight)}"`);
    }
    if (style.borderTop) {
        cellProps.push(`fo:border-top="${escapeXml(style.borderTop)}"`);
    }
    if (style.borderBottom) {
        cellProps.push(`fo:border-bottom="${escapeXml(style.borderBottom)}"`);
    }

    if (cellProps.length > 0) {
        xml += `      <style:table-cell-properties ${cellProps.join(' ')}/>\n`;
    }

    const textProps = [];
    if (style.color) {
        textProps.push(`fo:color="${escapeXml(style.color)}"`);
    }
    if (style.bold) {
        textProps.push(`fo:font-weight="bold"`);
    }
    if (style.italic) {
        textProps.push(`fo:font-style="italic"`);
    }
    if (style.fontSize) {
        const sz = typeof style.fontSize === 'number' ? `${style.fontSize}pt` : style.fontSize;
        textProps.push(`fo:font-size="${escapeXml(sz)}"`);
    }
    if (style.fontFamily) {
        textProps.push(`fo:font-family="${escapeXml(style.fontFamily)}"`);
    }

    if (textProps.length > 0) {
        xml += `      <style:text-properties ${textProps.join(' ')}/>\n`;
    }

    if (style.align) {
        const alignVal = style.align === 'center' ? 'center' : (style.align === 'right' ? 'end' : 'start');
        xml += `      <style:paragraph-properties fo:text-align="${alignVal}"/>\n`;
    }

    xml += `    </style:style>\n`;
    return xml;
}

/**
 * Build content.xml for OpenDocument Spreadsheet.
 */
function buildContentXml(sheets) {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
    xml += `<office:document-content ` +
        `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
        `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
        `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
        `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" ` +
        `office:version="1.2">\n`;

    // 1. Scan and register custom styles
    const stylesMap = new Map();
    let nextStyleId = 1;

    for (const sheet of sheets) {
        for (const row of sheet.rows) {
            if (!Array.isArray(row)) continue;
            for (const cell of row) {
                let val = cell;
                let isDate = false;
                if (cell && typeof cell === 'object') {
                    val = cell.value;
                    if (val instanceof Date || cell.type === Date) isDate = true;
                } else if (cell instanceof Date) {
                    isDate = true;
                }
                const st = extractCellStyle(cell, isDate);
                if (st) {
                    const key = JSON.stringify(st);
                    if (!stylesMap.has(key)) {
                        stylesMap.set(key, { name: `s_${nextStyleId++}`, style: st });
                    }
                }
            }
        }
    }

    // Automatic styles
    xml += `  <office:automatic-styles>\n`;
    xml += `    <number:date-style style:name="N_DATE">\n`;
    xml += `      <number:year number:style="long"/>\n`;
    xml += `      <number:text>-</number:text>\n`;
    xml += `      <number:month number:style="long"/>\n`;
    xml += `      <number:text>-</number:text>\n`;
    xml += `      <number:day number:style="long"/>\n`;
    xml += `    </number:date-style>\n`;

    // Column styles for widths
    const colStylesMap = new Map();
    let nextColStyleId = 1;
    for (const sheet of sheets) {
        if (sheet.columns && Array.isArray(sheet.columns)) {
            sheet.columns.forEach((col) => {
                if (col && typeof col.width === 'number') {
                    const widthCm = (col.width * 0.2).toFixed(2);
                    if (!colStylesMap.has(widthCm)) {
                        colStylesMap.set(widthCm, `co_${nextColStyleId++}`);
                    }
                }
            });
        }
    }

    for (const [widthCm, styleName] of colStylesMap.entries()) {
        xml += `    <style:style style:name="${styleName}" style:family="table-column">\n`;
        xml += `      <style:table-column-properties style:column-width="${widthCm}cm"/>\n`;
        xml += `    </style:style>\n`;
    }

    for (const { name, style } of stylesMap.values()) {
        xml += serializeOdsStyle(name, style);
    }
    xml += `  </office:automatic-styles>\n`;

    // Spreadsheet body
    xml += `  <office:body>\n`;
    xml += `    <office:spreadsheet>\n`;

    for (const sheet of sheets) {
        xml += `      <table:table table:name="${escapeXml(sheet.name)}">\n`;

        if (sheet.columns && Array.isArray(sheet.columns)) {
            sheet.columns.forEach(col => {
                if (col && typeof col.width === 'number') {
                    const widthCm = (col.width * 0.2).toFixed(2);
                    const styleName = colStylesMap.get(widthCm);
                    xml += `        <table:table-column table:style-name="${styleName}"/>\n`;
                } else {
                    xml += `        <table:table-column/>\n`;
                }
            });
        }

        const coveredCells = new Set();

        for (let r = 0; r < sheet.rows.length; r++) {
            const row = sheet.rows[r];
            xml += `        <table:table-row>\n`;

            if (Array.isArray(row)) {
                for (let c = 0; c < row.length; c++) {
                    const cellKey = `${r},${c}`;
                    if (coveredCells.has(cellKey)) {
                        xml += `          <table:covered-table-cell/>\n`;
                        continue;
                    }

                    const cell = row[c];
                    let val = cell;
                    let type = null;
                    let isDate = false;
                    let colSpan = 1;
                    let rowSpan = 1;

                    if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                        val = cell.value;
                        type = cell.type;
                        if (val instanceof Date || type === Date) isDate = true;
                        colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
                        rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);
                    } else if (cell instanceof Date) {
                        isDate = true;
                    }

                    let spanAttr = '';
                    if (colSpan > 1) spanAttr += ` table:number-columns-spanned="${colSpan}"`;
                    if (rowSpan > 1) spanAttr += ` table:number-rows-spanned="${rowSpan}"`;

                    if (colSpan > 1 || rowSpan > 1) {
                        for (let dr = 0; dr < rowSpan; dr++) {
                            for (let dc = 0; dc < colSpan; dc++) {
                                if (dr > 0 || dc > 0) {
                                    coveredCells.add(`${r + dr},${c + dc}`);
                                }
                            }
                        }
                    }

                    // Look up style name
                    let styleAttr = '';
                    const st = extractCellStyle(cell, isDate);
                    if (st) {
                        const styleEntry = stylesMap.get(JSON.stringify(st));
                        if (styleEntry) {
                            styleAttr = ` table:style-name="${styleEntry.name}"`;
                        }
                    }

                    if (val === null || val === undefined || val === '') {
                        xml += `          <table:table-cell${spanAttr}${styleAttr}/>\n`;
                    } else if (isDate) {
                        const d = val instanceof Date ? val : new Date(val);
                        const isoDate = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : '';
                        xml += `          <table:table-cell office:value-type="date" office:date-value="${isoDate}"${spanAttr}${styleAttr}><text:p>${isoDate}</text:p></table:table-cell>\n`;
                    } else if (typeof val === 'number' || type === Number) {
                        const num = Number(val);
                        xml += `          <table:table-cell office:value-type="float" office:value="${num}"${spanAttr}${styleAttr}><text:p>${num}</text:p></table:table-cell>\n`;
                    } else if (typeof val === 'boolean' || type === Boolean) {
                        const boolVal = Boolean(val);
                        xml += `          <table:table-cell office:value-type="boolean" office:boolean-value="${boolVal}"${spanAttr}${styleAttr}><text:p>${boolVal ? 'TRUE' : 'FALSE'}</text:p></table:table-cell>\n`;
                    } else {
                        const strVal = String(val);
                        xml += `          <table:table-cell office:value-type="string"${spanAttr}${styleAttr}><text:p>${escapeXml(strVal)}</text:p></table:table-cell>\n`;
                    }
                }
            }

            xml += `        </table:table-row>\n`;
        }

        xml += `      </table:table>\n`;
    }

    xml += `    </office:spreadsheet>\n`;
    xml += `  </office:body>\n`;
    xml += `</office:document-content>`;

    return xml;
}

/**
 * Universal in-memory ODS binary generator.
 * Produces a complete, valid OpenDocument Spreadsheet (.ods) as a Uint8Array.
 *
 * @param {any} data - 2D grid, array of objects, or TabularJS worksheets
 * @param {Object} [options={}]
 * @returns {Uint8Array}
 */
export function createODS(data, options = {}) {
    const sheets = normalizeWorkbookSheets(data, options);

    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>\n` +
        `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n` +
        `  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n` +
        `  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>\n` +
        `  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>\n` +
        `  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>\n` +
        `</manifest:manifest>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>\n` +
        `<office:document-styles ` +
        `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
        `xmlns:xlink="http://www.w3.org/1999/xlink" ` +
        `office:version="1.2">\n` +
        `  <office:styles>\n` +
        `    <style:default-style style:family="table-cell">\n` +
        `      <style:paragraph-properties fo:text-align="start"/>\n` +
        `    </style:default-style>\n` +
        `  </office:styles>\n` +
        `</office:document-styles>`;

    const metaXml = `<?xml version="1.0" encoding="UTF-8"?>\n` +
        `<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">\n` +
        `  <office:meta>\n` +
        `    <meta:generator>spreadsheet-utils</meta:generator>\n` +
        `    <meta:creation-date>${new Date().toISOString()}</meta:creation-date>\n` +
        `  </office:meta>\n` +
        `</office:document-meta>`;

    const contentXml = buildContentXml(sheets);

    const files = {
        'mimetype': [strToU8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }],
        'META-INF/manifest.xml': strToU8(manifestXml),
        'content.xml': strToU8(contentXml),
        'styles.xml': strToU8(stylesXml),
        'meta.xml': strToU8(metaXml)
    };

    return zipSync(files);
}

/**
 * Export spreadsheet data to OpenDocument Spreadsheet (.ods) format.
 *
 * In Node.js:
 * - Writes to file if `options.filePath` is provided and returns filePath.
 * - Returns Buffer by default.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Promise<Buffer|Uint8Array|string>}
 */
export async function writeODS(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const uint8Array = createODS(data, options);

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

export default writeODS;
