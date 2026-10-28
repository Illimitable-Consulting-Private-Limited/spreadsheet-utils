import fs from 'node:fs';
import { normalizeWorkbookSheets } from './ods.js';

/**
 * BIFF record type codes.
 */
const BIFF = {
    BOF: 0x0809,
    EOF: 0x000A,
    DIMENSIONS: 0x0200,
    BLANK: 0x0201,
    NUMBER: 0x0203,
    BOOLERR: 0x0205,
    STRING: 0x0207,
    BOUNDSHEET: 0x0085,
    SST: 0x00FC,
    LABELSST: 0x00FD,
    FONT: 0x0031,
    XF: 0x00E0,
    STYLE: 0x0293,
    CODEPAGE: 0x0042,
    WINDOW1: 0x003D,
    WINDOW2: 0x023E,
    MERGECELLS: 0x00E5
};

const PALETTE = [
    { idx: 0x08, r: 0x00, g: 0x00, b: 0x00 },
    { idx: 0x09, r: 0xFF, g: 0xFF, b: 0xFF },
    { idx: 0x0A, r: 0xFF, g: 0x00, b: 0x00 },
    { idx: 0x0B, r: 0x00, g: 0xFF, b: 0x00 },
    { idx: 0x0C, r: 0x00, g: 0x00, b: 0xFF },
    { idx: 0x0D, r: 0xFF, g: 0xFF, b: 0x00 },
    { idx: 0x0E, r: 0xFF, g: 0x00, b: 0xFF },
    { idx: 0x0F, r: 0x00, g: 0xFF, b: 0xFF },
    { idx: 0x10, r: 0x80, g: 0x00, b: 0x00 },
    { idx: 0x11, r: 0x00, g: 0x80, b: 0x00 },
    { idx: 0x12, r: 0x00, g: 0x00, b: 0x80 },
    { idx: 0x13, r: 0x80, g: 0x80, b: 0x00 },
    { idx: 0x14, r: 0x80, g: 0x00, b: 0x80 },
    { idx: 0x15, r: 0x00, g: 0x80, b: 0x80 },
    { idx: 0x16, r: 0xC0, g: 0xC0, b: 0xC0 },
    { idx: 0x17, r: 0x80, g: 0x80, b: 0x80 }
];

function hexToColorIndex(hex) {
    if (!hex) return 0x7FFF;
    const clean = String(hex).replace('#', '').trim();
    if (clean.length < 6) return 0x7FFF;
    const r = parseInt(clean.slice(0, 2), 16) || 0;
    const g = parseInt(clean.slice(2, 4), 16) || 0;
    const b = parseInt(clean.slice(4, 6), 16) || 0;

    let bestDist = Infinity;
    let bestIdx = 0x08;
    for (const p of PALETTE) {
        const d = (p.r - r) ** 2 + (p.g - g) ** 2 + (p.b - b) ** 2;
        if (d < bestDist) {
            bestDist = d;
            bestIdx = p.idx;
        }
    }
    return bestIdx;
}

function getBiffBorderStyle(style) {
    if (!style) return 0;
    const s = String(style).toLowerCase();
    if (s.includes('double')) return 6;
    if (s.includes('thick')) return 5;
    if (s.includes('dashed')) return 3;
    if (s.includes('dotted')) return 4;
    if (s.includes('medium')) return 2;
    return 1; // thin
}

/**
 * Helper to build a standard BIFF record:
 * - 2 bytes: type (uint16 LE)
 * - 2 bytes: length (uint16 LE)
 * - length bytes: payload
 */
function createRecord(type, data) {
    const payload = (data instanceof Uint8Array) ? data : new Uint8Array(data || 0);
    const len = payload.length;
    const rec = new Uint8Array(4 + len);
    const view = new DataView(rec.buffer, rec.byteOffset, rec.byteLength);

    view.setUint16(0, type, true);
    view.setUint16(2, len, true);
    rec.set(payload, 4);

    return rec;
}

/**
 * Concatenate multiple Uint8Arrays into one.
 */
function concatBuffers(buffers) {
    let totalLen = 0;
    for (const b of buffers) totalLen += b.length;
    const result = new Uint8Array(totalLen);
    let offset = 0;
    for (const b of buffers) {
        result.set(b, offset);
        offset += b.length;
    }
    return result;
}

function makeFontRecord(height, isBold, colorIndex = 0x7FFF) {
    const name = 'Arial';
    const data = new Uint8Array(14 + 2 + name.length);
    const view = new DataView(data.buffer);

    view.setUint16(0, height, true); // font height in twips (200 = 10pt)
    view.setUint16(2, 0, true);      // options
    view.setUint16(4, colorIndex, true); // color index
    view.setUint16(6, isBold ? 700 : 400, true); // weight
    view.setUint16(8, 0, true);      // escapement
    data[10] = 0; // underline
    data[11] = 0; // family
    data[12] = 0; // charset
    data[13] = 0; // reserved
    data[14] = name.length;
    data[15] = 0x00; // string flags (compressed ASCII)
    for (let i = 0; i < name.length; i++) {
        data[16 + i] = name.charCodeAt(i);
    }
    return createRecord(BIFF.FONT, data);
}

function createCellXFRecord(fontIndex, align, border, bgIndex) {
    const xfData = new Uint8Array(20);
    const view = new DataView(xfData.buffer);

    view.setUint16(0, fontIndex, true);
    view.setUint16(2, 0, true); // format 0
    view.setUint16(4, 0x0001, true); // cell XF

    // Byte 6: horizontal align
    let hAlign = 0;
    if (align === 'left') hAlign = 0;
    else if (align === 'center') hAlign = 1;
    else if (align === 'right') hAlign = 2;
    xfData[6] = hAlign & 0x07;

    // Bytes 10-13 (brdbkg1): Border styles
    if (border) {
        let bVal = 0;
        bVal |= (border.left & 0x0F);
        bVal |= (border.right & 0x0F) << 4;
        bVal |= (border.top & 0x0F) << 8;
        bVal |= (border.bottom & 0x0F) << 12;
        view.setUint32(10, bVal, true);
    }

    // Fill pattern (bytes 14-17: brdbkg2) & colors (bytes 18-19: brdbkg3)
    if (bgIndex !== undefined && bgIndex !== null && bgIndex !== 0x7FFF) {
        view.setUint32(14, 1 << 26, true);
        view.setUint16(18, bgIndex & 0x7F, true);
    }

    return createRecord(BIFF.XF, xfData);
}

/**
 * Build standard default font records for BIFF8.
 */
function buildFontRecords() {
    return [
        makeFontRecord(200, false), // Font 0: Arial 10pt regular
        makeFontRecord(200, true),  // Font 1: Arial 10pt bold
        makeFontRecord(200, false), // Font 2: Arial 10pt regular
        makeFontRecord(200, false)  // Font 3: Arial 10pt regular
    ];
}

/**
 * Build standard default XF records for BIFF8.
 */
function buildXFRecords() {
    const records = [];

    // First 15 are standard style XFs
    for (let i = 0; i < 15; i++) {
        const xfData = new Uint8Array(20);
        const view = new DataView(xfData.buffer);
        view.setUint16(0, 0, true); // font 0
        view.setUint16(2, 0, true); // format 0
        view.setUint16(4, 0xFFF0, true); // style XF flag
        records.push(createRecord(BIFF.XF, xfData));
    }

    // XF 15: Default cell format (regular)
    {
        const xfData = new Uint8Array(20);
        const view = new DataView(xfData.buffer);
        view.setUint16(0, 0, true); // font 0 (regular)
        view.setUint16(2, 0, true); // format 0
        view.setUint16(4, 0x0001, true); // cell XF
        records.push(createRecord(BIFF.XF, xfData));
    }

    // XF 16: Bold cell format
    {
        const xfData = new Uint8Array(20);
        const view = new DataView(xfData.buffer);
        view.setUint16(0, 1, true); // font 1 (bold)
        view.setUint16(2, 0, true); // format 0
        view.setUint16(4, 0x0001, true); // cell XF
        records.push(createRecord(BIFF.XF, xfData));
    }

    return records;
}

/**
 * Package a Workbook stream into a standard Microsoft Compound File Binary (CFB / OLE2) container.
 *
 * @param {Uint8Array} workbookStream
 * @returns {Uint8Array}
 */
export function packageCFB(workbookStream) {
    const SECTOR_SIZE = 512;
    const streamSize = workbookStream.length;
    const numStreamSectors = Math.max(1, Math.ceil(streamSize / SECTOR_SIZE));

    // Sector allocation layout:
    // Sector 0: Directory Sector
    // Sectors 1 .. numStreamSectors: Workbook Stream
    // Sector numStreamSectors + 1: FAT Sector
    const dirSector = 0;
    const fatSector = numStreamSectors + 1;
    const totalSectors = numStreamSectors + 2;

    const totalFileSize = 512 + (totalSectors * SECTOR_SIZE);
    const file = new Uint8Array(totalFileSize);
    const view = new DataView(file.buffer);

    // 1. CFB Header (512 bytes at offset 0)
    // Signature
    const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for (let i = 0; i < 8; i++) file[i] = sig[i];

    view.setUint16(0x18, 0x003E, true); // minor version
    view.setUint16(0x1A, 0x0003, true); // major version 3 (512-byte sectors)
    view.setUint16(0x1C, 0xFFFE, true); // byte order marker (little endian)
    view.setUint16(0x1E, 0x0009, true); // sector shift (9 -> 512)
    view.setUint16(0x20, 0x0006, true); // mini sector shift (6 -> 64)
    view.setUint32(0x2C, 1, true);      // number of FAT sectors = 1
    view.setUint32(0x30, dirSector, true); // first directory sector = 0
    view.setUint32(0x38, 0x00001000, true); // mini stream cutoff = 4096
    view.setUint32(0x3C, 0xFFFFFFFE, true); // first MiniFAT sector = ENDOFCHAIN
    view.setUint32(0x40, 0, true);          // number of MiniFAT sectors = 0
    view.setUint32(0x44, 0xFFFFFFFE, true); // first DIFAT sector = ENDOFCHAIN
    view.setUint32(0x48, 0, true);          // number of DIFAT sectors = 0

    // DIFAT entries (first 109 entries starting at offset 0x4C)
    view.setUint32(0x4C, fatSector, true); // First entry points to the FAT sector
    for (let i = 1; i < 109; i++) {
        view.setUint32(0x4C + (i * 4), 0xFFFFFFFF, true); // FREESECT
    }

    // 2. Directory Sector (Sector 0, file offset 512)
    const dirOffset = 512;

    // Entry 0: Root Entry (128 bytes)
    const rootName = 'Root Entry';
    for (let i = 0; i < rootName.length; i++) {
        view.setUint16(dirOffset + (i * 2), rootName.charCodeAt(i), true);
    }
    view.setUint16(dirOffset + 0x40, (rootName.length + 1) * 2, true); // name length
    file[dirOffset + 0x42] = 5; // Entry type: ROOT
    file[dirOffset + 0x43] = 1; // Color: Black
    view.setUint32(dirOffset + 0x44, 0xFFFFFFFF, true); // left sibling: NOSTREAM
    view.setUint32(dirOffset + 0x48, 0xFFFFFFFF, true); // right sibling: NOSTREAM
    view.setUint32(dirOffset + 0x4C, 1, true);          // child: Entry 1 (Workbook)
    view.setUint32(dirOffset + 0x74, 0xFFFFFFFE, true); // start sector: ENDOFCHAIN
    view.setUint32(dirOffset + 0x78, 0, true);          // size: 0

    // Entry 1: Workbook Stream (128 bytes)
    const wbEntryOffset = dirOffset + 128;
    const wbName = 'Workbook';
    for (let i = 0; i < wbName.length; i++) {
        view.setUint16(wbEntryOffset + (i * 2), wbName.charCodeAt(i), true);
    }
    view.setUint16(wbEntryOffset + 0x40, (wbName.length + 1) * 2, true); // name length
    file[wbEntryOffset + 0x42] = 2; // Entry type: STREAM
    file[wbEntryOffset + 0x43] = 1; // Color: Black
    view.setUint32(wbEntryOffset + 0x44, 0xFFFFFFFF, true); // left sibling: NOSTREAM
    view.setUint32(wbEntryOffset + 0x48, 0xFFFFFFFF, true); // right sibling: NOSTREAM
    view.setUint32(wbEntryOffset + 0x4C, 0xFFFFFFFF, true); // child: NOSTREAM
    view.setUint32(wbEntryOffset + 0x74, 1, true);          // starting sector: 1
    view.setUint32(wbEntryOffset + 0x78, streamSize, true); // stream size in bytes

    // Entries 2 and 3: Empty (128 bytes each, remaining bytes already 0)

    // 3. Workbook Stream Sectors (Sectors 1 .. numStreamSectors)
    const streamStartOffset = 512 + (1 * SECTOR_SIZE);
    file.set(workbookStream, streamStartOffset);

    // 4. FAT Sector (Sector fatSector)
    const fatOffset = 512 + (fatSector * SECTOR_SIZE);
    // Sector 0 (Dir) -> ENDOFCHAIN
    view.setUint32(fatOffset + 0, 0xFFFFFFFE, true);
    // Sectors 1 .. numStreamSectors -> Chain links
    for (let i = 1; i < numStreamSectors; i++) {
        view.setUint32(fatOffset + (i * 4), i + 1, true);
    }
    // Last sector of stream -> ENDOFCHAIN
    view.setUint32(fatOffset + (numStreamSectors * 4), 0xFFFFFFFE, true);
    // FAT sector itself -> FATSECT
    view.setUint32(fatOffset + (fatSector * 4), 0xFFFFFFFD, true);
    // Unused entries -> FREESECT (0xFFFFFFFF)
    for (let i = fatSector + 1; i < 128; i++) {
        view.setUint32(fatOffset + (i * 4), 0xFFFFFFFF, true);
    }

    return file;
}

/**
 * Universal in-memory Microsoft Excel 97-2004 (.xls) BIFF8 binary generator.
 * Produces a complete, valid binary .xls spreadsheet as a Uint8Array.
 *
 * @param {any} data - 2D grid, array of objects, or TabularJS worksheets
 * @param {Object} [options={}]
 * @returns {Uint8Array}
 */
export function createXLS(data, options = {}) {
    const sheets = normalizeWorkbookSheets(data, options);

    // 1. Collect all unique strings across the workbook for the SST
    const sstMap = new Map();
    const sstList = [];
    let totalStringRefs = 0;

    const getStringIndex = (str) => {
        totalStringRefs++;
        if (sstMap.has(str)) {
            return sstMap.get(str);
        }
        const idx = sstList.length;
        sstMap.set(str, idx);
        sstList.push(str);
        return idx;
    };

    // 2. Pre-scan sheets to collect strings into SST, fonts, and XFs
    const fontRecords = buildFontRecords();
    const xfRecords = buildXFRecords();
    const fontCache = new Map();
    const xfCache = new Map();

    const getFontIndex = (isBold, textColor) => {
        const colorIdx = textColor ? hexToColorIndex(textColor) : 0x7FFF;
        if (!textColor || colorIdx === 0x7FFF) {
            return isBold ? 1 : 0;
        }
        const key = `${isBold}:${colorIdx}`;
        if (fontCache.has(key)) return fontCache.get(key);
        const idx = fontRecords.length;
        fontRecords.push(makeFontRecord(200, isBold, colorIdx));
        fontCache.set(key, idx);
        return idx;
    };

    const getXFIndex = (cell) => {
        if (!cell || typeof cell !== 'object' || cell instanceof Date) {
            return 15;
        }
        const isBold = Boolean(cell.fontWeight === 'bold' || cell.bold);
        const textColor = cell.textColor || cell.color;
        const align = cell.align;
        const backgroundColor = cell.backgroundColor;
        const bgIdx = backgroundColor ? hexToColorIndex(backgroundColor) : 0x7FFF;

        let border = null;
        if (cell.borderStyle || cell.borderColor || cell.leftBorderStyle || cell.rightBorderStyle || cell.topBorderStyle || cell.bottomBorderStyle) {
            border = {
                left: getBiffBorderStyle(cell.leftBorderStyle || cell.borderStyle),
                right: getBiffBorderStyle(cell.rightBorderStyle || cell.borderStyle),
                top: getBiffBorderStyle(cell.topBorderStyle || cell.borderStyle),
                bottom: getBiffBorderStyle(cell.bottomBorderStyle || cell.borderStyle)
            };
        }

        if (!textColor && !align && !backgroundColor && !border) {
            return isBold ? 16 : 15;
        }

        const fontIdx = getFontIndex(isBold, textColor);
        const key = `${fontIdx}:${align || ''}:${bgIdx}:${border ? JSON.stringify(border) : ''}`;
        if (xfCache.has(key)) return xfCache.get(key);

        const xfIdx = xfRecords.length;
        xfRecords.push(createCellXFRecord(fontIdx, align, border, bgIdx));
        xfCache.set(key, xfIdx);
        return xfIdx;
    };

    for (const sheet of sheets) {
        for (const row of sheet.rows) {
            if (!Array.isArray(row)) continue;
            for (const cell of row) {
                let val = cell;
                let type = null;
                if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                    val = cell.value;
                    type = cell.type;
                    getXFIndex(cell); // pre-register custom font/XF
                }
                if (val !== null && val !== undefined && val !== '') {
                    if (val instanceof Date || type === Date) {
                        const d = val instanceof Date ? val : new Date(val);
                        const dateStr = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : String(val);
                        getStringIndex(dateStr);
                    } else if (typeof val !== 'number' && type !== Number && typeof val !== 'boolean' && type !== Boolean) {
                        getStringIndex(String(val));
                    }
                }
            }
        }
    }

    // 3. Build Globals Records
    const globalsRecords = [];

    // BOF Globals
    const bofGlobalsData = new Uint8Array([
        0x00, 0x06, // BIFF8
        0x05, 0x00, // Workbook Globals
        0xBB, 0x0D, // build
        0xCC, 0x07, // year
        0x00, 0x00, 0x00, 0x00,
        0x06, 0x00, 0x00, 0x00
    ]);
    globalsRecords.push(createRecord(BIFF.BOF, bofGlobalsData));

    // CODEPAGE (1252 ANSI)
    globalsRecords.push(createRecord(BIFF.CODEPAGE, new Uint8Array([0xE4, 0x04])));

    // WINDOW1
    globalsRecords.push(createRecord(BIFF.WINDOW1, new Uint8Array([
        0x00, 0x00, 0x00, 0x00, 0x00, 0x3F, 0x00, 0x24,
        0x38, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
        0x58, 0x02
    ])));

    // FONTS and XFS
    globalsRecords.push(...fontRecords);
    globalsRecords.push(...xfRecords);

    // Placeholders for BOUNDSHEET records (we will patch BOF offsets later)
    const boundsheetRecords = [];
    for (const sheet of sheets) {
        const name = sheet.name || 'Sheet1';
        const data = new Uint8Array(8 + name.length);
        const view = new DataView(data.buffer);
        view.setUint32(0, 0, true); // BOF offset (placeholder)
        data[4] = 0x00; // Visible
        data[5] = 0x00; // Worksheet
        data[6] = name.length;
        data[7] = 0x00; // Compressed ASCII
        for (let i = 0; i < name.length; i++) {
            data[8 + i] = name.charCodeAt(i);
        }
        const rec = createRecord(BIFF.BOUNDSHEET, data);
        boundsheetRecords.push(rec);
        globalsRecords.push(rec);
    }

    // SST (Shared String Table)
    let sstByteLength = 8;
    for (const str of sstList) {
        sstByteLength += 3 + str.length;
    }
    const sstData = new Uint8Array(sstByteLength);
    const sstView = new DataView(sstData.buffer);
    sstView.setUint32(0, totalStringRefs, true);
    sstView.setUint32(4, sstList.length, true);
    let sstOffset = 8;
    for (const str of sstList) {
        sstView.setUint16(sstOffset, str.length, true);
        sstData[sstOffset + 2] = 0x00; // 8-bit compressed characters
        for (let i = 0; i < str.length; i++) {
            sstData[sstOffset + 3 + i] = str.charCodeAt(i) & 0xFF;
        }
        sstOffset += 3 + str.length;
    }
    globalsRecords.push(createRecord(BIFF.SST, sstData));

    // EOF Globals
    globalsRecords.push(createRecord(BIFF.EOF, new Uint8Array(0)));

    // Calculate initial globals length
    let currentStreamOffset = globalsRecords.reduce((sum, r) => sum + r.length, 0);

    // 4. Build Worksheet Records
    const worksheetStreams = [];
    const sheetBofOffsets = [];

    for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
        sheetBofOffsets.push(currentStreamOffset);
        const sheet = sheets[sIdx];
        const sheetRecords = [];

        // Worksheet BOF
        const bofSheetData = new Uint8Array([
            0x00, 0x06, // BIFF8
            0x10, 0x00, // Worksheet
            0xBB, 0x0D,
            0xCC, 0x07,
            0x00, 0x00, 0x00, 0x00,
            0x06, 0x00, 0x00, 0x00
        ]);
        sheetRecords.push(createRecord(BIFF.BOF, bofSheetData));

        // Dimensions
        let maxRow = 0;
        let maxCol = 0;
        for (let r = 0; r < sheet.rows.length; r++) {
            const row = sheet.rows[r];
            if (Array.isArray(row)) {
                if (r > maxRow) maxRow = r;
                if (row.length > maxCol) maxCol = row.length;
            }
        }

        const dimData = new Uint8Array(14);
        const dimView = new DataView(dimData.buffer);
        dimView.setUint32(0, 0, true);
        dimView.setUint32(4, Math.max(1, maxRow + 1), true);
        dimView.setUint16(8, 0, true);
        dimView.setUint16(10, Math.max(1, maxCol), true);
        sheetRecords.push(createRecord(BIFF.DIMENSIONS, dimData));

        // Window2 (gridlines, etc.)
        const win2Data = new Uint8Array(18);
        new DataView(win2Data.buffer).setUint16(0, 0x02B6, true);
        sheetRecords.push(createRecord(BIFF.WINDOW2, win2Data));

        // Cell Records
        for (let r = 0; r < sheet.rows.length; r++) {
            const row = sheet.rows[r];
            if (!Array.isArray(row)) continue;

            for (let c = 0; c < row.length; c++) {
                const cell = row[c];
                let val = cell;
                let type = null;
                let isBold = false;

                if (cell && typeof cell === 'object' && !(cell instanceof Date)) {
                    val = cell.value;
                    type = cell.type;
                    if (cell.fontWeight === 'bold' || cell.bold) {
                        isBold = true;
                    }
                }

                if (val === null || val === undefined || val === '') {
                    continue;
                }

                const xfIndex = getXFIndex(cell);

                if (typeof val === 'number' || type === Number) {
                    // NUMBER record (0x0203)
                    const numData = new Uint8Array(14);
                    const view = new DataView(numData.buffer);
                    view.setUint16(0, r, true);
                    view.setUint16(2, c, true);
                    view.setUint16(4, xfIndex, true);
                    view.setFloat64(6, Number(val), true);
                    sheetRecords.push(createRecord(BIFF.NUMBER, numData));
                } else if (typeof val === 'boolean' || type === Boolean) {
                    // BOOLERR record (0x0205)
                    const boolData = new Uint8Array(8);
                    const view = new DataView(boolData.buffer);
                    view.setUint16(0, r, true);
                    view.setUint16(2, c, true);
                    view.setUint16(4, xfIndex, true);
                    boolData[6] = val ? 1 : 0;
                    boolData[7] = 0; // not error
                    sheetRecords.push(createRecord(BIFF.BOOLERR, boolData));
                } else if (val instanceof Date || type === Date) {
                    // Format date as string into SST
                    const d = val instanceof Date ? val : new Date(val);
                    const dateStr = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : String(val);
                    const sstIdx = sstMap.get(dateStr) !== undefined ? sstMap.get(dateStr) : getStringIndex(dateStr);
                    const labelData = new Uint8Array(10);
                    const view = new DataView(labelData.buffer);
                    view.setUint16(0, r, true);
                    view.setUint16(2, c, true);
                    view.setUint16(4, xfIndex, true);
                    view.setUint32(6, sstIdx, true);
                    sheetRecords.push(createRecord(BIFF.LABELSST, labelData));
                } else {
                    // LABELSST record (0x00FD)
                    const strVal = String(val);
                    const sstIdx = sstMap.get(strVal) !== undefined ? sstMap.get(strVal) : getStringIndex(strVal);
                    const labelData = new Uint8Array(10);
                    const view = new DataView(labelData.buffer);
                    view.setUint16(0, r, true);
                    view.setUint16(2, c, true);
                    view.setUint16(4, xfIndex, true);
                    view.setUint32(6, sstIdx, true);
                    sheetRecords.push(createRecord(BIFF.LABELSST, labelData));
                }
            }
        }

        // Merged cells (BIFF.MERGECELLS 0x00E5)
        const mergedCellsList = [];
        for (let r = 0; r < sheet.rows.length; r++) {
            const row = sheet.rows[r];
            if (!Array.isArray(row)) continue;
            for (let c = 0; c < row.length; c++) {
                const cell = row[c];
                if (cell && typeof cell === 'object') {
                    const colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
                    const rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);
                    if (colSpan > 1 || rowSpan > 1) {
                        mergedCellsList.push({
                            rowFirst: r,
                            rowLast: r + rowSpan - 1,
                            colFirst: c,
                            colLast: c + colSpan - 1
                        });
                    }
                }
            }
        }

        if (mergedCellsList.length > 0) {
            const count = mergedCellsList.length;
            const mergeData = new Uint8Array(2 + count * 8);
            const view = new DataView(mergeData.buffer);
            view.setUint16(0, count, true);
            for (let i = 0; i < count; i++) {
                const m = mergedCellsList[i];
                const offset = 2 + i * 8;
                view.setUint16(offset, m.rowFirst, true);
                view.setUint16(offset + 2, m.rowLast, true);
                view.setUint16(offset + 4, m.colFirst, true);
                view.setUint16(offset + 6, m.colLast, true);
            }
            sheetRecords.push(createRecord(BIFF.MERGECELLS, mergeData));
        }

        // EOF Sheet
        sheetRecords.push(createRecord(BIFF.EOF, new Uint8Array(0)));

        const sheetBytes = concatBuffers(sheetRecords);
        worksheetStreams.push(sheetBytes);
        currentStreamOffset += sheetBytes.length;
    }

    // 5. Patch BOUNDSHEET records with exact BOF offsets
    for (let sIdx = 0; sIdx < boundsheetRecords.length; sIdx++) {
        const bsRec = boundsheetRecords[sIdx];
        const offset = sheetBofOffsets[sIdx];
        // In record buffer, payload starts at byte 4
        new DataView(bsRec.buffer, bsRec.byteOffset + 4).setUint32(0, offset, true);
    }

    // 6. Concatenate full Workbook stream
    const workbookStream = concatBuffers([
        ...globalsRecords,
        ...worksheetStreams
    ]);

    // 7. Package Workbook stream into CFB OLE2 container
    return packageCFB(workbookStream);
}

/**
 * Export spreadsheet data to Microsoft Excel 97-2004 (.xls) BIFF8 binary format.
 *
 * In Node.js:
 * - Writes to file if `options.filePath` is provided and returns filePath.
 * - Returns Buffer by default.
 *
 * @param {any} data
 * @param {Object} [options={}]
 * @returns {Promise<Buffer|Uint8Array|string>}
 */
export async function writeXLS(data, options = {}) {
    if (typeof options === 'string') {
        options = { filePath: options };
    }

    const uint8Array = createXLS(data, options);

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

export default writeXLS;
