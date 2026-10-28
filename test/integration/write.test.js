import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    writeSpreadsheet,
    writeXlsxFile,
    getSheetData
} from '../../src/write/index.js';

import {
    readSpreadsheet,
    registerPlugin
} from '../../src/index.js';

import { textPlugin } from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tmpDir = path.resolve(__dirname, '../../tmp-test');

describe('Styled XLSX Writer (Integration Tests)', () => {

    before(() => {
        if (!fs.existsSync(tmpDir)) {
            fs.mkdirSync(tmpDir, { recursive: true });
        }
        registerPlugin(textPlugin);
    });

    after(() => {
        if (fs.existsSync(tmpDir)) {
            fs.rmSync(tmpDir, { recursive: true, force: true });
        }
    });

    it('should generate a valid XLSX buffer from a 2D array', async () => {
        const data = [
            [
                { value: 'Name', fontWeight: 'bold' },
                { value: 'Score', fontWeight: 'bold' }
            ],
            [
                { value: 'Alice', type: String },
                { value: 95, type: Number }
            ],
            [
                { value: 'Bob', type: String },
                { value: 87.5, type: Number }
            ]
        ];

        const buffer = await writeSpreadsheet(data, { buffer: true });
        expect(buffer).to.be.an.instanceOf(Uint8Array);
        expect(buffer.length).to.be.greaterThan(0);

        // Read it back with our reader
        const rows = await readSpreadsheet(buffer);
        expect(rows).to.deep.equal([
            ['Name', 'Score'],
            ['Alice', 95],
            ['Bob', 87.5]
        ]);
    });

    it('should write XLSX to disk using filePath option and read back', async () => {
        const filePath = path.join(tmpDir, 'styled-test.xlsx');
        const data = [
            [
                { value: 'Product', fontWeight: 'bold' },
                { value: 'Price', fontWeight: 'bold' }
            ],
            [
                { value: 'Gadget', type: String },
                { value: 19.99, type: Number }
            ]
        ];

        const savedPath = await writeSpreadsheet(data, { filePath });
        expect(savedPath).to.equal(filePath);
        expect(fs.existsSync(filePath)).to.be.true;

        const readRows = await readSpreadsheet(filePath);
        expect(readRows).to.deep.equal([
            ['Product', 'Price'],
            ['Gadget', 19.99]
        ]);
    });

    it('should support object array with schema mapping', async () => {
        const users = [
            { id: 101, name: 'Charlie', active: true, balance: 1250.50 },
            { id: 102, name: 'Dana', active: false, balance: 450.00 }
        ];

        const schema = {
            id: { column: 'User ID', type: Number },
            name: { column: 'Full Name', type: String },
            balance: { column: 'Account Balance', type: Number }
        };

        const buffer = await writeSpreadsheet(users, { schema, buffer: true });
        const rows = await readSpreadsheet(buffer);

        expect(rows).to.deep.equal([
            ['User ID', 'Full Name', 'Account Balance'],
            [101, 'Charlie', 1250.5],
            [102, 'Dana', 450]
        ]);
    });

    it('should support multi-sheet workbooks', async () => {
        const sheets = [
            {
                sheet: 'Users',
                data: [
                    [{ value: 'ID' }, { value: 'Name' }],
                    [{ value: 1 }, { value: 'Alice' }]
                ]
            },
            {
                sheet: 'Scores',
                data: [
                    [{ value: 'Subject' }, { value: 'Grade' }],
                    [{ value: 'Math' }, { value: 'A' }]
                ]
            }
        ];

        const buffer = await writeSpreadsheet(sheets, { buffer: true });
        const allSheets = await readSpreadsheet(buffer, { allSheets: true });

        expect(allSheets.length).to.equal(2);
        expect(allSheets[0].name).to.equal('Users');
        expect(allSheets[0].data).to.deep.equal([['ID', 'Name'], [1, 'Alice']]);
        expect(allSheets[1].name).to.equal('Scores');
        expect(allSheets[1].data).to.deep.equal([['Subject', 'Grade'], ['Math', 'A']]);
    });

    it('should perform full roundtrip: read CSV -> write styled XLSX -> read XLSX', async () => {
        const csvContent = 'SKU,Item,Qty,Price\nA1,Widget,10,5.99\nB2,Gadget,5,12.50';
        const csvBuffer = new TextEncoder().encode(csvContent);

        // 1. Read CSV
        const rawRows = await readSpreadsheet(csvBuffer, { extension: 'csv' });
        expect(rawRows.length).to.equal(3);

        // 2. Transform into styled cells
        const styledData = rawRows.map((row, rowIndex) => {
            return row.map((cell) => {
                if (rowIndex === 0) {
                    return { value: cell, fontWeight: 'bold' };
                }
                const num = Number(cell);
                if (!isNaN(num) && cell.trim() !== '') {
                    return { value: num, type: Number };
                }
                return { value: cell, type: String };
            });
        });

        // 3. Write styled XLSX
        const xlsxPath = path.join(tmpDir, 'roundtrip.xlsx');
        await writeSpreadsheet(styledData, { filePath: xlsxPath });

        // 4. Read back XLSX
        const result = await readSpreadsheet(xlsxPath);
        expect(result).to.deep.equal([
            ['SKU', 'Item', 'Qty', 'Price'],
            ['A1', 'Widget', 10, 5.99],
            ['B2', 'Gadget', 5, 12.5]
        ]);
    });

    it('should support writeXlsxFile drop-in API with toBuffer() and toFile()', async () => {
        const data = [
            [{ value: 'Direct API' }]
        ];

        const writer = writeXlsxFile(data);
        const buf = await writer.toBuffer();
        expect(buf).to.be.an.instanceOf(Buffer);

        const outPath = path.join(tmpDir, 'direct.xlsx');
        await writer.toFile(outPath);
        expect(fs.existsSync(outPath)).to.be.true;

        const rows = await readSpreadsheet(outPath);
        expect(rows).to.deep.equal([['Direct API']]);
    });
});
