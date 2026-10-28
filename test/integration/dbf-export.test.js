import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    writeSpreadsheet,
    writeDBF,
    createDBF,
    normalizeSchemaToDbfFields
} from '../../src/write/index.js';

import {
    readSpreadsheet,
    registerPlugin,
    unregisterPlugin,
    isPluginRegistered,
    getWriterForExtension
} from '../../src/index.js';

import {
    dbfPlugin,
    legacyPlugin
} from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tmpDir = path.resolve(__dirname, '../../tmp-test-dbf');

describe('DBF Export Plugin & Schema-Based Writer', () => {

    before(() => {
        if (!fs.existsSync(tmpDir)) {
            fs.mkdirSync(tmpDir, { recursive: true });
        }
    });

    after(() => {
        if (fs.existsSync(tmpDir)) {
            fs.rmSync(tmpDir, { recursive: true, force: true });
        }
        unregisterPlugin();
    });

    describe('Plugin registration & missing plugin diagnostics', () => {
        beforeEach(() => {
            unregisterPlugin();
        });

        it('should throw helpful missing-plugin error when writing DBF without plugin registered', async () => {
            const data = [
                { id: 1, name: 'Alice' }
            ];

            try {
                await writeSpreadsheet(data, { filePath: path.join(tmpDir, 'missing.dbf') });
                expect.fail('Should have thrown missing plugin error');
            } catch (err) {
                expect(err.code).to.equal('ERR_PLUGIN_REQUIRED');
                expect(err.format).to.equal('dbf');
                expect(err.message).to.include('legacyPlugin');
            }
        });

        it('should enable DBF export after registering dbfPlugin', async () => {
            registerPlugin(dbfPlugin);
            expect(isPluginRegistered('dbfPlugin')).to.be.true;
            expect(getWriterForExtension('dbf')).to.be.a('function');

            const filePath = path.join(tmpDir, 'dbf-plugin-test.dbf');
            const data = [
                { id: 1, name: 'Alice' },
                { id: 2, name: 'Bob' }
            ];

            const result = await writeSpreadsheet(data, { filePath });
            expect(result).to.equal(filePath);
            expect(fs.existsSync(filePath)).to.be.true;

            // Read it back
            const rows = await readSpreadsheet(filePath);
            expect(rows.length).to.equal(3); // header + 2 rows
            expect(rows[0]).to.deep.equal(['ID', 'NAME']);
            expect(rows[1]).to.deep.equal([1, 'Alice']);
            expect(rows[2]).to.deep.equal([2, 'Bob']);
        });

        it('should enable DBF export after registering legacyPlugin', async () => {
            registerPlugin(legacyPlugin);
            expect(isPluginRegistered('legacyPlugin')).to.be.true;
            expect(getWriterForExtension('dbf')).to.be.a('function');

            const filePath = path.join(tmpDir, 'legacy-plugin-test.dbf');
            const data = [
                { code: 'A1', score: 99.5 }
            ];

            await writeSpreadsheet(data, { filePath });
            expect(fs.existsSync(filePath)).to.be.true;

            const rows = await readSpreadsheet(filePath);
            expect(rows[0]).to.deep.equal(['CODE', 'SCORE']);
            expect(rows[1]).to.deep.equal(['A1', 99.5]);
        });

        it('should disable DBF export after unregistering plugin', async () => {
            registerPlugin(dbfPlugin);
            expect(isPluginRegistered('dbfPlugin')).to.be.true;

            unregisterPlugin('dbfPlugin');
            expect(isPluginRegistered('dbfPlugin')).to.be.false;
            expect(getWriterForExtension('dbf')).to.be.null;

            try {
                await writeSpreadsheet([{ a: 1 }], { filePath: path.join(tmpDir, 'unregistered.dbf') });
                expect.fail('Should have thrown missing plugin error');
            } catch (err) {
                expect(err.code).to.equal('ERR_PLUGIN_REQUIRED');
            }
        });
    });

    describe('Direct writeDBF and createDBF API', () => {
        it('should generate a valid DBF Uint8Array with createDBF', () => {
            const data = [
                { id: 10, title: 'Item 1' },
                { id: 20, title: 'Item 2' }
            ];

            const uint8 = createDBF(data);
            expect(uint8).to.be.an.instanceOf(Uint8Array);
            expect(uint8.length).to.be.greaterThan(32);
            // Check dBase III signature
            expect(uint8[0]).to.equal(0x03);
            // Check EOF marker
            expect(uint8[uint8.length - 1]).to.equal(0x1A);
        });

        it('should write DBF to Buffer using writeDBF without filePath', async () => {
            const data = [
                { id: 1, val: 'Direct' }
            ];

            const buf = await writeDBF(data);
            expect(buf).to.be.an.instanceOf(Buffer);

            // Read back with readSpreadsheet (register legacyPlugin first)
            registerPlugin(legacyPlugin);
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows[1]).to.deep.equal([1, 'Direct']);
        });

        it('should write DBF to file path using writeDBF', async () => {
            const filePath = path.join(tmpDir, 'direct-write.dbf');
            const data = [
                { id: 42, label: 'Answer' }
            ];

            const saved = await writeDBF(data, { filePath });
            expect(saved).to.equal(filePath);
            expect(fs.existsSync(filePath)).to.be.true;

            registerPlugin(dbfPlugin);
            const rows = await readSpreadsheet(filePath);
            expect(rows[1]).to.deep.equal([42, 'Answer']);
        });
    });

    describe('Schema-based writing (write-excel-file extended style)', () => {
        before(() => {
            registerPlugin(dbfPlugin);
        });

        it('should support object schema with types, sizes, decimals, and custom accessors', async () => {
            const employees = [
                { id: 1001, firstName: 'Ada', lastName: 'Lovelace', salary: 125000.75, active: true, hireDate: new Date('2020-05-15') },
                { id: 1002, firstName: 'Alan', lastName: 'Turing', salary: 135000.50, active: false, hireDate: new Date('2021-09-01') }
            ];

            const schema = {
                id: {
                    column: 'EMP_ID',
                    type: Number,
                    size: 6,
                    decimalPlaces: 0
                },
                fullName: {
                    column: 'NAME',
                    type: String,
                    size: 25,
                    value: row => `${row.firstName} ${row.lastName}`
                },
                salary: {
                    column: 'SALARY',
                    type: Number,
                    size: 12,
                    decimalPlaces: 2
                },
                active: {
                    column: 'ACTIVE',
                    type: Boolean
                },
                hireDate: {
                    column: 'HIRED',
                    type: Date
                }
            };

            const filePath = path.join(tmpDir, 'employees.dbf');
            await writeSpreadsheet(employees, { filePath, schema });

            const rows = await readSpreadsheet(filePath);
            expect(rows).to.deep.equal([
                ['EMP_ID', 'NAME', 'SALARY', 'ACTIVE', 'HIRED'],
                [1001, 'Ada Lovelace', 125000.75, true, '2020-05-15'],
                [1002, 'Alan Turing', 135000.5, false, '2021-09-01']
            ]);
        });

        it('should extract decimal places from format string in schema', async () => {
            const data = [
                { id: 1, rate: 3.14159, price: 19.99 }
            ];

            const schema = {
                id: { column: 'ID', type: Number },
                rate: { column: 'RATE', type: Number, format: '0.0000' },
                price: { column: 'PRICE', type: Number, format: '#,##0.00' }
            };

            const buf = await writeSpreadsheet(data, { schema, format: 'dbf', buffer: true });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            expect(rows[0]).to.deep.equal(['ID', 'RATE', 'PRICE']);
            expect(rows[1][0]).to.equal(1);
            expect(rows[1][1]).to.equal(3.1416); // 4 decimals
            expect(rows[1][2]).to.equal(19.99);  // 2 decimals
        });

        it('should support array-based schema definitions', async () => {
            const products = [
                { sku: 'WIDG-1', title: 'Blue Widget', cost: 12.5 },
                { sku: 'GADG-2', title: 'Red Gadget', cost: 45.0 }
            ];

            const schema = [
                { column: 'SKU', prop: 'sku', type: String, size: 10 },
                { column: 'PRODUCT', prop: 'title', type: String, size: 20 },
                { column: 'COST', prop: 'cost', type: Number, size: 8, decimalPlaces: 2 }
            ];

            const filePath = path.join(tmpDir, 'products.dbf');
            await writeSpreadsheet(products, { filePath, schema });

            const rows = await readSpreadsheet(filePath);
            expect(rows).to.deep.equal([
                ['SKU', 'PRODUCT', 'COST'],
                ['WIDG-1', 'Blue Widget', 12.5],
                ['GADG-2', 'Red Gadget', 45]
            ]);
        });

        it('should support reading back with schema after writing with schema', async () => {
            const users = [
                { id: 501, name: 'Carol', balance: 500.25 },
                { id: 502, name: 'Dan', balance: 750.00 }
            ];

            const writeSchema = {
                id: { column: 'ID', type: Number },
                name: { column: 'NAME', type: String, size: 20 },
                balance: { column: 'BALANCE', type: Number, size: 10, decimalPlaces: 2 }
            };

            const buf = await writeSpreadsheet(users, { schema: writeSchema, format: 'dbf', buffer: true });

            // Read back using read-excel-file schema format
            const readSchema = {
                id: { column: 'ID', type: Number },
                name: { column: 'NAME', type: String },
                balance: { column: 'BALANCE', type: Number }
            };

            const parsed = await readSpreadsheet(buf, { extension: 'dbf', schema: readSchema });
            expect(parsed.rows).to.deep.equal([
                { id: 501, name: 'Carol', balance: 500.25 },
                { id: 502, name: 'Dan', balance: 750 }
            ]);
        });
    });

    describe('yortus/DBFFile compatibility', () => {
        before(() => {
            registerPlugin(dbfPlugin);
        });

        it('should accept DBFFile-style fieldDescriptors directly', async () => {
            const fieldDescriptors = [
                { name: 'ID', type: 'I', size: 4 },
                { name: 'TITLE', type: 'C', size: 30 },
                { name: 'PRICE', type: 'N', size: 8, decimalPlaces: 2 },
                { name: 'ACTIVE', type: 'L', size: 1 },
                { name: 'STAMP', type: 'D', size: 8 }
            ];

            const records = [
                { ID: 1, TITLE: 'Book One', PRICE: 19.99, ACTIVE: true, STAMP: new Date('2026-01-01') },
                { ID: 2, TITLE: 'Book Two', PRICE: 25.50, ACTIVE: false, STAMP: new Date('2026-01-02') }
            ];

            const filePath = path.join(tmpDir, 'dbffile-compat.dbf');
            await writeDBF(records, { filePath, fieldDescriptors, fileVersion: 0x30 });

            const rows = await readSpreadsheet(filePath);
            expect(rows[0]).to.deep.equal(['ID', 'TITLE', 'PRICE', 'ACTIVE', 'STAMP']);
            expect(rows[1]).to.deep.equal([1, 'Book One', 19.99, true, '2026-01-01']);
            expect(rows[2]).to.deep.equal([2, 'Book Two', 25.5, false, '2026-01-02']);
        });

        it('should support record deletion flags', async () => {
            const fieldDescriptors = [
                { name: 'ID', type: 'N', size: 4 },
                { name: 'VAL', type: 'C', size: 10 }
            ];

            const records = [
                { ID: 1, VAL: 'Active 1' },
                { ID: 2, VAL: 'Deleted', _deleted: true },
                { ID: 3, VAL: 'Active 2' }
            ];

            const buf = await writeDBF(records, { fieldDescriptors });
            // By default readSpreadsheet skips deleted records
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows).to.deep.equal([
                ['ID', 'VAL'],
                [1, 'Active 1'],
                [3, 'Active 2']
            ]);
        });

        it('should support deletedRows option', async () => {
            const data = [
                { code: 'A' },
                { code: 'B' },
                { code: 'C' }
            ];

            const buf = await writeDBF(data, { deletedRows: [1] });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows).to.deep.equal([
                ['CODE'],
                ['A'],
                ['C']
            ]);
        });

        it('should support custom fileVersion and lastUpdate date', async () => {
            const data = [{ x: 1 }];
            const customDate = new Date(2025, 6, 4); // 2025-07-04

            const buf = await writeDBF(data, { fileVersion: 0x30, lastUpdate: customDate });
            expect(buf[0]).to.equal(0x30); // Visual FoxPro
            expect(buf[1]).to.equal(125);  // 2025 - 1900 = 125
            expect(buf[2]).to.equal(7);    // July
            expect(buf[3]).to.equal(4);    // Day 4
        });
    });

    describe('Auto-inference without schema', () => {
        before(() => {
            registerPlugin(dbfPlugin);
        });

        it('should auto-infer fields, types and sizes from array of objects', async () => {
            const data = [
                { name: 'Alpha', count: 10, ratio: 1.25, flag: true },
                { name: 'Beta Longer Name', count: 200, ratio: 4.5, flag: false }
            ];

            const buf = await writeDBF(data);
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            expect(rows[0]).to.deep.equal(['NAME', 'COUNT', 'RATIO', 'FLAG']);
            expect(rows[1]).to.deep.equal(['Alpha', 10, 1.25, true]);
            expect(rows[2]).to.deep.equal(['Beta Longer Name', 200, 4.5, false]);
        });

        it('should support 2D array grid data with headers in row 0', async () => {
            const grid = [
                ['Item', 'Qty', 'Available'],
                ['Widget', 15, true],
                ['Gizmo', 0, false]
            ];

            const buf = await writeDBF(grid);
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            expect(rows).to.deep.equal([
                ['ITEM', 'QTY', 'AVAILABLE'],
                ['Widget', 15, true],
                ['Gizmo', 0, false]
            ]);
        });

        it('should support TabularJS worksheets input format', async () => {
            const tabular = {
                worksheets: [
                    {
                        worksheetName: 'Sheet1',
                        data: [
                            ['CODE', 'VAL'],
                            ['X1', 100],
                            ['Y2', 200]
                        ]
                    }
                ]
            };

            const buf = await writeDBF(tabular);
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows).to.deep.equal([
                ['CODE', 'VAL'],
                ['X1', 100],
                ['Y2', 200]
            ]);
        });
    });

    describe('Field name sanitization & Edge cases', () => {
        before(() => {
            registerPlugin(dbfPlugin);
        });

        it('should sanitize invalid characters and truncate long column names', async () => {
            const data = [
                {
                    'Customer Account Identification': 1,
                    '1st Quarter Sales %': 500,
                    'Special!@#$%Chars': 'Test'
                }
            ];

            const buf = await writeDBF(data);
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            // Max 10 chars, uppercase, valid DBF identifier
            expect(rows[0][0].length).to.be.at.most(10);
            expect(rows[0][0]).to.equal('CUSTOMER_A');
            expect(rows[0][1].length).to.be.at.most(10);
            expect(rows[0][1]).to.equal('F_1ST_QUAR');
            expect(rows[0][2].length).to.be.at.most(10);
            expect(rows[0][2]).to.equal('SPECIALCHA');
        });

        it('should handle duplicate column names with unique suffixes', async () => {
            const schema = [
                { column: 'LONG_COLUMN_NAME_A', type: String },
                { column: 'LONG_COLUMN_NAME_B', type: String }
            ];
            const fields = normalizeSchemaToDbfFields(schema);
            expect(fields[0].name).to.equal('LONG_COLUM');
            expect(fields[1].name).to.equal('LONG_COL_2');
            expect(fields[0].name).to.not.equal(fields[1].name);
        });

        it('should handle null and undefined cell values safely', async () => {
            const data = [
                { id: 1, text: null, num: undefined, active: null, dt: null },
                { id: 2, text: 'Valid', num: 123.45, active: true, dt: new Date('2026-03-15') }
            ];

            const schema = {
                id: { column: 'ID', type: Number },
                text: { column: 'TEXT', type: String, size: 10 },
                num: { column: 'NUM', type: Number, size: 8, decimalPlaces: 2 },
                active: { column: 'ACTIVE', type: Boolean },
                dt: { column: 'DT', type: Date }
            };

            const buf = await writeDBF(data, { schema });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            expect(rows[1]).to.deep.equal([1, null, null, null, null]);
            expect(rows[2]).to.deep.equal([2, 'Valid', 123.45, true, '2026-03-15']);
        });

        it('should support empty records array with schema', async () => {
            const schema = {
                id: { column: 'ID', type: Number },
                name: { column: 'NAME', type: String, size: 20 }
            };

            const buf = await writeDBF([], { schema });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows.length).to.equal(1); // header only
            expect(rows[0]).to.deep.equal(['ID', 'NAME']);
        });
    });

    describe('Encoding and international text support', () => {
        before(() => {
            registerPlugin(dbfPlugin);
        });

        it('should write and read back UTF-8 international text', async () => {
            const data = [
                { city: 'München', country: 'Deutschland' },
                { city: 'São Paulo', country: 'Brasil' },
                { city: 'Tokyo', country: '日本' }
            ];

            const buf = await writeDBF(data, { encoding: 'utf-8' });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });

            expect(rows[1]).to.deep.equal(['München', 'Deutschland']);
            expect(rows[2]).to.deep.equal(['São Paulo', 'Brasil']);
            expect(rows[3]).to.deep.equal(['Tokyo', '日本']);
        });

        it('should support Latin-1 / ISO-8859-1 encoding', async () => {
            const data = [
                { word: 'Café', desc: 'Crème' }
            ];

            const buf = await writeDBF(data, { encoding: 'latin1' });
            const rows = await readSpreadsheet(buf, { extension: 'dbf' });
            expect(rows[1]).to.deep.equal(['Café', 'Crème']);
        });
    });
});
