import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    writeSpreadsheet,
    writeXLS,
    createXLS,
    writeODS,
    createODS,
    writeDBF
} from '../../src/write/index.js';

import {
    readSpreadsheet,
    registerPlugin,
    unregisterPlugin,
    isPluginRegistered,
    getWriterForExtension
} from '../../src/index.js';

import {
    xlsPlugin,
    markupPlugin,
    odsPlugin,
    dbfPlugin,
    parseXLS,
    parseODS
} from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tmpDir = path.resolve(__dirname, '../../tmp-test-ods-xls');

describe('XLS & ODS Export Plugins & Schema Writers', () => {

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

    describe('Plugin registration & missing-plugin diagnostics', () => {
        beforeEach(() => {
            unregisterPlugin();
        });

        it('should throw ERR_PLUGIN_REQUIRED when writing .xls without xlsPlugin', async () => {
            const data = [
                { id: 1, name: 'Alice' }
            ];

            try {
                await writeSpreadsheet(data, { filePath: path.join(tmpDir, 'test.xls') });
                expect.fail('Should have thrown missing plugin error');
            } catch (err) {
                expect(err.code).to.equal('ERR_PLUGIN_REQUIRED');
                expect(err.format).to.equal('xls');
                expect(err.message).to.include('xlsPlugin');
            }
        });

        it('should throw ERR_PLUGIN_REQUIRED when writing .ods without markupPlugin or odsPlugin', async () => {
            const data = [
                { id: 1, name: 'Alice' }
            ];

            try {
                await writeSpreadsheet(data, { filePath: path.join(tmpDir, 'test.ods') });
                expect.fail('Should have thrown missing plugin error');
            } catch (err) {
                expect(err.code).to.equal('ERR_PLUGIN_REQUIRED');
                expect(err.format).to.equal('ods');
                expect(err.message).to.include('markupPlugin');
            }
        });

        it('should successfully route .xls write after registerPlugin(xlsPlugin)', async () => {
            registerPlugin(xlsPlugin);
            expect(isPluginRegistered('xlsPlugin')).to.be.true;
            expect(getWriterForExtension('xls')).to.be.a('function');

            const outPath = path.join(tmpDir, 'plugin-registered.xls');
            const data = [
                ['Name', 'Score'],
                ['Alice', 95.5],
                ['Bob', 88.0]
            ];

            const res = await writeSpreadsheet(data, { filePath: outPath });
            expect(res).to.equal(outPath);
            expect(fs.existsSync(outPath)).to.be.true;

            // Header magic check: OLE2 container
            const buf = fs.readFileSync(outPath);
            expect(buf[0]).to.equal(0xD0);
            expect(buf[1]).to.equal(0xCF);
            expect(buf[2]).to.equal(0x11);
            expect(buf[3]).to.equal(0xE0);
        });

        it('should successfully route .ods write after registerPlugin(markupPlugin)', async () => {
            registerPlugin(markupPlugin);
            expect(isPluginRegistered('markupPlugin')).to.be.true;
            expect(getWriterForExtension('ods')).to.be.a('function');

            const outPath = path.join(tmpDir, 'plugin-markup.ods');
            const data = [
                ['Item', 'Qty'],
                ['Widgets', 12],
                ['Gadgets', 4]
            ];

            const res = await writeSpreadsheet(data, { filePath: outPath });
            expect(res).to.equal(outPath);
            expect(fs.existsSync(outPath)).to.be.true;

            // Header magic check: PK ZIP
            const buf = fs.readFileSync(outPath);
            expect(buf[0]).to.equal(0x50);
            expect(buf[1]).to.equal(0x4B);
        });

        it('should successfully route .ods write after registerPlugin(odsPlugin)', async () => {
            unregisterPlugin();
            registerPlugin(odsPlugin);
            expect(isPluginRegistered('odsPlugin')).to.be.true;
            expect(getWriterForExtension('ods')).to.be.a('function');

            const outPath = path.join(tmpDir, 'plugin-ods-standalone.ods');
            const data = [
                ['Product', 'Price'],
                ['Coffee', 4.5]
            ];

            const res = await writeSpreadsheet(data, { filePath: outPath });
            expect(res).to.equal(outPath);
            expect(fs.existsSync(outPath)).to.be.true;
        });

        it('should restore missing plugin error after unregisterPlugin', async () => {
            registerPlugin(xlsPlugin);
            expect(isPluginRegistered('xlsPlugin')).to.be.true;

            unregisterPlugin(xlsPlugin);
            expect(isPluginRegistered('xlsPlugin')).to.be.false;

            try {
                await writeSpreadsheet([['A']], { filePath: path.join(tmpDir, 'unregistered.xls') });
                expect.fail('Should fail after unregister');
            } catch (err) {
                expect(err.code).to.equal('ERR_PLUGIN_REQUIRED');
            }
        });
    });

    describe('Direct writers: writeXLS & createXLS', () => {
        it('createXLS should generate a valid Uint8Array buffer', () => {
            const data = [
                ['City', 'Population'],
                ['Tokyo', 37400000],
                ['Delhi', 29300000]
            ];

            const uint8 = createXLS(data);
            expect(uint8).to.be.instanceOf(Uint8Array);
            expect(uint8.length).to.be.greaterThan(512);

            // OLE2 magic
            expect(uint8[0]).to.equal(0xD0);
            expect(uint8[1]).to.equal(0xCF);
            expect(uint8[2]).to.equal(0x11);
            expect(uint8[3]).to.equal(0xE0);
        });

        it('writeXLS should return a Node.js Buffer when no filePath is provided', async () => {
            const data = [
                ['A', 'B'],
                [1, 2]
            ];
            const buf = await writeXLS(data);
            expect(Buffer.isBuffer(buf)).to.be.true;
            expect(buf.length).to.be.greaterThan(512);
        });

        it('writeXLS should return Uint8Array when uint8Array option is true', async () => {
            const data = [['A', 'B'], [1, 2]];
            const res = await writeXLS(data, { uint8Array: true });
            expect(res).to.be.instanceOf(Uint8Array);
        });

        it('writeXLS should write to file when filePath option is given', async () => {
            const outPath = path.join(tmpDir, 'direct-write.xls');
            const data = [
                ['Code', 'Active'],
                ['X1', true],
                ['X2', false]
            ];
            const res = await writeXLS(data, { filePath: outPath });
            expect(res).to.equal(outPath);
            expect(fs.existsSync(outPath)).to.be.true;
        });
    });

    describe('Direct writers: writeODS & createODS', () => {
        it('createODS should generate a valid Uint8Array buffer', () => {
            const data = [
                ['Title', 'Author'],
                ['1984', 'George Orwell'],
                ['Brave New World', 'Aldous Huxley']
            ];

            const uint8 = createODS(data);
            expect(uint8).to.be.instanceOf(Uint8Array);
            expect(uint8.length).to.be.greaterThan(100);

            // PK ZIP magic
            expect(uint8[0]).to.equal(0x50);
            expect(uint8[1]).to.equal(0x4B);
        });

        it('writeODS should return a Node.js Buffer when no filePath is provided', async () => {
            const data = [
                ['X', 'Y'],
                [10, 20]
            ];
            const buf = await writeODS(data);
            expect(Buffer.isBuffer(buf)).to.be.true;
            expect(buf.length).to.be.greaterThan(100);
        });

        it('writeODS should return Uint8Array when uint8Array option is true', async () => {
            const data = [['X', 'Y'], [10, 20]];
            const res = await writeODS(data, { uint8Array: true });
            expect(res).to.be.instanceOf(Uint8Array);
        });

        it('writeODS should write to file when filePath option is given', async () => {
            const outPath = path.join(tmpDir, 'direct-write.ods');
            const data = [
                ['Dept', 'Budget'],
                ['Marketing', 50000],
                ['Engineering', 120000]
            ];
            const res = await writeODS(data, { filePath: outPath });
            expect(res).to.equal(outPath);
            expect(fs.existsSync(outPath)).to.be.true;
        });
    });

    describe('Round-trip data integrity (Write then Read)', () => {
        before(() => {
            registerPlugin(xlsPlugin, markupPlugin);
        });

        it('should write .xls and read back with identical values and types', async () => {
            const filePath = path.join(tmpDir, 'roundtrip.xls');
            const sourceData = [
                ['Name', 'Age', 'Score', 'Active'],
                ['Alice', 30, 98.75, true],
                ['Bob', 25, 85.5, false],
                ['Charlie', 42, 100, true]
            ];

            await writeXLS(sourceData, { filePath });

            const rows = await readSpreadsheet(filePath);
            expect(rows).to.be.an('array').with.lengthOf(4);

            expect(rows[0]).to.deep.equal(['Name', 'Age', 'Score', 'Active']);
            expect(rows[1][0]).to.equal('Alice');
            expect(rows[1][1]).to.equal(30);
            expect(rows[1][2]).to.be.closeTo(98.75, 0.001);
            expect(rows[1][3]).to.equal(true);

            expect(rows[2][0]).to.equal('Bob');
            expect(rows[2][1]).to.equal(25);
            expect(rows[2][2]).to.be.closeTo(85.5, 0.001);
            expect(rows[2][3]).to.equal(false);
        });

        it('should write .ods and read back with identical values and types', async () => {
            const filePath = path.join(tmpDir, 'roundtrip.ods');
            const sourceData = [
                ['SKU', 'Price', 'InStock'],
                ['ITEM-101', 19.99, true],
                ['ITEM-202', 49.5, false]
            ];

            await writeODS(sourceData, { filePath });

            const rows = await readSpreadsheet(filePath);
            expect(rows).to.be.an('array').with.lengthOf(3);
            expect(rows[0]).to.deep.equal(['SKU', 'Price', 'InStock']);
            expect(rows[1][0]).to.equal('ITEM-101');
            expect(rows[1][1]).to.equal(19.99);
            expect(rows[1][2]).to.equal(true);
            expect(rows[2][0]).to.equal('ITEM-202');
            expect(rows[2][1]).to.equal(49.5);
            expect(rows[2][2]).to.equal(false);
        });
    });

    describe('Schema-based writing', () => {
        before(() => {
            registerPlugin(xlsPlugin, markupPlugin);
        });

        const records = [
            { id: 101, name: 'Alice Smith', salary: 75000, hireDate: new Date('2021-03-15') },
            { id: 102, name: 'Bob Johnson', salary: 62000, hireDate: new Date('2022-07-01') },
            { id: 103, name: 'Carol White', salary: 89000, hireDate: new Date('2020-01-10') }
        ];

        const schema = [
            {
                column: 'Employee ID',
                type: Number,
                value: row => row.id,
                fontWeight: 'bold'
            },
            {
                column: 'Full Name',
                type: String,
                value: row => row.name
            },
            {
                column: 'Annual Salary',
                type: Number,
                value: row => row.salary
            },
            {
                column: 'Hire Date',
                type: Date,
                value: row => row.hireDate
            }
        ];

        it('should write .xls with schema definition', async () => {
            const filePath = path.join(tmpDir, 'schema-export.xls');
            await writeSpreadsheet(records, {
                filePath,
                schema
            });

            const rows = await readSpreadsheet(filePath);

            // Header row generated from schema column names
            expect(rows[0]).to.deep.equal(['Employee ID', 'Full Name', 'Annual Salary', 'Hire Date']);
            // Data rows
            expect(rows[1][0]).to.equal(101);
            expect(rows[1][1]).to.equal('Alice Smith');
            expect(rows[1][2]).to.equal(75000);
            expect(rows[1][3]).to.equal('2021-03-15');

            expect(rows[2][0]).to.equal(102);
            expect(rows[2][1]).to.equal('Bob Johnson');
            expect(rows[2][2]).to.equal(62000);
            expect(rows[2][3]).to.equal('2022-07-01');

            // Also test schema-based parsing back into objects
            const readSchema = {
                'id': { column: 'Employee ID', type: Number },
                'name': { column: 'Full Name', type: String },
                'salary': { column: 'Annual Salary', type: Number }
            };
            const { objects } = await readSpreadsheet(filePath, { schema: readSchema });
            expect(objects).to.have.lengthOf(3);
            expect(objects[0]).to.deep.equal({ id: 101, name: 'Alice Smith', salary: 75000 });
        });

        it('should write .ods with schema definition', async () => {
            const filePath = path.join(tmpDir, 'schema-export.ods');
            await writeSpreadsheet(records, {
                filePath,
                schema
            });

            const rows = await readSpreadsheet(filePath);

            expect(rows[0]).to.deep.equal(['Employee ID', 'Full Name', 'Annual Salary', 'Hire Date']);
            expect(rows[1][0]).to.equal(101);
            expect(rows[1][1]).to.equal('Alice Smith');
            expect(rows[1][2]).to.equal(75000);
            expect(rows[1][3]).to.equal('2021-03-15');

            // Also test schema-based parsing back into objects
            const readSchema = {
                'id': { column: 'Employee ID', type: Number },
                'name': { column: 'Full Name', type: String },
                'salary': { column: 'Annual Salary', type: Number }
            };
            const { objects } = await readSpreadsheet(filePath, { schema: readSchema });
            expect(objects).to.have.lengthOf(3);
            expect(objects[0]).to.deep.equal({ id: 101, name: 'Alice Smith', salary: 75000 });
        });

        it('should auto-derive columns from object keys when schema is omitted', async () => {
            const plainObjects = [
                { team: 'Frontend', members: 8 },
                { team: 'Backend', members: 12 }
            ];

            const xlsPath = path.join(tmpDir, 'auto-schema.xls');
            await writeSpreadsheet(plainObjects, { filePath: xlsPath });

            const xlsRows = await readSpreadsheet(xlsPath);
            expect(xlsRows[0]).to.deep.equal(['team', 'members']);
            expect(xlsRows[1]).to.deep.equal(['Frontend', 8]);
            expect(xlsRows[2]).to.deep.equal(['Backend', 12]);

            const odsPath = path.join(tmpDir, 'auto-schema.ods');
            await writeSpreadsheet(plainObjects, { filePath: odsPath });

            const odsRows = await readSpreadsheet(odsPath);
            expect(odsRows[0]).to.deep.equal(['team', 'members']);
            expect(odsRows[1]).to.deep.equal(['Frontend', 8]);
            expect(odsRows[2]).to.deep.equal(['Backend', 12]);
        });
    });

    describe('Multi-sheet workbooks', () => {
        before(() => {
            registerPlugin(xlsPlugin, markupPlugin);
        });

        const multiSheetData = [
            {
                sheet: 'North America',
                data: [
                    ['Region', 'Sales'],
                    ['USA', 1500000],
                    ['Canada', 300000]
                ]
            },
            {
                sheet: 'Europe',
                data: [
                    ['Region', 'Sales'],
                    ['UK', 800000],
                    ['Germany', 950000]
                ]
            },
            {
                sheet: 'Asia Pacific',
                data: [
                    ['Region', 'Sales'],
                    ['Japan', 700000],
                    ['Australia', 450000]
                ]
            }
        ];

        it('should write multi-sheet .xls and read all sheets back', async () => {
            const filePath = path.join(tmpDir, 'multi-sheet.xls');
            await writeSpreadsheet(multiSheetData, { filePath });

            const sheets = await readSpreadsheet(filePath, { allSheets: true });
            expect(sheets).to.have.lengthOf(3);

            expect(sheets[0].name).to.equal('North America');
            expect(sheets[0].data[1]).to.deep.equal(['USA', 1500000]);

            expect(sheets[1].name).to.equal('Europe');
            expect(sheets[1].data[1]).to.deep.equal(['UK', 800000]);

            expect(sheets[2].name).to.equal('Asia Pacific');
            expect(sheets[2].data[1]).to.deep.equal(['Japan', 700000]);
        });

        it('should write multi-sheet .ods and read all sheets back', async () => {
            const filePath = path.join(tmpDir, 'multi-sheet.ods');
            await writeSpreadsheet(multiSheetData, { filePath });

            const sheets = await readSpreadsheet(filePath, { allSheets: true });
            expect(sheets).to.have.lengthOf(3);

            expect(sheets[0].name).to.equal('North America');
            expect(sheets[0].data[1]).to.deep.equal(['USA', 1500000]);

            expect(sheets[1].name).to.equal('Europe');
            expect(sheets[1].data[1]).to.deep.equal(['UK', 800000]);

            expect(sheets[2].name).to.equal('Asia Pacific');
            expect(sheets[2].data[1]).to.deep.equal(['Japan', 700000]);
        });

        it('should write multi-sheet TabularJS format ({ worksheets: [...] }) to .xls and .ods', async () => {
            const tabularData = {
                worksheets: [
                    { worksheetName: 'Summary', data: [['Metric', 'Val'], ['ARR', 10000000]] },
                    { worksheetName: 'Details', data: [['Cust', 'MRR'], ['Acme', 50000]] }
                ]
            };

            const xlsPath = path.join(tmpDir, 'tabular-multi.xls');
            await writeSpreadsheet(tabularData, { filePath: xlsPath });
            const xlsResult = await readSpreadsheet(xlsPath, { worksheets: true });
            expect(xlsResult.worksheets).to.have.lengthOf(2);
            expect(xlsResult.worksheets[0].worksheetName).to.equal('Summary');
            expect(xlsResult.worksheets[1].worksheetName).to.equal('Details');

            const odsPath = path.join(tmpDir, 'tabular-multi.ods');
            await writeSpreadsheet(tabularData, { filePath: odsPath });
            const odsResult = await readSpreadsheet(odsPath, { worksheets: true });
            expect(odsResult.worksheets).to.have.lengthOf(2);
            expect(odsResult.worksheets[0].worksheetName).to.equal('Summary');
            expect(odsResult.worksheets[1].worksheetName).to.equal('Details');
        });
    });

    describe('Cell styling: bold cells & attributes', () => {
        before(() => {
            registerPlugin(xlsPlugin, markupPlugin);
        });

        it('should preserve values when styled cell objects are passed to writeXLS', async () => {
            const data = [
                [
                    { value: 'Header Bold', fontWeight: 'bold' },
                    { value: 12345, bold: true }
                ],
                [
                    { value: 'Regular Text' },
                    { value: 67890 }
                ]
            ];

            const filePath = path.join(tmpDir, 'styled.xls');
            await writeXLS(data, { filePath });

            const rows = await readSpreadsheet(filePath);
            expect(rows[0][0]).to.equal('Header Bold');
            expect(rows[0][1]).to.equal(12345);
            expect(rows[1][0]).to.equal('Regular Text');
            expect(rows[1][1]).to.equal(67890);
        });

        it('should preserve values when styled cell objects are passed to writeODS', async () => {
            const data = [
                [
                    { value: 'Header Bold', fontWeight: 'bold' },
                    { value: 12345, bold: true }
                ],
                [
                    { value: 'Regular Text' },
                    { value: 67890 }
                ]
            ];

            const filePath = path.join(tmpDir, 'styled.ods');
            await writeODS(data, { filePath });

            const rows = await readSpreadsheet(filePath);
            expect(rows[0][0]).to.equal('Header Bold');
            expect(rows[0][1]).to.equal(12345);
            expect(rows[1][0]).to.equal('Regular Text');
            expect(rows[1][1]).to.equal(67890);
        });

        it('should support colors, borders, and merged cells in ODS export', async () => {
            const grid = [
                [
                    {
                        value: 'Quarterly Summary',
                        columnSpan: 2,
                        rowSpan: 1,
                        backgroundColor: '#FFFF00',
                        textColor: '#FF0000',
                        fontWeight: 'bold',
                        align: 'center',
                        borderColor: '#000000',
                        borderStyle: 'thin'
                    },
                    null
                ],
                [
                    { value: 'Q1', backgroundColor: '#00FF00', borderStyle: 'medium' },
                    { value: 'Q2', textColor: '#0000FF' }
                ]
            ];

            const filePath = path.join(tmpDir, 'styled-merged.ods');
            await writeODS(grid, { filePath });

            const tabular = await readSpreadsheet(filePath, { format: 'tabular' });
            const ws = tabular.worksheets[0];

            // Verify mergeCells
            expect(ws.mergeCells).to.be.an('object');
            expect(ws.mergeCells['A1']).to.deep.equal([2, 1]);

            // Verify styles table contains background-color, color, and borders
            expect(tabular.style).to.be.an('array');
            const styleA1 = tabular.style[ws.style['A1']];
            expect(styleA1).to.include('color: #FF0000');
            expect(styleA1).to.include('background-color: #FFFF00');
            expect(styleA1).to.include('border');
            expect(styleA1).to.include('text-align: center');
        });

        it('should support colors, borders, and merged cells in XLS export', async () => {
            const grid = [
                [
                    {
                        value: 'Annual Report',
                        columnSpan: 2,
                        rowSpan: 1,
                        backgroundColor: '#FFFF00',
                        textColor: '#FF0000',
                        fontWeight: 'bold',
                        align: 'center',
                        borderStyle: 'thin'
                    },
                    null
                ],
                [
                    { value: 'Revenue', backgroundColor: '#00FF00' },
                    { value: 500000, textColor: '#0000FF' }
                ]
            ];

            const filePath = path.join(tmpDir, 'styled-merged.xls');
            await writeXLS(grid, { filePath });

            const tabular = await readSpreadsheet(filePath, { format: 'tabular' });
            const ws = tabular.worksheets[0];

            // Verify mergeCells
            expect(ws.mergeCells).to.be.an('object');
            expect(ws.mergeCells['A1']).to.deep.equal([2, 1]);

            // Verify style contains background-color, color, and alignment
            expect(tabular.style).to.be.an('array');
            const styleA1 = tabular.style[ws.style['A1']];
            expect(styleA1).to.include('text-align: center');
            expect(styleA1).to.include('background-color: #FFFF00');
            expect(styleA1).to.include('color: #FF0000');
            expect(styleA1).to.include('border');
        });
    });

    describe('Buffer export across all supported formats', () => {
        before(() => {
            registerPlugin(xlsPlugin, markupPlugin, odsPlugin, dbfPlugin);
        });

        const formats = ['xlsx', 'xls', 'ods', 'dbf', 'csv', 'tsv', 'html', 'json'];

        for (const fmt of formats) {
            it(`should export .${fmt} as a Node.js Buffer when buffer option is true`, async () => {
                const sampleData = [
                    ['Product', 'Qty'],
                    ['Alpha', 10],
                    ['Beta', 20]
                ];
                const buf = await writeSpreadsheet(sampleData, { format: fmt, buffer: true });
                expect(Buffer.isBuffer(buf)).to.be.true;
                expect(buf.length).to.be.greaterThan(0);
            });
        }

        it('should export .dbf as a Node.js Buffer via direct writeDBF', async () => {
            const sampleData = [
                ['Product', 'Qty'],
                ['Alpha', 10]
            ];
            const directBuf = await writeDBF(sampleData);
            expect(Buffer.isBuffer(directBuf)).to.be.true;
            expect(directBuf.length).to.be.greaterThan(0);
        });
    });
});
