import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    writeSpreadsheet,
    writeXlsxFile,
    createHTML,
    writeHTML,
    writeXLS,
    createXLS,
    writeODS,
    createODS
} from '../../src/write/index.js';

import {
    readSpreadsheet,
    registerPlugin
} from '../../src/index.js';

import { textPlugin, xlsPlugin, markupPlugin, odsPlugin } from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tmpDir = path.resolve(__dirname, '../../tmp-schema-styling-test');

describe('Unified Schema Styling, Cell Overrides & createHTML', () => {

    before(() => {
        if (!fs.existsSync(tmpDir)) {
            fs.mkdirSync(tmpDir, { recursive: true });
        }
        registerPlugin(textPlugin, xlsPlugin, markupPlugin, odsPlugin);
    });

    after(() => {
        if (fs.existsSync(tmpDir)) {
            fs.rmSync(tmpDir, { recursive: true, force: true });
        }
    });

    const employees = [
        { id: 101, name: 'Alice Smith', dept: 'Engineering', salary: 120000, bonus: 15000, active: true },
        { id: 102, name: 'Bob Johnson', dept: 'Marketing', salary: 75000, bonus: -2000, active: false },
        { id: 103, name: 'Charlie Rose', dept: 'Design', salary: 95000, bonus: 6000, active: true }
    ];

    describe('1. createHTML & writeHTML', () => {

        it('should generate an HTML string synchronously with createHTML', () => {
            const schema = [
                {
                    header: {
                        value: 'Employee ID',
                        fontWeight: 'bold',
                        backgroundColor: '#1F4E79',
                        color: '#FFFFFF',
                        align: 'center'
                    },
                    prop: 'id',
                    width: 15
                },
                {
                    header: 'Full Name',
                    prop: 'name',
                    width: 25
                },
                {
                    header: 'Bonus',
                    cell: (row) => ({
                        value: row.bonus,
                        fontWeight: row.bonus > 5000 ? 'bold' : 'normal',
                        backgroundColor: row.bonus > 5000 ? '#D9EAD3' : (row.bonus < 0 ? '#FFC7CE' : undefined),
                        color: row.bonus > 5000 ? '#274E13' : (row.bonus < 0 ? '#9C0006' : undefined),
                        align: 'right'
                    }),
                    width: 18
                }
            ];

            const html = createHTML(employees, { schema });
            expect(html).to.be.a('string');
            expect(html).to.include('<table class="spreadsheet-table">');
            expect(html).to.include('<colgroup>');
            expect(html).to.include('<col style="width:15ch;">');
            expect(html).to.include('<thead>');
            expect(html).to.include('<th>');
            expect(html).to.include('background-color:#1F4E79');
            expect(html).to.include('color:#FFFFFF');
            expect(html).to.include('Employee ID');
            expect(html).to.include('<tbody>');
            expect(html).to.include('<td>Alice Smith</td>');
            // Check dynamic cell styling on positive and negative bonus
            expect(html).to.include('background-color:#D9EAD3');
            expect(html).to.include('color:#274E13');
            expect(html).to.include('background-color:#FFC7CE');
            expect(html).to.include('color:#9C0006');
        });

        it('should wrap in a full HTML document when fullDocument: true', () => {
            const html = createHTML([['A', 'B'], [1, 2]], { fullDocument: true, title: 'Company Report' });
            expect(html).to.include('<!DOCTYPE html>');
            expect(html).to.include('<title>Company Report</title>');
            expect(html).to.include('border-collapse: collapse;');
            expect(html).to.include('<td>1</td>');
        });

        it('should write HTML file to disk and read back', async () => {
            const filePath = path.join(tmpDir, 'export.html');
            const saved = await writeHTML([['Header1', 'Header2'], ['Value1', 'Value2']], { filePath });
            expect(saved).to.equal(filePath);
            expect(fs.existsSync(filePath)).to.be.true;

            const readBack = await readSpreadsheet(filePath, { extension: 'html', firstRowAsHeader: false });
            expect(readBack).to.deep.equal([
                ['Header1', 'Header2'],
                ['Value1', 'Value2']
            ]);
        });

        it('should return HTML Buffer when buffer: true', async () => {
            const buf = await writeHTML([['X', 'Y'], ['1', '2']], { buffer: true });
            expect(Buffer.isBuffer(buf)).to.be.true;
            expect(buf.toString('utf-8')).to.include('<table');
        });
    });

    describe('2. Header Object Styling in Schema', () => {

        const schemaWithStyledHeaders = [
            {
                header: {
                    value: 'Emp ID',
                    fontWeight: 'bold',
                    backgroundColor: '#002060',
                    color: '#FFFFFF',
                    align: 'center'
                },
                prop: 'id',
                type: Number,
                width: 12
            },
            {
                header: {
                    value: 'Name',
                    fontWeight: 'bold',
                    backgroundColor: '#2E75B6',
                    color: '#FFFFFF'
                },
                prop: 'name',
                type: String,
                width: 20
            },
            {
                column: 'Department',
                headerStyle: {
                    fontWeight: 'bold',
                    backgroundColor: '#5B9BD5',
                    color: '#FFFFFF'
                },
                prop: 'dept',
                type: String,
                width: 15
            }
        ];

        it('should write XLSX with styled header objects and read back correctly', async () => {
            const filePath = path.join(tmpDir, 'headers.xlsx');
            await writeSpreadsheet(employees, { schema: schemaWithStyledHeaders, filePath });
            expect(fs.existsSync(filePath)).to.be.true;

            const rows = await readSpreadsheet(filePath);
            expect(rows[0]).to.deep.equal(['Emp ID', 'Name', 'Department']);
            expect(rows[1][0]).to.equal(101);
            expect(rows[1][1]).to.equal('Alice Smith');
        });

        it('should write ODS with styled header objects and column widths', async () => {
            const filePath = path.join(tmpDir, 'headers.ods');
            await writeSpreadsheet(employees, { schema: schemaWithStyledHeaders, filePath });
            expect(fs.existsSync(filePath)).to.be.true;

            const rows = await readSpreadsheet(filePath);
            expect(rows[0]).to.deep.equal(['Emp ID', 'Name', 'Department']);
            expect(rows[1][0]).to.equal(101);
            expect(rows[1][1]).to.equal('Alice Smith');
        });

        it('should write XLS with styled header objects and column widths', async () => {
            const filePath = path.join(tmpDir, 'headers.xls');
            await writeSpreadsheet(employees, { schema: schemaWithStyledHeaders, filePath });
            expect(fs.existsSync(filePath)).to.be.true;

            const rows = await readSpreadsheet(filePath);
            expect(rows[0]).to.deep.equal(['Emp ID', 'Name', 'Department']);
            expect(rows[1][0]).to.equal(101);
            expect(rows[1][1]).to.equal('Alice Smith');
        });
    });

    describe('3. Global headerStyle Option', () => {

        const plainSchema = [
            { column: 'ID', prop: 'id', type: Number },
            { column: 'Name', prop: 'name', type: String },
            { column: 'Salary', prop: 'salary', type: Number }
        ];

        const globalHeaderStyle = {
            fontWeight: 'bold',
            backgroundColor: '#44546A',
            color: '#FFFFFF',
            align: 'center'
        };

        it('should apply global headerStyle across XLSX, XLS, ODS, and HTML', async () => {
            // HTML
            const html = createHTML(employees, { schema: plainSchema, headerStyle: globalHeaderStyle });
            expect(html).to.include('background-color:#44546A');
            expect(html).to.include('color:#FFFFFF');
            expect(html).to.include('font-weight:bold');

            // XLSX
            const xlsxBuf = await writeSpreadsheet(employees, { schema: plainSchema, headerStyle: globalHeaderStyle, buffer: true, format: 'xlsx' });
            expect(xlsxBuf.length).to.be.greaterThan(0);

            // XLS
            const xlsBuf = await writeSpreadsheet(employees, { schema: plainSchema, headerStyle: globalHeaderStyle, buffer: true, format: 'xls' });
            expect(xlsBuf.length).to.be.greaterThan(0);

            // ODS
            const odsBuf = await writeSpreadsheet(employees, { schema: plainSchema, headerStyle: globalHeaderStyle, buffer: true, format: 'ods' });
            expect(odsBuf.length).to.be.greaterThan(0);
        });
    });

    describe('4. Dynamic cell() Callbacks & Rich value() Return (Cell Overrides)', () => {

        const dynamicSchema = [
            {
                column: 'Employee',
                prop: 'name',
                type: String
            },
            // Dynamic cell callback with conditional styling
            {
                column: 'Net Salary',
                cell: (row) => ({
                    value: row.salary,
                    type: Number,
                    format: '$#,##0.00',
                    align: 'right',
                    fontWeight: row.salary > 100000 ? 'bold' : 'normal',
                    backgroundColor: row.salary > 100000 ? '#E2EFDA' : undefined,
                    color: row.salary > 100000 ? '#385723' : undefined
                })
            },
            // Rich value return
            {
                column: 'Status',
                value: (row) => ({
                    value: row.active ? 'Active' : 'Inactive',
                    fontWeight: 'bold',
                    backgroundColor: row.active ? '#C6EFCE' : '#FFC7CE',
                    color: row.active ? '#006100' : '#9C0006'
                })
            }
        ];

        it('should apply dynamic cell overrides across XLSX, XLS, ODS, and HTML', async () => {
            // HTML verification
            const html = createHTML(employees, { schema: dynamicSchema });
            expect(html).to.include('background-color:#E2EFDA'); // Alice salary > 100000
            expect(html).to.include('background-color:#C6EFCE'); // Active status
            expect(html).to.include('background-color:#FFC7CE'); // Inactive status

            // ODS export & read back
            const odsPath = path.join(tmpDir, 'dynamic.ods');
            await writeSpreadsheet(employees, { schema: dynamicSchema, filePath: odsPath });
            const odsRows = await readSpreadsheet(odsPath);
            expect(odsRows[1]).to.deep.equal(['Alice Smith', 120000, 'Active']);
            expect(odsRows[2]).to.deep.equal(['Bob Johnson', 75000, 'Inactive']);

            // XLS export & read back
            const xlsPath = path.join(tmpDir, 'dynamic.xls');
            await writeSpreadsheet(employees, { schema: dynamicSchema, filePath: xlsPath });
            const xlsRows = await readSpreadsheet(xlsPath);
            expect(xlsRows[1]).to.deep.equal(['Alice Smith', 120000, 'Active']);
            expect(xlsRows[2]).to.deep.equal(['Bob Johnson', 75000, 'Inactive']);

            // XLSX export & read back
            const xlsxPath = path.join(tmpDir, 'dynamic.xlsx');
            await writeSpreadsheet(employees, { schema: dynamicSchema, filePath: xlsxPath });
            const xlsxRows = await readSpreadsheet(xlsxPath);
            expect(xlsxRows[1]).to.deep.equal(['Alice Smith', 120000, 'Active']);
            expect(xlsxRows[2]).to.deep.equal(['Bob Johnson', 75000, 'Inactive']);
        });

        it('should support dynamic style function on columns', () => {
            const funcStyleSchema = [
                {
                    column: 'Bonus',
                    prop: 'bonus',
                    type: Number,
                    style: (val) => ({
                        color: val < 0 ? '#FF0000' : '#008800',
                        fontWeight: val < 0 ? 'bold' : 'normal'
                    })
                }
            ];

            const html = createHTML(employees, { schema: funcStyleSchema });
            expect(html).to.include('color:#FF0000');
            expect(html).to.include('color:#008800');
        });
    });
});
