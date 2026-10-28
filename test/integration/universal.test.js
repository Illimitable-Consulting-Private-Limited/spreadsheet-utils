import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';
import {
    readSpreadsheet,
    readXlsxFile,
    readSheet,
    tabularjs,
    parseTabular,
    detectFormat,
    registerPlugin,
    Integer
} from '../../src/index.js';
import { allPlugins } from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const testCasesDir = path.resolve(__dirname, '../read-excel-file/testCases');

describe('Unified Spreadsheet Reader (Integration Tests)', () => {
    before(() => {
        registerPlugin(...allPlugins);
    });

    describe('XLSX format with read-excel-file streaming engine', () => {
        const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
        const multipleSheetsXlsx = path.join(testCasesDir, 'multiple-sheets.xlsx');
        const schemaXlsx = path.join(testCasesDir, 'schema.xlsx');

        it('should read rows as 2D array by default', async () => {
            const rows = await readSpreadsheet(booleanXlsx);
            expect(rows).to.be.an('array');
            expect(rows).to.deep.equal([
                ['Boolean'],
                [true],
                [false],
                [1],
                [0]
            ]);
        });

        it('should support readXlsxFile alias', async () => {
            const rows = await readXlsxFile(booleanXlsx);
            expect(rows).to.deep.equal([
                ['Boolean'],
                [true],
                [false],
                [1],
                [0]
            ]);
        });

        it('should read multiple sheets when allSheets: true', async () => {
            const sheets = await readSpreadsheet(multipleSheetsXlsx, { allSheets: true });
            expect(sheets).to.be.an('array');
            expect(sheets.length).to.equal(2);
            expect(sheets[0]).to.have.property('name', 'sheet 1');
            expect(sheets[0].data).to.deep.equal([['First sheet']]);
            expect(sheets[1]).to.have.property('name', 'sheet 2');
            expect(sheets[1].data).to.deep.equal([['Second sheet']]);
        });

        it('should select specific sheet by name or 1-based index', async () => {
            const sheet2ByIndex = await readSpreadsheet(multipleSheetsXlsx, { sheet: 2 });
            expect(sheet2ByIndex).to.deep.equal([['Second sheet']]);

            const sheet2ByName = await readSpreadsheet(multipleSheetsXlsx, { sheet: 'sheet 2' });
            expect(sheet2ByName).to.deep.equal([['Second sheet']]);
        });

        it('should support schema validation and transformation on XLSX', async () => {
            const schema = {
                name: {
                    column: 'Name',
                    type: String,
                    required: true
                },
                income: {
                    column: 'Income',
                    type: Number,
                    required: true
                }
            };

            const result = await readSpreadsheet(schemaXlsx, { schema });
            expect(result).to.have.property('rows');
            expect(result.rows).to.be.an('array');
            expect(result.rows.length).to.be.greaterThan(0);
            expect(result.rows[0]).to.have.property('name', 'John Smith');
            expect(result.rows[0]).to.have.property('income', 120000);
        });

        it('should return TabularJS worksheet format when format: "tabular"', async () => {
            const result = await readSpreadsheet(booleanXlsx, { format: 'tabular' });
            expect(result).to.have.property('worksheets');
            expect(result.worksheets).to.be.an('array');
            expect(result.worksheets[0]).to.have.property('worksheetName', 'Sheet1');
            expect(result.worksheets[0]).to.have.property('columns');
            expect(result.worksheets[0].data).to.deep.equal([
                ['Boolean'],
                [true],
                [false],
                [1],
                [0]
            ]);
        });

        it('should return TabularJS worksheet format when called via tabularjs()', async () => {
            const result = await tabularjs(booleanXlsx);
            expect(result).to.have.property('worksheets');
            expect(result.worksheets[0].data).to.deep.equal([
                ['Boolean'],
                [true],
                [false],
                [1],
                [0]
            ]);
        });
    });

    describe('CSV and TSV formats extended with read-excel-file features', () => {
        const csvContent = 'ID,NAME,AGE\n1,Alice,30\n2,Bob,25\n3,Charlie,40';
        const csvBuffer = new TextEncoder().encode(csvContent);

        it('should read CSV as 2D array rows by default', async () => {
            const rows = await readSpreadsheet(csvBuffer, { extension: 'csv' });
            expect(rows).to.deep.equal([
                ['ID', 'NAME', 'AGE'],
                ['1', 'Alice', '30'],
                ['2', 'Bob', '25'],
                ['3', 'Charlie', '40']
            ]);
        });

        it('should apply schema parsing to CSV data', async () => {
            const schema = {
                id: {
                    column: 'ID',
                    type: Integer,
                    required: true
                },
                name: {
                    column: 'NAME',
                    type: String,
                    required: true
                },
                age: {
                    column: 'AGE',
                    type: Integer,
                    required: true
                }
            };

            const result = await readSpreadsheet(csvBuffer, { extension: 'csv', schema });
            expect(result.rows).to.deep.equal([
                { id: 1, name: 'Alice', age: 30 },
                { id: 2, name: 'Bob', age: 25 },
                { id: 3, name: 'Charlie', age: 40 }
            ]);
        });

        it('should return TabularJS format for CSV when requested', async () => {
            const result = await readSpreadsheet(csvBuffer, { extension: 'csv', format: 'tabular' });
            expect(result).to.have.property('worksheets');
            expect(result.worksheets[0].columns).to.have.lengthOf(3);
            expect(result.worksheets[0].columns[0].title).to.equal('A');
        });
    });

    describe('HTML Tables extended with read-excel-file features', () => {
        const html = `
            <table>
                <tr><th>City</th><th>Population</th></tr>
                <tr><td>New York</td><td>8000000</td></tr>
                <tr><td>London</td><td>9000000</td></tr>
            </table>
        `;
        const htmlBuffer = new TextEncoder().encode(html);

        it('should parse HTML table rows directly', async () => {
            const rows = await readSpreadsheet(htmlBuffer, { extension: 'html' });
            expect(rows).to.deep.equal([
                ['New York', '8000000'],
                ['London', '9000000']
            ]);
        });

        it('should apply schema parsing to HTML table data', async () => {
            const htmlWithHeaderData = `
                <table>
                    <tr><td>City</td><td>Population</td></tr>
                    <tr><td>New York</td><td>8000000</td></tr>
                    <tr><td>London</td><td>9000000</td></tr>
                </table>
            `;
            const schema = {
                city: { column: 'City', type: String },
                population: { column: 'Population', type: Integer }
            };
            const result = await readSpreadsheet(new TextEncoder().encode(htmlWithHeaderData), {
                extension: 'html',
                firstRowAsHeader: false,
                schema
            });
            expect(result.rows).to.deep.equal([
                { city: 'New York', population: 8000000 },
                { city: 'London', population: 9000000 }
            ]);
        });
    });

    describe('XML Spreadsheet 2003 (SpreadsheetML)', () => {
        const xml = `<?xml version="1.0"?>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
                <Worksheet ss:Name="Sheet1">
                    <Table>
                        <Row><Cell><Data ss:Type="String">Header</Data></Cell></Row>
                        <Row><Cell><Data ss:Type="Number">42</Data></Cell></Row>
                    </Table>
                </Worksheet>
            </Workbook>`;

        it('should parse XML spreadsheet rows', async () => {
            const rows = await readSpreadsheet(new TextEncoder().encode(xml), { extension: 'xml' });
            expect(rows).to.deep.equal([
                ['Header'],
                [42]
            ]);
        });
    });

    describe('Dependency & Architecture Verification', () => {
        it('must not depend on JSZip anywhere in the source code', () => {
            const srcDir = path.resolve(__dirname, '../../src');
            function checkNoJSZip(dir) {
                const entries = fs.readdirSync(dir, { withFileTypes: true });
                for (const entry of entries) {
                    const fullPath = path.join(dir, entry.name);
                    if (entry.isDirectory()) {
                        checkNoJSZip(fullPath);
                    } else if (entry.isFile() && entry.name.endsWith('.js')) {
                        const content = fs.readFileSync(fullPath, 'utf-8');
                        expect(content).to.not.include("'jszip'");
                        expect(content).to.not.include('"jszip"');
                    }
                }
            }
            checkNoJSZip(srcDir);
        });
    });
});
