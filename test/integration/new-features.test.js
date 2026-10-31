import { describe, it } from 'mocha';
import { expect } from 'chai';
import fs from 'node:fs';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

import {
    readSpreadsheet,
    convertRowsToObjects,
    iterateSpreadsheet,
    writeSpreadsheet,
    convertSpreadsheet,
    calculateAutoWidths,
    THEMES,
    getTheme,
    applyTheme,
    createHTML,
    streamCSV,
    streamTSV,
    streamJSONL,
    allPlugins,
    registerPlugin
} from '../../src/full.js';

const execFileAsync = promisify(execFile);
const tmpDir = path.resolve('test/integration/tmp-features');

describe('Pre-Release Enhancements (All 6 Features)', () => {
    before(() => {
        registerPlugin(...allPlugins);
        if (!fs.existsSync(tmpDir)) {
            fs.mkdirSync(tmpDir, { recursive: true });
        }
    });

    beforeEach(() => {
        registerPlugin(...allPlugins);
    });

    after(() => {
        if (fs.existsSync(tmpDir)) {
            fs.rmSync(tmpDir, { recursive: true, force: true });
        }
    });

    // -------------------------------------------------------------
    // Feature 1: Auto-Object Reading
    // -------------------------------------------------------------
    describe('1. Auto-Object Reading', () => {
        it('convertRowsToObjects should map 2D grid to objects using row 0 as keys', () => {
            const rows = [
                ['Name', 'Age', 'Role'],
                ['Alice', 30, 'Admin'],
                ['Bob', 25, 'Editor']
            ];
            const objs = convertRowsToObjects(rows);
            expect(objs).to.deep.equal([
                { Name: 'Alice', Age: 30, Role: 'Admin' },
                { Name: 'Bob', Age: 25, Role: 'Editor' }
            ]);
        });

        it('convertRowsToObjects should handle empty, null, or duplicate headers gracefully', () => {
            const rows = [
                ['ID', null, '', 'ID'],
                [1, 'Val1', 'Val2', 'DuplicateID']
            ];
            const objs = convertRowsToObjects(rows);
            expect(objs).to.deep.equal([
                { ID: 1, column2: 'Val1', column3: 'Val2', ID_2: 'DuplicateID' }
            ]);
        });

        it('readSpreadsheet with asObjects: true should return array of objects from CSV', async () => {
            const csvPath = path.join(tmpDir, 'test-objects.csv');
            fs.writeFileSync(csvPath, 'SKU,Price,InStock\nITEM-1,19.99,true\nITEM-2,45.00,false\n', 'utf-8');

            const objs = await readSpreadsheet(csvPath, { asObjects: true });
            expect(objs).to.have.lengthOf(2);
            expect(objs[0].SKU).to.equal('ITEM-1');
            expect(objs[0].Price).to.equal('19.99');
            expect(objs[1].SKU).to.equal('ITEM-2');
        });

        it('readSpreadsheet with header: true should work as alias for asObjects', async () => {
            const csvPath = path.join(tmpDir, 'test-header.csv');
            fs.writeFileSync(csvPath, 'Name,Score\nCharlie,98\nDavid,85\n', 'utf-8');

            const objs = await readSpreadsheet(csvPath, { header: true });
            expect(objs).to.have.lengthOf(2);
            expect(objs[0]).to.deep.equal({ Name: 'Charlie', Score: '98' });
        });

        it('iterateSpreadsheet with asObjects: true should yield row objects', async () => {
            const csvPath = path.join(tmpDir, 'test-iterate-objs.csv');
            fs.writeFileSync(csvPath, 'City,Country\nTokyo,Japan\nParis,France\n', 'utf-8');

            const results = [];
            for await (const row of iterateSpreadsheet(csvPath, { asObjects: true })) {
                results.push(row);
            }

            expect(results).to.deep.equal([
                { City: 'Tokyo', Country: 'Japan' },
                { City: 'Paris', Country: 'France' }
            ]);
        });
    });

    // -------------------------------------------------------------
    // Feature 2: Auto Column Widths with maxWidth Safeguard
    // -------------------------------------------------------------
    describe('2. Auto Column Widths', () => {
        it('calculateAutoWidths should compute column widths with minWidth and padding', () => {
            const rows = [
                ['Short', 'Medium Length Header'],
                ['A', '12345']
            ];
            const widths = calculateAutoWidths(rows, { autoWidth: true, minWidth: 8, padding: 2 });
            expect(widths).to.have.lengthOf(2);
            // Col 0: max length 5 + 2 = 7, minWidth 8 -> 8
            expect(widths[0].width).to.equal(8);
            // Col 1: max length 20 + 2 = 22 -> 22
            expect(widths[1].width).to.equal(22);
        });

        it('calculateAutoWidths should enforce maxWidth safeguard for long text', () => {
            const longText = 'This is a very long description paragraph that could potentially break spreadsheet table layout if not bounded properly by maxWidth.';
            const rows = [
                ['Title', 'Description'],
                ['Item', longText]
            ];
            const widths = calculateAutoWidths(rows, { autoWidth: { maxWidth: 35 } });
            expect(widths[1].width).to.equal(35);
        });

        it('calculateAutoWidths should preserve user explicit column widths unless force: true', () => {
            const rows = [
                ['Col A', 'Col B'],
                ['Content', 'Content']
            ];
            const existing = [{ width: 50 }, {}];
            const widths = calculateAutoWidths(rows, { autoWidth: true }, existing);
            expect(widths[0].width).to.equal(50);
            expect(widths[1].width).to.be.below(50);
        });

        it('createHTML with autoWidth should generate <colgroup> with responsive widths', () => {
            const data = [
                ['Name', 'Bio'],
                ['Alice', 'Developer working on spreadsheets']
            ];
            const html = createHTML(data, { autoWidth: true });
            expect(html).to.include('<colgroup>');
            expect(html).to.include('<col style="width:');
        });
    });

    // -------------------------------------------------------------
    // Feature 3: Ready-to-Use Styling Themes
    // -------------------------------------------------------------
    describe('3. Ready-to-Use Styling Themes', () => {
        it('getTheme should resolve built-in themes by name', () => {
            for (const name of ['ocean', 'forest', 'slate', 'sunset', 'striped']) {
                const theme = getTheme(name);
                expect(theme).to.be.an('object');
                expect(theme.header).to.be.an('object');
                expect(theme.header.backgroundColor).to.be.a('string');
            }
        });

        it('applyTheme should apply header palette and alternating zebra background', () => {
            const rows = [
                ['ID', 'Name'],
                [1, 'Alice'],
                [2, 'Bob'],
                [3, 'Charlie']
            ];
            const themed = applyTheme(rows, 'ocean');
            // Header row
            expect(themed[0][0].backgroundColor).to.equal('#1B365D');
            expect(themed[0][0].color).to.equal('#FFFFFF');
            expect(themed[0][0].fontWeight).to.equal('bold');

            // Row 1 (first data row, rIdx=1): plain
            expect(themed[1][0]).to.equal(1);

            // Row 2 (second data row, rIdx=2): alternating row
            expect(themed[2][0].backgroundColor).to.equal('#F0F4F8');
            expect(themed[2][0].value).to.equal(2);

            // Row 3 (third data row, rIdx=3): plain
            expect(themed[3][0]).to.equal(3);
        });

        it('applyTheme should preserve custom cell status overrides', () => {
            const rows = [
                ['Name', 'Status'],
                ['Task 1', { value: 'Critical', backgroundColor: '#FF0000' }],
                ['Task 2', { value: 'Done', backgroundColor: '#00FF00' }]
            ];
            const themed = applyTheme(rows, 'slate');
            // Row 2 (rIdx=2) would normally get alternate background, but cell's backgroundColor wins
            expect(themed[2][1].backgroundColor).to.equal('#00FF00');
        });

        it('writeSpreadsheet with theme across XLSX, XLS, ODS, HTML should render successfully', async () => {
            const data = [
                ['Product', 'Revenue'],
                ['Widget A', 50000],
                ['Widget B', 75000],
                ['Widget C', 32000]
            ];

            // XLSX
            const xlsxPath = path.join(tmpDir, 'themed.xlsx');
            await writeSpreadsheet(data, { filePath: xlsxPath, theme: 'ocean', autoWidth: true });
            expect(fs.existsSync(xlsxPath)).to.be.true;

            // ODS
            const odsPath = path.join(tmpDir, 'themed.ods');
            await writeSpreadsheet(data, { filePath: odsPath, theme: 'forest', autoWidth: true });
            expect(fs.existsSync(odsPath)).to.be.true;

            // XLS
            const xlsPath = path.join(tmpDir, 'themed.xls');
            await writeSpreadsheet(data, { filePath: xlsPath, theme: 'slate', autoWidth: true });
            expect(fs.existsSync(xlsPath)).to.be.true;

            // HTML
            const htmlPath = path.join(tmpDir, 'themed.html');
            await writeSpreadsheet(data, { filePath: htmlPath, theme: 'sunset', fullDocument: true });
            expect(fs.existsSync(htmlPath)).to.be.true;
            const htmlContent = fs.readFileSync(htmlPath, 'utf-8');
            expect(htmlContent).to.include('background-color:#9C4123'); // Sunset header
            expect(htmlContent).to.include('background-color:#FFF7ED'); // Sunset alternate
        });
    });

    // -------------------------------------------------------------
    // Feature 4: One-Liner Universal File Converter
    // -------------------------------------------------------------
    describe('4. Universal File Converter (convertSpreadsheet)', () => {
        it('should convert CSV to styled XLSX', async () => {
            const csvPath = path.join(tmpDir, 'conv-in.csv');
            const xlsxPath = path.join(tmpDir, 'conv-out.xlsx');
            fs.writeFileSync(csvPath, 'Item,Qty,Price\nApples,50,1.25\nOranges,30,1.80\n', 'utf-8');

            await convertSpreadsheet(csvPath, xlsxPath, { theme: 'ocean', autoWidth: true });
            expect(fs.existsSync(xlsxPath)).to.be.true;

            const readBack = await readSpreadsheet(xlsxPath);
            expect(readBack).to.have.lengthOf(3);
            expect(readBack[0]).to.deep.equal(['Item', 'Qty', 'Price']);
            expect(readBack[1][0]).to.equal('Apples');
        });

        it('should convert XLSX to ODS', async () => {
            const xlsxPath = path.join(tmpDir, 'conv-out.xlsx');
            const odsPath = path.join(tmpDir, 'conv-out.ods');

            await convertSpreadsheet(xlsxPath, odsPath, { theme: 'forest' });
            expect(fs.existsSync(odsPath)).to.be.true;

            const readBack = await readSpreadsheet(odsPath);
            expect(readBack).to.have.lengthOf(3);
            expect(readBack[0]).to.deep.equal(['Item', 'Qty', 'Price']);
        });

        it('should convert ODS to HTML table', async () => {
            const odsPath = path.join(tmpDir, 'conv-out.ods');
            const htmlPath = path.join(tmpDir, 'conv-out.html');

            await convertSpreadsheet(odsPath, htmlPath, { theme: 'slate', fullDocument: true });
            expect(fs.existsSync(htmlPath)).to.be.true;

            const html = fs.readFileSync(htmlPath, 'utf-8');
            expect(html).to.include('<table');
            expect(html).to.include('Apples');
        });

        it('should convert memory buffer to buffer', async () => {
            const csvStr = 'A,B\n10,20\n';
            const csvBuf = Buffer.from(csvStr, 'utf-8');

            const xlsxBuf = await convertSpreadsheet(csvBuf, {
                from: 'csv',
                format: 'xlsx',
                buffer: true,
                theme: 'striped'
            });

            expect(Buffer.isBuffer(xlsxBuf)).to.be.true;
            const readBack = await readSpreadsheet(xlsxBuf);
            expect(readBack[0]).to.deep.equal(['A', 'B']);
            expect(readBack[1]).to.deep.equal(['10', '20']);
        });
    });

    // -------------------------------------------------------------
    // Feature 5: CLI Executable
    // -------------------------------------------------------------
    describe('5. CLI Tool (bin/cli.js)', () => {
        const cliPath = path.resolve('bin/cli.js');

        it('should display help with --help', async () => {
            const { stdout } = await execFileAsync('node', [cliPath, '--help']);
            expect(stdout).to.include('spreadsheet-utils v');
            expect(stdout).to.include('convert <input> <output>');
            expect(stdout).to.include('cat <input>');
            expect(stdout).to.include('info <input>');
        });

        it('should display version with --version', async () => {
            const { stdout } = await execFileAsync('node', [cliPath, '--version']);
            const { version } = JSON.parse(fs.readFileSync(path.resolve('package.json'), 'utf8'));
            expect(stdout).to.include('spreadsheet-utils v' + version);
        });

        it('should inspect file with info command', async () => {
            const testXlsx = path.join(tmpDir, 'conv-out.xlsx');
            const { stdout } = await execFileAsync('node', [cliPath, 'info', testXlsx]);
            expect(stdout).to.include('Spreadsheet Info:');
            expect(stdout).to.include('Detected Format: xlsx');
            expect(stdout).to.include('3 rows, 3 columns');
        });

        it('should print table with cat command', async () => {
            const testCsv = path.join(tmpDir, 'conv-in.csv');
            const { stdout } = await execFileAsync('node', [cliPath, 'cat', testCsv]);
            expect(stdout).to.include('Item');
            expect(stdout).to.include('Apples');
            expect(stdout).to.include('Oranges');
        });

        it('should print json objects with cat --objects', async () => {
            const testCsv = path.join(tmpDir, 'conv-in.csv');
            const { stdout } = await execFileAsync('node', [cliPath, 'cat', testCsv, '--objects']);
            const parsed = JSON.parse(stdout);
            expect(parsed).to.be.an('array');
            expect(parsed[0].Item).to.equal('Apples');
        });

        it('should convert files via cli convert command', async () => {
            const src = path.join(tmpDir, 'conv-in.csv');
            const dst = path.join(tmpDir, 'cli-converted.xlsx');

            const { stdout } = await execFileAsync('node', [cliPath, 'convert', src, dst, '--theme', 'ocean', '--auto-width']);
            expect(stdout).to.include('✔ Converted');
            expect(fs.existsSync(dst)).to.be.true;
        });
    });

    // -------------------------------------------------------------
    // Feature 6: Streaming Exporter for Massive Datasets
    // -------------------------------------------------------------
    describe('6. Streaming Exporter (O(1) memory)', () => {
        it('streamCSV should stream rows from an async generator directly to file', async () => {
            async function* generateRows() {
                yield ['ID', 'Value', 'Notes'];
                for (let i = 1; i <= 100; i++) {
                    yield [i, `Val-${i}`, `Line\nBreak ${i}`];
                }
            }

            const outPath = path.join(tmpDir, 'streamed.csv');
            await streamCSV(generateRows(), outPath);
            expect(fs.existsSync(outPath)).to.be.true;

            const content = fs.readFileSync(outPath, 'utf-8');
            expect(content).to.include('ID,Value,Notes\n');
            expect(content).to.include('1,Val-1,"Line\nBreak 1"');
            expect(content).to.include('100,Val-100,"Line\nBreak 100"');
        });

        it('streamCSV should handle objects and auto-emit header row', async () => {
            async function* generateObjects() {
                yield { sku: 'A1', price: 10, inStock: true };
                yield { sku: 'B2', price: 20, inStock: false };
            }

            const outPath = path.join(tmpDir, 'streamed-objs.csv');
            await streamCSV(generateObjects(), outPath);
            expect(fs.existsSync(outPath)).to.be.true;

            const lines = fs.readFileSync(outPath, 'utf-8').trim().split('\n');
            expect(lines[0]).to.equal('sku,price,inStock');
            expect(lines[1]).to.equal('A1,10,true');
            expect(lines[2]).to.equal('B2,20,false');
        });

        it('streamTSV should stream with tab delimiters', async () => {
            const rows = [
                ['A', 'B'],
                ['1', '2']
            ];
            const outPath = path.join(tmpDir, 'streamed.tsv');
            await streamTSV(rows, outPath);

            const content = fs.readFileSync(outPath, 'utf-8');
            expect(content).to.equal('A\tB\n1\t2\n');
        });

        it('streamJSONL should stream records to JSON Lines format', async () => {
            async function* generateRecords() {
                yield { id: 1, name: 'Alice' };
                yield { id: 2, name: 'Bob' };
            }

            const outPath = path.join(tmpDir, 'streamed.jsonl');
            await streamJSONL(generateRecords(), outPath);
            expect(fs.existsSync(outPath)).to.be.true;

            const lines = fs.readFileSync(outPath, 'utf-8').trim().split('\n');
            expect(lines).to.have.lengthOf(2);
            expect(JSON.parse(lines[0])).to.deep.equal({ id: 1, name: 'Alice' });
            expect(JSON.parse(lines[1])).to.deep.equal({ id: 2, name: 'Bob' });
        });
    });
});
