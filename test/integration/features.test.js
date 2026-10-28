import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    readSpreadsheet,
    iterateSpreadsheet,
    registerPlugin
} from '../../src/index.js';

import {
    writeSpreadsheet,
    writeCSV,
    writeTSV,
    writeHTML,
    writeJSON
} from '../../src/write/index.js';

import { readInWorker } from '../../src/worker/index.js';
import { textPlugin } from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const testCasesDir = path.resolve(__dirname, '../read-excel-file/testCases');
const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
const tmpDir = path.resolve(__dirname, '../../tmp-features');

describe('Advanced Enterprise Features (Integration Tests)', () => {

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

    describe('Feature 1: Streaming Row Consumer (iterateSpreadsheet & onRow)', () => {
        it('should yield rows asynchronously with iterateSpreadsheet (XLSX)', async () => {
            const collected = [];
            for await (const row of iterateSpreadsheet(booleanXlsx)) {
                collected.push(row);
            }

            expect(collected.length).to.equal(5);
            expect(collected[0]).to.deep.equal(['Boolean']);
            expect(collected[1]).to.deep.equal([true]);
        });

        it('should yield rows asynchronously with iterateSpreadsheet (CSV)', async () => {
            const csv = new TextEncoder().encode('A,B\n1,2\n3,4');
            const collected = [];
            for await (const row of iterateSpreadsheet(csv, { extension: 'csv' })) {
                collected.push(row);
            }

            expect(collected).to.deep.equal([
                ['A', 'B'],
                ['1', '2'],
                ['3', '4']
            ]);
        });

        it('should invoke onRow callback during readSpreadsheet', async () => {
            const seenRows = [];
            const rows = await readSpreadsheet(booleanXlsx, {
                onRow: (row, index) => {
                    seenRows.push({ row, index });
                }
            });

            expect(seenRows.length).to.equal(rows.length);
            expect(seenRows[0].row).to.deep.equal(['Boolean']);
        });
    });

    describe('Feature 2: Multi-Format Exporter (CSV, TSV, HTML, JSON)', () => {
        const sampleData = [
            [
                { value: 'Product', fontWeight: 'bold', backgroundColor: '#e2e8f0' },
                { value: 'Price', fontWeight: 'bold', backgroundColor: '#e2e8f0' }
            ],
            [
                { value: 'Widget, Standard', type: String },
                { value: 19.99, type: Number }
            ],
            [
                { value: 'Gadget "Pro"', type: String },
                { value: 49.95, type: Number }
            ]
        ];

        it('should export RFC 4180 compliant CSV', async () => {
            const csv = await writeCSV(sampleData);
            expect(csv).to.include('"Widget, Standard",19.99');
            expect(csv).to.include('"Gadget ""Pro""",49.95');

            // Roundtrip verify by reading back with CSV parser
            const readBack = await readSpreadsheet(csv, { extension: 'csv' });
            expect(readBack[1][0]).to.equal('Widget, Standard');
            expect(readBack[2][0]).to.equal('Gadget "Pro"');
        });

        it('should export TSV with tab delimiters', async () => {
            const tsv = await writeTSV(sampleData);
            expect(tsv).to.include('Product\tPrice');
            expect(tsv).to.include('19.99');
        });

        it('should export HTML table with inline styles', async () => {
            const html = await writeHTML(sampleData, { fullDocument: true });
            expect(html).to.include('<!DOCTYPE html>');
            expect(html).to.include('<table class="spreadsheet-table">');
            expect(html).to.include('background-color:#e2e8f0');
            expect(html).to.include('Widget, Standard');
        });

        it('should export JSON array of rows or objects', async () => {
            const jsonRows = await writeJSON(sampleData);
            const parsed = JSON.parse(jsonRows);
            expect(parsed.length).to.equal(3);
            expect(parsed[0]).to.deep.equal(['Product', 'Price']);

            const jsonObjects = await writeJSON(sampleData, { asObjects: true });
            const parsedObjs = JSON.parse(jsonObjects);
            expect(parsedObjs.length).to.equal(2);
            expect(parsedObjs[0]).to.deep.equal({ Product: 'Widget, Standard', Price: 19.99 });
        });

        it('should route writeSpreadsheet automatically based on filePath extension or format option', async () => {
            const csvFile = path.join(tmpDir, 'auto-routed.csv');
            await writeSpreadsheet(sampleData, { filePath: csvFile });
            expect(fs.existsSync(csvFile)).to.be.true;

            const content = fs.readFileSync(csvFile, 'utf-8');
            expect(content).to.include('"Widget, Standard",19.99');

            const htmlFile = path.join(tmpDir, 'auto-routed.html');
            await writeSpreadsheet(sampleData, { filePath: htmlFile });
            expect(fs.existsSync(htmlFile)).to.be.true;
            const htmlContent = fs.readFileSync(htmlFile, 'utf-8');
            expect(htmlContent).to.include('<table');
        });
    });

    describe('Feature 3: Security & Input Limits Hardening', () => {
        it('should reject files exceeding maxFileSize', async () => {
            const dummyBuffer = new Uint8Array(1024 * 100); // 100 KB
            try {
                await readSpreadsheet(dummyBuffer, { extension: 'csv', maxFileSize: 1024 * 50 }); // 50 KB cap
                expect.fail('Should have rejected exceeding maxFileSize');
            } catch (err) {
                expect(err.message).to.include('File size limit');
            }
        });

        it('should enforce maxRows limit', async () => {
            const csv = new TextEncoder().encode('A\n1\n2\n3\n4\n5\n6\n7\n8\n9\n10');
            const rows = await readSpreadsheet(csv, { extension: 'csv', maxRows: 3 });
            expect(rows.length).to.equal(3);
            expect(rows).to.deep.equal([['A'], ['1'], ['2']]);
        });

        it('should enforce maxColumns limit', async () => {
            const csv = new TextEncoder().encode('1,2,3,4,5,6,7,8');
            const rows = await readSpreadsheet(csv, { extension: 'csv', maxColumns: 4 });
            expect(rows[0].length).to.equal(4);
            expect(rows[0]).to.deep.equal(['1', '2', '3', '4']);
        });

        it('should reject zip bomb with maxUncompressedSize limit', async () => {
            // Test with boolean.xlsx and an artificially tiny maxUncompressedSize (e.g. 50 bytes)
            try {
                await readSpreadsheet(booleanXlsx, { maxUncompressedSize: 50 });
                expect.fail('Should have thrown zip bomb limit error');
            } catch (err) {
                expect(err.message).to.include('Zip decompression limit');
            }
        });
    });

    describe('Feature 4: Background Thread & Worker Offloader', () => {
        it('should parse spreadsheet offloaded to background worker thread', async () => {
            const rows = await readInWorker(booleanXlsx);
            expect(rows).to.be.an('array');
            expect(rows.length).to.equal(5);
            expect(rows[0]).to.deep.equal(['Boolean']);
        });
    });
});
