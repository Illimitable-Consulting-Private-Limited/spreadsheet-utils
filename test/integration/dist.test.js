import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const testCasesDir = path.resolve(__dirname, '../read-excel-file/testCases');

describe('Distribution Builds Verification', () => {

    describe('ESM Builds (dist/esm/)', () => {
        it('should load core reader from dist/esm/index.mjs and parse XLSX', async () => {
            const { readSpreadsheet, readXlsxFile, parseTabular } = await import('../../dist/esm/index.mjs');
            expect(readSpreadsheet).to.be.a('function');
            expect(readXlsxFile).to.be.a('function');
            expect(parseTabular).to.be.a('function');

            const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
            const rows = await readSpreadsheet(booleanXlsx);
            expect(rows[0]).to.deep.equal(['Boolean']);
        });

        it('should load plugins from dist/esm/plugins.mjs and register textPlugin', async () => {
            const { readSpreadsheet, registerPlugin } = await import('../../dist/esm/index.mjs');
            const { textPlugin } = await import('../../dist/esm/plugins.mjs');

            registerPlugin(textPlugin);
            const csv = new TextEncoder().encode('Name,Age\nAlice,30');
            const rows = await readSpreadsheet(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['Name', 'Age'], ['Alice', '30']]);
        });

        it('should load writer from dist/esm/write.mjs and generate XLSX buffer', async () => {
            const { writeSpreadsheet } = await import('../../dist/esm/write.mjs');
            const { readSpreadsheet } = await import('../../dist/esm/index.mjs');

            const buf = await writeSpreadsheet([[{ value: 'ESM Write', fontWeight: 'bold' }]], { buffer: true });
            expect(buf).to.be.an.instanceOf(Uint8Array);

            const rows = await readSpreadsheet(buf);
            expect(rows).to.deep.equal([['ESM Write']]);
        });

        it('should load all-in-one preset from dist/esm/full.mjs', async () => {
            const { readSpreadsheet, writeSpreadsheet, writeCSV, writeJSON } = await import('../../dist/esm/full.mjs');

            // Reads CSV out-of-the-box (pre-registered)
            const csv = new TextEncoder().encode('Col1,Col2\n100,200');
            const rows = await readSpreadsheet(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['Col1', 'Col2'], ['100', '200']]);

            // Writes styled XLSX
            const buf = await writeSpreadsheet(rows, { buffer: true });
            expect(buf.length).to.be.greaterThan(0);

            // Writes CSV & JSON
            const csvOut = await writeCSV(rows);
            expect(csvOut).to.include('Col1,Col2');
            const jsonOut = await writeJSON(rows);
            expect(jsonOut).to.include('100');
        });

        it('should load worker offloader from dist/esm/worker.mjs', async () => {
            const { readInWorker } = await import('../../dist/esm/worker.mjs');
            expect(readInWorker).to.be.a('function');

            const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
            const rows = await readInWorker(booleanXlsx);
            expect(rows[0]).to.deep.equal(['Boolean']);
        });
    });

    describe('CommonJS Builds (dist/cjs/)', () => {
        it('should require core reader from dist/cjs/index.cjs', async () => {
            const cjs = require('../../dist/cjs/index.cjs');
            expect(cjs.readSpreadsheet).to.be.a('function');

            const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
            const rows = await cjs.readSpreadsheet(booleanXlsx);
            expect(rows[0]).to.deep.equal(['Boolean']);
        });

        it('should require plugins from dist/cjs/plugins.cjs and register textPlugin', async () => {
            const { readSpreadsheet, registerPlugin } = require('../../dist/cjs/index.cjs');
            const { textPlugin } = require('../../dist/cjs/plugins.cjs');

            registerPlugin(textPlugin);
            const csv = Buffer.from('ColA,ColB\n1,2');
            const rows = await readSpreadsheet(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['ColA', 'ColB'], ['1', '2']]);
        });

        it('should require writer from dist/cjs/write.cjs', async () => {
            const { writeSpreadsheet } = require('../../dist/cjs/write.cjs');
            const { readSpreadsheet } = require('../../dist/cjs/index.cjs');

            const buf = await writeSpreadsheet([[{ value: 'CJS Write' }]], { buffer: true });
            const rows = await readSpreadsheet(buf);
            expect(rows).to.deep.equal([['CJS Write']]);
        });

        it('should require full bundle from dist/cjs/full.cjs', async () => {
            const full = require('../../dist/cjs/full.cjs');
            expect(full.readSpreadsheet).to.be.a('function');
            expect(full.writeSpreadsheet).to.be.a('function');
            expect(full.writeCSV).to.be.a('function');
            expect(full.writeJSON).to.be.a('function');

            const csv = Buffer.from('A\n10');
            const rows = await full.readSpreadsheet(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['A'], ['10']]);
            const csvOut = await full.writeCSV(rows);
            expect(csvOut).to.include('A');
        });

        it('should require worker offloader from dist/cjs/worker.cjs', async () => {
            const worker = require('../../dist/cjs/worker.cjs');
            expect(worker.readInWorker).to.be.a('function');

            const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
            const rows = await worker.readInWorker(booleanXlsx);
            expect(rows[0]).to.deep.equal(['Boolean']);
        });
    });

    describe('Browser Standalone Bundles (dist/browser/)', () => {
        it('should have valid browser standalone bundle files', () => {
            const browserFiles = [
                'spreadsheet-utils.min.js',
                'spreadsheet-utils-plugins.min.js',
                'spreadsheet-utils-write.min.js',
                'spreadsheet-utils-full.min.js'
            ];

            for (const file of browserFiles) {
                const fullPath = path.resolve('dist/browser', file);
                expect(fs.existsSync(fullPath)).to.be.true;
                const stats = fs.statSync(fullPath);
                expect(stats.size).to.be.greaterThan(1000);
            }
        });
    });
});
