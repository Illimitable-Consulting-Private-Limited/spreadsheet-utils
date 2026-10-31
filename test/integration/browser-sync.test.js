import { expect } from 'chai';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { randomBytes } from 'node:crypto';
import esbuild from 'esbuild';
import { unzipSync, strFromU8 } from 'fflate';

import * as sync from '../../src/browser-sync.js';
import { readSpreadsheet, registerPlugin } from '../../src/index.js';
import { xlsPlugin, odsPlugin } from '../../src/plugins/index.js';
import { writeSpreadsheet, writeODS, writeXLS } from '../../src/write/index.js';

// `spreadsheet-utils/browser-sync`: the browser entry for pages under a Content Security Policy that blocks `blob:`
// Workers. fflate's async zip()/unzip() start one for larger files; this entry uses zipSync()/unzipSync() only, so
// neither its source bundle nor its built file may mention `Worker`.

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '../..');
const builtFile = path.join(root, 'dist/browser/spreadsheet-utils-sync.min.js');

// rows of random text, so the archive stays large after compression
function bigRows(count) {
    const rows = [['id', 'text', 'amount']];
    for (let i = 1; i <= count; i++) {
        rows.push([i, randomBytes(24).toString('hex'), i * 1.5]);
    }
    return rows;
}

function asCells(rows) {
    return rows.map((row) => row.map((value) => ({ value })));
}

function toArrayBuffer(buffer) {
    return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

describe('browser-sync entry (Integration Tests)', function () {
    this.timeout(60000);

    before(() => {
        registerPlugin(xlsPlugin, odsPlugin);
    });

    describe('no Worker', () => {
        it('is declared in the exports map with types', () => {
            const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
            const entry = pkg.exports['./browser-sync'];
            expect(entry, 'exports["./browser-sync"]').to.be.an('object');
            expect(entry.import).to.be.a('string');
            expect(entry.types).to.be.a('string');
            expect(fs.existsSync(path.join(root, entry.types)), entry.types).to.equal(true);
        });

        it('bundles for the browser without a Worker', async () => {
            const result = await esbuild.build({
                entryPoints: [path.join(root, 'src/browser-sync.js')],
                bundle: true,
                format: 'esm',
                platform: 'browser',
                write: false,
                minify: false,
                logLevel: 'silent'
            });
            const code = result.outputFiles[0].text;
            expect(code).to.not.match(/\bWorker\b/);
            expect(code).to.not.match(/worker_threads/);
        });

        it('builds dist/browser/spreadsheet-utils-sync.min.js without a Worker', () => {
            expect(fs.existsSync(builtFile), builtFile).to.equal(true);
            const code = fs.readFileSync(builtFile, 'utf8');
            expect(code).to.not.match(/\bWorker\b/);
            expect(code).to.match(/spreadsheetUtilsSync/);
        });
    });

    describe('API', () => {
        it('exports the readers and the writer', () => {
            for (const name of ['readSpreadsheet', 'readSheet', 'readXlsxFile', 'parseTabular', 'writeXlsxFile']) {
                expect(sync[name], name).to.be.a('function');
            }
        });
    });

    describe('reading', () => {
        it('reads an .xlsx over 1 MB (fflate would use a Worker for it) to the same rows', async () => {
            const rows = bigRows(25000);
            const buffer = await writeSpreadsheet(asCells(rows), { buffer: true });
            expect(buffer.length, 'archive size').to.be.greaterThan(1024 * 1024);
            const expected = await readSpreadsheet(buffer);
            expect(await sync.readSpreadsheet(toArrayBuffer(buffer))).to.deep.equal(expected);
            expect(await sync.readSpreadsheet(new Uint8Array(buffer))).to.deep.equal(expected);
            expect(await sync.readSpreadsheet(new Blob([buffer]))).to.deep.equal(expected);
        });

        it('readXlsxFile and readSheet give the first sheet\'s rows', async () => {
            const rows = [['a', 'b'], [1, 'x'], [2, 'y']];
            const buffer = await writeSpreadsheet(asCells(rows), { buffer: true });
            const viaReadSheet = await sync.readSheet(new Uint8Array(buffer));
            expect(viaReadSheet).to.deep.equal(rows);
            const viaReadXlsx = await sync.readXlsxFile(new Uint8Array(buffer));
            // read-excel-file compat: rows, or [{sheet, data}] for every sheet
            const data = Array.isArray(viaReadXlsx[0]) ? viaReadXlsx : viaReadXlsx[0].data;
            expect(data).to.deep.equal(rows);
        });

        it('reads .xls and .ods through parseTabular to the same rows as the main entry', async () => {
            const rows = [['a', 'b'], [1, 'x'], [2, 'y']];
            for (const [name, buffer] of [['xls', await writeXLS(rows, {})], ['ods', await writeODS(rows, {})]]) {
                const book = await sync.parseTabular(new Uint8Array(buffer));
                expect(book.worksheets[0].data, name).to.deep.equal(rows);
            }
        });

        it('reads an .ods over 1 MB', async () => {
            const rows = bigRows(25000);
            const buffer = await writeODS(rows, {});
            expect(buffer.length, 'archive size').to.be.greaterThan(1024 * 1024);
            const book = await sync.parseTabular(new Uint8Array(buffer));
            expect(book.worksheets[0].data.length).to.equal(rows.length);
            expect(book.worksheets[0].data[1]).to.deep.equal(rows[1]);
        });
    });

    describe('writing', () => {
        it('writes over 1 MB synchronously and reads back the same rows', async () => {
            const rows = bigRows(25000);
            const writer = sync.writeXlsxFile(asCells(rows));
            const bytes = writer.toUint8Array();
            expect(bytes, 'toUint8Array is synchronous').to.be.instanceOf(Uint8Array);
            expect(bytes.length).to.be.greaterThan(1024 * 1024);
            expect(await readSpreadsheet(Buffer.from(bytes))).to.deep.equal(rows);
            const blob = await writer.toBlob();
            expect(blob).to.be.instanceOf(Blob);
            expect(blob.size).to.equal(bytes.length);
        });

        it('takes the multi-sheet form and rowBreaks', async () => {
            const bytes = sync.writeXlsxFile([
                { sheet: 'One', data: asCells([['a'], ['b'], ['c']]), rowBreaks: [2] },
                { sheet: 'Two', data: asCells([['z']]) }
            ]).toUint8Array();
            const files = unzipSync(bytes);
            expect(strFromU8(files['xl/worksheets/sheet1.xml'])).to.match(/<brk\b[^>]*\bid="2"/);
            expect(strFromU8(files['xl/workbook.xml'])).to.match(/name="Two"/);
        });
    });
});
