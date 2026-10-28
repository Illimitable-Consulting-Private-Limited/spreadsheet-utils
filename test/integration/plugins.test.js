import { expect } from 'chai';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

import {
    readSpreadsheet,
    createReader,
    registerPlugin,
    unregisterPlugin,
    isPluginRegistered
} from '../../src/index.js';

import {
    textPlugin,
    xlsPlugin,
    markupPlugin,
    legacyPlugin,
    allPlugins
} from '../../src/plugins/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const testCasesDir = path.resolve(__dirname, '../read-excel-file/testCases');

describe('Plugin Architecture & Extension System', () => {

    beforeEach(() => {
        // Reset registered non-core plugins before each test
        unregisterPlugin();
    });

    afterEach(() => {
        // Clean up
        unregisterPlugin();
    });

    describe('Core Reader isolation', () => {
        it('should read XLSX out of the box without any plugins', async () => {
            const booleanXlsx = path.join(testCasesDir, 'boolean.xlsx');
            const rows = await readSpreadsheet(booleanXlsx);
            expect(rows).to.be.an('array');
            expect(rows[0]).to.deep.equal(['Boolean']);
        });

        it('should throw helpful missing-plugin error when reading CSV without textPlugin', async () => {
            const csv = new TextEncoder().encode('A,B\n1,2');
            try {
                await readSpreadsheet(csv, { extension: 'csv' });
                expect.fail('Should have thrown missing plugin error');
            } catch (err) {
                expect(err.message).to.include('textPlugin');
                expect(err.message).to.include('registerPlugin(textPlugin)');
            }
        });

        it('should throw helpful missing-plugin error when reading XLS without xlsPlugin', async () => {
            try {
                await readSpreadsheet('test.xls');
                expect.fail('Should have thrown');
            } catch (err) {
                expect(err.message).to.include('xlsPlugin');
            }
        });

        it('should throw helpful missing-plugin error when reading ODS without markupPlugin', async () => {
            try {
                await readSpreadsheet('test.ods');
                expect.fail('Should have thrown');
            } catch (err) {
                expect(err.message).to.include('markupPlugin');
            }
        });

        it('should throw helpful missing-plugin error when reading DBF without legacyPlugin', async () => {
            try {
                await readSpreadsheet('test.dbf');
                expect.fail('Should have thrown');
            } catch (err) {
                expect(err.message).to.include('legacyPlugin');
            }
        });
    });

    describe('Dynamic plugin registration', () => {
        it('should enable CSV reading after registering textPlugin', async () => {
            registerPlugin(textPlugin);
            expect(isPluginRegistered('textPlugin')).to.be.true;

            const csv = new TextEncoder().encode('Name,Score\nAlice,95\nBob,88');
            const rows = await readSpreadsheet(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([
                ['Name', 'Score'],
                ['Alice', '95'],
                ['Bob', '88']
            ]);
        });

        it('should enable markup formats after registering markupPlugin', async () => {
            registerPlugin(markupPlugin);
            const html = '<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>';
            const rows = await readSpreadsheet(html, { extension: 'html', firstRowAsHeader: false });
            expect(rows).to.deep.equal([
                ['A', 'B'],
                ['1', '2']
            ]);
        });

        it('should support unregistering a plugin', async () => {
            registerPlugin(textPlugin);
            expect(isPluginRegistered('textPlugin')).to.be.true;

            unregisterPlugin('textPlugin');
            expect(isPluginRegistered('textPlugin')).to.be.false;

            const csv = new TextEncoder().encode('A,B\n1,2');
            try {
                await readSpreadsheet(csv, { extension: 'csv' });
                expect.fail('Should have thrown');
            } catch (err) {
                expect(err.message).to.include('textPlugin');
            }
        });
    });

    describe('Scoped reader factory (createReader)', () => {
        it('should create an isolated reader with specific plugins without global pollution', async () => {
            // Global registry has no plugins
            expect(isPluginRegistered('textPlugin')).to.be.false;

            const scopedReader = createReader([textPlugin]);
            const csv = new TextEncoder().encode('X,Y\n10,20');

            // Scoped reader works
            const rows = await scopedReader(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['X', 'Y'], ['10', '20']]);

            // Global reader still rejects CSV
            try {
                await readSpreadsheet(csv, { extension: 'csv' });
                expect.fail('Global reader should still reject');
            } catch (err) {
                expect(err.message).to.include('textPlugin');
            }
        });

        it('should support chaining .register() on scoped reader', async () => {
            const scoped = createReader();
            scoped.register(textPlugin);

            const csv = new TextEncoder().encode('A\n1');
            const rows = await scoped(csv, { extension: 'csv' });
            expect(rows).to.deep.equal([['A'], ['1']]);
        });
    });
});
