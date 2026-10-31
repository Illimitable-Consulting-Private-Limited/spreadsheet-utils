import { expect } from 'chai';
import { unzipSync, strFromU8 } from 'fflate';

import { writeSpreadsheet, writeXlsxFile } from '../../src/write/index.js';
import { readSpreadsheet } from '../../src/index.js';

// `rowBreaks`: manual page breaks after the given 1-based rows, written the way Excel (and ExcelJS's
// `row.addPageBreak()`) writes them, so a sheet printed by a group starts each group on a new page.

function rows(count) {
    const data = [];
    for (let i = 1; i <= count; i++) {
        data.push([{ value: 'row ' + i }, { value: i }]);
    }
    return data;
}

function sheetXml(buffer, sheetNumber = 1) {
    const files = unzipSync(new Uint8Array(buffer));
    const name = 'xl/worksheets/sheet' + sheetNumber + '.xml';
    expect(files, name + ' in the archive').to.have.property(name);
    return strFromU8(files[name]);
}

function breakIds(xml) {
    return [...xml.matchAll(/<brk\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
}

describe('XLSX writer: rowBreaks (Integration Tests)', () => {

    it('writes one <brk> per row, in Excel\'s shape', async () => {
        const buffer = await writeSpreadsheet(rows(10), { rowBreaks: [3, 7], buffer: true });
        const xml = sheetXml(buffer);
        const element = xml.match(/<rowBreaks\b[^>]*>.*?<\/rowBreaks>/);
        expect(element, '<rowBreaks> element').to.not.equal(null);
        expect(element[0]).to.match(/<rowBreaks\b[^>]*\bcount="2"/);
        expect(element[0]).to.match(/<rowBreaks\b[^>]*\bmanualBreakCount="2"/);
        const brks = element[0].match(/<brk\b[^>]*\/>/g);
        expect(brks).to.have.length(2);
        for (const brk of brks) {
            expect(brk).to.match(/\bmax="16383"/);
            expect(brk).to.match(/\bman="1"/);
        }
        expect(breakIds(xml)).to.deep.equal([3, 7]);
    });

    it('places <rowBreaks> after <pageSetup> and before <drawing>, as the schema orders them', async () => {
        const buffer = await writeSpreadsheet(rows(5), { rowBreaks: [2], buffer: true });
        const xml = sheetXml(buffer);
        const at = (tag) => xml.indexOf('<' + tag);
        expect(at('rowBreaks')).to.be.greaterThan(at('sheetData'));
        expect(at('rowBreaks')).to.be.greaterThan(at('pageMargins'));
        if (at('pageSetup') !== -1) {
            expect(at('rowBreaks')).to.be.greaterThan(at('pageSetup'));
        }
        if (at('drawing') !== -1) {
            expect(at('rowBreaks')).to.be.lessThan(at('drawing'));
        }
    });

    it('comes after <mergeCells> when the sheet has merges', async () => {
        const data = rows(6);
        data[0] = [{ value: 'title', columnSpan: 2 }, null];
        const buffer = await writeSpreadsheet(data, { rowBreaks: [4], buffer: true });
        const xml = sheetXml(buffer);
        expect(xml.indexOf('<mergeCells')).to.be.greaterThan(-1);
        expect(xml.indexOf('<rowBreaks')).to.be.greaterThan(xml.indexOf('<mergeCells'));
    });

    it('sorts and de-duplicates the rows', async () => {
        const buffer = await writeSpreadsheet(rows(12), { rowBreaks: [9, 2, 9, 5], buffer: true });
        const xml = sheetXml(buffer);
        expect(breakIds(xml)).to.deep.equal([2, 5, 9]);
        expect(xml).to.match(/<rowBreaks\b[^>]*\bcount="3"/);
    });

    it('writes no <rowBreaks> when the option is missing or empty', async () => {
        for (const options of [{}, { rowBreaks: [] }]) {
            const buffer = await writeSpreadsheet(rows(3), { ...options, buffer: true });
            expect(sheetXml(buffer)).to.not.contain('<rowBreaks');
        }
    });

    it('throws on a row that is not a positive integer', async () => {
        for (const bad of [[0], [-1], [2.5], ['3'], [NaN]]) {
            let error;
            try {
                await writeSpreadsheet(rows(3), { rowBreaks: bad, buffer: true });
            } catch (e) {
                error = e;
            }
            expect(error, JSON.stringify(bad) + ' should throw').to.be.an('error');
            expect(error.message).to.match(/rowBreaks/);
        }
    });

    it('applies per sheet in the multi-sheet form', async () => {
        const buffer = await writeXlsxFile([
            { sheet: 'One', data: rows(6), rowBreaks: [2] },
            { sheet: 'Two', data: rows(6) },
            { sheet: 'Three', data: rows(6), rowBreaks: [1, 4] }
        ]).toBuffer();
        expect(breakIds(sheetXml(buffer, 1))).to.deep.equal([2]);
        expect(sheetXml(buffer, 2)).to.not.contain('<rowBreaks');
        expect(breakIds(sheetXml(buffer, 3))).to.deep.equal([1, 4]);
    });

    it('passes through writeXlsxFile(data, options) for a single sheet', async () => {
        const buffer = await writeXlsxFile(rows(4), { rowBreaks: [3] }).toBuffer();
        expect(breakIds(sheetXml(buffer))).to.deep.equal([3]);
    });

    it('leaves the cells as they were', async () => {
        const buffer = await writeSpreadsheet(rows(5), { rowBreaks: [2, 4], buffer: true });
        const read = await readSpreadsheet(buffer);
        expect(read).to.deep.equal([1, 2, 3, 4, 5].map((i) => ['row ' + i, i]));
    });
});
