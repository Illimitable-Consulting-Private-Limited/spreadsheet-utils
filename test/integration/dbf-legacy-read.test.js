import { expect } from 'chai';

import { parseDBF, parseDBFBuffer } from '../../src/parser/dbf.js';

// Read options that give the values yortus/DBFFile 1.x gives (what legacy DBF importers were written against):
// text decoded as latin1, only trailing spaces cut, an empty text field as '', a D field as a Date at UTC midnight.
const LEGACY_READ = { encoding: 'latin1', trimText: 'end', emptyText: '', dateType: 'date' };

const FIELDS = [
    { name: 'CODE', type: 'C', size: 6 },
    { name: 'QTY', type: 'N', size: 8, decimalPlaces: 2 },
    { name: 'DATE', type: 'D', size: 8 },
    { name: 'NAME', type: 'C', size: 6 }
];

// A dBase III file built by hand: header, descriptors, 0x0D 0x00, records (each a deletion flag and the field
// texts, already padded to their sizes), 0x1A.
function dbf(fields, records) {
    const recordLength = 1 + fields.reduce((sum, f) => sum + f.size, 0);
    const headerLength = 32 + fields.length * 32 + 2;
    const header = Buffer.alloc(32);
    header[0] = 0x03;
    header[1] = 126;
    header[2] = 11;
    header[3] = 12;
    header.writeUInt32LE(records.length, 4);
    header.writeUInt16LE(headerLength, 8);
    header.writeUInt16LE(recordLength, 10);
    const descriptors = fields.map((field) => {
        const d = Buffer.alloc(32);
        d.write(field.name, 0, 'latin1');
        d[11] = field.type.charCodeAt(0);
        d[16] = field.size;
        d[17] = field.decimalPlaces || 0;
        d[20] = 0x01;
        return d;
    });
    const rows = records.map(({ deleted, texts }) => {
        const bytes = texts.map((text) => (Buffer.isBuffer(text) ? text : Buffer.from(text, 'latin1')));
        bytes.forEach((b, i) => expect(b.length, `field ${fields[i].name}`).to.equal(fields[i].size));
        return Buffer.concat([Buffer.from([deleted ? 0x2A : 0x20]), ...bytes]);
    });
    return Buffer.concat([header, ...descriptors, Buffer.from([0x0D, 0x00]), ...rows, Buffer.from([0x1A])]);
}

const FILE = dbf(FIELDS, [
    { texts: ['  12  ', '   12.50', '20260401', 'Ab    '] },
    { texts: ['      ', '        ', '        ', '      '] },
    // 0xE9 is é in latin1 and not valid UTF-8 on its own; 0xC3 0xA9 is é in UTF-8 but Ã© in latin1
    { texts: [Buffer.from([0x43, 0x61, 0x66, 0xE9, 0x20, 0x20]), '   -0.50', '20251231', Buffer.from([0xC3, 0xA9, 0x20, 0x20, 0x20, 0x20])] },
    { deleted: true, texts: ['GONE  ', '    1.00', '20260101', 'x     '] }
]);

function rows(result) {
    return result.worksheets[0].data;
}

describe('DBF legacy read options (encoding, trimText, emptyText, dateType)', () => {

    describe('defaults are unchanged', () => {
        it('text trimmed both sides, empty values null, dates as YYYY-MM-DD, UTF-8 tried first', () => {
            const data = rows(parseDBFBuffer(FILE));
            expect(data[0]).to.deep.equal(['CODE', 'QTY', 'DATE', 'NAME']);
            expect(data[1]).to.deep.equal(['12', 12.5, '2026-04-01', 'Ab']);
            expect(data[2]).to.deep.equal([null, null, null, null]);
            expect(data[3][2]).to.equal('2025-12-31');
            expect(data[3][3]).to.equal('é');
        });

        it('deleted records are left out', () => {
            expect(rows(parseDBFBuffer(FILE))).to.have.length(4);
        });
    });

    describe('all four together (the dbffile 1.x reading)', () => {
        it('gives the values dbffile gives', () => {
            const data = rows(parseDBFBuffer(FILE, LEGACY_READ));
            expect(data[0]).to.deep.equal(['CODE', 'QTY', 'DATE', 'NAME']);
            expect(data[1]).to.deep.equal(['  12', 12.5, new Date(Date.UTC(2026, 3, 1)), 'Ab']);
            expect(data[2]).to.deep.equal(['', null, null, '']);
            expect(data[3]).to.deep.equal(['Café', -0.5, new Date(Date.UTC(2025, 11, 31)), 'Ã©']);
            expect(data).to.have.length(4);
        });

        it('parseDBF passes the options through', async () => {
            const data = rows(await parseDBF(FILE, LEGACY_READ));
            expect(data[1][0]).to.equal('  12');
            expect(data[1][2]).to.be.instanceOf(Date);
        });
    });

    describe('encoding', () => {
        it("'latin1' decodes one byte a character", () => {
            const data = rows(parseDBFBuffer(FILE, { encoding: 'latin1' }));
            expect(data[3][0]).to.equal('Café');
            expect(data[3][3]).to.equal('Ã©');
        });

        it("'utf-8' is the default's first try", () => {
            const data = rows(parseDBFBuffer(FILE, { encoding: 'utf-8' }));
            expect(data[3][3]).to.equal('é');
        });

        it('field names are unaffected', () => {
            expect(rows(parseDBFBuffer(FILE, { encoding: 'latin1' }))[0]).to.deep.equal(['CODE', 'QTY', 'DATE', 'NAME']);
        });
    });

    describe('trimText', () => {
        it("'end' cuts trailing spaces only and keeps leading ones", () => {
            const data = rows(parseDBFBuffer(FILE, { trimText: 'end' }));
            expect(data[1][0]).to.equal('  12');
            expect(data[1][3]).to.equal('Ab');
        });

        it("'both' is the default", () => {
            expect(rows(parseDBFBuffer(FILE, { trimText: 'both' }))[1][0]).to.equal('12');
        });

        it('numbers and dates still read with leading spaces', () => {
            const data = rows(parseDBFBuffer(FILE, { trimText: 'end' }));
            expect(data[1][1]).to.equal(12.5);
            expect(data[3][1]).to.equal(-0.5);
            expect(data[2][1]).to.equal(null);
        });
    });

    describe('emptyText', () => {
        it("'' gives an all-space C field as '' (other types stay null)", () => {
            const data = rows(parseDBFBuffer(FILE, { emptyText: '' }));
            expect(data[2]).to.deep.equal(['', null, null, '']);
        });

        it('null is the default', () => {
            expect(rows(parseDBFBuffer(FILE, { emptyText: null }))[2][0]).to.equal(null);
        });
    });

    describe('dateType', () => {
        it("'date' gives a Date at UTC midnight, a blank date null", () => {
            const data = rows(parseDBFBuffer(FILE, { dateType: 'date' }));
            expect(data[1][2]).to.be.instanceOf(Date);
            expect(data[1][2].toISOString()).to.equal('2026-04-01T00:00:00.000Z');
            expect(data[2][2]).to.equal(null);
        });

        it("'string' is the default", () => {
            expect(rows(parseDBFBuffer(FILE, { dateType: 'string' }))[1][2]).to.equal('2026-04-01');
        });

        it("a D field that is not a date is null with 'date'", () => {
            const file = dbf([{ name: 'D', type: 'D', size: 8 }], [{ texts: ['2026XX01'] }]);
            expect(rows(parseDBFBuffer(file, { dateType: 'date' }))[1][0]).to.equal(null);
        });
    });

    describe('bad option values throw', () => {
        const BAD = [
            ['encoding', 'cp1252'],
            ['trimText', 'start'],
            ['trimText', true],
            ['emptyText', 'x'],
            ['dateType', 'iso'],
        ];
        for (const [name, value] of BAD) {
            it(`${name}: ${JSON.stringify(value)}`, () => {
                expect(() => parseDBFBuffer(FILE, { [name]: value })).to.throw(new RegExp(`Invalid DBF option ${name}`));
            });
        }
    });
});
