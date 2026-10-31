import { expect } from 'chai';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

import { createDBF, writeDBF } from '../../src/write/index.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

// Options that reproduce the byte layout of yortus/DBFFile 0.7 (what legacy FoxPro consumers have been reading),
// apart from the header date, which stays the real month.
const LEGACY = {
    numberText: 'raw',
    headerNullByte: true,
    workAreaId: 1,
    encoding: 'latin1',
    languageDriver: 0,
    strictEncoding: true
};

const FIELDS = [
    { name: 'NAME', type: 'C', size: 5 },
    // decimalPlaces stated: left out, su infers it from the values (2 here) and writes that in the descriptor.
    { name: 'QTY', type: 'N', size: 6, decimalPlaces: 0 },
    { name: 'ATT', type: 'N', size: 3, decimalPlaces: 1 }
];

const RECORDS = [
    { NAME: 'Ab', QTY: 1.25, ATT: 12 },
    { NAME: null, QTY: null, ATT: 0.5 },
    { NAME: 'Toolong', QTY: 1234567.5, ATT: -1 }
];

const OCT_31 = new Date(2026, 9, 31);

function ascii(buffer, start, length) {
    return Buffer.from(buffer.subarray(start, start + length)).toString('latin1');
}

function headerLength(buffer) {
    return buffer[8] | (buffer[9] << 8);
}

function recordLength(buffer) {
    return buffer[10] | (buffer[11] << 8);
}

// The text of field `f` in record `r`, by the sizes in FIELDS.
function fieldText(buffer, fields, r, f) {
    let offset = headerLength(buffer) + r * recordLength(buffer) + 1;
    for (let i = 0; i < f; i++) offset += fields[i].size;
    return ascii(buffer, offset, fields[f].size);
}

// The whole file, written out by hand: dBase III header, three descriptors, 0x0D 0x00, three records, 0x1A.
function expectedLegacyFile() {
    const header = Buffer.alloc(32);
    header[0] = 0x03;
    header[1] = 126;
    header[2] = 10;
    header[3] = 31;
    header.writeUInt32LE(3, 4);
    header.writeUInt16LE(32 + 3 * 32 + 2, 8);
    header.writeUInt16LE(1 + 5 + 6 + 3, 10);
    const descriptors = FIELDS.map((field) => {
        const d = Buffer.alloc(32);
        d.write(field.name, 0, 'latin1');
        d[11] = field.type.charCodeAt(0);
        d[16] = field.size;
        d[17] = field.decimalPlaces || 0;
        d[20] = 0x01;
        return d;
    });
    const records = Buffer.from(
        ' ' + 'Ab   ' + '  1.25' + ' 12' +
        ' ' + '     ' + '      ' + '0.5' +
        ' ' + 'Toolo' + '123456' + ' -1',
        'latin1'
    );
    return Buffer.concat([header, ...descriptors, Buffer.from([0x0D, 0x00]), records, Buffer.from([0x1A])]);
}

describe('DBF legacy layout options (numberText, headerNullByte, workAreaId, strictEncoding)', () => {

    describe('defaults are unchanged', () => {
        it('header is 33 + 32n bytes, work area 0, N values rounded to the decimal places', () => {
            const buffer = createDBF(RECORDS, { fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            expect(headerLength(buffer)).to.equal(32 + 3 * 32 + 1);
            expect(buffer[headerLength(buffer) - 1]).to.equal(0x0D);
            for (let i = 0; i < FIELDS.length; i++) {
                expect(buffer[32 + i * 32 + 20]).to.equal(0x00);
            }
            expect(fieldText(buffer, FIELDS, 0, 1)).to.equal('     1');
            // "12.0" does not fit N(3,1)
            expect(fieldText(buffer, FIELDS, 0, 2)).to.equal('***');
        });

        it('a number wider than its field is filled with *', () => {
            const buffer = createDBF(RECORDS, { fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            expect(fieldText(buffer, FIELDS, 2, 1)).to.equal('******');
        });

        it('header month is the real month (October = 10)', () => {
            const buffer = createDBF(RECORDS, { fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            expect([buffer[1], buffer[2], buffer[3]]).to.deep.equal([126, 10, 31]);
        });
    });

    describe('headerNullByte', () => {
        it('writes 0x00 after the 0x0D terminator and counts it in the header length', () => {
            const buffer = createDBF(RECORDS, { fieldDescriptors: FIELDS, headerNullByte: true });
            const length = headerLength(buffer);
            expect(length).to.equal(32 + 3 * 32 + 2);
            expect(buffer[length - 2]).to.equal(0x0D);
            expect(buffer[length - 1]).to.equal(0x00);
            expect(buffer[length]).to.equal(0x20);
            expect(buffer.length).to.equal(length + 3 * recordLength(buffer) + 1);
            expect(buffer[buffer.length - 1]).to.equal(0x1A);
        });

        it('with no records the file is the header and 0x1A', () => {
            const buffer = createDBF([], { fieldDescriptors: FIELDS, headerNullByte: true });
            expect(buffer.length).to.equal(32 + 3 * 32 + 2 + 1);
            expect(buffer[4] | buffer[5] | buffer[6] | buffer[7]).to.equal(0);
            expect(buffer[buffer.length - 1]).to.equal(0x1A);
        });
    });

    describe('workAreaId', () => {
        it('sets byte 0x14 of every field descriptor', () => {
            const buffer = createDBF(RECORDS, { fieldDescriptors: FIELDS, workAreaId: 1 });
            for (let i = 0; i < FIELDS.length; i++) {
                expect(buffer[32 + i * 32 + 20]).to.equal(0x01);
            }
        });

        for (const bad of [-1, 256, 1.5, '1']) {
            it(`throws for ${JSON.stringify(bad)}`, () => {
                expect(() => createDBF(RECORDS, { fieldDescriptors: FIELDS, workAreaId: bad })).to.throw(/workAreaId/);
            });
        }
    });

    describe('numberText: "raw"', () => {
        const options = { fieldDescriptors: FIELDS, numberText: 'raw' };

        it('writes String(value), left-padded, ignoring the decimal places', () => {
            const buffer = createDBF(RECORDS, options);
            expect(fieldText(buffer, FIELDS, 0, 1)).to.equal('  1.25');
            expect(fieldText(buffer, FIELDS, 0, 2)).to.equal(' 12');
            expect(fieldText(buffer, FIELDS, 1, 2)).to.equal('0.5');
            expect(fieldText(buffer, FIELDS, 2, 2)).to.equal(' -1');
        });

        it('cuts a wider number to the field size from the left, no * fill', () => {
            const buffer = createDBF(RECORDS, options);
            expect(fieldText(buffer, FIELDS, 2, 1)).to.equal('123456');
        });

        it('null, undefined, NaN and Infinity are blank', () => {
            const rows = [{ QTY: null }, { QTY: undefined }, { QTY: NaN }, { QTY: Infinity }, {}];
            const buffer = createDBF(rows, options);
            for (let r = 0; r < rows.length; r++) {
                expect(fieldText(buffer, FIELDS, r, 1)).to.equal('      ');
            }
        });

        it('a numeric string is written as its number', () => {
            const buffer = createDBF([{ QTY: '42' }], options);
            expect(fieldText(buffer, FIELDS, 0, 1)).to.equal('    42');
        });

        it('applies to F fields too', () => {
            const fields = [{ name: 'F1', type: 'F', size: 8, decimalPlaces: 2 }];
            const buffer = createDBF([{ F1: 3.14159 }], { fieldDescriptors: fields, numberText: 'raw' });
            expect(fieldText(buffer, fields, 0, 0)).to.equal(' 3.14159');
        });

        it('"fixed" is the default behaviour', () => {
            const fixed = createDBF(RECORDS, { fieldDescriptors: FIELDS, numberText: 'fixed', lastUpdate: OCT_31 });
            const plain = createDBF(RECORDS, { fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            expect(Buffer.from(fixed).equals(Buffer.from(plain))).to.be.true;
        });

        it('throws for an unknown value', () => {
            expect(() => createDBF(RECORDS, { fieldDescriptors: FIELDS, numberText: 'string' })).to.throw(/numberText/);
        });
    });

    describe('strictEncoding', () => {
        const fields = [{ name: 'EMPL_NAME', type: 'C', size: 10 }];

        it('latin1 characters up to 0xFF are written as one byte each', () => {
            const buffer = createDBF([{ EMPL_NAME: 'Café' }], { fieldDescriptors: fields, encoding: 'latin1', strictEncoding: true });
            const start = headerLength(buffer) + 1;
            expect(Array.from(buffer.subarray(start, start + 4))).to.deep.equal([0x43, 0x61, 0x66, 0xE9]);
        });

        it('a character over 0xFF in latin1 throws, naming the field and the record', () => {
            const rows = [{ EMPL_NAME: 'Ram' }, { EMPL_NAME: 'Ram' }, { EMPL_NAME: 'রাম' }];
            expect(() => createDBF(rows, { fieldDescriptors: fields, encoding: 'latin1', strictEncoding: true }))
                .to.throw(/EMPL_NAME.*record 2|record 2.*EMPL_NAME/);
        });

        it('a character over 0x7F in ascii throws', () => {
            expect(() => createDBF([{ EMPL_NAME: 'Café' }], { fieldDescriptors: fields, encoding: 'ascii', strictEncoding: true }))
                .to.throw(/EMPL_NAME/);
        });

        it('only the characters inside the field size count (they are the ones written)', () => {
            const rows = [{ EMPL_NAME: 'Ramchandra রাম' }];
            expect(() => createDBF(rows, { fieldDescriptors: fields, encoding: 'latin1', strictEncoding: true })).to.not.throw();
        });

        it('without strictEncoding latin1 does not throw (unchanged)', () => {
            expect(() => createDBF([{ EMPL_NAME: 'রাম' }], { fieldDescriptors: fields, encoding: 'latin1' })).to.not.throw();
        });

        it('has no effect with utf-8', () => {
            expect(() => createDBF([{ EMPL_NAME: 'রাম' }], { fieldDescriptors: fields, strictEncoding: true })).to.not.throw();
        });
    });

    describe('all legacy options together', () => {
        it('produce the expected file byte for byte', () => {
            const buffer = createDBF(RECORDS, { ...LEGACY, fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            const expected = expectedLegacyFile();
            expect(Buffer.from(buffer).toString('hex')).to.equal(expected.toString('hex'));
        });

        it('writeDBF writes the same bytes to a file', async () => {
            const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'su-dbf-legacy-'));
            try {
                const filePath = path.join(dir, 'legacy.dbf');
                await writeDBF(RECORDS, { ...LEGACY, fieldDescriptors: FIELDS, lastUpdate: OCT_31, filePath });
                expect(fs.readFileSync(filePath).equals(expectedLegacyFile())).to.be.true;
            } finally {
                fs.rmSync(dir, { recursive: true, force: true });
            }
        });

        it('the built CommonJS write entry gives the same bytes (dist rebuilt)', () => {
            const cjs = require(path.resolve(__dirname, '../../dist/cjs/write.cjs'));
            const buffer = cjs.createDBF(RECORDS, { ...LEGACY, fieldDescriptors: FIELDS, lastUpdate: OCT_31 });
            expect(Buffer.from(buffer).equals(expectedLegacyFile())).to.be.true;
        });

        it('write.d.ts declares the options', () => {
            const types = fs.readFileSync(path.resolve(__dirname, '../../write.d.ts'), 'utf8');
            for (const name of ['numberText', 'headerNullByte', 'workAreaId', 'strictEncoding']) {
                expect(types).to.include(`${name}?:`);
            }
        });
    });
});
