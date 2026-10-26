import { readSpreadsheet, parseTabular } from './core/dispatcher.js';
import { detectFormat } from './core/detector.js';

// Schema exports from read-excel-file
import parseSheetData from './schema/parseSheetData.js';
import Integer from './schema/types/additional/Integer.js';
import Email from './schema/types/additional/Email.js';
import URL from './schema/types/additional/URL.js';

// Errors
import InvalidInputError from './xlsx/file/InvalidInputError.js';
import InvalidSpreadsheetError from './xlsx/InvalidSpreadsheetError.js';
import SheetNotFoundError from './xlsx/SheetNotFoundError.js';

// Specific format parsers
import { parseCSVFile, parseCSVString } from './parser/csv.js';
import { parseXLS, parseXLSBuffer } from './parser/xls.js';
import { parseXLSX } from './parser/xlsx.js';
import { parseODS } from './parser/ods.js';
import { parseDBF, parseDBFBuffer } from './parser/dbf.js';
import { parseDIF, parseDIFString } from './parser/dif.js';
import { parseLotus, parseLotusBuffer, parseWK1, parseWK3, parseWK4 } from './parser/lotus.js';
import { parseSLK, parseSLKString } from './parser/slk.js';
import { parseHTMLTable, parseAllHTMLTables } from './parser/table.js';
import { parseXMLSpreadsheet, parseXMLSpreadsheetString } from './parser/xml.js';
import { parseNumbers, isNumbersFile } from './parser/numbers.js';

// read-excel-file backwards compatibility
export function readSheet(input, sheet, options) {
    if (!options && sheet && typeof sheet !== 'number' && typeof sheet !== 'string') {
        options = sheet;
        sheet = undefined;
    }
    return readSpreadsheet(input, { ...options, sheet });
}

export const readXlsxFile = readSpreadsheet;
export const tabularjs = parseTabular;

export {
    readSpreadsheet,
    parseTabular,
    detectFormat,
    // read-excel-file schema
    parseSheetData,
    Integer,
    Email,
    URL,
    // errors
    InvalidInputError,
    InvalidSpreadsheetError,
    SheetNotFoundError,
    // individual parsers
    parseCSVFile,
    parseCSVString,
    parseXLS,
    parseXLSBuffer,
    parseXLSX,
    parseODS,
    parseDBF,
    parseDBFBuffer,
    parseDIF,
    parseDIFString,
    parseLotus,
    parseLotusBuffer,
    parseWK1,
    parseWK3,
    parseWK4,
    parseSLK,
    parseSLKString,
    parseHTMLTable,
    parseAllHTMLTables,
    parseXMLSpreadsheet,
    parseXMLSpreadsheetString,
    parseNumbers,
    isNumbersFile
};

export default readSpreadsheet;
