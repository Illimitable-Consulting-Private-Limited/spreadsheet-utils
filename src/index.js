import { readSpreadsheet, parseTabular, createReader, iterateSpreadsheet } from './core/dispatcher.js';
import { detectFormat } from './core/detector.js';
import { registerPlugin, unregisterPlugin, isPluginRegistered } from './core/registry.js';

// Schema exports from read-excel-file
import parseSheetData from './schema/parseSheetData.js';
import Integer from './schema/types/additional/Integer.js';
import Email from './schema/types/additional/Email.js';
import URL from './schema/types/additional/URL.js';

// Errors
import InvalidInputError from './xlsx/file/InvalidInputError.js';
import InvalidSpreadsheetError from './xlsx/InvalidSpreadsheetError.js';
import SheetNotFoundError from './xlsx/SheetNotFoundError.js';

// Core XLSX parser
import { parseXLSX } from './parser/xlsx.js';

// read-excel-file backwards compatibility
export function readSheet(input, sheet, options) {
    if (!options && sheet && typeof sheet !== 'number' && typeof sheet !== 'string') {
        options = sheet;
        sheet = undefined;
    }
    return readSpreadsheet(input, { ...options, sheet });
}

const readXlsxFile = readSpreadsheet;
const tabularjs = parseTabular;
const parser = parseTabular;

export {
    // Core readers
    readSpreadsheet,
    readSpreadsheet as default,
    readXlsxFile,
    iterateSpreadsheet,
    parseTabular,
    tabularjs,
    parser,
    createReader,
    detectFormat,

    // Plugin registry
    registerPlugin,
    unregisterPlugin,
    isPluginRegistered,

    // Schema validation and mapping
    parseSheetData,
    Integer,
    Email,
    URL,

    // Errors
    InvalidInputError,
    InvalidSpreadsheetError,
    SheetNotFoundError,

    // Core parser
    parseXLSX
};
