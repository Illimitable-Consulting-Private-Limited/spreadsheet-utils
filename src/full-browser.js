/**
 * spreadsheet-utils/full for Browser environments.
 *
 * Includes:
 * - Core streaming XLSX reader & schema parser
 * - All format plugins (CSV, TSV, XLS, ODS, HTML tables, XML SpreadsheetML, DBF, Lotus, DIF, SLK, Numbers)
 * - Browser styled XLSX writer with file download support
 */

import { registerPlugin } from './core/registry.js';
import { allPlugins, textPlugin, xlsPlugin, markupPlugin, legacyPlugin, dbfPlugin, odsPlugin } from './plugins/index.js';

// Auto-register all format plugins
registerPlugin(...allPlugins);

// Core re-exports
export * from './index.js';

// Plugin exports
export {
    textPlugin,
    xlsPlugin,
    markupPlugin,
    legacyPlugin,
    dbfPlugin,
    odsPlugin,
    allPlugins
};

// Writer exports (browser version)
export {
    writeSpreadsheet,
    writeXlsxFile,
    getSheetData,
    normalizeSchemaToColumns,
    normalizeSchemaToDbfFields,
    writeCSV,
    writeTSV,
    writeHTML,
    createHTML,
    writeJSON,
    writeDBF,
    createDBF,
    writeODS,
    createODS,
    writeXLS,
    createXLS
} from './write/browser.js';

import { readSpreadsheet } from './index.js';
export default readSpreadsheet;
