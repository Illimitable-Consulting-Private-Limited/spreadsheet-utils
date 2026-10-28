/**
 * spreadsheet-utils/full
 *
 * All-in-one preset that includes:
 * - Core streaming XLSX reader & schema parser
 * - All format plugins (CSV, TSV, XLS, ODS, HTML tables, XML SpreadsheetML, DBF, Lotus 1-2-3, DIF, SLK, Numbers)
 * - Styled XLSX writer with formatting, cell styles, and schema mapping
 */

import { registerPlugin } from './core/registry.js';
import { allPlugins, textPlugin, xlsPlugin, markupPlugin, legacyPlugin } from './plugins/index.js';

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
    allPlugins
};

// Writer exports
export {
    writeSpreadsheet,
    writeXlsxFile,
    getSheetData,
    writeCSV,
    writeTSV,
    writeHTML,
    writeJSON
} from './write/index.js';

// Default export is the universal reader
import { readSpreadsheet } from './index.js';
export default readSpreadsheet;
