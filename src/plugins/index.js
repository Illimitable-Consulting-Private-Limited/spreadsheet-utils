/**
 * TabularJS format plugins for spreadsheet-utils.
 *
 * Provides grouped format plugins that extend the core XLSX reader:
 * - textPlugin: CSV, TSV, TAB, TXT
 * - xlsPlugin: XLS (BIFF5/BIFF8) + CFB + PTG decoder
 * - markupPlugin: ODS, HTML tables, XML SpreadsheetML 2003
 * - legacyPlugin: DBF, Lotus 1-2-3 (WKS/WK1/WK3/WK4/123), DIF, SLK, Numbers
 */

export { textPlugin, parseCSVFile, parseCSVString } from './text.js';
export { xlsPlugin, parseXLS, parseXLSBuffer } from './xls.js';
export { markupPlugin, parseODS, parseHTMLTable, parseAllHTMLTables, parseXMLSpreadsheet, parseXMLSpreadsheetString } from './markup.js';
export { legacyPlugin, parseDBF, parseLotus, parseDIF, parseSLK, parseNumbers } from './legacy.js';

import { textPlugin } from './text.js';
import { xlsPlugin } from './xls.js';
import { markupPlugin } from './markup.js';
import { legacyPlugin } from './legacy.js';

export const allPlugins = [
    textPlugin,
    xlsPlugin,
    markupPlugin,
    legacyPlugin
];

export { registerPlugin, unregisterPlugin, isPluginRegistered } from '../core/registry.js';
