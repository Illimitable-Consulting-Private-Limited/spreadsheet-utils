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
export { xlsPlugin, parseXLS, parseXLSBuffer, writeXLS, createXLS } from './xls.js';
export { markupPlugin, parseODS, parseHTMLTable, parseAllHTMLTables, parseXMLSpreadsheet, parseXMLSpreadsheetString, writeODS, createODS } from './markup.js';
export { legacyPlugin, parseDBF, parseLotus, parseDIF, parseSLK, parseNumbers } from './legacy.js';
export { dbfPlugin, writeDBF, createDBF, normalizeSchemaToDbfFields } from './dbf.js';
export { odsPlugin } from './ods.js';

import { textPlugin } from './text.js';
import { xlsPlugin } from './xls.js';
import { markupPlugin } from './markup.js';
import { legacyPlugin } from './legacy.js';
import { dbfPlugin } from './dbf.js';
import { odsPlugin } from './ods.js';

export const allPlugins = [
    textPlugin,
    xlsPlugin,
    markupPlugin,
    legacyPlugin,
    dbfPlugin,
    odsPlugin
];

export { registerPlugin, unregisterPlugin, isPluginRegistered, getWriterForExtension } from '../core/registry.js';
