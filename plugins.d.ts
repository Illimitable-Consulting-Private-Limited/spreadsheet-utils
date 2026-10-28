import { SpreadsheetPlugin, TabularSpreadsheet } from './index.js';

/**
 * Text and Delimited formats plugin.
 * Adds support for CSV, TSV, TAB, and TXT files.
 */
export const textPlugin: SpreadsheetPlugin;

/**
 * Legacy Excel format plugin.
 * Adds support for Microsoft Excel 97-2004 (.xls) BIFF5/BIFF8 binary spreadsheets.
 */
export const xlsPlugin: SpreadsheetPlugin;

/**
 * Structured Documents & Markup format plugin.
 * Adds support for OpenDocument (.ods), HTML tables (.html, .htm), and XML Spreadsheet 2003 (.xml).
 */
export const markupPlugin: SpreadsheetPlugin;

/**
 * Databases & Legacy Spreadsheets plugin.
 * Adds support for DBF (dBase), Lotus 1-2-3 (WKS/WK1/WK3/WK4/123), DIF, SLK, and Apple Numbers.
 */
export const legacyPlugin: SpreadsheetPlugin;

/**
 * All format plugins array.
 */
export const allPlugins: SpreadsheetPlugin[];

export function registerPlugin(...plugins: SpreadsheetPlugin[]): void;
export function unregisterPlugin(target?: string | SpreadsheetPlugin): void;
export function isPluginRegistered(name: string): boolean;

export function parseCSVFile(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseCSVString(content: string, delimiter?: string): TabularSpreadsheet;
export function parseXLS(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXLSBuffer(buffer: Uint8Array): TabularSpreadsheet;
export function parseODS(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseHTMLTable(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseAllHTMLTables(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXMLSpreadsheet(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXMLSpreadsheetString(fileContent: string): TabularSpreadsheet;
export function parseDBF(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseLotus(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseDIF(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseSLK(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseNumbers(input: any, options?: any): Promise<TabularSpreadsheet>;
