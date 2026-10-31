import {
    CommonOptions,
    OptionsWithSchema,
    ParsedSchemaResult,
    Row,
    SheetEntry,
    TabularSpreadsheet,
    InvalidInputError,
    InvalidSpreadsheetError,
    SheetNotFoundError,
    convertRowsToObjects
} from './index.js';
import { SheetSpec, WriteGrid, WriteSheetOptions } from './write.js';

/**
 * `spreadsheet-utils/browser-sync`: reads xlsx, xls and ods and writes xlsx without ever starting a Web Worker
 * (fflate's `zipSync()` / `unzipSync()` only), for pages whose Content Security Policy blocks `blob:` Workers.
 */

/** A file path or a stream is not accepted here: the browser has neither. */
export type BrowserSyncInput = File | Blob | Uint8Array | ArrayBuffer;

export function readSpreadsheet(file: BrowserSyncInput, options: { format: 'tabular' } & CommonOptions): Promise<TabularSpreadsheet>;
export function readSpreadsheet(file: BrowserSyncInput, options: { worksheets: true } & CommonOptions): Promise<TabularSpreadsheet>;
export function readSpreadsheet(file: BrowserSyncInput, options: { allSheets: true } & CommonOptions): Promise<SheetEntry[]>;
export function readSpreadsheet(file: BrowserSyncInput, options: ({ asObjects: true } | { header: true }) & CommonOptions): Promise<Array<Record<string, any>>>;
export function readSpreadsheet<T = any>(file: BrowserSyncInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
export function readSpreadsheet(file: BrowserSyncInput, options?: CommonOptions): Promise<Row[]>;
export default readSpreadsheet;

export const readXlsxFile: typeof readSpreadsheet;

export function readSheet(file: BrowserSyncInput, sheet?: string | number, options?: CommonOptions): Promise<Row[]>;
export function readSheet<T = any>(file: BrowserSyncInput, sheet: string | number, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
export function readSheet<T = any>(file: BrowserSyncInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;

export function parseTabular(file: BrowserSyncInput, options?: CommonOptions): Promise<TabularSpreadsheet>;

export interface SyncWriterInstance {
    /** Synchronous: zips on the calling thread. */
    toUint8Array: () => Uint8Array;
    toBlob: () => Promise<Blob>;
    /** Downloads the file in the browser. */
    toFile: (fileName: string) => Promise<void>;
}

/** The arguments of `writeXlsxFile()` of `spreadsheet-utils/write`, including `rowBreaks`. */
export function writeXlsxFile(
    data: WriteGrid | SheetSpec[],
    options?: WriteSheetOptions,
    extraOptions?: any
): SyncWriterInstance;

export { convertRowsToObjects, InvalidInputError, InvalidSpreadsheetError, SheetNotFoundError };
