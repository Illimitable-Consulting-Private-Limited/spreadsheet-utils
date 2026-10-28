import { SpreadsheetInput, CommonOptions, OptionsWithSchema, ParsedSchemaResult, Row } from './index.js';

/**
 * Offloads spreadsheet parsing to a background worker thread.
 *
 * In Node.js: Uses node:worker_threads.
 * In Browser: Uses Web Worker API.
 *
 * @param file - File path, Buffer, ArrayBuffer, Blob, or File
 * @param options - Parser options
 * @returns Parsed spreadsheet rows or schema-mapped objects
 */
export function readInWorker<T = any>(file: SpreadsheetInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
export function readInWorker(file: SpreadsheetInput, options?: CommonOptions): Promise<Row[]>;
export default readInWorker;
