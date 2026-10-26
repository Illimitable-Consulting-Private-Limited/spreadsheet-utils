export type CellValue = string | number | boolean | Date | null;
export type Row = CellValue[];

export interface ColumnDefinition {
    title: string;
    width?: number;
    style?: any;
}

export interface Worksheet {
    worksheetName?: string;
    sheet?: string;
    data: Row[];
    columns?: ColumnDefinition[];
    mergeCells?: Record<string, [number, number]>;
    style?: Record<string, any> | string[];
    comments?: Record<string, any>;
    cells?: Record<string, any>;
    rows?: any[];
    [key: string]: any;
}

export interface TabularSpreadsheet {
    worksheets: Worksheet[];
    definedNames?: Record<string, string>;
    style?: string[];
    [key: string]: any;
}

export interface SheetEntry {
    name: string;
    data: Row[];
}

export interface SchemaProperty {
    prop: string;
    type?: any;
    required?: boolean;
    validate?: (value: any) => void;
    parse?: (value: any) => any;
}

export type Schema = Record<string, SchemaProperty>;

export interface SheetError {
    error: string;
    row?: number;
    column?: string;
    value?: any;
    type?: any;
    reason?: string;
}

export interface ParsedSchemaResult<T = any> {
    rows: T[];
    errors: SheetError[];
}

export interface CommonOptions {
    extension?: string;
    format?: 'rows' | 'tabular' | string;
    worksheets?: boolean;
    sheet?: string | number;
    allSheets?: boolean;
    delimiter?: string;
    tableIndex?: number;
    firstRowAsHeader?: boolean;
    dateFormat?: string;
    parseNumber?: (value: string) => any;
    trim?: boolean;
    onload?: (result: any) => void;
    onerror?: (error: Error) => void;
}

export interface OptionsWithSchema<T = any> extends CommonOptions {
    schema: Schema;
}

export type SpreadsheetInput =
    | string
    | Uint8Array
    | ArrayBuffer
    | Blob
    | any; // File, Buffer, Stream

export function readSpreadsheet(file: SpreadsheetInput, options: { format: 'tabular' } & CommonOptions): Promise<TabularSpreadsheet>;
export function readSpreadsheet(file: SpreadsheetInput, options: { worksheets: true } & CommonOptions): Promise<TabularSpreadsheet>;
export function readSpreadsheet(file: SpreadsheetInput, options: { allSheets: true } & CommonOptions): Promise<SheetEntry[]>;
export function readSpreadsheet<T = any>(file: SpreadsheetInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
export function readSpreadsheet(file: SpreadsheetInput, options?: CommonOptions): Promise<Row[]>;
export default readSpreadsheet;

export const readXlsxFile: typeof readSpreadsheet;

export function readSheet(file: SpreadsheetInput, sheet?: string | number, options?: CommonOptions): Promise<Row[]>;
export function readSheet<T = any>(file: SpreadsheetInput, sheet: string | number, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;
export function readSheet<T = any>(file: SpreadsheetInput, options: OptionsWithSchema<T>): Promise<ParsedSchemaResult<T>>;

export function tabularjs(file: SpreadsheetInput, options?: CommonOptions): Promise<TabularSpreadsheet>;
export function parseTabular(file: SpreadsheetInput, options?: CommonOptions): Promise<TabularSpreadsheet>;

export function detectFormat(buffer: Uint8Array): Promise<string | null>;

export function parseSheetData<T = any>(data: Row[], schema: Schema, options?: any): ParsedSchemaResult<T>;

export function Integer(value: any): number;
export function Email(value: any): string;
export function URL(value: any): string;

export class InvalidInputError extends Error {
    constructor(message: string, cause?: any);
}
export class InvalidSpreadsheetError extends Error {
    constructor(message: string);
}
export class SheetNotFoundError extends Error {
    constructor(sheet: string | number, availableSheets?: string[]);
}

export function parseCSVFile(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseCSVString(content: string, delimiter?: string): TabularSpreadsheet;
export function parseXLS(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXLSBuffer(buffer: Uint8Array): TabularSpreadsheet;
export function parseXLSX(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseODS(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseDBF(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseDBFBuffer(buffer: Uint8Array): TabularSpreadsheet;
export function parseDIF(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseDIFString(fileContent: string): TabularSpreadsheet;
export function parseLotus(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseLotusBuffer(buffer: Uint8Array): TabularSpreadsheet;
export function parseSLK(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseSLKString(fileContent: string): TabularSpreadsheet;
export function parseHTMLTable(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseAllHTMLTables(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXMLSpreadsheet(input: any, options?: any): Promise<TabularSpreadsheet>;
export function parseXMLSpreadsheetString(fileContent: string): TabularSpreadsheet;
export function parseNumbers(input: any, options?: any): Promise<TabularSpreadsheet>;
export function isNumbersFile(input: any): Promise<boolean>;
