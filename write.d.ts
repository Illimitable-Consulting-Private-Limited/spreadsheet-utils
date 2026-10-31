import { Schema } from './index.js';

export interface CellStyle {
    fontWeight?: 'bold' | string;
    fontStyle?: 'italic' | string;
    fontSize?: number;
    fontFamily?: string;
    textColor?: string;
    backgroundColor?: string;
    borderColor?: string;
    borderStyle?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
    align?: 'left' | 'center' | 'right';
    verticalAlign?: 'top' | 'center' | 'bottom';
    wrap?: boolean;
    format?: string;
}

export interface CellData extends CellStyle {
    value: any;
    type?: any;
    format?: string;
    columnSpan?: number;
    rowSpan?: number;
}

export type WriteRow = (CellData | any)[];
export type WriteGrid = WriteRow[];

export interface HeaderCell extends CellStyle {
    value: string;
    [key: string]: any;
}

export interface WriteColumn extends CellStyle {
    header?: string | HeaderCell;
    column?: string;
    name?: string;
    prop?: string;
    headerStyle?: CellStyle;
    cell?: (row: any, rowIndex?: number) => CellData | any;
    value?: (row: any, rowIndex?: number) => any;
    style?: CellStyle | ((value: any, row: any, rowIndex?: number) => CellStyle);
    width?: number;
    type?: any;
    format?: string;
}

export interface WriteSheetOptions {
    sheet?: string;
    columns?: WriteColumn[];
    schema?: Schema | any[] | Record<string, any>;
    headerStyle?: CellStyle;
    theme?: 'ocean' | 'forest' | 'slate' | 'sunset' | 'striped' | Record<string, any>;
    zebra?: boolean;
    autoWidth?: boolean | { minWidth?: number; maxWidth?: number; padding?: number; force?: boolean };
    minWidth?: number;
    maxWidth?: number;
    padding?: number;
    hasHeader?: boolean;
    tableClass?: string;
    fullDocument?: boolean;
    filePath?: string;
    fileName?: string;
    buffer?: boolean;
    blob?: boolean;
    format?: 'xlsx' | 'csv' | 'tsv' | 'html' | 'json' | 'dbf' | 'ods' | 'xls' | string;
    delimiter?: string;
    dateFormat?: string;
    fileVersion?: number | string;
    encoding?: string;
    lastUpdate?: Date;
    languageDriver?: number;
    fieldDescriptors?: DbfFieldDescriptor[];
    isRowDeleted?: (row: any, index: number) => boolean;
    deletedRows?: number[];
    features?: Record<string, any>;
    /**
     * xlsx only. 1-based row numbers a manual page break follows (`[10, 20]` starts a page at rows 11 and 21), written
     * as `<rowBreaks>`; sorted and de-duplicated. A value that is not a positive integer throws.
     */
    rowBreaks?: number[];
    [key: string]: any;
}

export interface DbfFieldDescriptor {
    name: string;
    type: 'C' | 'N' | 'F' | 'I' | 'L' | 'D' | 'T' | 'B' | 'Y' | 'M' | string;
    size: number;
    decimalPlaces?: number;
    value?: (row: any, index?: number) => any;
    prop?: string;
}

export interface SheetSpec {
    sheet: string;
    data: WriteGrid;
    columns?: WriteColumn[];
    /** See `WriteSheetOptions.rowBreaks`. */
    rowBreaks?: number[];
    [key: string]: any;
}

export interface WriterInstance {
    toBuffer: () => Promise<Uint8Array | Buffer>;
    toBlob: () => Promise<Blob>;
    toFile: (pathOrFileName: string) => Promise<void>;
    toStream?: (writable?: any) => Promise<any>;
}

export function writeSpreadsheet(
    data: WriteGrid | Record<string, any>[] | SheetSpec[] | { worksheets: any[] },
    options?: WriteSheetOptions | string,
    extraOptions?: any
): Promise<Uint8Array | Buffer | Blob | string | void>;

export default writeSpreadsheet;

export function writeXlsxFile(
    data: WriteGrid | Record<string, any>[] | SheetSpec[],
    options?: WriteSheetOptions,
    extraOptions?: any
): WriterInstance;

export function writeCSV(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
export function writeTSV(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
export function writeHTML(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
export function createHTML(data: any, options?: WriteSheetOptions | string): string;
export function writeJSON(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;
export function writeDBF(data: any, options?: WriteSheetOptions | string): Promise<Uint8Array | Buffer | string | void>;
export function createDBF(data: any, options?: WriteSheetOptions | string): Uint8Array;
export function writeODS(data: any, options?: WriteSheetOptions | string): Promise<Uint8Array | Buffer | string | void>;
export function createODS(data: any, options?: WriteSheetOptions | string): Uint8Array;
export function writeXLS(data: any, options?: WriteSheetOptions | string): Promise<Uint8Array | Buffer | string | void>;
export function createXLS(data: any, options?: WriteSheetOptions | string): Uint8Array;

export function getSheetData(objects: Record<string, any>[], columns: WriteColumn[]): WriteGrid;
export function normalizeSchemaToColumns(schema: Schema | any[]): WriteColumn[];
export function normalizeSchemaToDbfFields(schema?: any, data?: any, options?: any): DbfFieldDescriptor[];

export function calculateAutoWidths(rows: any[][], options?: any, existingColumns?: WriteColumn[]): WriteColumn[];
export const THEMES: Record<string, any>;
export function getTheme(theme: string | object): any;
export function applyTheme(rows: any[][], theme: string | object, options?: any): any[][];

export function streamCSV(source: any, destination: any, options?: any): Promise<any>;
export function streamTSV(source: any, destination: any, options?: any): Promise<any>;
export function streamJSONL(source: any, destination: any, options?: any): Promise<any>;
export class CSVTransformStream {
    constructor(options?: any);
}
export class JSONLTransformStream {
    constructor(options?: any);
}
