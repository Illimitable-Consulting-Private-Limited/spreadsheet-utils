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

export interface WriteColumn {
    header?: string;
    cell: (row: any, rowIndex?: number) => CellData | any;
    width?: number;
}

export interface WriteSheetOptions {
    sheet?: string;
    columns?: WriteColumn[];
    schema?: Schema | any[];
    filePath?: string;
    fileName?: string;
    buffer?: boolean;
    blob?: boolean;
    format?: 'xlsx' | 'csv' | 'tsv' | 'html' | 'json' | string;
    delimiter?: string;
    dateFormat?: string;
    features?: Record<string, any>;
    [key: string]: any;
}

export interface SheetSpec {
    sheet: string;
    data: WriteGrid;
    columns?: WriteColumn[];
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
export function writeJSON(data: any, options?: WriteSheetOptions | string): Promise<string | Buffer>;

export function getSheetData(objects: Record<string, any>[], columns: WriteColumn[]): WriteGrid;
export function normalizeSchemaToColumns(schema: Schema | any[]): WriteColumn[];
