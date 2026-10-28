import fs from 'node:fs';
import { Readable, Transform, pipeline } from 'node:stream';

/**
 * Format a single line of CSV cells with proper RFC 4180 escaping.
 *
 * @param {Array<any>} cells
 * @param {string} delimiter
 * @returns {string}
 */
export function formatCSVLine(cells, delimiter = ',') {
    if (!Array.isArray(cells)) return '';
    return cells.map(cell => {
        const val = (cell && typeof cell === 'object' && !(cell instanceof Date) && 'value' in cell)
            ? cell.value
            : cell;

        if (val === null || val === undefined) return '';
        if (val instanceof Date) return val.toISOString();

        const str = String(val);
        if (str.includes(delimiter) || str.includes('"') || str.includes('\n') || str.includes('\r')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    }).join(delimiter);
}

/**
 * Node.js Transform stream converting row arrays or objects to CSV chunks.
 */
export class CSVTransformStream extends Transform {
    constructor(options = {}) {
        super({ objectMode: true });
        this.delimiter = options.delimiter || ',';
        this.hasHeader = options.hasHeader ?? true;
        this.headers = options.headers || null;
        this.headerEmitted = false;
    }

    _transform(chunk, encoding, callback) {
        try {
            if (Array.isArray(chunk)) {
                this.push(formatCSVLine(chunk, this.delimiter) + '\n');
            } else if (chunk && typeof chunk === 'object') {
                if (!this.headerEmitted && this.hasHeader) {
                    const keys = this.headers || Object.keys(chunk);
                    this.push(formatCSVLine(keys, this.delimiter) + '\n');
                    this.headerEmitted = true;
                }
                const keys = this.headers || Object.keys(chunk);
                const values = keys.map(k => chunk[k]);
                this.push(formatCSVLine(values, this.delimiter) + '\n');
            }
            callback();
        } catch (err) {
            callback(err);
        }
    }
}

/**
 * Node.js Transform stream converting objects or arrays to JSON Lines (JSONL).
 */
export class JSONLTransformStream extends Transform {
    constructor(options = {}) {
        super({ objectMode: true });
    }

    _transform(chunk, encoding, callback) {
        try {
            this.push(JSON.stringify(chunk) + '\n');
            callback();
        } catch (err) {
            callback(err);
        }
    }
}

function runPipeline(readable, transform, writable) {
    return new Promise((resolve, reject) => {
        pipeline(readable, transform, writable, (err) => {
            if (err) reject(err);
            else resolve();
        });
    });
}

/**
 * Stream massive datasets directly to a CSV file or writable stream with constant O(1) memory.
 *
 * @param {AsyncIterable<any>|Iterable<any>|Readable} source - Async generator, stream, or iterable
 * @param {string|Writable} destination - Destination file path or writable stream
 * @param {Object} [options={}] - Streaming options
 * @param {string} [options.delimiter=','] - Cell delimiter
 * @param {boolean} [options.hasHeader=true] - Auto-emit header row when objects are received
 * @param {Array<string>} [options.headers] - Explicit header keys
 * @returns {Promise<string|Writable>}
 */
export async function streamCSV(source, destination, options = {}) {
    const readable = (source instanceof Readable)
        ? source
        : Readable.from(source);

    const transform = new CSVTransformStream(options);

    const isPath = typeof destination === 'string';
    const writable = isPath
        ? fs.createWriteStream(destination, { encoding: 'utf-8' })
        : destination;

    await runPipeline(readable, transform, writable);
    return isPath ? destination : writable;
}

/**
 * Stream massive datasets directly to a TSV file or writable stream with constant O(1) memory.
 *
 * @param {AsyncIterable<any>|Iterable<any>|Readable} source
 * @param {string|Writable} destination
 * @param {Object} [options={}]
 * @returns {Promise<string|Writable>}
 */
export async function streamTSV(source, destination, options = {}) {
    return streamCSV(source, destination, { ...options, delimiter: '\t' });
}

/**
 * Stream massive datasets directly to a JSON Lines (JSONL) file or writable stream with constant O(1) memory.
 *
 * @param {AsyncIterable<any>|Iterable<any>|Readable} source
 * @param {string|Writable} destination
 * @param {Object} [options={}]
 * @returns {Promise<string|Writable>}
 */
export async function streamJSONL(source, destination, options = {}) {
    const readable = (source instanceof Readable)
        ? source
        : Readable.from(source);

    const transform = new JSONLTransformStream(options);

    const isPath = typeof destination === 'string';
    const writable = isPath
        ? fs.createWriteStream(destination, { encoding: 'utf-8' })
        : destination;

    await runPipeline(readable, transform, writable);
    return isPath ? destination : writable;
}
