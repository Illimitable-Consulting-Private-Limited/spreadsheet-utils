/**
 * Web Worker & Thread offloader for spreadsheet-utils.
 * Offloads CPU-intensive spreadsheet parsing to background worker threads.
 */

import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';
import { readSpreadsheet } from '../index.js';
import { allPlugins } from '../plugins/index.js';
import { registerPlugin } from '../core/registry.js';

// If running inside a Node.js worker thread:
if (!isMainThread && parentPort) {
    registerPlugin(...allPlugins);
    (async () => {
        try {
            const result = await readSpreadsheet(workerData.input, workerData.options);
            parentPort.postMessage({ ok: true, data: result });
        } catch (err) {
            parentPort.postMessage({ ok: false, error: err.message, stack: err.stack });
        }
    })();
}

function getWorkerLocation() {
    try {
        if (typeof __filename === 'string' && __filename) {
            return __filename;
        }
    } catch {}
    try {
        if (typeof import.meta !== 'undefined' && import.meta && import.meta.url) {
            return new URL(import.meta.url);
        }
    } catch {}
    return null;
}

/**
 * Read and parse a spreadsheet in a background worker thread.
 *
 * In Node.js: Uses node:worker_threads.
 * In Browser: Uses Web Worker API.
 *
 * @param {any} input - File path, Buffer, ArrayBuffer, Blob, or File
 * @param {Object} [options={}] - Options passed to readSpreadsheet
 * @returns {Promise<any>} Parsed spreadsheet rows or data
 */
export async function readInWorker(input, options = {}) {
    // 1. Node.js worker_threads
    if (typeof process !== 'undefined' && process.versions && process.versions.node) {
        return new Promise((resolve, reject) => {
            try {
                const workerLocation = getWorkerLocation();
                if (!workerLocation) {
                    throw new Error('Worker script location could not be determined');
                }

                let workerInput = input;
                if (input instanceof Uint8Array) {
                    workerInput = new Uint8Array(input);
                }

                const worker = new Worker(workerLocation, {
                    workerData: { input: workerInput, options }
                });

                worker.on('message', (msg) => {
                    worker.terminate();
                    if (msg.ok) {
                        resolve(msg.data);
                    } else {
                        reject(new Error(msg.error));
                    }
                });

                worker.on('error', (err) => {
                    worker.terminate();
                    // Fallback to main thread
                    registerPlugin(...allPlugins);
                    readSpreadsheet(input, options).then(resolve).catch(reject);
                });

                worker.on('exit', (code) => {
                    if (code !== 0) {
                        registerPlugin(...allPlugins);
                        readSpreadsheet(input, options).then(resolve).catch(reject);
                    }
                });
            } catch (err) {
                // Fallback to main thread
                registerPlugin(...allPlugins);
                readSpreadsheet(input, options).then(resolve).catch(reject);
            }
        });
    }

    // 2. Browser Web Worker
    if (typeof Worker !== 'undefined' && typeof Blob !== 'undefined' && typeof URL !== 'undefined') {
        return new Promise((resolve, reject) => {
            const workerCode = `
                self.onmessage = async function(e) {
                    try {
                        var bundle = e.data.bundleUrl || 'spreadsheet-utils-full.min.js';
                        if (typeof spreadsheetUtilsFull === 'undefined') {
                            importScripts(bundle);
                        }
                        var res = await self.spreadsheetUtilsFull.readSpreadsheet(e.data.input, e.data.options);
                        self.postMessage({ ok: true, data: res });
                    } catch(err) {
                        self.postMessage({ ok: false, error: err.message });
                    }
                };
            `;
            const blob = new Blob([workerCode], { type: 'application/javascript' });
            const workerUrl = URL.createObjectURL(blob);
            const worker = new Worker(workerUrl);

            worker.onmessage = (e) => {
                worker.terminate();
                URL.revokeObjectURL(workerUrl);
                if (e.data.ok) resolve(e.data.data);
                else reject(new Error(e.data.error));
            };

            worker.onerror = (err) => {
                worker.terminate();
                URL.revokeObjectURL(workerUrl);
                reject(err);
            };

            worker.postMessage({ input, options });
        });
    }

    // 3. Transparent Fallback
    registerPlugin(...allPlugins);
    return readSpreadsheet(input, options);
}

export default readInWorker;
