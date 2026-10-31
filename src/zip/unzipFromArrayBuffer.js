// `fflate` uses "synchronous" decompressor on `.zip` files that're under `512KB`
// (or when the compression level is too low), and an "asynchronous" (web worker)
// decompressor otherwise, as it can be seen from the source code of the `unzip()` function.
// This is because using web workers comes with an overhead of
// both creating a new worker and sending data back and forth.
//
import { unzip } from 'fflate'

import { createUnzipError } from './UnzipError.js'

/**
 * Reads `*.zip` file contents. Ignores anything besides `.xml` or `.xml.rels` files.
 * @param  {ArrayBuffer} input
 * @param  {object} [options] — `filter({ path })` and `maxUncompressedSize`.
 * @return {Promise<Record<string,Uint8Array>>} Resolves to an object holding `*.zip` file entries.
 */
export default function unzipFromArrayBuffer(input, options) {
	return unzipFromArrayBufferUsingFunction(input, options, unzipAsync, true)
}

/**
 * Same, with the `unzip` function to use.
 * @param  {ArrayBuffer} input
 * @param  {object} [options]
 * @param  {({ path: string }) => boolean} [options.filter] — Which entries to read.
 * @param  {number} [options.maxUncompressedSize] — Total uncompressed size limit (zip bomb guard).
 * @param  {(Uint8Array, object) => Record<string, Uint8Array> | Promise<Record<string, Uint8Array>>} unzip — Called with the archive and `fflate`'s options (`{ filter }`).
 * @param  {boolean} isAsync — Should be `true` when `unzip()` returns a `Promise`, `false` when it returns the files.
 * @return {Promise<Record<string,Uint8Array>> | Record<string,Uint8Array>} The entries of the `*.zip` file (a Promise of them when `isAsync`).
 */
export function unzipFromArrayBufferUsingFunction(input, options = {}, unzip, isAsync) {
	const filter = options.filter;
	const maxBytes = options.maxUncompressedSize || options.maxBytes || (256 * 1024 * 1024);
	let totalUncompressed = 0;

	const unzipOptions = {
		// Ignore certain types of files and protect against zip bombs
		filter: (file) => {
			if (filter) {
				if (!filter({ path: file.name })) {
					return false;
				}
			}
			if (file.originalSize) {
				totalUncompressed += file.originalSize;
				if (totalUncompressed > maxBytes) {
					throw new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`);
				}
			}
			return true;
		}
	}

	// If `fflate` throws its specific error then it implies that the `.zip` file is not valid.
	const convertError = (error) => {
		if (isFlateError(error)) {
			return createUnzipError(error)
		}
		return error
	}

	const archive = new Uint8Array(input)

	if (!isAsync) {
		try {
			return unzip(archive, unzipOptions)
		} catch (error) {
			throw convertError(error)
		}
	}

	// `unzip()` throws synchronously too (the `filter` runs before it returns), so it is called inside `Promise.resolve().then()`.
	return Promise.resolve().then(() => unzip(archive, unzipOptions)).then(
		result => result,
		(error) => {
			throw convertError(error)
		}
	)
}

function unzipAsync(archive, options) {
  return new Promise((resolve, reject) => {
		// `unzip()` will resort to "synchronous" decompression in two edge cases:
    // * When the archive size is less than `512KB`.
    // * When the data is barely compressed, i.e. the compression ratio is less than 20% reduction in size.
    unzip(archive, options, (error, files) => {
      if (error) {
        reject(error)
      } else {
        resolve(files)
      }
    })
  })
}

// This function attempts to guess if a given `error` was thrown by `fflate`.
function isFlateError(error) {
	// `fflate` doesn't export a `FlateError` class.
	// https://github.com/101arrowz/fflate/issues/290
	// return error instanceof FlateError

	// Here, it attempts to guess if an `error` is a `FlateError` by checking if `error.code` is a `number`.
	return typeof error.code === 'number'
}
