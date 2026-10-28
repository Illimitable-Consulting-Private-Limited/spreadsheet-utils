// Considerations:
//
// * `unzipper` unpacks a `.zip` file using Node.js's "native" `zlib` module.
//   This means that it's more performant than any equivalent pure-javascript decompressor.
//
// * `unzipper` is a true "streaming" unpacker: it works in a "pull" fashion rather than a "push" one.
//   New chunks of archived data are only read and decompressed as the receiver is ready to process them.
//
// * `unzipper` can completely skip a given compressed file entry if it's of no interest.
//   This means that a `.zip` archive could be read in multiple passes without performance penalty.

import { Parse, ZipFileError } from 'unzipper-esm'

import { createUnzipError } from './UnzipError.js'

const PROMISE_RESOLVE_VALUE = undefined

/**
 * Reads `*.zip` file contents.
 * @param  {Stream} stream
 * @param  {function} onFile
 * @param  {function} onFileData
 * @param  {function} onFileDataEnd
 * @return {Promise<void>}
 */
export default function unzipFromStream_(stream, onFile, onFileData, onFileDataEnd) {
	return new Promise((resolve, reject) => {
		const promises = []

		let errored = false

		const unzipperStream = Parse()

		const onError = (error) => {
			if (error instanceof ZipFileError) {
				error = createUnzipError(error)
			}
			if (!errored) {
				errored = true
				try {
					if (typeof unzipperStream.destroy === 'function') {
						unzipperStream.destroy(error)
					}
				} catch {}
				try {
					if (typeof stream.destroy === 'function') {
						stream.destroy(error)
					}
				} catch {}
				reject(error)
			}
		}

		stream
			.on('error', onError)
			.pipe(unzipperStream)
			.on('error', (err) => {
				if (!errored) onError(err)
			})
			.on('finish', () => {
				if (!errored) {
					Promise.all(promises).then(() => {
						resolve(PROMISE_RESOLVE_VALUE)
					}, onError)
				}
			})
			.on('entry', (entry) => {
				if (errored) {
					if (typeof entry.autodrain === 'function') {
						entry.autodrain().on('error', () => {})
					}
					return
				}

				let ignore = false
				if (entry.type === 'Directory') {
					ignore = true
				}

				const uncompressedFileSizeIsKnown = !(entry.vars.flags & 0x08) || entry.vars.compressedSize > 0
				const uncompressedFileSize = uncompressedFileSizeIsKnown ? entry.vars.uncompressedSize : undefined
				try {
					if (onFile(entry.path, uncompressedFileSize) === false) {
						ignore = true
					}
				} catch (err) {
					onError(err)
					return
				}

				if (ignore) {
					entry.autodrain().on('error', (err) => {
						if (!errored) onError(err)
					})
					return
				}

				promises.push(new Promise((resolve) => {
					entry
						.on('data', (data) => {
							if (errored) return
							try {
								onFileData(entry.path, data)
							} catch (err) {
								try {
									if (typeof entry.destroy === 'function') {
										entry.destroy(err)
									} else if (typeof entry.autodrain === 'function') {
										entry.autodrain().on('error', () => {})
									}
								} catch {}
								onError(err)
								resolve()
							}
						})
						.on('error', (error) => {
							if (!errored) onError(error)
							resolve()
						})
						.on('finish', () => {
							if (!errored) {
								try {
									onFileDataEnd(entry.path)
								} catch (err) {
									onError(err)
								}
							}
							resolve()
						})
				}))
			})
	})
}
