/**
 * Converts file content to a `Uint8Array` universally.
 * Accepts Uint8Array, ArrayBuffer, Blob, or Buffer.
 *
 * @param {Uint8Array|ArrayBuffer|Blob|Buffer} fileContent
 * @returns {Promise<Uint8Array>}
 */
export default function convertFileContentToUint8Array(fileContent) {
  if (fileContent instanceof Uint8Array) {
    return Promise.resolve(fileContent);
  }
  if (typeof ArrayBuffer !== 'undefined' && fileContent instanceof ArrayBuffer) {
    return Promise.resolve(new Uint8Array(fileContent));
  }
  if (typeof Blob !== 'undefined' && fileContent instanceof Blob) {
    return blobToUint8Array(fileContent);
  }
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer(fileContent)) {
    return Promise.resolve(new Uint8Array(fileContent));
  }
  throw new Error('Unsupported file content type. Expected a `Blob`, `Buffer`, `Uint8Array`, or `ArrayBuffer`');
}

function blobToUint8Array(blob) {
  return blob.arrayBuffer().then(function (arrayBuffer) {
    return new Uint8Array(arrayBuffer);
  });
}
