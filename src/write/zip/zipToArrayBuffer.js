import { zip } from 'fflate';

/**
 * Creates a `*.zip` file from a map of files asynchronously.
 * @param {Record<string,Uint8Array>} files
 * @return {Promise<ArrayBuffer>} Promise of `*.zip` file data.
 */
export default function zipToArrayBuffer(files) {
  return zipAsync(files).then(function (uint8Array) {
    if (uint8Array.byteOffset === 0 && uint8Array.byteLength === uint8Array.buffer.byteLength) {
      return uint8Array.buffer;
    }
    return uint8Array.buffer.slice(uint8Array.byteOffset, uint8Array.byteOffset + uint8Array.byteLength);
  });
}

function zipAsync(files) {
  return new Promise(function (resolve, reject) {
    zip(files, function (error, archive) {
      if (error) {
        reject(error);
      } else {
        resolve(archive);
      }
    });
  });
}
