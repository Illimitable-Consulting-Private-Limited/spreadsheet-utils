import { zipSync } from 'fflate';

/**
 * Creates a `*.zip` file from a map of files synchronously.
 * @param {Record<string,Uint8Array>} files
 * @return {ArrayBuffer} `*.zip` file data.
 */
export default function zipToArrayBufferSync(files) {
  var uint8Array = zipSync(files);
  if (uint8Array.byteOffset === 0 && uint8Array.byteLength === uint8Array.buffer.byteLength) {
    return uint8Array.buffer;
  }
  return uint8Array.buffer.slice(uint8Array.byteOffset, uint8Array.byteOffset + uint8Array.byteLength);
}
