import { describe, it } from 'mocha'
import { expect } from 'chai'

import { zipSync, strToU8 } from 'fflate'

import unzipFromArrayBuffer from './unzipFromArrayBuffer.js'
import unzipFromArrayBufferSync from './unzipFromArrayBufferSync.js'

const archive = () => zipSync({
	'a.xml': strToU8('<a/>'),
	'b.bin': strToU8('binary'),
	'c.xml': strToU8('<c>' + 'x'.repeat(2000) + '</c>')
})

const onlyXml = (entry) => entry.path.endsWith('.xml')

describe('unzipFromArrayBuffer', () => {
	it('should read a `.zip` archive', async () => {
		const zip = zipSync({ 'a.xml': strToU8('<a/>') })
		const files = await unzipFromArrayBuffer(zip)
		expect(Object.keys(files)).to.deep.equal(['a.xml'])
	})

	it('should apply the `filter` option', async () => {
		const files = await unzipFromArrayBuffer(archive(), { filter: onlyXml })
		expect(Object.keys(files).sort()).to.deep.equal(['a.xml', 'c.xml'])
	})

	it('should enforce `maxUncompressedSize` (zip bomb guard)', async () => {
		let error
		try {
			await unzipFromArrayBuffer(archive(), { maxUncompressedSize: 1000 })
		} catch (e) {
			error = e
		}
		expect(error, 'should reject').to.be.an('error')
		expect(error.message).to.match(/limit/i)
	})
})

describe('unzipFromArrayBufferSync', () => {
	it('should read a `.zip` archive synchronously', () => {
		const files = unzipFromArrayBufferSync(archive())
		expect(files).to.not.be.instanceOf(Promise)
		expect(Object.keys(files).sort()).to.deep.equal(['a.xml', 'b.bin', 'c.xml'])
	})

	it('should apply the `filter` option', () => {
		const files = unzipFromArrayBufferSync(archive(), { filter: onlyXml })
		expect(Object.keys(files).sort()).to.deep.equal(['a.xml', 'c.xml'])
	})

	it('should enforce `maxUncompressedSize` (zip bomb guard)', () => {
		expect(() => unzipFromArrayBufferSync(archive(), { maxUncompressedSize: 1000 })).to.throw(/limit/i)
	})

	it('should throw on a file that is not a `.zip`', () => {
		expect(() => unzipFromArrayBufferSync(strToU8('not a zip at all, just text'))).to.throw()
	})
})
