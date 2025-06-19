/**
 * Zip/Unzip Module
 */

// For minify
const U = Uint8Array, V = DataView;

const crc32table = Uint32Array.from({ length: 256 }, (_, i) => {
	let c = i, j = 8
	while (j--) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : (c >>> 1);
	return c;
});

/**
 * Calculate CRC-32 from a byte array.
 * @param {Uint8Array<ArrayBuffer>} bytes Byte array
 * @returns {number} CRC-32
 */
export const calcCrc32 = bytes => {
	let c = 0xffffffff;
	for (const b of bytes) c = crc32table[(c ^ b) & 0xff] ^ (c >>> 8);
	return c ^ 0xffffffff;
};

/**
 * Formats CRC-32 (int32) to hex string.
 * @param {number} int32 CRC-32
 * @returns {string} Hex string
 */
export const fmtCrc32 = int32 => {
	const octets = [24, 16, 8, 0].map(shift => int32 >>> shift & 0xff);
	return octets.map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
};

/**
 * Class for unzipping file
 */
export class Extractor {
	/** @type {EndOfCentralDirectoryRecord} */ eocd;
	/** @type {CentralDirectoryEntry[]} */ cd;
	/** @type {Array<{ header: LocalFileHeader, body: Uint8Array<ArrayBuffer> }> } */ contents = [];

	/**
	 * Creates a Zip Extractor object from an array buffer.
	 * @param {ArrayBuffer} buffer array buffer
	 */
	constructor(buffer) {
		const bytes = new U(buffer);
		const eocdOffset = bytes.findLastIndex((_, i, a) => [0x50, 0x4b, 5, 6].every((v, j) => v === a[i + j]));
		if (eocdOffset < 0) throw new Error('Invalid format: End of central directory record is not found.');
		this.eocd = EndOfCentralDirectoryRecord.from(new V(buffer, eocdOffset));
		let cdOffset = this.eocd.cdOffset;
		this.cd = Array.from({ length: this.eocd.numOfFiles }, () => {
			const view = new V(buffer, cdOffset, eocdOffset - cdOffset);
			const entry = CentralDirectoryEntry.from(view);
			cdOffset += entry.length;
			return entry;
		});
		const offsets = this.cd.map(cd => cd.headerOffset);
		for (let i = 0; i < offsets.length; i++) {
			const cd = this.cd[i];
			const thisOffset = offsets[i];
			const nextOffset = offsets.at(i + 1) || this.eocd.cdOffset;
			const view = new V(buffer, thisOffset, nextOffset - thisOffset);
			const header = LocalFileHeader.from(view);
			if (header.hasDataDescriptor && header.isCentralDirectoryEncrypted) {
				const dataDescriptor = new Uint32Array(buffer, nextOffset - 12, 12);
				cd.crc32 = dataDescriptor[1];
				cd.compressedSize = dataDescriptor[2];
				cd.uncompressedSize = dataDescriptor[3];
			}
			const body = new U(view.buffer, cd.headerOffset + header.length, cd.compressedSize);
			this.contents.push({ header, body });
		}
	}

	/**
	 * Picks a content at index.
	 * @param {number} index 
	 * @returns {Response} decompressed response
	 */
	pick(index) {
		const content = this.contents.at(index);
		if (!content) {
			const reason = 'No content at the index: ' + index;
			throw new Error(reason);
		}
		const stream = new Blob([content.body]).stream();
		switch (content.header.method) {
			case 0:
				return new Response(stream);
			case 8:
				const decompression = new DecompressionStream('deflate-raw');
				return new Response(stream.pipeThrough(decompression));
			default:
				const reason = 'Unsupported compression method: ' + content.header.method;
				return new Response(null, { status: 418, statusText: reason });
		}
	}

	/**
	 * Gets a list of file names with specific encoding.
	 * @param {string} [encoding] encoding charset name
	 * @returns {string[]} list of file names
	 */
	getNames(encoding) {
		const decoder = new TextDecoder(encoding);
		return this.cd.map(cd => decoder.decode(cd.fileNameBytes));
	}

	*[Symbol.iterator]() {
		for (let i = 0; i < this.contents.length; i++) {
			yield {
				entry: this.cd[i],
				header: this.contents[i].header,
				body: this.pick(i),
			};
		}
	}
}

/**
 * Class for zipping file
 */
export class Builder {
	static textDecoder = new TextDecoder();
	static textEncoder = new TextEncoder();

	/** @type {CentralDirectoryEntry[]} */
	centralDirectory = [];
	/** @type { Array<{ header: LocalFileHeader, body: Uint8Array<ArrayBuffer> }> } */
	contents = [];

	/**
	 * @typedef {object} ZipBuilderAppendOptions
	 * @prop {string} filepath File path
	 * @prop {0|8} [method] Compression method
	 * @prop {number} [lastModified] Last modified Unix timestamp
	 * @prop {Uint8Array<ArrayBuffer>} [extraField] Byte array of extra field
	 * @prop {string} [comment] Comment of each item
	 */
	
	/**
	 * Appends a file (buffer) to this package and returns index asynchronously.
	 * @overload
	 * @param {ArrayBuffer} buf Uncompressed array buffer
	 * @param {string} opt File path
	 * @return {Promise<number>} Index
	 */
	/**
	 * Appends a file (buffer) to this package and returns index asynchronously.
	 * @overload
	 * @param {ArrayBuffer} buf Uncompressed array buffer
	 * @param {ZipBuilderAppendOptions} opt Compression options
	 * @return {Promise<number>} Index
	 */
	/**
	 * @param {ArrayBuffer} buf
	 * @param {string|ZipBuilderAppendOptions} opt
	 * @returns {Promise<number>}
	 */
	async append(buf, opt) {
		const uncompressed = new U(buf);
		const options = typeof opt === 'string' ? { filepath: opt } : opt;
		const cd = new CentralDirectoryEntry();
		const header = new LocalFileHeader();
		const method = options.method ?? 8;
		cd.versionNeeded = header.versionNeeded = method === 8 ? 20 : 10;
		cd.versionMadeBy = 20;
		cd.method = header.method = method;
		cd.lastModified = header.lastModified = ZipDateTime.fromDate(options.lastModified ? new Date(options.lastModified) : new Date());
		cd.isUtf8 = header.isUtf8 = true;
		const fileNameBytes = Builder.textEncoder.encode(options.filepath);
		cd.fileNameBytes = header.fileNameBytes = fileNameBytes;
		cd.extraFieldBytes = header.extraFieldBytes = options.extraField ?? new U();
		cd.commentBytes = Builder.textEncoder.encode(options.comment);
		cd.crc32 = header.crc32 = calcCrc32(uncompressed);
		cd.uncompressedSize = header.uncompressedSize = uncompressed.length;
		switch (method) {
			case 0:
				cd.compressedSize = header.compressedSize = uncompressed.length;
				this.centralDirectory.push(cd);
				return this.contents.push({ header, body: uncompressed });
			case 8:
				const stream = new Blob([uncompressed]).stream();
				const compression = new CompressionStream('deflate-raw');
				const response = new Response(stream.pipeThrough(compression));
				const compressed = await response.bytes();
				cd.compressedSize = header.compressedSize = compressed.length;
				this.centralDirectory.push(cd);
				return this.contents.push({ header, body: compressed });
			default:
				const reason = 'Unsupported compression method: ' + method;
				throw new Error(reason);
		}
	}

	/**
	 * Removes the items at the index from this archive.
	 * @param {number} index Index of the first item to remove
	 * @param {number} [count=1] Count of items to remove (Default: 1)
	 * @returns {string[]} List of removed files names
	 */
	remove(index, count = 1) {
		const cd = this.centralDirectory.splice(index, count);
		this.contents.splice(index, 1);
		return cd.map(entry => Builder.textDecoder.decode(entry.fileNameBytes));
	}

	/**
	 * Builds the package.
	 * @param {object} [options] options
	 * @param {string} [options.comment] package comment
	 * @returns {Blob} result blob
	 */
	build(options) {
		/** @type {BufferSource[]} */
		const blobParts = [];
		let offset = 0;
		const len = this.contents.length;
		for (let i = 0; i < len; i++) {
			this.centralDirectory[i].headerOffset = offset;
			const { header, body } = this.contents[i];
			const parts = header.toBlobParts();
			blobParts.push(...parts, body);
			offset += parts.reduce((a, c) => a + c.byteLength, 0) + body.length;
		}
		const cdBytes = this.centralDirectory.flatMap(cd => cd.toBlobParts());
		blobParts.push(...cdBytes);
		const eocd = new EndOfCentralDirectoryRecord();
		eocd.numOfFiles = len;
		eocd.totalNumOfFiles = len;
		eocd.cdSize = cdBytes.reduce((a, c) => a + c.byteLength, 0);
		eocd.cdOffset = offset;
		eocd.commentBytes = Builder.textEncoder.encode(options?.comment);
		blobParts.push(...eocd.toBlobParts());
		return new Blob(blobParts, { type: 'application/zip' });
	}
}

class ZipDateTime {
	/** @type {number} */ year;
	/** @type {number} */ month;
	/** @type {number} */ day;
	/** @type {number} */ hour;
	/** @type {number} */ minute;
	/** @type {number} */ second;

	toInt() {
		return (this.year - 1980 << 25)
			+ (this.month << 21)
			+ (this.day << 16)
			+ (this.hour << 11)
			+ (this.minute << 5)
			+ Math.floor(this.second / 2);
	}

	toDate() {
		return new Date(Date.UTC(this.year, this.month - 1, this.day, this.hour, this.minute, this.second));
	}

	/**
	 * Creates a new ZipDateTime object from an unsigned integer.
	 * @param {number} i unsigned integer
	 * @returns {ZipDateTime}
	 */
	static fromInt(i) {
		const dt = new ZipDateTime();
		// year_month_day_hour_minute_second
		// 0b1111111_1111_11111_11111_111111_11111
		dt.year = 1980 + (i >>> 25);
		dt.month = (i >>> 21 & 0b1111);
		dt.day = (i >>> 16) & 0b11111;
		dt.hour = (i >>> 11) & 0b11111;
		dt.minute = (i >>> 5) & 0b111111;
		dt.second = (i & 0b11111) * 2;
		return dt;
	}

	/**
	 * Creates a new ZipDateTime object from a Date object.
	 * @param {Date} d date object
	 * @returns {ZipDateTime}
	 */
	static fromDate(d = new Date()) {
		const dt = new ZipDateTime();
		dt.year = d.getFullYear();
		dt.month = d.getMonth() + 1;
		dt.day = d.getDate();
		dt.hour = d.getHours();
		dt.minute = d.getMinutes();
		dt.second = d.getSeconds();
		return dt;
	}
}

/**
 * @interface
 * @typedef PackageRecord
 * @prop {number} signature
 * @prop {() => BufferSource[]} toBlobParts
 * @static @prop {(view: DataView<ArrayBuffer>) => PackageRecord} from
 */

/**
 * Class for Local File Header
 * @implements PackageRecord
 */
class LocalFileHeader {
	/** @type {number} */ signature = 0x04034b50;
	/** @type {number} */ versionNeeded = 20;
	/** @type {number} */ flags = 1024;
	/** @type {number} */ method = 8;
	/** @type {ZipDateTime} */ lastModified = ZipDateTime.fromDate();
	/** @type {number} */ crc32 = 0;
	/** @type {number} */ compressedSize = 0;
	/** @type {number} */ uncompressedSize = 0;
	/** @type {Uint8Array<ArrayBuffer>} */ fileNameBytes;
	/** @type {Uint8Array<ArrayBuffer>} */ extraFieldBytes;
	/** @type {number} */ length;

	/**
	 * Creates a new "Local File Header" record from bytes.
	 * @param {DataView<ArrayBuffer>} view byte array cursor
	 * @returns {LocalFileHeader} new "Local File Header" record
	 */
	static from(view) {
		const record = new LocalFileHeader();
		record.signature = view.getUint32(0, true);
		record.versionNeeded = view.getUint16(4, true);
		record.flags = view.getUint16(6, true);
		record.method = view.getUint16(8, true);
		record.lastModified = ZipDateTime.fromInt(view.getUint32(10, true));
		record.crc32 = view.getUint32(14, true);
		record.compressedSize = view.getUint32(18, true);
		record.uncompressedSize = view.getUint32(22, true);

		const fileNameLength = view.getUint16(26, true);
		const extraFieldLength = view.getUint16(28, true);
		const offset = view.byteOffset + 30;
		record.fileNameBytes = new U(view.buffer, offset, fileNameLength);
		record.extraFieldBytes = new U(view.buffer, offset + fileNameLength, extraFieldLength);
		record.length = 30 + fileNameLength + extraFieldLength;
		return record;
	}

	toBlobParts() {
		const view = new V(new ArrayBuffer(30));
		view.setUint32(0, this.signature, true);
		view.setUint16(4, this.versionNeeded, true);
		view.setUint16(6, this.flags, true);
		view.setUint16(8, this.method, true);
		view.setUint32(10, this.lastModified.toInt(), true);
		view.setUint32(14, this.crc32, true);
		view.setUint32(18, this.compressedSize, true);
		view.setUint32(22, this.uncompressedSize, true);
		view.setUint16(26, this.fileNameBytes.length, true);
		view.setUint16(28, this.extraFieldBytes.length, true);
		return [view, this.fileNameBytes, this.extraFieldBytes];
	}
	
	/**
	 * 
	 * @param {number} i index
	 * @returns {boolean} flag value
	 */
	getFlag(i) {
		return (this.flags >> i & 1) === 1;
	}
	/**
	 * 
	 * @param {number} i index
	 * @param {boolean} b flag value
	 */
	setFlag(i, b) {
		const v = 1 << i;
		if (b) this.flags |= v;
		else this.flags &= ~v;
	}

	get isEncrypted() {
		return this.getFlag(0);
	}
	set isEncrypted(b) {
		this.setFlag(0, b);
	}
	get hasDataDescriptor() {
		return this.getFlag(3);
	}
	set hasDataDescriptor(b) {
		this.setFlag(3, b);
	}
	get isEnhancedDeflating() {
		return this.getFlag(4);
	}
	set isEnhancedDeflating(b) {
		this.setFlag(4, b);
	}
	get isCompressedPatchedData() {
		return this.getFlag(5);
	}
	set isCompressedPatchedData(b) {
		this.setFlag(5, b);
	}
	get isEncryptedStrongly() {
		return this.getFlag(6);
	}
	set isEncryptedStrongly(b) {
		this.setFlag(6, b);
	}
	get isUtf8() {
		return this.getFlag(11);
	}
	set isUtf8(b) {
		this.setFlag(11, b);
	}
	get isCentralDirectoryEncrypted() {
		return this.getFlag(13);
	}
	set isCentralDirectoryEncrypted(b) {
		this.setFlag(13, b);
	}
}

/**
 * Class for Central Directory entry
 */
class CentralDirectoryEntry extends LocalFileHeader {
	/** @type {number} */ signature = 0x02014b50;
	/** @type {number} */ versionMadeBy = 20;
	/** @type {number} */ diskIdStart = 0;
	/** @type {number} */ internalAttributes = 0;
	/** @type {number} */ externalAttributes = 0;
	/** @type {number} */ headerOffset;
	/** @type {Uint8Array<ArrayBuffer>} */ commentBytes;

	/**
	 * Creates a new "Central Directory" entry from bytes.
	 * @param {DataView<ArrayBuffer>} view byte array cursor
	 * @returns {CentralDirectoryEntry} new "Central Directory" entry
	 */
	static from(view) {
		const record = new CentralDirectoryEntry();
		record.signature = view.getUint32(0, true);
		record.versionMadeBy = view.getUint16(4, true);
		record.versionNeeded = view.getUint16(6, true);
		record.flags = view.getUint16(8, true);
		record.method = view.getUint16(10, true);
		record.lastModified = ZipDateTime.fromInt(view.getUint32(12, true));
		record.crc32 = view.getUint32(16, true);
		record.compressedSize = view.getUint32(20, true);
		record.uncompressedSize = view.getUint32(24, true);
		const fLen = view.getUint16(28, true);
		const eLen = view.getUint16(30, true);
		const cLen = view.getUint16(32, true);
		record.diskIdStart = view.getUint16(34, true);
		record.internalAttributes = view.getUint16(36, true);
		record.externalAttributes = view.getUint32(38, true);
		record.headerOffset = view.getUint32(42, true);
		const offset = view.byteOffset + 46;
		record.fileNameBytes = new U(view.buffer, offset, fLen);
		record.extraFieldBytes = new U(view.buffer, offset + fLen, eLen);
		record.commentBytes = new U(view.buffer, offset + fLen + eLen, cLen);
		record.length = 46 + fLen + eLen + cLen;
		return record;
	}

	toBlobParts() {
		const view = new V(new ArrayBuffer(46));
		view.setUint32(0, this.signature, true);
		view.setUint16(4, this.versionMadeBy, true);
		view.setUint16(6, this.versionNeeded, true);
		view.setUint16(8, this.flags, true);
		view.setUint16(10, this.method, true);
		view.setUint32(12, this.lastModified.toInt(), true);
		view.setUint32(16, this.crc32, true);
		view.setUint32(20, this.compressedSize, true);
		view.setUint32(24, this.uncompressedSize, true);
		view.setUint16(28, this.fileNameBytes.length, true);
		view.setUint16(30, this.extraFieldBytes.length, true);
		view.setUint16(32, this.commentBytes.length, true);
		view.setUint16(34, this.diskIdStart, true);
		view.setUint16(36, this.internalAttributes, true);
		view.setUint32(38, this.externalAttributes, true);
		view.setUint32(42, this.headerOffset, true);
		return [view, this.fileNameBytes, this.extraFieldBytes, this.commentBytes];
	}
}

/**
 * Class for End of Central Directory (EOCD) record
 * @implements PackageRecord
 */
class EndOfCentralDirectoryRecord {
	/** @type {number} */ signature = 0x06054b50;
	/** @type {number} */ diskId = 0;
	/** @type {number} */ firstDiskId = 0;
	/** @type {number} */ numOfFiles = 1;
	/** @type {number} */ totalNumOfFiles = 0;
	/** @type {number} */ cdSize;
	/** @type {number} */ cdOffset;
	/** @type {Uint8Array<ArrayBuffer>} */ commentBytes;
	/** @type {number} */ length;

	/**
	 * Creates a new <abbr>EOCD</abbr> (End of Central Directory) record from bytes.
	 * @param {DataView<ArrayBuffer>} view byte array cursor
	 * @returns {EndOfCentralDirectoryRecord} new EOCD record
	 */
	static from(view) {
		const record = new EndOfCentralDirectoryRecord();
		record.signature = view.getUint32(0, true);
		record.diskId = view.getUint16(4, true);
		record.firstDiskId = view.getUint16(6, true);
		record.numOfFiles = view.getUint16(8, true);
		record.totalNumOfFiles = view.getUint16(10, true);
		record.cdSize = view.getUint32(12, true);
		record.cdOffset = view.getUint32(16, true);
		const cLen = view.getUint16(20, true);
		const offset = view.byteOffset + 22;
		record.commentBytes = new U(view.buffer, offset, cLen);
		record.length = 22 + cLen;
		return record;
	}

	toBlobParts() {
		const view = new V(new ArrayBuffer(22));
		view.setUint32(0, this.signature, true);
		view.setUint16(4, this.diskId, true);
		view.setUint16(6, this.firstDiskId, true);
		view.setUint16(8, this.numOfFiles, true);
		view.setUint16(10, this.totalNumOfFiles, true);
		view.setUint32(12, this.cdSize, true);
		view.setUint32(16, this.cdOffset, true);
		view.setUint16(20, this.commentBytes.length, true);
		return [view, this.commentBytes];
	}
}