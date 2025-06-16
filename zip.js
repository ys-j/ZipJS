/**
 * Zip/Unzip Module
 */

export const acceptMime = [
	'application/zip', // zip
	'application/x-zip-compressed', // zip
	// Office Open XML
	'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // docx
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
	'application/vnd.openxmlformats-officedocument.presentationml.presentation', // pptx
	// OpenDocument Format
	'application/vnd.oasis.opendocument.text', // odt
	'application/vnd.oasis.opendocument.spreadsheet', // ods
	'application/vnd.oasis.opendocument.presentation', // odp
	'application/vnd.oasis.opendocument.graphics', // odg
	// Others
	'application/epub+zip', // epub
];

export const compressionMethods = {
	0: 'Store',
	1: 'Shrunk',
	2: 'Reduce 1',
	3: 'Reduce 2',
	4: 'Reduce 3',
	5: 'Reduce 4',
	6: 'Implode',
	7: 'Tokenize',
	8: 'Deflate',
	9: 'Deflate64',
	10: 'TERSE (old)',
	12: 'BZIP2',
	14: 'LZMA',
	16: 'z/OS CMPSC',
	18: 'TERSE (new)',
	19: 'LZ77 z',
	93: 'Zstandard',
	94: 'MP3',
	95: 'XZ',
	96: 'JPEG',
	97: 'WavPack',
	98: 'PPMd ver.I rev.1',
	99: 'AE-x encryption',
};

/**
 * Class for CRC-32
 */
class Crc32 {
	#table = new Uint32Array(256);
	constructor() {
		for (let i = 0; i < 256; i++) {
			let c = i;
			for (let j = 0; j < 8; j++) {
				const mc = c >>> 1;
				c = c & 1 ? 0xedb88320 ^ mc : mc;
			}
			this.#table[i] = c;
		}
	}
	/**
	 * Calculate CRC-32 from a byte array.
	 * @param {Uint8Array<ArrayBuffer>} bytes Byte array
	 * @returns {number} CRC-32
	 */
	calc(bytes) {
		let c = 0xffffffff;
		for (const b of bytes) {
			c = this.#table[(c ^ b) & 0xff] ^ (c >>> 8);
		}
		return c ^ 0xffffffff;
	}
	/**
	 * Formats CRC-32 (int32) to hex string.
	 * @param {number} int32 CRC-32
	 * @returns {string} Hex string
	 */
	static fmt(int32) {
		const octets = [24, 16, 8, 0].map(shift => int32 >>> shift & 0xff);
		return octets.map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
	}
}

export const fmtCrc32 = Crc32.fmt;

/**
 * Class for unzipping file
 */
export class Extractor {
	/** @type {EndOfCentralDirectoryRecord} */ eocd;
	/** @type {Array<CentralDirectoryEntry>} */ centralDirectory;
	/** @type {Array<{ header: LocalFileHeader, body: Uint8Array<ArrayBuffer> }> } */ contents;

	/**
	 * Creates a Zip Extractor object from an array buffer.
	 * @param {ArrayBuffer} buffer array buffer
	 */
	constructor(buffer) {
		const bytes = new Uint8Array(buffer);
		const cursor = new ByteArrayCursor(bytes);
		const offsetEocd = bytes.findLastIndex((_, i, a) => [0x50, 0x4b, 5, 6].every((v, j) => v === a[i + j]));
		cursor.moveTo(offsetEocd);
		this.eocd = EndOfCentralDirectoryRecord.from(cursor);
		cursor.moveTo(this.eocd.cdOffset);
		this.centralDirectory = Array.from({ length: this.eocd.numOfFiles }, () => CentralDirectoryEntry.from(cursor));
		this.contents = this.centralDirectory.map(record => {
			cursor.moveTo(record.headerOffset);
			const header = LocalFileHeader.from(cursor);
			const body = cursor.subarray(record.compressedSize);
			if (header.hasDataDescriptor) {
				const maySignature = cursor.readAsInt();
				header.crc32 = maySignature === 0x08074b50 ? cursor.readAsInt() : maySignature;
				header.compressedSize = cursor.readAsInt();
				header.uncompressedSize = cursor.readAsInt();
			}
			return { header, body };
		});
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
		const stream = new Blob([content.body.slice()]).stream();
		switch (content.header.method) {
			case 0:
				return new Response(stream);
			case 8:
				const decompression = new DecompressionStream('deflate-raw');
				return new Response(stream.pipeThrough(decompression));
			default:
				const reason = 'Unsupported compression method: ' + content.header.method;
				throw new Error(reason);
		}
	}

	/**
	 * Gets a list of file names with specific encoding.
	 * @param {string} [encoding] encoding charset name
	 * @returns {string[]} list of file names
	 */
	getPaths(encoding) {
		const decoder = new TextDecoder(encoding);
		return this.contents.map(content => decoder.decode(content.header.fileNameBytes))
	}
}

/**
 * Class for zipping file
 */
export class Builder {
	static textDecoder = new TextDecoder();
	static textEncoder = new TextEncoder();
	static crc32 = new Crc32();

	/** @type {CentralDirectoryEntry[]} */
	centralDirectory = [];
	/** @type { Array<{ header: LocalFileHeader, body: Uint8Array<ArrayBuffer> }> } */
	contents = [];

	/**
	 * @typedef {object} ZipBuilderAppendOptions
	 * @prop {string} filepath File path
	 * @prop {0|8} [method] Compression method
	 * @prop {number} [lastModified] Last modified Unix timestamp
	 * @prop {Uint8Array} [extraField] Byte array of extra field
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
		const uncompressed = new Uint8Array(buf);
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
		cd.extraFieldBytes = header.extraFieldBytes = options.extraField ?? new Uint8Array();
		cd.commentBytes = Builder.textEncoder.encode(options.comment);
		cd.crc32 = header.crc32 = Builder.crc32.calc(uncompressed);
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
		/** @type {Uint8Array<ArrayBuffer>[]} */
		const blobPart = [];
		let offset = 0;
		const len = this.contents.length;
		for (let i = 0; i < len; i++) {
			this.centralDirectory[i].headerOffset = offset;
			const { header, body } = this.contents[i];
			const headerBytes = header.toBytes();
			blobPart.push(headerBytes, body);
			offset += headerBytes.length + body.length;
		}
		const cdBytes = this.centralDirectory.map(cd => cd.toBytes());
		blobPart.push(...cdBytes);
		const eocd = new EndOfCentralDirectoryRecord();
		eocd.numOfFiles = len;
		eocd.totalNumOfFiles = len;
		eocd.cdSize = cdBytes.reduce((a, c) => a + c.length, 0);
		eocd.cdOffset = offset;
		eocd.commentBytes = Builder.textEncoder.encode(options?.comment);
		blobPart.push(eocd.toBytes());
		return new Blob(blobPart, { type: 'application/zip' });
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
 * @prop {() => Uint8Array<ArrayBuffer>} toBytes
 * @static @prop {(cursor: ByteArrayCursor) => PackageRecord} from
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
	/** @type {Uint8Array} */ fileNameBytes;
	/** @type {Uint8Array} */ extraFieldBytes;

	/**
	 * Creates a new "Local File Header" record from bytes.
	 * @param {ByteArrayCursor} cursor byte array cursor
	 * @returns {LocalFileHeader} new "Local File Header" record
	 */
	static from(cursor) {
		const record = new LocalFileHeader();
		record.signature = cursor.readAsUint();
		record.versionNeeded = cursor.readAsUshort();
		record.flags = cursor.readAsUshort();
		record.method = cursor.readAsUshort();
		record.lastModified = ZipDateTime.fromInt(cursor.readAsUint());
		record.crc32 = cursor.readAsUint();
		record.compressedSize = cursor.readAsUint();
		record.uncompressedSize = cursor.readAsUint();

		const fileNameLength = cursor.readAsUshort();
		const extraFieldLength = cursor.readAsUshort();
		record.fileNameBytes = cursor.subarray(fileNameLength);
		record.extraFieldBytes = cursor.subarray(extraFieldLength);
		return record;
	}

	toBytes() {
		const fileNameLength = this.fileNameBytes.length;
		const extraFieldLength = this.extraFieldBytes.length;
		const bytes = new Uint8Array(30 + fileNameLength + extraFieldLength);
		const cursor = new ByteArrayCursor(bytes);
		cursor.writeInt(this.signature);
		cursor.writeUshort(this.versionNeeded);
		cursor.writeUshort(this.flags);
		cursor.writeUshort(this.method);
		cursor.writeInt(this.lastModified.toInt());
		cursor.writeInt(this.crc32);
		cursor.writeInt(this.compressedSize);
		cursor.writeInt(this.uncompressedSize);
		cursor.writeUshort(fileNameLength);
		cursor.writeUshort(extraFieldLength);
		cursor.set(this.fileNameBytes);
		cursor.set(this.extraFieldBytes);
		return bytes;
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
 * @extends LocalFileHeader
 */
class CentralDirectoryEntry extends LocalFileHeader {
	/** @type {number} */ signature = 0x02014b50;
	/** @type {number} */ versionMadeBy = 20;
	/** @type {number} */ diskIdStart = 0;
	/** @type {number} */ internalAttributes = 0;
	/** @type {number} */ externalAttributes = 0;
	/** @type {number} */ headerOffset;
	/** @type {Uint8Array} */ commentBytes;

	/**
	 * Creates a new "Central Directory" entry from bytes.
	 * @param {ByteArrayCursor} cursor byte array cursor
	 * @returns {CentralDirectoryEntry} new "Central Directory" entry
	 */
	static from(cursor) {
		const record = new CentralDirectoryEntry();
		record.signature = cursor.readAsUint();
		record.versionMadeBy = cursor.readAsUshort();
		record.versionNeeded = cursor.readAsUshort();
		record.flags = cursor.readAsUshort();
		record.method = cursor.readAsUshort();
		record.lastModified = ZipDateTime.fromInt(cursor.readAsUint());
		record.crc32 = cursor.readAsUint();
		record.compressedSize = cursor.readAsUint();
		record.uncompressedSize = cursor.readAsUint();
		const fileNameLength = cursor.readAsUshort();
		const extraFieldLength = cursor.readAsUshort();
		const commentLength = cursor.readAsUshort();
		record.diskIdStart = cursor.readAsUshort();
		record.internalAttributes = cursor.readAsUshort();
		record.externalAttributes = cursor.readAsUint();
		record.headerOffset = cursor.readAsUint();
		record.fileNameBytes = cursor.subarray(fileNameLength);
		record.extraFieldBytes = cursor.subarray(extraFieldLength);
		record.commentBytes = cursor.subarray(commentLength);
		return record;
	}

	toBytes() {
		const fileNameLength = this.fileNameBytes.length;
		const extraFieldLength = this.extraFieldBytes.length;
		const commentLength = this.commentBytes.length;
		const bytes = new Uint8Array(46 + fileNameLength + extraFieldLength + commentLength);
		const cursor = new ByteArrayCursor(bytes);
		cursor.writeInt(this.signature);
		cursor.writeUshort(this.versionMadeBy);
		cursor.writeUshort(this.versionNeeded);
		cursor.writeUshort(this.flags);
		cursor.writeUshort(this.method);
		cursor.writeInt(this.lastModified.toInt());
		cursor.writeInt(this.crc32);
		cursor.writeInt(this.compressedSize);
		cursor.writeInt(this.uncompressedSize);
		cursor.writeUshort(fileNameLength);
		cursor.writeUshort(extraFieldLength);
		cursor.writeUshort(commentLength);
		cursor.writeUshort(this.diskIdStart);
		cursor.writeUshort(this.internalAttributes);
		cursor.writeInt(this.externalAttributes);
		cursor.writeInt(this.headerOffset);
		cursor.set(this.fileNameBytes);
		cursor.set(this.extraFieldBytes);
		cursor.set(this.commentBytes);
		return bytes;
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
	/** @type {Uint8Array} */ commentBytes;

	/**
	 * Creates a new <abbr>EOCD</abbr> (End of Central Directory) record from bytes.
	 * @param {ByteArrayCursor} cursor byte array cursor
	 * @returns {EndOfCentralDirectoryRecord} new EOCD record
	 */
	static from(cursor) {
		const record = new EndOfCentralDirectoryRecord();
		record.signature = cursor.readAsUint();
		record.diskId = cursor.readAsUshort();
		record.firstDiskId = cursor.readAsUshort();
		record.numOfFiles = cursor.readAsUshort();
		record.totalNumOfFiles = cursor.readAsUshort();
		record.cdSize = cursor.readAsUint();
		record.cdOffset = cursor.readAsUint();
		const commentLength = cursor.readAsUshort();
		record.commentBytes = cursor.subarray(commentLength);
		return record;
	}

	toBytes() {
		const commentLength = this.commentBytes.length;
		const bytes = new Uint8Array(22 + commentLength);
		const cursor = new ByteArrayCursor(bytes);
		cursor.writeInt(this.signature);
		cursor.writeUshort(this.diskId);
		cursor.writeUshort(this.firstDiskId);
		cursor.writeUshort(this.numOfFiles);
		cursor.writeUshort(this.totalNumOfFiles);
		cursor.writeInt(this.cdSize);
		cursor.writeInt(this.cdOffset);
		cursor.writeUshort(commentLength);
		cursor.set(this.commentBytes);
		return bytes;
	}
}

/**
 * Class for viewing and editing ArrayBuffer in little endian
 */
class ByteArrayCursor {
	/** @type {number} */ #offset = 0;
	/** @type {Uint8Array<ArrayBuffer>} */ bytes;

	/**
	 * @returns {number} Offset index
	 */
	get offset() {
		return this.#offset;
	}

	/**
	 * @param {number} v New offset index
	 */
	set offset(v) {
		if (v <= this.bytes.length) this.#offset = v;
		else throw new Error('Index is invalid: ' + v);
	}

	/**
	 * Creates a cursor of byte array.
	 * @param {Uint8Array<ArrayBuffer>} bytes Byte array
	 */
	constructor(bytes) {
		this.bytes = bytes;
	}

	/**
	 * Reads the value at the current offset as a byte.
	 * @returns {number} Byte
	 */
	read() {
		return this.bytes[this.offset++];
	}

	/**
	 * Writes a byte at the current offset.
	 * @param {number} b Byte
	 */
	write(b) {
		this.bytes[this.offset++] = b;
	}

	/**
	 * Reads the value at the current offset as an unsigned short.
	 * @returns {number} Unsigned short
	 */
	readAsUshort() {
		return this.read() + (this.read() << 8);
	}

	/**
	 * Writes an unsigned short at the current offset.
	 * @param {number} s Unsigned short
	 */
	writeUshort(s) {
		this.write(s & 0xff);
		this.write(s >>> 8 & 0xff);
	}

	/**
	 * Reads the value at the current offset as an signed integer.
	 * @returns {number} Signed integer
	 */
	readAsInt() {
		return this.read() + (this.read() << 8) + (this.read() << 16) + (this.read() << 24);
	}

	/**
	 * Writes a signed/an unsigned integer at the current offset.
	 * @param {number} i Signed/unsigned integer
	 */
	writeInt(i) {
		this.write(i & 0xff);
		this.write(i >>> 8 & 0xff);
		this.write(i >>> 16 & 0xff);
		this.write(i >>> 24 & 0xff);
	}

	/**
	 * Reads the value at the current offset as an unsigned integer.
	 * @returns {number} Unsigned integer
	 */
	readAsUint() {
		return this.readAsInt() >>> 0;
	}

	/**
	 * Gets a new Uint8Array view at the current offset.
	 * @param {number} length Length of subarray
	 */
	subarray(length) {
		const sub = this.bytes.subarray(this.offset, this.offset + length);
		this.offset += length;
		return sub;
	}

	/**
	 * Sets a byte array at the current offset.
	 * @param {Uint8Array} bytes Byte array
	 */
	set(bytes) {
		const nextOffset = bytes.length + this.offset;
		if (this.bytes.length < nextOffset) {
			throw new Error('Index is invalid: ' + nextOffset);
		}
		this.bytes.set(bytes, this.offset);
		this.offset = nextOffset;
	}

	/**
	 * Moves cursor to index.
	 * @param {number} index Index of byte array
	 */
	moveTo(index) {
		this.offset = index;
	}
}