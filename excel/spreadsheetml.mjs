import * as ZipFile from '../zip.js';
import { Theme } from './drawingml.mjs';

const domParser = new DOMParser();

const MESSAGE = {
	RANGE_ERROR_POSITIVE_INT: 'The value must be a positive integer.',
	RANGE_ERROR_POSITIVE_INT_OR_ZERO: 'The value must be a positive integer or zero.',
	RANGE_ERROR_POSITIVE_INT_OR_NULL: 'The value must be a positive integer or null.',
};

const NAMESPACES = {
	STRICT: {
		sml: 'http://purl.oclc.org/ooxml/spreadsheetml/main',
	},
	TRANSITIONAL: {
		sml: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	},
}

const ns = NAMESPACES.STRICT;

/**
 * 
 * @param {string} text XML string
 * @returns XML document
 */
const xml = text => domParser.parseFromString(text, 'application/xml');

export class WorkbookPackage {
	/**
	 * 
	 * @param {ZipFile.Extractor} zf 
	 */
	constructor(zf) {
		this.package = zf;
		this.filepaths = zf.getNames();
	}
	/**
	 * 
	 * @param {string} path filepath
	 * @returns {Response} content response
	 */
	fetch(path) {
		const index = this.filepaths.indexOf(path);
		if (index < 0) throw new Error('Not found: ' + path);
		return this.package.pick(index);
	}
	async getWorkbook() {
		if (!this.workbook) {
			const wbPath = 'xl/workbook.xml';
			const doc = await this.fetch(wbPath).text().then(xml);
			this.workbook = WorkbookPart.from(this, doc, wbPath);
		}
		return this.workbook;
	}
}

export class WorkbookPart {
	/** @type {WorkbookPackage} */ package;
	/** @type {Document} */ doc;
	/** @type {string[]} */ path = [];
	/** @type {Array<{ name: string?, sheetId: string?, id: string?, part?: WorksheetPart }>} */ sheets = [];
	/** @type {Theme[]} */ themes = [];
	/** @type {Array<{ id: string, type: string, target: string, document: Document }>} */ parts = [];
	/** @type {StylesPart?} */ stylesPart = null;
	/** @type {StringItem[]} */ sharedStringList = [];
	/**
	 * @param {WorkbookPackage} pkg 
	 * @param {Document} doc
	 * @param {string} path  
	 */
	static from(pkg, doc, path) {
		const self = new WorkbookPart;
		self.package = pkg;
		self.doc = doc;
		self.path = path.split('/');
		self.sheets = Array.from(doc.querySelectorAll('sheet'), s => {
			const name = s.getAttribute('name');
			const sheetId = s.getAttribute('sheetId');
			const id = s.getAttribute('r:id');
			return { name, sheetId, id };
		});
		/** @type {Theme[]} */
		self.themes = [];
		return self;
	}

	async fetchRels() {
		if (!this.rels) {
			const path = this.path.with(-1, '_rels/' +  this.path.at(-1) + '.rels').join('/');
			const doc = await this.package.fetch(path).text().then(xml);
			this.rels = doc;
			return doc;
		}
		return this.rels;
	}

	async fetchParts() {
		this.parts = [];
		const rels = (await this.fetchRels()).documentElement;
		for (const rel of rels.children) {
			const id = rel.getAttribute('Id');
			const type = rel.getAttribute('Type')?.split('/').at(-1);
			const target = rel.getAttribute('Target');
			if (id && type && target) {
				const path = target.startsWith('/') ? target.substring(1) : 'xl/' + target;
				const doc = await this.package.fetch(path).text().then(xml);
				switch (type) {
					case 'styles':
						this.stylesPart = StylesPart.from(doc, this);
						break;
					case 'sharedStrings':
						this.sharedStringList = Array.from(doc.querySelectorAll('si'), si => new StringItem(si));
						break;
					case 'theme':
						this.themes.push(new Theme(doc.documentElement));
						break;
					case 'worksheet':
						const sheet = this.sheets.find(s => s.id === id);
						if (sheet) sheet.part = WorksheetPart.from(this, doc, path);
				}
				this.parts.push({ id, type, target, document: doc });
			}
		}
		return this.parts;
	}
}

export class WorksheetPart {
	/** @type {WorkbookPart} */ workbook;
	/** @type {Document} */ doc;
	/** @type {string[]} */ path = [];
	/**
	 * @param {WorkbookPart} wb 
	 * @param {Document} doc
	 * @param {string} path  
	 */
	static from(wb, doc, path) {
		const self = new WorksheetPart;
		self.workbook = wb;
		self.doc = doc;
		self.path = path.split('/');
		return self;
	}

	get dimension() {
		const el = this.doc.querySelector('dimension');
		const ref = el?.getAttribute('ref') || '';
		return new Range(ref);
	}
	set dimension(range) {
		const el = this.doc.querySelector('dimension');
		if (el) {
			el.setAttribute('ref', `${range}`);
		} else {
			const newEl = this.doc.createElement('dimension');
			newEl.setAttribute('ref', `${range}`);
			this.doc.documentElement.appendChild(newEl);
		} 
	}
	get cols() {
		const el = this.doc.querySelector('cols');
		return el ? Array.from(el.children, child => new Col(child)) : [];
	}
	get sheetData() {
		const el = /** @type {Element} */ (this.doc.querySelector('sheetData'));
		return Array.from(el.children, child => new Row(child));
	}
	get mergeCells() {
		const el = this.doc.querySelector('mergeCells');
		return el ? Array.from(el.children, child => new MergeCell(child)) : [];
	}
}

export class StylesPart {
	/** @type {Document} */ doc;
	/** @type {WorkbookPart} */ workbook;

	/**
	 * 
	 * @param {Document} doc 
	 * @param {WorkbookPart} wb 
	 */
	static from(doc, wb) {
		const self = new StylesPart;
		self.doc = doc;
		self.workbook = wb;
		return self;
	}

	get borders() {
		const borders = this.doc.querySelector('borders');
		return Array.from(borders?.children || [], border => new Border(border));
	}

	get cellFormats() {
		const cellXfs = this.doc.querySelector('cellXfs');
		return Array.from(cellXfs?.children || [], xf => new Format(xf));
	}

	get fonts() {
		const fonts = this.doc.querySelector('fonts');
		return Array.from(fonts?.children || [], font => new Font(font, this.workbook));
	}
}

export class Col {
	/** @type {Element} */ element;

	/**
	 * 
	 * @param {Element} [col] `<col>` element
	 */
	constructor(col) {
		this.element = col ?? document.createElementNS(ns.sml, 'col');
	}

	get min() {
		const min = this.element.getAttribute('min');
		return min ? Number.parseInt(min) : null;
	}
	set min(i) {
		if (i && i > 0) this.element.setAttribute('min', i.toFixed());
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT);
	}
	
	get max() {
		const max = this.element.getAttribute('max');
		return max ? Number.parseInt(max) : null;
	}
	set max(i) {
		if (i && i > 0) this.element.setAttribute('max', i.toFixed());
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT);
	}

	get width() {
		const width = this.element.getAttribute('width');
		return width ? Number.parseFloat(width) : null;
	}
	set width(d) {
		if (d && d > 0) this.element.setAttribute('width', `${d}`);
		else if (d === null) this.element?.removeAttribute('width');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get style() {
		const style = this.element.getAttribute('style');
		return style ? Number.parseInt(style) : 0;
	}
	set style(i) {
		if (i > 0) this.element.setAttribute('style', i.toFixed());
		else if (!i) this.element.removeAttribute('style');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_ZERO);
	}

	get hidden() {
		return hasAttributeTruthy(this.element, 'hidden');
	}
	set hidden(b) {
		if (b) this.element.setAttribute('hidden', '1');
		else this.element.removeAttribute('hidden');
	}

	get bestFit() {
		return hasAttributeTruthy(this.element, 'bestFit');
	}
	set bestFit(b) {
		if (b) this.element.setAttribute('bestFit', '1');
		else this.element.removeAttribute('bestFit');
	}

	get customWidth() {
		return hasAttributeTruthy(this.element, 'customWidth');
	}
	set customWidth(b) {
		if (b) this.element.setAttribute('customWidth', '1');
		else this.element.removeAttribute('customWidth');
	}

	get phonetic() {
		return hasAttributeTruthy(this.element, 'phonetic');
	}
	set phonetic(b) {
		if (b) this.element.setAttribute('phonetic', '1');
		else this.element.removeAttribute('phonetic');
	}

	get outlineLevel() {
		const ol = this.element.getAttribute('outlineLevel');
		return ol ? Number.parseInt(ol) : 0;
	}
	set outlineLevel(i) {
		if (i > 0) this.element.setAttribute('outlineLevel', i.toFixed());
		else if (!i) this.element.removeAttribute('outlineLevel');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_ZERO);
	}

	get collapsed() {
		return hasAttributeTruthy(this.element, 'collapsed');
	}
	set collapsed(b) {
		if (b) this.element.setAttribute('collapsed', '1');
		else this.element.removeAttribute('collapsed');
	}
}

export class Row {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [row] `<row>` element
	 */
	constructor(row) {
		this.element = row ?? document.createElementNS(ns.sml, 'r');
	}

	get index() {
		const r = this.element.getAttribute('r');
		return r ? Number.parseInt(r) : null;
	}
	set index(i) {
		if (i && i > 0) this.element.setAttribute('r', i.toFixed());
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT)
	}

	get span() {
		return this.element.getAttribute('spans')?.split(':').map(Number.parseInt) || [];
	}
	set span([start, end]) {
		if (start && end && start <= end) this.element.setAttribute('span', `${start}:${end}`);
		else if (start === null) this.element.removeAttribute('span');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL)
	}

	get height() {
		const ht = this.element.getAttribute('ht');
		return ht ? Number.parseFloat(ht) : null;
	}
	set height(d) {
		if (d && d > 0) this.element.setAttribute('ht', `${d}`);
		else if (d === null) this.element.removeAttribute('ht');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get hidden() {
		return hasAttributeTruthy(this.element, 'hidden');
	}
	set hidden(b) {
		if (b) this.element.setAttribute('hidden', '1');
		else this.element.removeAttribute('hidden');
	}

	get outlineLevel() {
		const ol = this.element.getAttribute('outlineLevel');
		return ol ? Number.parseInt(ol) : null;
	}
	set outlineLevel(i) {
		if (i && i > 0) this.element.setAttribute('outlineLevel', i.toFixed());
		else if (i === null) this.element.removeAttribute('outlineLevel');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get collapsed() {
		return hasAttributeTruthy(this.element, 'collapsed');
	}
	set collapsed(b) {
		if (b) this.element.setAttribute('collapsed', '1');
		else this.element.removeAttribute('collapsed');
	}

	get cells() {
		return Array.from(this.element.children || [], c => new Cell(c));
	}
}

/**
 * @typedef {"b" | "d" | "e" | "inlineStr" | "n" | "s" | "str"} CellDataType
 */

export class Cell {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [cell] 
	 */
	constructor(cell) {
		this.element = cell ?? document.createElementNS(ns.sml, 'c');
	}

	get ref() {
		const r = this.element.getAttribute('r');
		return r ? new CellReference(r) : null;
	}
	set ref(s) {
		this.element.setAttribute('r', `${s}`);
	}

	get dataType() {
		const t = /** @type {CellDataType?} */ (this.element.getAttribute('t'));
		return t || 'n';
	}
	set dataType(s) {
		this.element.setAttribute('t', s);
	}

	get styleIndex() {
		const s = this.element.getAttribute('s');
		return s ? Number.parseInt(s) : null;
	}
	set styleIndex(i) {
		if (i && i > 0) this.element.setAttribute('s', `${i}`);
		else if (i === null) this.element.removeAttribute('s');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get showPhonetic() {
		return hasAttributeTruthy(this.element, 'ph');
	}
	set showPhonetic(b) {
		if (b) this.element.setAttribute('ph', '1');
		else this.element.removeAttribute('ph');
	}

	get value() {
		const v = this.element.querySelector('v');
		const text = v?.textContent || null;
		switch (this.dataType) {
			case 'n':
				return text ? Number.parseFloat(text) : NaN;
			case 'b':
				return false;
			case 'd':
				return text ? new Date(text) : null;
			case 'e':
			case 'str':
				return text;
			case 's':
				return text ? Number.parseFloat(text) : null;
			case 'inlineStr':
				return this.element.querySelector('is') || null;
				
		}
	}

	/**
	 * @param {boolean | Date | Element | number | string | null} val 
	 */
	set value(val) {
		if (!this.element) return;
		let v = this.element.querySelector('v');
		if (val === null) {
			v?.remove();
			return;
		}
		if (!v) {
			v = this.element.ownerDocument.createElement('v');
			this.element.appendChild(v);
		}
		if (val instanceof Date) {
			this.dataType = 'd';
			v.textContent = val.toISOString();
			return;
		} else if (val instanceof Element) {
			this.dataType = 'inlineStr';
			v.appendChild(val);
			return;
		}
		switch (typeof val) {
			case 'string':
				this.dataType = 'str';
				v.textContent = val;
				return;
			case 'number':
				this.dataType = 'n';
				v.textContent = val.toString();
				return;
			case 'boolean':
				this.dataType = 'b';
				v.textContent = val ? 'true' : 'false';
				return;
		}
	}

	get valueAsString() {
		return this.value?.toString() || '';
	}
}

export class MergeCell {
	/** @type {Element} */ element;

	/**
	 * @param {Element} [element] 
	 */
	constructor(element) {
		this.element = element ?? document.createElementNS(ns.sml, 'mergeCell');
	}

	get ref() {
		const ref = this.element.getAttribute('ref');
		return ref ? new Range(ref) : null;
	}
	set ref(range) {
		this.element.setAttribute('ref', `${range}`);
	}
}

export class CellReference {
	static RE_A1_FORMAT = /^(\$?)([A-Z]+)(\$?)(\d+)$/i;
	static RE_RC_FORMAT = /^R(\d+|\[\-?\d+\])C(\d+|\[\-?\d+\])$/i;
	r = -1;
	c = -1;
	abs = { r: false, c: false };
	/**
	 * @overload
	 * @param {string} address A1-style or R1C1-style address
	 */
	/**
	 * @overload
	 * @param {number} r row index
	 * @param {number} c column index
	 */
	/**
	 * @param {string[] | number[]} args 
	 */
	constructor(...args) {
		const [arg1, arg2] = args;
		switch (typeof arg1) {
			case 'string':
				let match = CellReference.RE_A1_FORMAT.exec(arg1);
				if (match) {
					const [_, $c, c, $r, r] = match;
					this.abs.r = $r !== '';
					this.abs.c = $c !== '';
					this.r = Number.parseInt(r);
					const ord = Array.from(c.toUpperCase(), ch => ch.charCodeAt(0) - 64);
					this.c = ord.reduce((a, c) => 26 * a + c, 0);
					return this;
				}
				match = CellReference.RE_RC_FORMAT.exec(arg1);
				if (match) {
					const [_, r, c] = match;
					this.abs.r = !r.startsWith('[');
					this.abs.c = !c.startsWith('[');
					this.r = Number.parseInt(this.abs.r ? r : r.substring(1, r.length - 1));
					this.c = Number.parseInt(this.abs.c ? c : c.substring(1, c.length - 1));
					return this;
				}
				break;
			case 'number':
				if (arg1 > 0 && typeof arg2 === 'number' && arg2 > 0) {
					this.abs.r = this.abs.c = true;
					this.r = arg1;
					this.c = arg2;
					return this;
				}
		}
		throw new Error(`Invalid arguments inputted: ${args.join()}`);
	}

	getColName() {
		/** @type {number[]} */
		const codes = [];
		let c = this.c - 1;
		let mod = c % 26 + 1
		do {
			codes.unshift(mod + 64);
			c /= 26;
			mod = c % 26;
		} while (c >= 1);
		return String.fromCharCode(...codes);
	}

	/**
	 * @param {"A1" | "RC"} style 
	 */
	toString(style = 'A1') {
		switch (style) {
			case 'A1':
				return this.getColName() + this.r;
			case 'RC':
				return '';
		}
	}
}

export class CellText {
	/** @type {Element} */ element;
	/** @type {string?} */ content = null;
	/** @type {boolean} */ spacePreserve = false;
	/**
	 * @param {Element} element `<t>` element
	 */
	constructor(element) {
		this.element = element;
		this.content = element.textContent;
		this.spacePreserve = element.getAttribute('xml:space') === 'preserve';
	}

	toHTMLElement() {
		const span = document.createElement('span');
		if (this.spacePreserve) span.style.whiteSpace = 'pre-wrap';
		span.textContent = this.content;
		return span;
	}
}

export class RichTextRun {
	/** @type {RunProperties?} */ props = null;
	/** @type {CellText?} */ text = null;
	/**
	 * @param {Element} element `<r>` element
	 * @param {WorkbookPart} wb workbook part
	 */
	constructor(element, wb) {
		this.element = element;
		const rPr = element.querySelector('rPr');
		if (rPr) this.props = new RunProperties(rPr, wb);
		const t = element.querySelector('t');
		if (t) this.text = new CellText(t);
	}

	toHTMLElement() {
		/** @type {"sup" | "sub" | "span"} */
		const tagName = ({ superscript: 'sup', subscript: 'sub' })[this.props?.verticalAlign] || 'span';
		/** @type {HTMLElement} */
		const span = document.createElement(tagName);
		if (this.props?.name) span.style.fontFamily = `"${this.props.name}"`;
		if (this.props?.bold) span.style.fontWeight = 'bold';
		if (this.props?.italic) span.style.fontStyle = 'italic';
		if (this.props?.strike) span.style.textDecorationLine = 'line-through';
		if (this.props?.color) span.style.color = this.props.color.toString();
		if (this.props?.size) span.style.fontSize = `${this.props.size}pt`;
		switch (this.props?.underline) {
			case 'single':
			case 'singleAccounting':
				span.style.borderBottom = 'thin solid';
				break;
			case 'double':
			case 'doubleAccounting':
				span.style.borderBottom = 'medium double'
		}
		if (this.text) {
			if (this.text.spacePreserve) span.style.whiteSpace = 'pre-wrap';
			span.textContent = this.text.content;
		}
		return span;
	}
}

export class Range {
	static RE_RANGE_FORMAT = /^(?:'?(.+)'?\!)?(.+):(.+)$/i;
	/**
	 * 
	 * @param {string} range 
	 */
	constructor(range) {
		const match = Range.RE_RANGE_FORMAT.exec(range);
		if (!match) throw new Error('Invalid range format: ' + range);
		const [_, sheetName, begin, end] = match;
		this.sheetName = sheetName || null;
		this.begin = new CellReference(begin);
		this.end = new CellReference(end);
		this.width = this.end.c - this.begin.c + 1;
		this.height = this.end.r - this.begin.r + 1;
	}

	toString() {
		return `${this.begin}:${this.end}`;
	}

	*[Symbol.iterator]() {
		for (let r = this.begin.r; r <= this.end.r; r++) {
			for (let c = this.begin.c; c <= this.end.c; c++) {
				yield new CellReference(r, c);
			}
		}
	}
}

/**
 * @typedef {"none" | "major" | "minor"} FontSchemeValuesEnum
 * @typedef {"center" | "centerContinuous" | "distributed" | "fill" | "general" | "justify" | "left" | "right"} HorizontalAlignmentValuesEnum
 * @typedef {"single" | "double" | "singleAccounting" | "doubleAccounting" | "none"} UnderlineValuesEnum
 * @typedef {"baseline" | "superscript" | "subscript"} VerticalAlignmentRunValuesEnum
 * @typedef {"bottom" | "center" | "distributed" | "justify" | "top"} VerticalAlignmentValuesEnum
 */

export class Alignment {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [element] `<alignment>` element
	 */
	constructor (element) {
		this.element = element ?? document.createElementNS(ns.sml, 'alignment');
	}

	get horizontal() {
		return /** @type {HorizontalAlignmentValuesEnum?} */ (this.element.getAttribute('horizontal'));
	}
	set horizontal(s) {
		if (s) this.element.setAttribute('horizontal', s);
		else this.element.removeAttribute('horizontal');
	}

	get indent() {
		const indent = this.element.getAttribute('indent');
		return indent ? Number.parseInt(indent) : null;
	}
	set indent(i) {
		if (i && i > 0) this.element.setAttribute('indent', i.toFixed());
		else if (i === null) this.element.removeAttribute('indent');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get justifyLastLine() {
		return hasAttributeTruthy(this.element, 'justifyLastLine');
	}
	set justifyLastLine(b) {
		if (b) this.element.setAttribute('justifyLastLine', '1');
		else this.element.removeAttribute('justifyLastLine');
	}

	get readingOrder() {
		const ro = this.element.getAttribute('readingOrder');
		return ro ? Number.parseInt(ro) : null;
	}
	set readingOrder(i) {
		if (i && i > 0) this.element.setAttribute('readingOrder', i.toFixed());
		else if (!i) this.element.removeAttribute('readingOrder');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_ZERO);
	}

	get relativeIndent() {
		const ri = this.element.getAttribute('relativeIndent');
		return ri ? Number.parseInt(ri) : null;
	}
	set relativeIndent(i) {
		if (i && i > 0) this.element.setAttribute('relativeIndent', i.toFixed());
		else if (!i) this.element.removeAttribute('relativeIndent');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_ZERO);
	}

	get shrinkToFit() {
		return hasAttributeTruthy(this.element, 'shrinkToFit');
	}
	set shrinkToFit(b) {
		if (b) this.element.setAttribute('shrinkToFit', '1');
		else this.element.removeAttribute('shrinkToFit');
	}

	get textRotation() {
		const tr = this.element.getAttribute('textRotation');
		return tr ? Number.parseInt(tr) : null;
	}
	set textRotation(i) {
		if (i && 0 < i && i <= 90) this.element.setAttribute('textRotation', i.toFixed());
		else if (i && 90 < i && i <= 180) this.element.setAttribute('textRotation', (90 - i).toFixed());
		else if (!i) this.element.removeAttribute('textRotation');
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_ZERO);
	}

	get vertical() {
		return /** @type {VerticalAlignmentValuesEnum?} */ (this.element.getAttribute('vertical'));
	}
	set vertical(s) {
		if (s) this.element.setAttribute('vertical', s);
		else this.element.removeAttribute('vertical');
	}

	get wrapText() {
		return hasAttributeTruthy(this.element, 'wrapText');
	}
	set wrapText(b) {
		if (b) this.element.setAttribute('wrapText', '1');
		else this.element.removeAttribute('wrapText');
	}
}

export class Border {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [element] 
	 */
	constructor(element) {
		this.element = element ?? document.createElementNS(ns.sml, 'border');
	}

	/**
	 * 
	 * @param {"start" | "end" | "top" | "bottom" | "diagonal" | "vertical" | "horizontal" | "left" | "right"} dir 
	 */
	getProperty(dir) {
		const prop = this.element.querySelector(dir);
		if (prop) {
			const style = prop.getAttribute('style');
			const color = prop.querySelector('color');
			return { style, color };
		}
		return null;
	}

	get left() {
		return this.getProperty('start') || this.getProperty('left');
	}
	get right() {
		return this.getProperty('end') || this.getProperty('right');
	}
	get top() {
		return this.getProperty('top');
	}
	get bottom() {
		return this.getProperty('bottom');
	}
}

export class Color {
	/** @type {Element?} */ element = null;
	/** @type {WorkbookPart} */ workbook;
	/** @type {number?} */ indexed = null;
	/** @type {number[]?} */ rgb = null;
	/** @type {number?} */ themeIndex = null;
	/**
	 * 
	 * @param {Element} color 
	 * @param {WorkbookPart} wb
	 */
	static from(color, wb) {
		const self = new Color;
		self.element = color;
		self.workbook = wb;
		const rgb = color?.getAttribute('rgb');
		if (rgb) {
			const [a, r, g, b] = rgb.split(/([0-9A-Fa-f]{2})/).filter(s => s.length == 2).map(n => Number.parseInt(n, 16));
			self.rgb = [r, g, b, a];
			return self;
		}
		const theme = color?.getAttribute('theme');
		if (theme) {
			self.themeIndex = Number.parseInt(theme);
			return self;
		}
		return self;
	}
	
	toString() {
		if (this.rgb) return `rgb(${this.rgb.join(' ')})`;
		if (this.themeIndex) {
			const theme = this.workbook.themes[this.themeIndex - 1];
			return theme.themeElements?.colorScheme?.dk1?.toString() || '';
		}
		return '';
	}
}

export class RunProperties  {
	static CharacterSet = {
		0: 'ascii', // ANSI
		1: 'utf-8', // DEFAULT
		2: 'x-user-defined', // SYMBOL ?
		77: 'mac', // MAC
		128: 'sjis', // SHIFTJIS
		129: 'korean', // HANGUL
		130: 'windows-1361', // JOHAB (not supported)
		134: 'gb2312', // GB2312
		136: 'big5', // CHINESEBIG5
		161: 'cp1253', // GREEK
		162: 'cp1254', // TURKISH
		163: 'cp1258', // VIETNAMESE
		177: 'cp1255', // HEBREW
		178: 'cp1256', // ARABIC
		186: 'cp1257', // BALTIC
		204: 'koi8-r', // RUSSIAN
		222: 'dos-874', // THAI
		238: 'cp1250', // EASTEUROPE
		255: 'x-user-defined', // OEM
	};

	/** @type {Element} */ element;
	/** @type {WorkbookPart} */ workbook;
	/**
	 * @param {Element} element `<rPr>` element
	 * @param {WorkbookPart} wb workbook part
	 */
	constructor(element, wb) {
		this.element = element;
		this.workbook = wb;
	}
	
	get name() {
		return getChildAttribute(this.element, 'rFont');
	}
	set name(s) {
		setOrRemoveChildAttribute(this.element, 'rFont', s);
	}

	get charset() {
		const val = getChildAttribute(this.element, 'charset');
		return val ? /** @type {keyof RunProperties.CharacterSet} */ (Number.parseInt(val)) : null;
	}
	set charset(i) {
		const el = this.element?.querySelector('charset');
		if (i && i > 0) el?.setAttribute('val', `${i}`);
		else if (i === null) el?.remove();
		else throw new RangeError(MESSAGE.RANGE_ERROR_POSITIVE_INT_OR_NULL);
	}

	get family() {
		return getChildAttribute(this.element, 'family');
	}
	set family(s) {
		setOrRemoveChildAttribute(this.element, 'family', s);
	}

	get bold() {
		return !!this.element?.querySelector('b');
	}
	set bold(b) {
		createOrRemoveChild(this.element, 'b', b);
	}

	get italic() {
		return !!this.element?.querySelector('i');
	}
	set italic(b) {
		createOrRemoveChild(this.element, 'i', b);
	}

	get strike() {
		return !!this.element?.querySelector('strike');
	}
	set strike(b) {
		createOrRemoveChild(this.element, 'strike', b);
	}

	get outline() {
		return !!this.element?.querySelector('outline');
	}
	set outline(b) {
		createOrRemoveChild(this.element, 'outline', b);
	}

	get shadow() {
		return !!this.element?.querySelector('shadow');
	}
	set shadow(b) {
		createOrRemoveChild(this.element, 'shadow', b);
	}
	
	get condense() {
		return !!this.element?.querySelector('condense');
	}
	set condense(b) {
		createOrRemoveChild(this.element, 'condense', b);
	}

	get extend() {
		return !!this.element?.querySelector('extend');
	}
	set extend(b) {
		createOrRemoveChild(this.element, 'extend', b);
	}

	get color() {
		const el = this.element?.querySelector('color');
		return el ? Color.from(el, this.workbook) : null;
	}
	set color(c) {
		const el = this.element?.querySelector('color');
		el?.remove();
		if (c?.element) this.element?.appendChild(c.element);
	}

	get size() {
		const sz = getChildAttribute(this.element, 'sz');
		return sz ? Number.parseFloat(sz) : null;
	}
	set size(d) {
		setOrRemoveChildAttribute(this.element, 'sz', d ? `${d}` : null);
	}

	get underline() {
		return /** @type {UnderlineValuesEnum?} */ (getChildAttribute(this.element, 'u'));
	}
	set underline(s) {
		setOrRemoveChildAttribute(this.element, 'u', s);
	}

	get verticalAlign() {
		return /** @type {VerticalAlignmentRunValuesEnum?} */ (getChildAttribute(this.element, 'vertAlign'));
	}
	set verticalAlign(s) {
		setOrRemoveChildAttribute(this.element, 'vertAlign', s);
	}

	get scheme() {
		return /** @type {FontSchemeValuesEnum?} */ (getChildAttribute(this.element, 'scheme'));
	}
	set scheme(s) {
		setOrRemoveChildAttribute(this.element, 'scheme', s);
	}
}

export class Font extends RunProperties {
	/**
	 * 
	 * @param {Element} element `<font>` element
	 * @param {WorkbookPart} wb workbook part
	 */
	constructor(element, wb) {
		super(element, wb);
	}
	
	/**
	 * @override
	 */
	get name() {
		return getChildAttribute(this.element, 'name');
	}
	/**
	 * @override
	 */
	set name(s) {
		setOrRemoveChildAttribute(this.element, 'name', s);
	}
}

export class Format {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [element] `<xf>` element
	 */
	constructor(element) {
		this.element = element ?? document.createElementNS(ns.sml, 'xf');
	}

	get alignment() {
		const apply = hasAttributeTruthy(this.element, 'applyAlignment');
		if (apply) {
			const alignment = this.element.querySelector('alignment');
			return alignment ? new Alignment(alignment) : null;
		}
		return null;
	}

	get borderId() {
		const apply = hasAttributeTruthy(this.element, 'applyBorder');
		if (apply) {
			const id = this.element.getAttribute('borderId');
			return id ? Number.parseInt(id) : 0;
		}
		return 0;
	}

	get fillId() {
		const apply = hasAttributeTruthy(this.element, 'applyFill');
		if (apply) {
			const id = this.element.getAttribute('fillId');
			return id ? Number.parseInt(id) : 0;
		}
		return 0;
	}

	get fontId() {
		const apply = hasAttributeTruthy(this.element, 'applyFont');
		if (apply) {
			const id = this.element.getAttribute('fontId');
			return id ? Number.parseInt(id) : 0;
		}
		return 0;
	}

	get numberFormatId() {
		const apply = hasAttributeTruthy(this.element, 'applyNumberFormat');
		if (apply) {
			const id = this.element.getAttribute('numFmtId');
			return id ? Number.parseInt(id) : 0;
		}
		return 0;
	}

	get pivotButton() {
		return hasAttributeTruthy(this.element, 'pivotButton');
	}

	get quotePrefix() {
		return hasAttributeTruthy(this.element, 'quotePrefix');
	}

	get formatId() {
		const xfId = this.element.getAttribute('xfId');
		return xfId ? Number.parseInt(xfId) : 0;
	}

	get isProtected() {
		return hasAttributeTruthy(this.element, 'applyProtection');
	}
}

export class StringItem {
	/** @type {Element} */ element;
	/**
	 * 
	 * @param {Element} [element] `<si>` element
	 */
	constructor(element) {
		this.element = element ?? document.createElementNS(ns.sml, 'si');
	}

	/**
	 * 
	 * @param {WorkbookPart} wb 
	 */
	toHTMLElement(wb) {
		const fragment = document.createDocumentFragment();
		for (const child of this.element.children) {
			switch (child.tagName) {
				case 't':
					const text = new CellText(child);
					return text.toHTMLElement();
				case 'r':
					const richText = new RichTextRun(child, wb);
					const el = richText.toHTMLElement();
					const last = /** @type {HTMLElement?} */ (fragment.lastElementChild);
					if (el.style.cssText === last?.style.cssText) {
						last.appendChild(new Text(el.textContent || ''));
					} else {
						fragment.appendChild(el);
					}
			}
		}
		return fragment;
	}
}

/**
 * Gets attribute of child element.
 * @param {Element?} parent this element
 * @param {string} query query string
 * @param {string} [attrName="val"] attribute name (default is `val`)
 * @returns {string?} attribute value
 */
function getChildAttribute(parent, query, attrName = 'val') {
	const el = parent?.querySelector(query);
	return el?.getAttribute(attrName) ?? null;
}

/**
 * Sets attribute of child element.
 * @param {Element} parent this element
 * @param {string} query query string
 * @param {string?} value attribute value
 * @param {string} [attrName="val"] attribute name (default is `val`)
 */
function setOrRemoveChildAttribute(parent, query, value, attrName = 'val') {
	let el = parent.querySelector(query);
	if (!el && value) {
		el = parent.ownerDocument.createElement(query);
		parent.appendChild(el);
	}
	if (value) el?.setAttribute(attrName, value);
	else el?.remove();
}

/**
 * Creates or removes child element.
 * @param {Element} parent this element
 * @param {string} tagName tag name of child element
 * @param {boolean} create create (`true`) or remove (`false`)
 */
function createOrRemoveChild(parent, tagName, create) {
	const child = parent.querySelector(tagName);
	if (create) {
		if (!child) {
			const newChild = parent.ownerDocument.createElement(tagName);
			parent.appendChild(newChild);
		}
	} else {
		child?.remove();
	}
}

/**
 * 
 * @param {Element} element 
 * @param {string} attrName 
 */
function hasAttributeTruthy(element, attrName) {
	const attr = element.getAttribute(attrName);
	return attr === '1' || attr === 'true';
}