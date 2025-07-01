import * as ZipFile from '../zip.js';
import { WorkbookPackage, WorkbookPart, WorksheetPart, CellReference, Row, Cell, StylesPart, RichTextRun, CellText } from './spreadsheetml.mjs'

const $main = /** @type {} */ document.getElementById('main');
const $tablist = document.getElementById('tablist')

/** @type {HTMLFormElement} */
const $form = document.forms['form'];

/** @type {HTMLInputElement} */
const $file = $form.elements['file'];

$form.addEventListener('submit', async e => {
	e.preventDefault();
	if ($main) $main.innerHTML = '';
	if ($tablist) $tablist.innerHTML = '';
	
	const file = $file.files?.[0];
	if (!file) return;
	
	const zf = new ZipFile.Extractor(await file.arrayBuffer());
	const pkg = new WorkbookPackage(zf);
	const wb = await pkg.getWorkbook();
	
	const sections = wb.sheets.map(s => {
		const $section = document.createElement('section');
		$section.id = `sheet-${s.sheetId}`;
		$section.dataset.sheetName = `${s.name}`;
		return $section;
	});
	$main?.append(...sections);

	const buttons = wb.sheets.map(s => {
		const $button = document.createElement('button');
		$button.textContent = s.name;
		$button.id = `button-${s.sheetId}`;
		const hash = `#sheet-${s.sheetId}`;
		$button.addEventListener('click', async () => {
			const section = /** @type {HTMLElement?} */ (document.querySelector(hash));
			if (s.part && section) {
				buttons.forEach(b => b.disabled = true);
				try {
					document.body.style.cursor = 'wait';
					await renderWorksheet(s.part, section);
				} catch (e) {
					alert('Sorry, failed to open the workbook.\n' + e.message);
				} finally {
					document.body.style.cursor = '';
				}
			}
		}, { once: true });
		$button.addEventListener('click', () => {
			location.hash = hash;
			buttons.forEach(b => b.disabled = b.id === $button.id);
		});
		return $button;
	});
	$tablist?.append(...buttons);

	await wb.fetchParts();
	console.log(wb);
	buttons[0].click();
});

/**
 * Renders a worksheet as `<section>`.
 * @param {WorksheetPart} wsPart worksheet part
 * @param {HTMLElement} $section output `<section>` element
 */
async function renderWorksheet(wsPart, $section) {
	await sleep(0);
	const dimensionEnd = wsPart.dimension.end;

	// layout
	const $topleft = document.createElement('div');
	$topleft.classList.add('header-topleft');
	const $topright = document.createElement('div');
	$topright.classList.add('header-columns');
	const $bottomleft = document.createElement('div');
	$bottomleft.classList.add('header-rows');

	// heading row
	for (let c = 1; c <= dimensionEnd.c; c++) {
		const $cell = document.createElement('div');
		$cell.classList.add('header');
		$cell.dataset.text = new CellReference(1, c).getColName();
		$topright.appendChild($cell);
	}

	// colgroup
	const cols = wsPart.cols;
	if (cols.length > 0) {
		const colWidths = new Array(dimensionEnd.c + 1).fill('72px');
		colWidths[0] = 'auto';
		for (const col of wsPart.cols) {
			let c = col.min;
			const cm = Math.min(col.max || Infinity, dimensionEnd.c);
			while (c && cm && c <= cm) {
				if (col.customWidth) colWidths[c] = `calc(${col.width}ch + 5px)`;
				c++;
			}
		}
		$section.style.gridTemplateColumns = colWidths.join(' ');
	} else {
		$section.style.gridTemplateColumns = `auto repeat(${dimensionEnd.c}, 72px)`;
	}
	
	$section.append($topleft, $topright, $bottomleft);

	const $contents = document.createElement('div');
	$contents.classList.add('contents');
	
	// data rows
	for (let r = 1; r <= dimensionEnd.r; r++) {
		const $headCell = document.createElement('div');
		$headCell.classList.add('header');
		$headCell.dataset.text = `${r}`;
		$bottomleft.appendChild($headCell);
		for (let c = 1; c <= dimensionEnd.c; c++) {
			const $cell = document.createElement('div');
			$cell.classList.add('cell');
			const ref = new CellReference(r, c);
			$cell.id = `cell-${ref}`;
			$cell.style.gridArea = `${ref.r + 1} / ${ref.c + 1}`;
			$contents.appendChild($cell);
		}
	}

	// merged cells
	for (const mergeCell of wsPart.mergeCells) {
		const range = mergeCell.ref;
		if (range) {
			const $beginCell = /** @type {HTMLElement?} */ ($contents.querySelector(`#cell-${range.begin}`));
			if ($beginCell) {
				$beginCell.classList.add('merged');
				$beginCell.style.gridArea = `${range.begin.r + 1} / ${range.begin.c + 1} / span ${range.height} / span ${range.width}`;
				let w = range.width, h = range.begin.r;
				while (w-- > 1) $beginCell.nextElementSibling?.remove();
				while (++h <= range.end.r) {
					const ref = new CellReference(h, range.begin.c);
					const $cell = $contents.querySelector(`#cell-${ref}`);
					w = range.width;
					while (w-- > 1) $cell?.nextElementSibling?.remove();
					$cell?.remove();
				}
			}
		}
	}
	$section.appendChild($contents);

	// cell value
	const rowHeights = new Array(dimensionEnd.r + 1).fill('25px');
	const sheetData = wsPart.sheetData;
	for (const row of sheetData) {
		if (row.index && row.height) {
			rowHeights[row.index] = `${row.height}pt`;
		}
		for (const c of row.cells) {
			const $cell = /** @type {HTMLElement?} */ ($contents.querySelector(`#cell-${c.ref}`));
			if (!$cell) continue;
			$cell.dataset.type = c.dataType;
			$cell.dataset.style = `${c.styleIndex}`;
			formatCell($cell, c, wsPart.workbook);
		}
	}
	
	$section.style.gridTemplateRows = rowHeights.some(ht => ht !== '25px') ? rowHeights.join(' ') : `25px repeat(${dimensionEnd.r}, 25px)`;
}

/**
 * Creates a cell value element.
 * @param {HTMLElement} $cell HTML cell element
 * @param {Cell} cell Cell object of SpreadsheetML cell element
 * @param {WorkbookPart} wbpart workbook part
 */
function formatCell($cell, cell, wbpart) {
	const v = cell.value;
	if (!v) return;
	const stylesPart = wbpart.stylesPart;
	let text;
	if (cell.dataType === 's') {
		const si = wbpart.sharedStringList[/** @type {number} */ (v)];
		const el = si.toHTMLElement(wbpart);
		if (el instanceof DocumentFragment) {
			$cell.appendChild(el);
			return;
		} else {
			text = el.textContent;
		}
	} else if (cell.dataType === 'inlineStr') {
		const fragment = document.createDocumentFragment();
		for (const child of /** @type {Element} */ (v).children) {
			switch (child.tagName) {
				case 't':
					fragment.appendChild(new CellText(child).toHTMLElement());
					break;
				case 'r':
					const last = /** @type {HTMLElement?} */ (fragment.lastElementChild);
					if (last && child?.style.cssText === last?.style.cssText) {
						last.appendChild(new Text(child.textContent || ''));
					} else {
						fragment.appendChild(new RichTextRun(child, wbpart).toHTMLElement());
					}
			}
		}
		$cell.appendChild(fragment);
		return;
	}
	text ||= cell.valueAsString;
	if (cell.styleIndex && stylesPart) {
		const xf = stylesPart.cellFormats[cell.styleIndex];
		const font = stylesPart.fonts[xf.fontId];
		/** @type {"sup" | "sub" | "span"} */
		const tagName = ({ superscript: 'sup', subscript: 'sub' })[font.verticalAlign] || 'span';
		/** @type {HTMLElement} */
		const span = document.createElement(tagName);
		if (xf.fontId > 0) {
			if (font.name) span.style.fontFamily = `"${font.name}"`;
			if (font.bold) span.style.fontWeight = 'bold';
			if (font.italic) span.style.fontStyle = 'italic';
			if (font.strike) span.style.textDecorationLine = 'line-through';
			if (font.color) span.style.color = font.color.toString();
			if (font.size) span.style.fontSize = `${font.size}pt`;
			switch (font.underline) {
				case 'single':
				case 'singleAccounting':
					span.style.borderBottom = 'thin solid';
					break;
				case 'double':
				case 'doubleAccounting':
					span.style.borderBottom = 'medium double'
			}
		}
		if (xf.alignment) {
			const horizontal = xf.alignment.horizontal;
			switch (horizontal) {
				case 'left':
				case 'center':
				case 'right':
				case 'justify':
					span.style.textAlign = horizontal;
					break;
				case 'fill':
				case 'distributed':
					span.style.textAlign = 'justify';
					break;
				case 'centerContinuous':
					span.style.textAlign = 'center';
					break;
			}
			const vertical = xf.alignment.vertical;
			if (vertical) span.classList.add(`vertical-${vertical}`);
			const wrapText = xf.alignment.wrapText;
			if (wrapText) span.classList.add('wrap');
		}
		if (xf.borderId > 0) {
			const border = stylesPart.borders[xf.borderId];
			const borderMap = {
				none: 'none',
				thin: 'thin solid',
				medium: 'medium solid',
				dashed: 'thin dashed',
				dotted: 'thin dotted',
				thick: 'thick solid',
				double: 'medium double',
				hair: '.5px solid',
				mediumDashed: 'medium dashed',
				dashDot: 'thin dashed',
				mediumDashdot: 'medium dashed',
				dashDotDot: 'thin dashed',
				mediumDashDotDot: 'medium dashed',
				slantDashDot: 'none',
			};
			for (const dir of ['top', 'right', 'bottom', 'left']) {
				const { style, color } = border[dir];
				$cell.style.setProperty(`border-${dir}`, borderMap[style] + (color ? ' ' + color.toString() : ''));
				if (style !== 'none') $cell.classList.add(`border-${dir}`);
			}
		}
		span.textContent = text;
		$cell.appendChild(span);
		return;
	} else {
		$cell.appendChild(new Text(text));
		return;
	}
}

/**
 * Waits a few seconds.
 * Use with `await`.
 * @param {number} ms milliseconds
 */
function sleep(ms) {
	return new Promise(resolve => {
		setTimeout(resolve, ms);
	});
}