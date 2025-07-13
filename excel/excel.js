import * as ZipFile from '../zip.js';
import { WorkbookPackage, WorkbookPart, WorksheetPart, CellReference, Row, Cell, StylesPart, RichTextRun, CellText, Color } from './spreadsheetml.mjs'

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
					console.error(e);
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
	const root = $section.attachShadow({ mode: 'closed' });
	await sleep(0);

	// style
	const $link = document.createElement('link');
	$link.rel = 'stylesheet';
	$link.href = 'sheet.css';
	root.append($link);

	const css = new CSSStyleSheet();
	const stylesPart = wsPart.workbook.stylesPart;
	if (stylesPart)
	stylesPart.cellFormats.forEach((xf, i) => {
		/** @type {Map<string, string>} */
		const props = new Map();
		if (xf.fontId > 0) {
			const font = stylesPart.fonts[xf.fontId];
			if (font.name) props.set('font-family', `"${font.name}"`);
			if (font.bold) props.set('font-weight', 'bold');
			if (font.italic) props.set('font-style', 'italic');
			if (font.strike) props.set('text-decoration-line', 'line-through');
			if (font.color) props.set('color', font.color.toString());
			if (font.size) props.set('font-size', `${font.size}pt`);
			switch (font.underline) {
				case 'single':
				case 'singleAccounting':
					props.set('border-bottom', 'thin solid');
					break;
				case 'double':
				case 'doubleAccounting':
					props.set('border-bottom', 'medium double');
			}
			switch (font.verticalAlign) {
				case 'subscript':
					props.set('vertical-align', 'sub');
					break;
				case 'superscript':
					props.set('vertical-align', 'super');
			}
		}
		if (xf.alignment) {
			/** @type {Map<string, string>} */
			const childProps = new Map();
			const horizontal = xf.alignment.horizontal;
			switch (horizontal) {
				case 'left':
				case 'center':
				case 'right':
				case 'justify':
					props.set('text-align', horizontal);
					childProps.set('text-align', horizontal);
					break;
				case 'fill':
				case 'distributed':
					props.set('text-align', 'justify');
					childProps.set('text-align', 'justify');
					break;
				case 'centerContinuous':
					props.set('text-align', 'center');
					childProps.set('text-align', 'center');
					break;
			}
			switch (xf.alignment.vertical) {
				case 'top':
					childProps.set('margin-bottom', 'auto');
					break;
				case 'center':
					childProps.set('margin-bottom', 'auto');
				case 'top':
					childProps.set('margin-top', 'auto');
					break;
			}
			if (xf.alignment.wrapText) props.set('white-space', 'pre-wrap !important');
			css.insertRule(`[data-style="${i}"]>span{${Array.from(childProps.entries(), ([k, v]) => k + ':' + v).join(';')}}`);
		}
		if (xf.borderId > 0) {
			const border = stylesPart.borders[xf.borderId];
			const borderDict = {
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
				/** @type { { style: string?, color?: Color } } */
				const { style, color } = border[dir];
				if (style && style !== 'none') {
					props.set(`border-${dir}`, borderDict[style] + (color ? ' ' + color.toString() : ''));
				}
			}
		}
		css.insertRule(`[data-style="${i}"]{${Array.from(props.entries(), ([k, v]) => k + ':' + v).join(';')}}`);
	});
	root.adoptedStyleSheets.push(css);

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
	
	root.append($topleft, $topright, $bottomleft);

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
		$contents.lastElementChild?.classList.add('right');
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
				if (range.end.c === dimensionEnd.c) $beginCell.classList.add('right');
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
	root.appendChild($contents);

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
	let content;
	if (v !== null) {
		if (cell.dataType === 's') {
			const si = wbpart.sharedStringList[/** @type {number} */ (v)];
			const el = si.toHTMLElement(wbpart);
			if (el instanceof DocumentFragment) {
				content = el;
			} else {
				content = el.textContent;
			}
		} else if (cell.dataType === 'inlineStr') {
			const fragment = document.createDocumentFragment();
			for (const child of /** @type {Element} */ (v).children) {
				switch (child.tagName) {
					case 't':
						fragment.appendChild(new CellText(child).toHTMLElement());
						break;
					case 'r':
						// const last = /** @type {HTMLElement?} */ (fragment.lastElementChild);
						fragment.appendChild(new RichTextRun(child, wbpart).toHTMLElement());
				}
			}
			$cell.appendChild(fragment);
		} else {
			console.log(cell.valueAsString);
			content ||= cell.valueAsString;
		}
	}
	if (content)
	if (cell.styleIndex) {
		const span = document.createElement('span');
		span.append(content);
		$cell.appendChild(span);
	} else {
		$cell.append(content);
	}
	// if ($cell.classList.contains('border-top')) {
	// 	const $upperCell = $cell.parentElement?.querySelector(`#cell-${cell.ref?.getColName()}${cell.ref?.r}`);
	// 	if ($upperCell) $upperCell.classList.remove('border-bottom');
	// }
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