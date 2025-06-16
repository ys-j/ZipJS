import * as ZipFile from '../zip.js';

/** @type {HTMLFormElement} */
const $form = document.forms['f'];
$form.addEventListener('submit', onSubmit);
$form.elements['download']?.addEventListener('click', onClickDownload);

/** @type {HTMLTableSectionElement?} */ //@ts-ignore
const $tbody = document.getElementById('tbody');

const builder = new ZipFile.Builder();

/**
 * @param {SubmitEvent} e 
 */
async function onSubmit(e) {
	e.preventDefault();
	const file = /** @type {HTMLInputElement} */ ($form.file)?.files?.[0];
	if (!file) return;

	document.body.style.cursor = 'wait';
	await builder.append(await file.arrayBuffer(), {
		filepath: file.name,
		lastModified: file.lastModified,
		method: ZipFile.acceptMime.includes(file.type) ? 0 : 8,
	});
	document.body.style.cursor = 'auto';

	if (!$tbody) return;
	$tbody.innerHTML = '';
	builder.centralDirectory.forEach((dir, i) => {	
		if (!$tbody) return;
		const $tr = $tbody.insertRow();
		const $operation = $tr.insertCell();
		const $rename = document.createElement('button');
		$rename.classList.add('rename');
		$rename.addEventListener('click', onClickRename);
		const $delete = document.createElement('button');
		$delete.classList.add('delete');
		$delete.addEventListener('click', onClickDelete);
		$operation.append($rename, $delete);

		const $filepath = $tr.insertCell();
		$filepath.append(ZipFile.Builder.textDecoder.decode(dir.fileNameBytes));
	
		const $uncompSize = $tr.insertCell();
		$uncompSize.classList.add('r');
		$uncompSize.dataset.suffix = 'B';
		$uncompSize.append(dir.uncompressedSize.toLocaleString());
	
		const $compSize = $tr.insertCell();
		$compSize.classList.add('r');
		$compSize.dataset.suffix = 'B';
		$compSize.append(dir.compressedSize.toLocaleString());
	
		const $lastModified = $tr.insertCell();
		const $time = document.createElement('time');
		$time.append(dir.lastModified.toDate().toISOString().substring(0, 19));
		$lastModified.append($time);
	
		const $method = $tr.insertCell();
		$method.classList.add('c');
		$method.append(ZipFile.compressionMethods[dir.method]);
	
		const $version = $tr.insertCell();
		$version.classList.add('c');
		$version.append(dir.versionNeeded.toString());
	
		const $crc32 = $tr.insertCell();
		$crc32.classList.add('c');
		if (dir.crc32) $crc32.append(ZipFile.fmtCrc32(dir.crc32));
	});
}

function onClickDownload() {
	const blob = builder.build();
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.download = '';
	a.href = url;
	a.click();
	URL.revokeObjectURL(url);
}

/**
 * @this {HTMLButtonElement}
 */
function onClickRename() {
	const $tr = this.closest('tr');
	if (!$tr) return;
	const $tbody = /** @type {HTMLTableSectionElement} */ ($tr.parentElement);
	const index = Array.from($tbody.rows).indexOf($tr);
	if (index < 0) return;
	const cd = builder.centralDirectory.at(index);
	const content = builder.contents.at(index);
	if (!cd || !content) return;
	const newName = prompt('Enter new name', ZipFile.Builder.textDecoder.decode(cd.fileNameBytes));
	if (!newName) return;
	const bytes = ZipFile.Builder.textEncoder.encode(newName);
	cd.fileNameBytes = content.header.fileNameBytes = bytes;
	$tr.children[1]?.firstChild?.remove();
	$tr.children[1]?.append(newName);
}

/**
 * @this {HTMLButtonElement}
 */
function onClickDelete() {
	const $tr = this.closest('tr');
	if (!$tr) return;
	const $tbody = /** @type {HTMLTableSectionElement} */ ($tr.parentElement);
	const index = Array.from($tbody.rows).indexOf($tr);
	if (index < 0) return;
	builder.remove(index);
	$tbody.removeChild($tr);
}