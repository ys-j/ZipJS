import * as ZipFile from '../zip.js';

/** @type {HTMLFormElement} */
const $form = document.forms['f'];
$form.addEventListener('submit', onSubmit);
$form.elements['file'].accept = ZipFile.acceptMime.join();

/** @type {HTMLTableSectionElement?} */ //@ts-ignore
const $tbody = document.getElementById('tbody');
const utf8Decoder = new TextDecoder();

/**
 * @param {SubmitEvent} e 
 */
async function onSubmit(e) {
	e.preventDefault();
	const file = /** @type {HTMLInputElement} */ ($form.file)?.files?.[0];
	if (!file) return;

	document.body.style.cursor = 'wait';
	const zf = new ZipFile.Extractor(await file.arrayBuffer());
	document.body.style.cursor = 'auto';

	if (!$tbody) return;
	$tbody.innerHTML = '';
	zf.centralDirectory.forEach((dir, i) => {
		let fileName;
		if (dir.isUtf8) {
			fileName = utf8Decoder.decode(dir.fileNameBytes);
		} else {
			const charsets = ['utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp'];
			for (const charset of charsets) {
				try {
					const decoder = new TextDecoder(charset, { fatal: true });
					fileName = decoder.decode(dir.fileNameBytes);
					break;
				} catch {
					// continue;
				}
			}
			if (!fileName) {
				console.error('Failed to detect charset of file name.');
				fileName = utf8Decoder.decode(dir.fileNameBytes);
			}
		}
		const filePaths = fileName.split('/');
	
		if (!$tbody) return;
		const $tr = $tbody.insertRow();
		const $filepath = $tr.insertCell();
		if (dir.crc32) {
			const $a = document.createElement('a');
			$a.text = filePaths.pop() || '';
			$a.href = '#';
			$a.onclick = async e => {
				if ($a.getAttribute('href') === '#') {
					e.preventDefault();
					try {
						const response = zf.pick(i);
						const blob = await response.blob();
						$a.href = URL.createObjectURL(blob);
						$a.download = $a.text;
						$a.click();
					} catch (e) {
						console.error(e);
						alert(e);
					}
				}
			};
			const path = filePaths.join('/');
			if (path) $filepath.append(path + '/');
			$filepath.append($a);
		} else {
			$filepath.append(filePaths.join('/'));
		}
	
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