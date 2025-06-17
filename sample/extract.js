import * as ZipFile from '../zip.js';

const acceptMime = [
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
	'application/java+archive', // jar
	'application/vnd.android.package-archive', // apk
];

const compressionMethods = {
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

/** @type {HTMLFormElement} */
const $form = document.forms['f'];
$form.addEventListener('submit', onSubmit);
$form.elements['file'].accept = acceptMime.join();

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
	const zip = new ZipFile.Extractor(await file.arrayBuffer());
	document.body.style.cursor = 'auto';

	if (!$tbody) return;
	$tbody.innerHTML = '';
	
	for (const content of zip) {
		const cd = content.entry;
		let fileName;
		if (cd.isUtf8) {
			fileName = utf8Decoder.decode(cd.fileNameBytes);
		} else {
			const charsets = ['utf-8', 'shift_jis', 'euc-jp', 'iso-2022-jp'];
			for (const charset of charsets) {
				try {
					const decoder = new TextDecoder(charset, { fatal: true });
					fileName = decoder.decode(cd.fileNameBytes);
					break;
				} catch {
					// continue;
				}
			}
			if (!fileName) {
				console.error('Failed to detect charset of file name.');
				fileName = utf8Decoder.decode(cd.fileNameBytes);
			}
		}
		const filePaths = fileName.split('/');
	
		if (!$tbody) return;
		const $tr = $tbody.insertRow();
		const $filepath = $tr.insertCell();
		if (cd.crc32) {
			const $a = document.createElement('a');
			$a.text = filePaths.pop() || '';
			$a.href = '#';
			$a.onclick = async e => {
				if ($a.getAttribute('href') === '#') {
					e.preventDefault();
					try {
						if (cd.isEncrypted) throw new Error('Encypted file is not supported.');
						if (cd.versionNeeded > 20) throw new Error('Version greater than 20 is not supported: ' + cd.versionNeeded);
						const blob = await content.body.blob();
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
		$uncompSize.append(cd.uncompressedSize.toLocaleString());
	
		const $compSize = $tr.insertCell();
		$compSize.classList.add('r');
		$compSize.dataset.suffix = 'B';
		$compSize.append(cd.compressedSize.toLocaleString());
	
		const $lastModified = $tr.insertCell();
		const $time = document.createElement('time');
		$time.append(cd.lastModified.toDate().toISOString().substring(0, 19));
		$lastModified.append($time);
	
		const $method = $tr.insertCell();
		$method.classList.add('c');
		$method.append(compressionMethods[cd.method]);
	
		const $version = $tr.insertCell();
		$version.classList.add('c');
		$version.append(cd.versionNeeded.toString());
	
		const $crc32 = $tr.insertCell();
		$crc32.classList.add('c', 'tt');
		if (cd.crc32) $crc32.append(ZipFile.fmtCrc32(cd.crc32));
	}
}