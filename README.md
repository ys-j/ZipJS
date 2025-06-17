ZipJS
===

ES Module for compressing/decompressing ZIP files\
ZIPファイルを圧縮/解凍するES Module

Examples
---

```js
import * as ZipFile from './zip.js'; // your path

// Decompression: Picks uncompressed contents
const zf = new ZipFile.Extractor(arrayBuffer);
for (const content of zf) {
    const blob = await content.body.blob();
}

// Compression
const builder = new ZipFile.Builder();
await builder.append(arrayBuffer, filename);
const blob = builder.build();
```

Demo
---

- [GitHub Pages](https://ys-j.github.io/ZipJS/)
