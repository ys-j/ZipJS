ZipJS
===

ES Module for compressing/decompressing ZIP files\
ZIPファイルを圧縮/解凍するES Module

## How to use

```js
import * as ZipFile from './zip.js'; // your path

// Decompression: Picks uncompressed contents
const zf = new ZipFile.Extractor(arrayBuffer);
for (let i = 0; i < zf.contents.length; i++) {
    const response = zf.pick(i);
    const content = await response.blob();
}

// Compression
const builder = new ZipFile.Builder();
await builder.append(arrayBuffer, filename);
const blob = builder.build();
```

## Sample

- [GitHub Pages](https://ys-j.github.io/ZipJS/)
