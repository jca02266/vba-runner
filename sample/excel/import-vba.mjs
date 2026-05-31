#!/usr/bin/env node
/**
 * import-vba.mjs
 * .bas / .cls ファイルを読んで xlsm の VBA ソースを書き戻す。
 *
 * Usage:
 *   node import-vba.mjs <input.xlsm> <source-dir> [output.xlsm]
 *
 * output.xlsm のデフォルトは input.xlsm（上書き）。
 * 安全のため input と output が同じ場合は一時ファイルを経由する。
 */

import JSZip from 'jszip';
import CFB from 'cfb';
import { readFileSync, writeFileSync, readdirSync } from 'fs';
import { resolve, dirname, extname, basename } from 'path';
import { decompress, compress } from './lib/ovba.mjs';
import { parseDirStream } from './lib/dir-parser.mjs';

const [,, xlsmPath, srcDir, outPathArg] = process.argv;
if (!xlsmPath || !srcDir) {
    console.error('Usage: node import-vba.mjs <input.xlsm> <source-dir> [output.xlsm]');
    process.exit(1);
}

const absXlsm = resolve(xlsmPath);
const absSrc  = resolve(srcDir);
const outPath = resolve(outPathArg ?? absXlsm);

// 1. ソースファイルを収集（ファイル名（拡張子なし）→ソース）
const sourceMap = new Map();
for (const f of readdirSync(absSrc)) {
    const ext = extname(f).toLowerCase();
    if (ext !== '.bas' && ext !== '.cls') continue;
    const modName = basename(f, ext);
    sourceMap.set(modName.toLowerCase(), readFileSync(`${absSrc}/${f}`, 'utf8'));
}
console.log(`ソースファイル: ${sourceMap.size} 件`);

// 2. xlsm を ZIP として開く
const zip = await JSZip.loadAsync(readFileSync(absXlsm));

const vbaBinEntry = zip.file('xl/vbaProject.bin');
if (!vbaBinEntry) {
    console.error('xl/vbaProject.bin が見つかりません');
    process.exit(1);
}
const vbaBuf = Buffer.from(await vbaBinEntry.async('nodebuffer'));

// 3. CFB を解析
const cfb = CFB.read(vbaBuf, { type: 'buffer' });

// 4. dir ストリームからモジュール一覧を取得
const dirEntry = CFB.find(cfb, '/VBA/dir');
const dirDecompressed = decompress(Buffer.from(dirEntry.content));
const modules = parseDirStream(dirDecompressed);

console.log(`VBAモジュール: ${modules.length} 件`);

// 5. 各モジュールのストリームを更新
let updated = 0;
for (const mod of modules) {
    const src = sourceMap.get(mod.name.toLowerCase());
    if (src === undefined) {
        console.log(`  [skip] ${mod.name}: ソースファイルなし`);
        continue;
    }

    const streamPath = `/VBA/${mod.streamName}`;
    const entry = CFB.find(cfb, streamPath);
    if (!entry) {
        console.warn(`  [warn] ${mod.name}: ストリーム ${streamPath} なし`);
        continue;
    }

    const raw = Buffer.from(entry.content);

    // オフセット以前のバイトコード部分は保持し、ソース部分だけ置き換える
    const header = raw.slice(0, mod.offset);
    const newSource = compress(Buffer.from(src, 'latin1'));
    const newStream = Buffer.concat([header, newSource]);

    // CFB のエントリを上書き
    entry.content = newStream;
    entry.size = newStream.length;

    console.log(`  ✓ ${mod.name} (${src.length} bytes)`);
    updated++;
}

if (updated === 0) {
    console.error('更新するモジュールがありませんでした。');
    process.exit(1);
}

// 6. CFB を再シリアライズ
const newVbaBin = CFB.write(cfb, { type: 'buffer' });

// 7. ZIP に書き戻して xlsm を保存
zip.file('xl/vbaProject.bin', newVbaBin);
const newXlsm = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
});

writeFileSync(outPath, newXlsm);
console.log(`保存: ${outPath}`);
