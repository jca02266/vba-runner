#!/usr/bin/env node
/**
 * extract-vba.mjs
 * xlsm ファイルから VBA ソースを抽出して .bas / .cls ファイルに書き出す。
 *
 * Usage:
 *   node extract-vba.mjs <input.xlsm> [output-dir]
 *
 * output-dir のデフォルトは input.xlsm と同じディレクトリ。
 */

import JSZip from 'jszip';
import CFB from 'cfb';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';
import { resolve, dirname, basename } from 'path';
import { decompress } from './lib/ovba.mjs';
import { parseDirStream } from './lib/dir-parser.mjs';

const [,, xlsmPath, outDirArg] = process.argv;
if (!xlsmPath) {
    console.error('Usage: node extract-vba.mjs <input.xlsm> [output-dir]');
    process.exit(1);
}

const absXlsm = resolve(xlsmPath);
const outDir = resolve(outDirArg ?? dirname(absXlsm));

if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

// 1. xlsm を ZIP として開く
const zip = await JSZip.loadAsync(readFileSync(absXlsm));

const vbaBinEntry = zip.file('xl/vbaProject.bin');
if (!vbaBinEntry) {
    console.error('xl/vbaProject.bin が見つかりません（マクロなし xlsm？）');
    process.exit(1);
}

const vbaBuf = Buffer.from(await vbaBinEntry.async('nodebuffer'));

// 2. CFB（Compound File Binary）として解析
const cfb = CFB.read(vbaBuf, { type: 'buffer' });

// 3. dir ストリームを展開してモジュール一覧を取得
const dirEntry = CFB.find(cfb, '/VBA/dir');
if (!dirEntry) throw new Error('/VBA/dir が見つかりません');

const dirDecompressed = decompress(Buffer.from(dirEntry.content));
const modules = parseDirStream(dirDecompressed);

console.log(`モジュール数: ${modules.length}`);

// 4. 各モジュールのソースを抽出
for (const mod of modules) {
    const streamPath = `/VBA/${mod.streamName}`;
    const entry = CFB.find(cfb, streamPath);
    if (!entry) {
        console.warn(`  [skip] ${mod.name}: ストリーム ${streamPath} なし`);
        continue;
    }

    const raw = Buffer.from(entry.content);
    // ソースは offset バイト目以降（それ以前はコンパイル済みバイトコード）
    const compressed = raw.slice(mod.offset);
    const source = decompress(compressed).toString('latin1');

    const ext = mod.isClass ? '.cls' : '.bas';
    const outPath = `${outDir}/${mod.name}${ext}`;
    writeFileSync(outPath, source, 'utf8');
    console.log(`  → ${basename(outPath)} (${source.length} bytes)`);
}

console.log('完了');
