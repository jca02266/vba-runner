#!/usr/bin/env node
/**
 * import-vba.mjs
 * UTF-8 の .bas / .cls ファイルを読んで xlsm の VBA ソースを書き戻す。
 * コードページは xlsm の dir ストリーム（PROJECTCODEPAGE）から自動検出し、
 * UTF-8 から変換して格納する。
 *
 * Usage:
 *   node import-vba.mjs <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>]
 *
 * Options:
 *   --encoding <cp>  コードページを明示指定（例: cp932, cp1252）。
 *                    省略時は xlsm 内の PROJECTCODEPAGE を使用。
 *
 * output.xlsm のデフォルトは input.xlsm（上書き）。
 */

import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import { readFileSync, writeFileSync, readdirSync } from 'fs';
import { resolve, extname, basename } from 'path';
import { decompress, compress } from './lib/ovba.mjs';
import { parseDirStream } from './lib/dir-parser.mjs';

// ── 引数パース ──────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length < 2 || args[0] === '--help') {
    console.error('Usage: node import-vba.mjs <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>]');
    process.exit(1);
}

let xlsmPath, srcDir, outPathArg, encodingOverride;
for (let i = 0; i < args.length; i++) {
    if (args[i] === '--encoding') {
        encodingOverride = args[++i];
    } else if (!xlsmPath) {
        xlsmPath = args[i];
    } else if (!srcDir) {
        srcDir = args[i];
    } else {
        outPathArg = args[i];
    }
}

const absXlsm = resolve(xlsmPath);
const absSrc  = resolve(srcDir);
const outPath = resolve(outPathArg ?? absXlsm);

// ── ソースファイルを収集（UTF-8 として読み込む）──────────────────────────────
const sourceMap = new Map();
for (const f of readdirSync(absSrc)) {
    const ext = extname(f).toLowerCase();
    if (ext !== '.bas' && ext !== '.cls') continue;
    const modName = basename(f, ext);
    sourceMap.set(modName.toLowerCase(), readFileSync(`${absSrc}/${f}`, 'utf8'));
}
console.log(`ソースファイル: ${sourceMap.size} 件`);

// ── xlsm を ZIP として開く ──────────────────────────────────────────────────
const zip = await JSZip.loadAsync(readFileSync(absXlsm));

const vbaBinEntry = zip.file('xl/vbaProject.bin');
if (!vbaBinEntry) {
    console.error('xl/vbaProject.bin が見つかりません');
    process.exit(1);
}
const vbaBuf = Buffer.from(await vbaBinEntry.async('nodebuffer'));

// ── CFB を解析 ─────────────────────────────────────────────────────────────
const cfb = CFB.read(vbaBuf, { type: 'buffer' });

// ── dir ストリームからモジュール一覧・コードページを取得 ──────────────────────
const dirEntry = CFB.find(cfb, '/VBA/dir');
const dirDecompressed = decompress(Buffer.from(dirEntry.content));
const { codePage, modules } = parseDirStream(dirDecompressed);

// エンコーディングの決定: コマンドライン引数 > dir ストリームの PROJECTCODEPAGE
if (!encodingOverride && codePage === null) {
    console.warn('警告: PROJECTCODEPAGE が見つかりませんでした。--encoding で指定してください。');
    process.exit(1);
}
const encoding = encodingOverride ?? `cp${codePage}`;
console.log(`VBAモジュール: ${modules.length} 件`);
console.log(`エンコーディング: ${encoding}${codePage !== null ? ` (PROJECTCODEPAGE=${codePage})` : ' (--encoding 指定)'}`);

// ── 各モジュールのストリームを更新 ────────────────────────────────────────
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

    // UTF-8 → コードページ変換してから OVBA 圧縮
    const srcBytes  = iconv.encode(src, encoding);
    const newSource = compress(srcBytes);

    // offset 以前のバイトコード部分は保持し、ソース部分だけ置き換える
    const newStream = Buffer.concat([raw.slice(0, mod.offset), newSource]);

    entry.content = newStream;
    entry.size    = newStream.length;

    console.log(`  ✓ ${mod.name} (${srcBytes.length} bytes as ${encoding})`);
    updated++;
}

if (updated === 0) {
    console.error('更新するモジュールがありませんでした。');
    process.exit(1);
}

// ── CFB を再シリアライズして xlsm を保存 ────────────────────────────────────
const newVbaBin = CFB.write(cfb, { type: 'buffer' });

zip.file('xl/vbaProject.bin', newVbaBin);
const newXlsm = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
});

writeFileSync(outPath, newXlsm);
console.log(`保存: ${outPath}`);
