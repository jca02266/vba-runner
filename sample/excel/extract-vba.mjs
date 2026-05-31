#!/usr/bin/env node
/**
 * extract-vba.mjs
 * xlsm ファイルから VBA ソースを抽出して .bas / .cls ファイルに書き出す。
 * コードページは xlsm の dir ストリーム（PROJECTCODEPAGE）から自動検出し、
 * UTF-8 に変換して出力する。
 *
 * Usage:
 *   node extract-vba.mjs <input.xlsm> [output-dir] [--encoding <cp>]
 *
 * Options:
 *   --encoding <cp>  コードページを明示指定（例: cp932, cp1252）。
 *                    省略時は xlsm 内の PROJECTCODEPAGE を使用。
 *
 * output-dir のデフォルトは input.xlsm と同じディレクトリ。
 */

import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';
import { resolve, dirname, basename } from 'path';
import { decompress } from './lib/ovba.mjs';
import { parseDirStream } from './lib/dir-parser.mjs';

// ── 引数パース ──────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length === 0 || args[0] === '--help') {
    console.error('Usage: node extract-vba.mjs <input.xlsm> [output-dir] [--encoding <cp>]');
    process.exit(1);
}

let xlsmPath, outDirArg, encodingOverride;
for (let i = 0; i < args.length; i++) {
    if (args[i] === '--encoding') {
        encodingOverride = args[++i];
    } else if (!xlsmPath) {
        xlsmPath = args[i];
    } else {
        outDirArg = args[i];
    }
}

const absXlsm = resolve(xlsmPath);
const outDir  = resolve(outDirArg ?? dirname(absXlsm));

if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

// ── xlsm を ZIP として開く ──────────────────────────────────────────────────
const zip = await JSZip.loadAsync(readFileSync(absXlsm));

const vbaBinEntry = zip.file('xl/vbaProject.bin');
if (!vbaBinEntry) {
    console.error('xl/vbaProject.bin が見つかりません（マクロなし xlsm？）');
    process.exit(1);
}

const vbaBuf = Buffer.from(await vbaBinEntry.async('nodebuffer'));

// ── CFB（Compound File Binary）として解析 ──────────────────────────────────
const cfb = CFB.read(vbaBuf, { type: 'buffer' });

// ── dir ストリームを伸長してモジュール一覧・コードページを取得 ────────────────
const dirEntry = CFB.find(cfb, '/VBA/dir');
if (!dirEntry) throw new Error('/VBA/dir が見つかりません');

const dirDecompressed = decompress(Buffer.from(dirEntry.content));
const { codePage, modules } = parseDirStream(dirDecompressed);

// エンコーディングの決定: コマンドライン引数 > dir ストリームの PROJECTCODEPAGE
if (!encodingOverride && codePage === null) {
    console.warn('警告: PROJECTCODEPAGE が見つかりませんでした。--encoding で指定してください。');
    process.exit(1);
}
const encoding = encodingOverride ?? `cp${codePage}`;
console.log(`モジュール数: ${modules.length}`);
console.log(`エンコーディング: ${encoding}${codePage !== null ? ` (PROJECTCODEPAGE=${codePage})` : ' (--encoding 指定)'}`);

// ── 各モジュールのソースを抽出 ─────────────────────────────────────────────
for (const mod of modules) {
    const streamPath = `/VBA/${mod.streamName}`;
    const entry = CFB.find(cfb, streamPath);
    if (!entry) {
        console.warn(`  [skip] ${mod.name}: ストリーム ${streamPath} なし`);
        continue;
    }

    const raw        = Buffer.from(entry.content);
    const compressed = raw.slice(mod.offset);
    const srcBytes   = decompress(compressed);

    // コードページ → UTF-8 変換
    const source = iconv.decode(srcBytes, encoding);

    const ext     = mod.isClass ? '.cls' : '.bas';
    const outPath = `${outDir}/${mod.name}${ext}`;
    writeFileSync(outPath, source, 'utf8');
    console.log(`  → ${basename(outPath)} (${source.length} chars)`);
}

console.log('完了');
