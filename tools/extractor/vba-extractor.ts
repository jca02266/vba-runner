import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from 'fs';
import { resolve, dirname, basename, extname } from 'path';
import { decompress, compress } from './ovba.js';
import { parseDirStream } from './dir-parser.js';

function printUsage(): void {
    console.error('Usage:');
    console.error('  vba-extractor export <input.xlsm> [output-dir] [--encoding <cp>]');
    console.error('  vba-extractor import <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>]');
}

function parseEncoding(args: string[]): { encoding: string | undefined; rest: string[] } {
    const rest: string[] = [];
    let encoding: string | undefined;
    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--encoding') {
            encoding = args[++i];
        } else {
            rest.push(args[i]);
        }
    }
    return { encoding, rest };
}

async function openXlsm(xlsmPath: string) {
    const zip = await JSZip.loadAsync(readFileSync(xlsmPath));
    const entry = zip.file('xl/vbaProject.bin');
    if (!entry) {
        console.error('xl/vbaProject.bin が見つかりません（マクロなし xlsm？）');
        process.exit(1);
    }
    const vbaBuf = Buffer.from(await entry.async('nodebuffer'));
    const cfb = CFB.read(vbaBuf, { type: 'buffer' });
    const dirEntry = CFB.find(cfb, '/VBA/dir');
    if (!dirEntry) throw new Error('/VBA/dir が見つかりません');
    const dirDecompressed = decompress(Buffer.from(dirEntry.content as unknown as ArrayBuffer));
    const { codePage, modules } = parseDirStream(dirDecompressed);
    return { zip, cfb, codePage, modules };
}

function resolveEncoding(encodingOverride: string | undefined, codePage: number | null): string {
    if (!encodingOverride && codePage === null) {
        console.error('警告: PROJECTCODEPAGE が見つかりませんでした。--encoding で指定してください。');
        process.exit(1);
    }
    return encodingOverride ?? `cp${codePage}`;
}

async function runExport(args: string[]): Promise<void> {
    const { encoding: encodingOverride, rest } = parseEncoding(args);
    const [xlsmArg, outDirArg] = rest;
    if (!xlsmArg) { printUsage(); process.exit(1); }

    const absXlsm = resolve(xlsmArg);
    const outDir  = resolve(outDirArg ?? dirname(absXlsm));
    if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

    const { cfb, codePage, modules } = await openXlsm(absXlsm);
    const encoding = resolveEncoding(encodingOverride, codePage);
    console.log(`モジュール数: ${modules.length}`);
    console.log(`エンコーディング: ${encoding}`);

    for (const mod of modules) {
        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) { console.warn(`  [skip] ${mod.name}: ストリームなし`); continue; }

        const raw        = Buffer.from(entry.content as unknown as ArrayBuffer);
        const compressed = raw.subarray(mod.offset);
        const source     = iconv.decode(decompress(compressed), encoding);

        const ext     = mod.isClass ? '.cls' : '.bas';
        const outPath = `${outDir}/${mod.name}${ext}`;
        writeFileSync(outPath, source, 'utf8');
        console.log(`  → ${basename(outPath)} (${source.length} chars)`);
    }
    console.log('完了');
}

async function runImport(args: string[]): Promise<void> {
    const { encoding: encodingOverride, rest } = parseEncoding(args);
    const [xlsmArg, srcDirArg, outPathArg] = rest;
    if (!xlsmArg || !srcDirArg) { printUsage(); process.exit(1); }

    const absXlsm = resolve(xlsmArg);
    const absSrc  = resolve(srcDirArg);
    const outPath = resolve(outPathArg ?? absXlsm);

    const sourceMap = new Map<string, string>();
    for (const f of readdirSync(absSrc)) {
        const ext = extname(f).toLowerCase();
        if (ext !== '.bas' && ext !== '.cls') continue;
        sourceMap.set(basename(f, ext).toLowerCase(), readFileSync(`${absSrc}/${f}`, 'utf8'));
    }
    console.log(`ソースファイル: ${sourceMap.size} 件`);

    const { zip, cfb, codePage, modules } = await openXlsm(absXlsm);
    const encoding = resolveEncoding(encodingOverride, codePage);
    console.log(`VBAモジュール: ${modules.length} 件`);
    console.log(`エンコーディング: ${encoding}`);

    let updated = 0;
    for (const mod of modules) {
        const src = sourceMap.get(mod.name.toLowerCase());
        if (src === undefined) { console.log(`  [skip] ${mod.name}: ソースファイルなし`); continue; }

        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) { console.warn(`  [warn] ${mod.name}: ストリームなし`); continue; }

        const raw       = Buffer.from(entry.content as unknown as ArrayBuffer);
        const srcBytes  = iconv.encode(src, encoding);
        const newStream = Buffer.concat([raw.subarray(0, mod.offset), compress(srcBytes)]);

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (entry as any).content = newStream;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (entry as any).size    = newStream.length;
        console.log(`  ✓ ${mod.name} (${srcBytes.length} bytes as ${encoding})`);
        updated++;
    }

    if (updated === 0) { console.error('更新するモジュールがありませんでした。'); process.exit(1); }

    const newVbaBin = CFB.write(cfb, { type: 'buffer' });
    zip.file('xl/vbaProject.bin', newVbaBin as unknown as Buffer, { createFolders: false });
    const newXlsm = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    writeFileSync(outPath, newXlsm);
    console.log(`保存: ${outPath}`);
}

async function main(): Promise<void> {
    const [,, mode, ...rest] = process.argv;

    if (mode !== 'export' && mode !== 'import') {
        printUsage();
        process.exit(1);
    }

    if (mode === 'export') {
        await runExport(rest);
    } else {
        await runImport(rest);
    }
}

main().catch(err => { console.error(err); process.exit(1); });
