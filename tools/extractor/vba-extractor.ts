import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import { readFileSync, writeFileSync, copyFileSync, mkdirSync, existsSync, readdirSync } from 'fs';
import { createInterface } from 'readline';
import { resolve, dirname, basename, extname, join } from 'path';
import { decompress, compress } from './ovba.js';
import { parseDirStream, parseDirStreamFull, VbaModuleFull } from './dir-parser.js';
import { buildDirStream } from './dir-builder.js';

// tsx from tools/extractor/: ../../build/extractor/package.json
// built CJS in build/extractor/dist/bin/: ../../package.json
const VERSION = (() => {
    for (const rel of ['../../build/extractor/package.json', '../../package.json']) {
        try {
            return JSON.parse(readFileSync(join(import.meta.dirname, rel), 'utf8')).version as string;
        } catch { /* try next */ }
    }
    return 'unknown';
})();

/**
 * Remove deleted modules from the PROJECT text stream and add new module lines.
 * Handles Module=, Class=, Package=+BaseClass= (UserForm), and [Workspace] entries.
 */
function patchProjectStream(
    text: string,
    deletedLower: Set<string>,
    added: { name: string; isClass: boolean }[],
): string {
    const lines = text.split(/\r\n|\r|\n/);
    const out: string[] = [];
    let inWorkspace = false;
    let pendingPackage: string | null = null;
    let insertIdx = -1; // position in `out` after last module declaration

    for (const line of lines) {
        const t = line.trim();

        if (!inWorkspace && t.toLowerCase() === '[workspace]') inWorkspace = true;

        if (inWorkspace) {
            const eq = t.indexOf('=');
            if (eq > 0 && deletedLower.has(t.slice(0, eq).toLowerCase())) continue;
            out.push(line);
            continue;
        }

        // Buffer Package= line; drop it if the following BaseClass= is for a deleted module
        if (t.startsWith('Package=')) { pendingPackage = line; continue; }

        if (t.startsWith('BaseClass=')) {
            const name = t.slice('BaseClass='.length).trim();
            if (deletedLower.has(name.toLowerCase())) { pendingPackage = null; continue; }
            if (pendingPackage !== null) { out.push(pendingPackage); pendingPackage = null; }
            out.push(line);
            // Do NOT update insertIdx here — new modules must go before Package=+BaseClass= block
            continue;
        }

        if (pendingPackage !== null) { out.push(pendingPackage); pendingPackage = null; }

        if (t.startsWith('Module=') || t.startsWith('Class=')) {
            const name = t.slice(t.indexOf('=') + 1).trim();
            if (deletedLower.has(name.toLowerCase())) continue;
            out.push(line);
            insertIdx = out.length;
            continue;
        }

        out.push(line);
    }

    if (added.length > 0) {
        const newLines = added.map(m => (m.isClass ? 'Class=' : 'Module=') + m.name);
        out.splice(insertIdx >= 0 ? insertIdx : 1, 0, ...newLines);
        // Add [Workspace] entries for new modules
        const wsIdx = out.findIndex(l => l.trim().toLowerCase() === '[workspace]');
        if (wsIdx >= 0) {
            const wsEntries = added.map(m => `${m.name}=0, 0, 2000, 1000, C`);
            out.splice(wsIdx + 1, 0, ...wsEntries);
        }
    }

    return out.join('\r\n');
}

/**
 * Remove deleted modules from the PROJECTwm binary stream and append new entries.
 * Format: (ASCII-name NUL UTF16LE-name NUL NUL)* NUL NUL
 */
function patchProjectWm(data: Buffer, deletedLower: Set<string>, addedNames: string[]): Buffer {
    const parts: Buffer[] = [];
    let i = 0;
    while (i < data.length) {
        if (data[i] === 0) break; // stream terminator
        let e = i;
        while (e < data.length && data[e] !== 0) e++;
        const ascii = data.subarray(i, e).toString('latin1');
        i = e + 1;
        let u = i;
        while (u + 1 < data.length && (data[u] !== 0 || data[u + 1] !== 0)) u += 2;
        const unicodeEnd = u + 2;
        i = unicodeEnd;
        if (!deletedLower.has(ascii.toLowerCase())) {
            parts.push(Buffer.from(ascii + '\0', 'latin1'));
            parts.push(data.subarray(e + 1, unicodeEnd));
        }
    }
    for (const name of addedNames) {
        parts.push(Buffer.from(name + '\0', 'latin1'));
        parts.push(Buffer.from(name, 'utf16le'));
        parts.push(Buffer.from([0, 0]));
    }
    parts.push(Buffer.from([0, 0])); // stream terminator
    return Buffer.concat(parts);
}

async function promptYesNo(question: string): Promise<boolean> {
    const rl = createInterface({ input: process.stdin, output: process.stdout });
    return new Promise(resolve => {
        rl.question(question, answer => {
            rl.close();
            resolve(answer.trim().toLowerCase() === 'y');
        });
    });
}

function printUsage(): void {
    console.log('Usage:');
    console.log('  vba-extractor export <input.xlsm> [output-dir] [--encoding <cp>]');
    console.log('    output-dir defaults to: <input.xlsm directory>/src');
    console.log('  vba-extractor import <input.xlsm> [source-dir] [output.xlsm] [--encoding <cp>]');
    console.log('    source-dir defaults to: <input.xlsm directory>/src');
    console.log('');
    console.log('Options:');
    console.log('  --version    Show version');
    console.log('  --help       Show this help');
}

function parseEncoding(args: string[]): { encoding: string | undefined; rest: string[] } {
    const rest: string[] = [];
    let encoding: string | undefined;
    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--encoding') {
            encoding = args[++i];
        } else if (args[i].startsWith('-') && args[i] !== '-') {
            // Any other token that looks like a flag is an unknown option.
            // (Positional paths are not expected to start with '-'.)
            console.error(`vba-extractor: unknown option '${args[i]}'`);
            printUsage();
            process.exit(1);
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
        console.error('xl/vbaProject.bin not found (macro-free xlsm?)');
        process.exit(1);
    }
    const vbaBuf = Buffer.from(await entry.async('nodebuffer'));
    const cfb = CFB.read(vbaBuf, { type: 'buffer' });
    const dirEntry = CFB.find(cfb, '/VBA/dir');
    if (!dirEntry) throw new Error('/VBA/dir not found');
    const dirDecompressed = decompress(Buffer.from(dirEntry.content as unknown as ArrayBuffer));
    const { codePage, modules } = parseDirStream(dirDecompressed);
    return { zip, cfb, codePage, modules, dirEntry, dirDecompressed };
}

function resolveEncoding(encodingOverride: string | undefined, codePage: number | null): string {
    if (!encodingOverride && codePage === null) {
        console.error('Error: PROJECTCODEPAGE not found. Specify encoding with --encoding.');
        process.exit(1);
    }
    return encodingOverride ?? `cp${codePage}`;
}

async function runExport(args: string[]): Promise<void> {
    const { encoding: encodingOverride, rest } = parseEncoding(args);
    const [xlsmArg, outDirArg] = rest;
    if (!xlsmArg) { printUsage(); process.exit(1); }

    const absXlsm = resolve(xlsmArg);
    const outDir  = resolve(outDirArg ?? join(dirname(absXlsm), 'src'));
    if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });

    const { cfb, codePage, modules } = await openXlsm(absXlsm);
    const encoding = resolveEncoding(encodingOverride, codePage);
    console.log(`Modules  : ${modules.length}`);
    console.log(`Encoding : ${encoding}`);

    for (const mod of modules) {
        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) { console.warn(`  [skip] ${mod.name}: stream not found`); continue; }

        const raw        = Buffer.from(entry.content as unknown as ArrayBuffer);
        const compressed = raw.subarray(mod.offset);
        const source     = iconv.decode(decompress(compressed), encoding);

        const ext     = mod.isClass ? '.cls' : '.bas';
        const outPath = `${outDir}/${mod.name}${ext}`;
        writeFileSync(outPath, source, 'utf8');
        console.log(`  → ${basename(outPath)} (${source.length} chars)`);
    }
    console.log('Done.');
}

async function runImport(args: string[]): Promise<void> {
    const { encoding: encodingOverride, rest } = parseEncoding(args);
    const [xlsmArg, srcDirArg, outPathArg] = rest;
    if (!xlsmArg) { printUsage(); process.exit(1); }

    const absXlsm = resolve(xlsmArg);
    const absSrc  = resolve(srcDirArg ?? join(dirname(absXlsm), 'src'));
    const outPath = resolve(outPathArg ?? absXlsm);
    const backupPath = absXlsm + '.bak';
    const willOverwrite = outPath === absXlsm;

    console.log('');
    console.log('⚠️  Warning: import directly modifies the Excel file.');
    console.log('   The file may become corrupted. It is strongly recommended to back up before proceeding.');
    console.log(`   Input  : ${absXlsm}`);
    console.log(`   Output : ${outPath}`);
    if (willOverwrite) {
        console.log(`   Backup : ${backupPath}`);
    }
    console.log('');

    const question = willOverwrite
        ? 'Create a backup and run import? [y/N]: '
        : 'Run import? [y/N]: ';
    const confirmed = await promptYesNo(question);

    if (!confirmed) {
        console.log('Cancelled.');
        process.exit(0);
    }

    if (willOverwrite) {
        copyFileSync(absXlsm, backupPath);
        console.log(`Backup created: ${backupPath}`);
    }

    // Build a map of source files: lowercase name → source text
    const sourceFileNames = new Map<string, string>(); // lowercase name → original filename (no ext)
    const sourceMap = new Map<string, string>();        // lowercase name → source text
    for (const f of readdirSync(absSrc)) {
        const ext = extname(f).toLowerCase();
        if (ext !== '.bas' && ext !== '.cls') continue;
        const baseName = basename(f, ext);
        sourceMap.set(baseName.toLowerCase(), readFileSync(`${absSrc}/${f}`, 'utf8'));
        sourceFileNames.set(baseName.toLowerCase(), baseName);
    }
    console.log(`Source files : ${sourceMap.size}`);

    const { zip, cfb, codePage, dirEntry, dirDecompressed } = await openXlsm(absXlsm);
    const encoding = resolveEncoding(encodingOverride, codePage);

    // Parse the dir stream fully so we can rebuild it
    const dirData = parseDirStreamFull(dirDecompressed);
    console.log(`VBA modules  : ${dirData.modules.length}`);
    console.log(`Encoding     : ${encoding}`);

    // Determine which module names from dir appear in sourceMap (update),
    // which are missing from sourceMap (delete), and which source files are new (add).
    const existingNames = new Set(dirData.modules.map(m => m.name.toLowerCase()));
    const newModuleNames = [...sourceMap.keys()].filter(k => !existingNames.has(k));

    // Check for stream name collisions before adding new modules
    const existingStreamNames = new Set(dirData.modules.map(m => m.streamName.toLowerCase()));
    for (const k of newModuleNames) {
        const origName = sourceFileNames.get(k)!;
        if (existingStreamNames.has(origName.toLowerCase())) {
            console.error(`Error: stream name collision for new module '${origName}'`);
            process.exit(1);
        }
    }

    // Refuse to add a brand-new designer/document module: UserForms, worksheets
    // (Sheet*), and ThisWorkbook. These are bound to a host object that lives
    // outside vbaProject.bin — a UserForm needs its binary form-layout storage
    // (<name>/f, <name>/o, <name>/VBFrame), and a Sheet/Workbook module must map to
    // a real sheet/workbook declared in the xlsx parts (workbook.xml codeName, the
    // worksheet XML). Their .cls source alone cannot reconstruct that, so adding one
    // crashes or corrupts Excel. All of them carry the designer-only VB_Base
    // attribute, which a normal class/standard module never has — use it to detect.
    for (const k of newModuleNames) {
        const origName = sourceFileNames.get(k)!;
        if (/^\s*Attribute\s+VB_Base\s*=/im.test(sourceMap.get(k)!)) {
            console.error(
                `Error: cannot add document/designer module '${origName}' (UserForm, ` +
                `worksheet, or ThisWorkbook). These are bound to a host object that a .cls ` +
                `source cannot recreate. Add the form/sheet in Excel first, then use import ` +
                `to update its code-behind.`,
            );
            process.exit(1);
        }
    }

    // Process existing modules: update source or mark for deletion
    const updatedModules: VbaModuleFull[] = [];
    const deletedNames: string[] = [];
    let updated = 0;
    let deleted = 0;

    for (const mod of dirData.modules) {
        const src = sourceMap.get(mod.name.toLowerCase());

        if (src === undefined) {
            deletedNames.push(mod.name);
            // Delete VBA source code stream
            const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
            if (entry) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (entry as any).type = 0; // STGTY_INVALID → cfb.js skips it on write
            }
            // Delete form layout storage if present (UserForm has a matching root-level storage)
            // e.g. Root Entry/UserForm1/, Root Entry/UserForm1/f, Root Entry/UserForm1/o, ...
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const cfbAny3 = cfb as any;
            const formPrefix = `Root Entry/${mod.streamName}/`;
            let formRemoved = false;
            for (let idx = 0; idx < cfbAny3.FullPaths.length; idx++) {
                const p: string = cfbAny3.FullPaths[idx];
                if (p === formPrefix || p.startsWith(formPrefix)) {
                    cfbAny3.FileIndex[idx].type = 0;
                    formRemoved = true;
                }
            }
            console.log(`  - ${mod.name}: removed${formRemoved ? ' (including form layout)' : ''}`);
            deleted++;
            // Do not push to updatedModules — excluded from dir stream
            continue;
        }

        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) { console.warn(`  [warn] ${mod.name}: stream not found`); continue; }

        const srcBytes = iconv.encode(src, encoding);
        const compressed = compress(srcBytes);

        // Decompile the module: the stream becomes compressed source only and
        // MODULEOFFSET is set to 0 (no p-code preamble). _VBA_PROJECT (rewritten below)
        // forces Excel to recompile from source on open, so any stale preamble would be
        // ignored anyway. This is the clean source-only project EPPlus produces; Excel
        // opens it, recompiles, and supports all operations including sheet copy.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (entry as any).content = compressed;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (entry as any).size = compressed.length;
        console.log(`  ✓ ${mod.name} (${srcBytes.length} bytes as ${encoding})`);
        updatedModules.push({ ...mod, offset: 0 });
        updated++;
    }

    // Add new modules
    for (const k of newModuleNames) {
        const origName = sourceFileNames.get(k)!;
        const src = sourceMap.get(k)!;
        const srcBytes = iconv.encode(src, encoding);
        const compressed = compress(srcBytes);

        // Determine if class module by checking file extension
        const f = readdirSync(absSrc).find(file => basename(file, extname(file)).toLowerCase() === k);
        const isClass = f ? extname(f).toLowerCase() === '.cls' : false;

        // Use cfb_add so cfb.js builds a valid red-black directory tree (cfb_gc).
        // Manually pushing FileIndex/FullPaths leaves the tree inconsistent → Excel "repair".
        CFB.utils.cfb_add(cfb, `/VBA/${origName}`, Buffer.from(compressed));

        console.log(`  + ${origName} (new, ${srcBytes.length} bytes as ${encoding})`);
        updated++;

        // New module: empty rawBlock signals dir-builder to synthesize records
        updatedModules.push({
            name: origName,
            streamName: origName,
            offset: 0,
            isClass,
            rawBlock: Buffer.alloc(0),
        });
    }

    if (updated === 0 && deleted === 0) {
        console.error('Error: no modules were updated or removed.');
        process.exit(1);
    }

    // Patch PROJECT stream (module list + workspace entries)
    const projEntry = CFB.find(cfb, '/PROJECT');
    if (projEntry) {
        const projText = iconv.decode(Buffer.from(projEntry.content as unknown as ArrayBuffer), encoding);
        const deletedLower = new Set(deletedNames.map(n => n.toLowerCase()));
        const addedMods = newModuleNames.map(k => {
            const f = readdirSync(absSrc).find(file => basename(file, extname(file)).toLowerCase() === k);
            return { name: sourceFileNames.get(k)!, isClass: f ? extname(f).toLowerCase() === '.cls' : false };
        });
        const newText = patchProjectStream(projText, deletedLower, addedMods);
        const newBytes = iconv.encode(newText, encoding);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projEntry as any).content = newBytes;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projEntry as any).size = newBytes.length;
    }

    // Patch PROJECTwm stream (module name unicode mappings)
    const projWmEntry = CFB.find(cfb, '/PROJECTwm');
    if (projWmEntry) {
        const wmData = Buffer.from(projWmEntry.content as unknown as ArrayBuffer);
        const deletedLower = new Set(deletedNames.map(n => n.toLowerCase()));
        const addedNames2 = newModuleNames.map(k => sourceFileNames.get(k)!);
        const newWm = patchProjectWm(wmData, deletedLower, addedNames2);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projWmEntry as any).content = newWm;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projWmEntry as any).size = newWm.length;
    }

    // Replace _VBA_PROJECT with the canonical 7-byte source-only header so Excel
    // discards any compiled p-code and recompiles from the module source. This is
    // exactly what EPPlus (a production library whose generated .xlsm files open,
    // run, and support sheet operations in Excel) writes [MS-OVBA 2.3.4.1]:
    //   CC 61  — Reserved1 (0x61CC)
    //   FF FF  — Version (0xFFFF: matches no real VBA version → forces recompile)
    //   00     — Reserved3
    //   00 00  — Reserved4
    // The stream length is exactly 7, so PerformanceCache is empty (len = size - 7 = 0).
    // (A full-size zero-filled _VBA_PROJECT or a 0x0000 version is what previously
    //  caused Excel's "sheet cannot be copied" failure.)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const vbaProjectEntry = CFB.find(cfb, '/VBA/_VBA_PROJECT') as any;
    if (vbaProjectEntry) {
        vbaProjectEntry.content = Buffer.from([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]);
        vbaProjectEntry.size = 7;
    }
    // Delete every __SRP_* performance-cache stream. EPPlus never writes these;
    // leaving stale per-module caches behind corrupts Excel's project operations.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const cfbAnySrp = cfb as any;
    for (let i = 0; i < cfbAnySrp.FullPaths.length; i++) {
        if ((cfbAnySrp.FullPaths[i] as string).includes('/__SRP_')) {
            cfbAnySrp.FileIndex[i].type = 0; // STGTY_INVALID → cfb.js skips on write
        }
    }

    // Rebuild the dir stream: every module is now decompiled (MODULEOFFSET = 0) and the
    // module list may have changed (add/delete), so the dir is regenerated from scratch.
    const newDirUncompressed = buildDirStream(dirData, updatedModules, encoding);
    const newDirCompressed = compress(newDirUncompressed);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (dirEntry as any).content = newDirCompressed;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (dirEntry as any).size = newDirCompressed.length;

    // Note: cfb.js always seeds a harmless 'Sh33tJ5' signature stream during
    // write (it cannot be suppressed — every rebuild re-adds it). Leave it as a valid
    // stream; Office ignores unknown CFB streams. Do NOT try to "remove" it by zeroing
    // its directory entry type — that leaves dangling red-black tree pointers and makes
    // Excel reject the whole vbaProject.bin ("unreadable content" / cannot copy sheet).
    const newVbaBin = Buffer.from(CFB.write(cfb, { type: 'buffer' }) as unknown as ArrayBuffer);
    zip.file('xl/vbaProject.bin', newVbaBin, { createFolders: false });
    const newXlsm = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    writeFileSync(outPath, newXlsm);
    console.log(`Saved: ${outPath}`);
}

async function main(): Promise<void> {
    const [,, mode, ...rest] = process.argv;

    if (mode === '--version' || mode === '-v') {
        console.log(VERSION);
        process.exit(0);
    }

    if (!mode || mode === '--help' || mode === '-h') {
        printUsage();
        process.exit(mode ? 0 : 1);
    }

    if (mode !== 'export' && mode !== 'import') {
        console.error(`vba-extractor: unknown command '${mode}'`);
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
