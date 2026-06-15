#!/usr/bin/env tsx
/**
 * Integration test for vba-extractor import.
 *
 * Generates test1.xlsm – test6.xlsm in sample/excel/.
 * Open each in Excel and verify as noted per test case.
 *
 * Usage:
 *   npx tsx tools/extractor/run-import-tests.ts
 *
 * Test cases:
 *   test1 — Content size increase past 4096 bytes compressed (miniSAT → FAT boundary).
 *            A Function is appended that builds a string via 300 "..." & _ continuation
 *            lines, each with a unique rotated alphabet — OVBA compresses to ~4667 bytes.
 *   test2 — Content size decrease: Module1 replaced with minimal code (FAT → miniSAT).
 *   test3 — Add new module: NewModule.bas
 *   test4 — Remove module: Class1.cls deleted
 *   test5 — Remove UserForm: UserForm1.cls deleted (form layout storage also removed)
 *   test6 — Add UserForm (code-behind only): UserForm2.cls added
 *            (no binary form layout — appears as class module in VBA Editor)
 */

import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import {
    readFileSync, writeFileSync, mkdirSync, existsSync,
    readdirSync, copyFileSync, rmSync,
} from 'fs';
import { resolve, basename, extname, join } from 'path';
import { decompress, compress } from './ovba.js';
import { parseDirStream, parseDirStreamFull, VbaModuleFull } from './dir-parser.js';
import { buildDirStream, patchModuleOffset } from './dir-builder.js';

const XLSM_IN  = resolve('sample/excel/test.xlsm');
const XLSM_OUT = resolve('sample/excel');
const TMP_ROOT = '/tmp/vba-extractor-test';

// ── PROJECT stream patching ───────────────────────────────────────────────────

function patchProjectStream(text: string, deletedLower: Set<string>, added: { name: string; isClass: boolean }[]): string {
    const lines = text.split(/\r\n|\r|\n/);
    const out: string[] = [];
    let inWorkspace = false;
    let pendingPackage: string | null = null;
    let insertIdx = -1;
    for (const line of lines) {
        const t = line.trim();
        if (!inWorkspace && t.toLowerCase() === '[workspace]') inWorkspace = true;
        if (inWorkspace) {
            const eq = t.indexOf('=');
            if (eq > 0 && deletedLower.has(t.slice(0, eq).toLowerCase())) continue;
            out.push(line); continue;
        }
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
            out.push(line); insertIdx = out.length; continue;
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

function patchProjectWm(data: Buffer, deletedLower: Set<string>, addedNames: string[]): Buffer {
    const parts: Buffer[] = [];
    let i = 0;
    while (i < data.length) {
        if (data[i] === 0) break;
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
    parts.push(Buffer.from([0, 0]));
    return Buffer.concat(parts);
}

// ── Core helpers ─────────────────────────────────────────────────────────────

async function openVba(xlsmPath: string) {
    const zip = await JSZip.loadAsync(readFileSync(xlsmPath));
    const bin = zip.file('xl/vbaProject.bin');
    if (!bin) throw new Error(`xl/vbaProject.bin not found in ${xlsmPath}`);
    const vbaBuf   = Buffer.from(await bin.async('nodebuffer'));
    const cfb      = CFB.read(vbaBuf, { type: 'buffer' });
    const dirEntry = CFB.find(cfb, '/VBA/dir');
    if (!dirEntry) throw new Error('/VBA/dir not found');
    const dirRaw      = decompress(Buffer.from(dirEntry.content as unknown as ArrayBuffer));
    const { codePage } = parseDirStream(dirRaw);
    const encoding    = `cp${codePage ?? 1252}`;
    const dirData     = parseDirStreamFull(dirRaw);
    return { zip, cfb, dirEntry, dirRaw, dirData, encoding };
}

async function doExport(xlsmPath: string, outDir: string): Promise<string> {
    mkdirSync(outDir, { recursive: true });
    const { cfb, dirData, encoding } = await openVba(xlsmPath);
    for (const mod of dirData.modules) {
        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) continue;
        const raw = Buffer.from(entry.content as unknown as ArrayBuffer);
        const src = iconv.decode(decompress(raw.subarray(mod.offset)), encoding);
        writeFileSync(join(outDir, mod.name + (mod.isClass ? '.cls' : '.bas')), src, 'utf8');
    }
    return encoding;
}

async function doImport(srcXlsm: string, srcDir: string, outXlsm: string): Promise<void> {
    const { zip, cfb, dirEntry, dirData, encoding } = await openVba(srcXlsm);

    // Build source file maps
    const sourceMap     = new Map<string, string>();   // lower name → text
    const sourceNames   = new Map<string, string>();   // lower name → original name
    const sourceIsClass = new Map<string, boolean>();  // lower name → isClass
    for (const f of readdirSync(srcDir)) {
        const e = extname(f).toLowerCase();
        if (e !== '.bas' && e !== '.cls') continue;
        const n = basename(f, e);
        sourceMap.set(n.toLowerCase(), readFileSync(join(srcDir, f), 'utf8'));
        sourceNames.set(n.toLowerCase(), n);
        sourceIsClass.set(n.toLowerCase(), e === '.cls');
    }

    const existingKeys = new Set(dirData.modules.map(m => m.name.toLowerCase()));
    const newKeys      = [...sourceMap.keys()].filter(k => !existingKeys.has(k));
    const updatedMods: VbaModuleFull[] = [];
    const deletedNames: string[] = [];

    // Refuse to add a brand-new UserForm (no binary form layout in a .cls source).
    for (const k of newKeys) {
        if (/^\s*Attribute\s+VB_Base\s*=/im.test(sourceMap.get(k)!)) {
            throw new Error(
                `cannot add UserForm '${sourceNames.get(k)!}' — UserForms require a binary ` +
                `form layout that is not present in the .cls source`,
            );
        }
    }

    // Process existing modules: update or delete
    for (const mod of dirData.modules) {
        const src = sourceMap.get(mod.name.toLowerCase());

        if (src === undefined) {
            deletedNames.push(mod.name);
            // Delete VBA source stream
            const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
            if (entry) (entry as any).type = 0;
            // Delete form layout storage if present (Root Entry/<streamName>/)
            const prefix = `Root Entry/${mod.streamName}/`;
            let hadLayout = false;
            for (let i = 0; i < (cfb as any).FullPaths.length; i++) {
                if (((cfb as any).FullPaths[i] as string).startsWith(prefix)) {
                    (cfb as any).FileIndex[i].type = 0;
                    hadLayout = true;
                }
            }
            console.log(`  - ${mod.name}${hadLayout ? ' (+ form layout)' : ''}`);
            continue;
        }

        const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
        if (!entry) { console.warn(`  [warn] ${mod.name}: stream not found`); continue; }

        const srcBytes   = iconv.encode(src, encoding);
        const compressed = compress(srcBytes);
        (entry as any).content = compressed;
        (entry as any).size    = compressed.length;
        console.log(`  ✓ ${mod.name} (${srcBytes.length} source bytes → ${compressed.length} compressed)`);
        updatedMods.push({ ...mod, offset: 0, rawBlock: patchModuleOffset(mod.rawBlock, 0) });
    }

    // Add new modules
    for (const k of newKeys) {
        const name     = sourceNames.get(k)!;
        const src      = sourceMap.get(k)!;
        const isClass  = sourceIsClass.get(k)!;
        const srcBytes   = iconv.encode(src, encoding);
        const compressed = compress(srcBytes);
        const content  = Buffer.from(compressed);

        // Use cfb_add so cfb.js builds a valid red-black directory tree (cfb_gc).
        // Manually pushing FileIndex/FullPaths leaves the tree inconsistent → Excel "repair".
        CFB.utils.cfb_add(cfb, `/VBA/${name}`, content);
        console.log(`  + ${name} (new, ${srcBytes.length} source bytes → ${compressed.length} compressed)`);
        updatedMods.push({ name, streamName: name, offset: 0, isClass, rawBlock: Buffer.alloc(0) });
    }

    // Patch PROJECT stream (module list + workspace)
    const projEntry = CFB.find(cfb, '/PROJECT');
    if (projEntry) {
        const projText = iconv.decode(Buffer.from(projEntry.content as unknown as ArrayBuffer), encoding);
        const deletedLower = new Set(deletedNames.map(n => n.toLowerCase()));
        const addedMods = newKeys.map(k => ({ name: sourceNames.get(k)!, isClass: sourceIsClass.get(k)! }));
        const newText = patchProjectStream(projText, deletedLower, addedMods);
        const newBytes = iconv.encode(newText, encoding);
        (projEntry as any).content = newBytes;
        (projEntry as any).size    = newBytes.length;
    }

    // Patch PROJECTwm stream (unicode name mappings)
    const projWmEntry = CFB.find(cfb, '/PROJECTwm');
    if (projWmEntry) {
        const wmData = Buffer.from(projWmEntry.content as unknown as ArrayBuffer);
        const deletedLower = new Set(deletedNames.map(n => n.toLowerCase()));
        const newWm = patchProjectWm(wmData, deletedLower, newKeys.map(k => sourceNames.get(k)!));
        (projWmEntry as any).content = newWm;
        (projWmEntry as any).size    = newWm.length;
    }

    // Invalidate the compiled p-code cache so Excel rebuilds it from source.
    // Proven approach: keep the stream SIZE unchanged and keep the first 2 magic bytes,
    // but zero every byte from offset 2 onward. Same size → no FAT/miniSAT boundary change
    // and no "repair" prompt; magic preserved → the VBE can still load the project;
    // zeroed body → Excel discards the stale cache and recompiles (no macro crash).
    // The same treatment applies to every __SRP_* performance-cache stream.
    const zeroCacheBody = (entry: any) => {
        const buf = Buffer.from(entry.content);
        buf.fill(0, 2);
        entry.content = buf;
        entry.size = buf.length;
    };
    const vbaProjectEntry = CFB.find(cfb, '/VBA/_VBA_PROJECT');
    if (vbaProjectEntry) zeroCacheBody(vbaProjectEntry);
    const cfbAnySrp = cfb as any;
    for (let i = 0; i < cfbAnySrp.FullPaths.length; i++) {
        if ((cfbAnySrp.FullPaths[i] as string).includes('/__SRP_')) {
            zeroCacheBody(cfbAnySrp.FileIndex[i]);
        }
    }

    // Rebuild dir stream
    const newDir = compress(buildDirStream(dirData, updatedMods, encoding));
    (dirEntry as any).content = newDir;
    (dirEntry as any).size    = newDir.length;

    const newBin  = CFB.write(cfb, { type: 'buffer' });
    zip.file('xl/vbaProject.bin', newBin as unknown as Buffer, { createFolders: false });
    const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    writeFileSync(outXlsm, out);
    console.log(`  → ${outXlsm}`);
}

// ─── Utility ─────────────────────────────────────────────────────────────────

function copyDir(src: string, dst: string) {
    if (existsSync(dst)) rmSync(dst, { recursive: true });
    mkdirSync(dst, { recursive: true });
    for (const f of readdirSync(src)) copyFileSync(join(src, f), join(dst, f));
}

/**
 * Generate a VBA Function body (no module header) whose OVBA-compressed size
 * exceeds 4096 bytes when appended to an existing module.
 *
 * Uses 300 `& _` continuation lines, each with a unique rotated-alphabet string,
 * so the content is hard to compress. Empirically: ~22,635 source bytes →
 * ~4,667 compressed bytes at 300 lines (>4096 threshold).
 *
 * Groups of 20 lines form one VBA statement to stay within VBA's logical-line limit.
 * The caller should append this to existing module source — do NOT use as a standalone
 * module, because class modules require VB_Base / VB_GlobalNameSpace etc. attributes
 * that only the original source contains.
 */
function generateLargeStringFunction(lineCount = 300): string {
    const alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const GROUP = 20;
    const lines: string[] = [];

    for (let i = 0; i < lineCount; i++) {
        const num     = String(i + 1).padStart(3, '0');
        const offset  = i % alpha.length;
        const rotated = (alpha + alpha).slice(offset, offset + 50);
        const content = `L${num}:${rotated}_${num}`; // unique: index embedded at both ends

        const isGroupStart = i % GROUP === 0;
        const isEnd        = (i + 1) % GROUP === 0 || i === lineCount - 1;

        if (isGroupStart) {
            const kw = i === 0 ? '    s = ' : '    s = s & ';
            lines.push(`${kw}"${content}"${isEnd ? '' : ' & _'}`);
        } else {
            lines.push(`        "${content}"${isEnd ? '' : ' & _'}`);
        }
    }

    return [
        '',
        '\'  GetTestData: 300-line & _ string concat to push stream past 4096-byte CFB boundary.',
        'Function GetTestData() As String',
        '    Dim s As String',
        ...lines,
        '    GetTestData = s',
        'End Function',
        '',
    ].join('\r\n');
}

function banner(label: string, desc: string) {
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`${label}: ${desc}`);
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
    if (existsSync(TMP_ROOT)) rmSync(TMP_ROOT, { recursive: true });

    // ── Export base modules ──────────────────────────────────────────────────
    const baseDir = join(TMP_ROOT, 'base');
    banner('Export', XLSM_IN);
    const encoding = await doExport(XLSM_IN, baseDir);
    const baseFiles = readdirSync(baseDir);
    console.log('Modules :', baseFiles.join(', '));
    console.log('Encoding:', encoding);

    // Detect module roles from CFB structure
    const { cfb: cfb0, dirData: dd0 } = await openVba(XLSM_IN);
    const hasFormStorage = (streamName: string) =>
        (cfb0 as any).FullPaths.some((p: string) => p.startsWith(`Root Entry/${streamName}/`));
    const isBuiltin = (name: string) =>
        name.toLowerCase() === 'thisworkbook' || name.toLowerCase().startsWith('sheet');

    const stdModule  = dd0.modules.find(m => !m.isClass);
    const clsModule  = dd0.modules.find(m => m.isClass && !isBuiltin(m.name) && !hasFormStorage(m.streamName));
    const formModule = dd0.modules.find(m => m.isClass && hasFormStorage(m.streamName));

    console.log('Standard module :', stdModule?.name ?? '(none)');
    console.log('Class module    :', clsModule?.name ?? '(none)');
    console.log('UserForm module :', formModule?.name ?? '(none)');

    // ── test1: Size increase past 4096 bytes compressed (miniSAT → FAT) ───────
    banner('test1', 'Size increase: & _ string concat (300 lines) → compressed >4096 bytes (miniSAT→FAT)');
    if (clsModule) {
        const dir = join(TMP_ROOT, 'test1');
        copyDir(baseDir, dir);
        const target   = join(dir, clsModule.name + '.cls');
        const original = readFileSync(target, 'utf8');
        const appended = generateLargeStringFunction();
        const newSrc   = original + appended;
        writeFileSync(target, newSrc, 'utf8');
        const srcBytes = iconv.encode(newSrc, encoding);
        const cmpSize  = compress(srcBytes).length;
        console.log(`  Appended GetTestData() to ${clsModule.name}.cls: ${srcBytes.length} source bytes → ${cmpSize} compressed (${cmpSize > 4096 ? '✓ >4096' : '✗ <4096'})`);
        await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test1.xlsm'));
        console.log(`  ✔ Verify in Excel: ${clsModule.name} has GetTestData() function with 300-line string concat`);
    } else {
        console.log('  SKIPPED (no class module found)');
    }

    // ── test2: Size decrease — Module1 replaced with minimal code ───────────
    banner('test2', 'Size decrease: standard module replaced with minimal code (FAT → miniSAT)');
    if (stdModule) {
        const dir = join(TMP_ROOT, 'test2');
        copyDir(baseDir, dir);
        const target = join(dir, stdModule.name + '.bas');
        writeFileSync(target,
            `Attribute VB_Name = "${stdModule.name}"\r\n` +
            'Sub Hello()\r\n' +
            '    MsgBox "Hello from test2 (minimal module)"\r\n' +
            'End Sub\r\n', 'utf8');
        console.log(`  Replaced ${stdModule.name}.bas with minimal code`);
        await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test2.xlsm'));
        console.log(`  ✔ Verify in Excel: ${stdModule.name} contains only Hello() sub, macro runs`);
    } else {
        console.log('  SKIPPED (no standard module found)');
    }

    // ── test3: Add new module ────────────────────────────────────────────────
    banner('test3', 'Add new module: NewModule.bas');
    {
        const dir = join(TMP_ROOT, 'test3');
        copyDir(baseDir, dir);
        writeFileSync(join(dir, 'NewModule.bas'),
            'Attribute VB_Name = "NewModule"\r\n' +
            'Sub SayHello()\r\n' +
            '    MsgBox "Hello from NewModule (added by vba-extractor)!"\r\n' +
            'End Sub\r\n', 'utf8');
        console.log('  Added NewModule.bas');
        await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test3.xlsm'));
        console.log('  ✔ Verify in Excel: NewModule appears in Project Explorer, SayHello() runs');
    }

    // ── test4: Remove class module ───────────────────────────────────────────
    banner('test4', `Remove class module: ${clsModule?.name ?? '(none)'}.cls`);
    if (clsModule) {
        const dir = join(TMP_ROOT, 'test4');
        copyDir(baseDir, dir);
        rmSync(join(dir, clsModule.name + '.cls'));
        console.log(`  Deleted ${clsModule.name}.cls from source dir`);
        await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test4.xlsm'));
        console.log(`  ✔ Verify in Excel: ${clsModule.name} is gone from Project Explorer`);
    } else {
        console.log('  SKIPPED (no class module found)');
    }

    // ── test5: Remove UserForm (source + form layout) ───────────────────────
    banner('test5', `Remove UserForm: ${formModule?.name ?? '(none)'}.cls + form layout storage`);
    if (formModule) {
        const dir = join(TMP_ROOT, 'test5');
        copyDir(baseDir, dir);
        rmSync(join(dir, formModule.name + '.cls'));
        console.log(`  Deleted ${formModule.name}.cls from source dir`);
        await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test5.xlsm'));
        console.log(`  ✔ Verify in Excel: ${formModule.name} is completely gone (no zombie in Project Explorer)`);
    } else {
        console.log('  SKIPPED (no UserForm found)');
    }

    // ── test6: Adding a brand-new UserForm must be REFUSED ──────────────────
    // A UserForm needs its binary form layout (f/o/VBFrame), absent from a .cls source.
    // import must reject it with a clear error rather than produce a file that crashes Excel.
    banner('test6', 'Add UserForm2.cls (code-behind only) — must be refused');
    {
        const dir = join(TMP_ROOT, 'test6');
        copyDir(baseDir, dir);
        // Use the real exported UserForm source (carries the designer VB_Base attribute).
        const frmSrc = formModule
            ? readFileSync(join(baseDir, formModule.name + '.cls'), 'utf8')
                .replace(new RegExp(`VB_Name = "${formModule.name}"`, 'i'), 'VB_Name = "UserForm2"')
            : 'Attribute VB_Name = "UserForm2"\r\n' +
              'Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}{...}"\r\n' +
              'Private Sub UserForm_Initialize()\r\nEnd Sub\r\n';
        writeFileSync(join(dir, 'UserForm2.cls'), frmSrc, 'utf8');
        let refused = false;
        try {
            await doImport(XLSM_IN, dir, join(XLSM_OUT, 'test6.xlsm'));
        } catch (e) {
            refused = true;
            console.log(`  ✔ Correctly refused: ${(e as Error).message}`);
        }
        if (!refused) console.log('  ✗ FAIL: import should have refused to add a UserForm');
        // Remove any stale test6.xlsm from earlier runs so it isn't mistaken for output.
        rmSync(join(XLSM_OUT, 'test6.xlsm'), { force: true });
    }

    // ── Summary ──────────────────────────────────────────────────────────────
    console.log(`\n${'═'.repeat(60)}`);
    console.log('All tests completed. Files written to sample/excel/');
    const written = readdirSync(XLSM_OUT).filter(f => /^test\d+\.xlsm$/.test(f)).sort();
    for (const f of written) console.log(`  ${f}`);
    console.log('\nOpen each file in Excel and verify as described above.');
}

main().catch(err => { console.error(err); process.exit(1); });
