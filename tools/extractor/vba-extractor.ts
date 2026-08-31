import JSZip from 'jszip';
import CFB from 'cfb';
import iconv from 'iconv-lite';
import { readFileSync, writeFileSync, copyFileSync, mkdirSync, existsSync, readdirSync, mkdtempSync, rmSync } from 'fs';
import { createInterface } from 'readline';
import { resolve, dirname, basename, extname, join } from 'path';
import { randomUUID } from 'crypto';
import { tmpdir } from 'os';
import { decompress, compress } from './ovba.js';
import { parseDirStream, parseDirStreamFull, VbaModuleFull } from './dir-parser.js';
import { buildDirStream, buildMinimalDirStream } from './dir-builder.js';

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

        if (t.startsWith('Document=')) {
            out.push(line);
            // New standard/class declarations belong after all host Document
            // declarations, even when the input project has no other modules.
            insertIdx = out.length;
            continue;
        }

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

/** Return workbook/document module names that are backed by the OOXML host. */
async function hostDocumentNames(zip: JSZip, documentNames: Set<string>): Promise<Set<string>> {
    const names = new Set<string>(['thisworkbook']);
    const workbook = zip.file('xl/workbook.xml');
    if (!workbook) return names;
    const workbookXml = await workbook.async('string');
    for (const m of workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/gi)) {
        if (documentNames.has(m[1].toLowerCase())) names.add(m[1].toLowerCase());
    }
    for (const m of workbookXml.matchAll(/<workbookPr\b[^>]*\bcodeName="([^"]+)"/gi)) {
        if (documentNames.has(m[1].toLowerCase())) names.add(m[1].toLowerCase());
    }
    // Worksheet codeName is optional and is the name used by PROJECT Document=.
    for (const path of Object.keys(zip.files)) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(path)) continue;
        const file = zip.file(path);
        if (!file) continue;
        const xml = await file.async('string');
        for (const m of xml.matchAll(/<sheetPr\b[^>]*\bcodeName="([^"]+)"/gi)) {
            if (documentNames.has(m[1].toLowerCase())) names.add(m[1].toLowerCase());
        }
    }
    return names;
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
    console.log('  vba-extractor import <input.xlsm> [source-dir] [output.xlsm] [--encoding <cp>] [--yes]');
    console.log('    source-dir defaults to: <input.xlsm directory>/src');
    console.log('');
    console.log('Options:');
    console.log('  --encoding <cp>   Override the source encoding (e.g. cp932)');
    console.log('  --yes, -y         Skip the confirmation prompt before import (for scripts/CI)');
    console.log('  --version         Show version');
    console.log('  --help            Show this help');
}

function parseEncoding(args: string[]): { encoding: string | undefined; yes: boolean; rest: string[] } {
    const rest: string[] = [];
    let encoding: string | undefined;
    let yes = false;
    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--encoding') {
            encoding = args[++i];
        } else if (args[i] === '--yes' || args[i] === '-y') {
            yes = true;
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
    return { encoding, yes, rest };
}

async function openXlsm(xlsmPath: string) {
    const zip = await JSZip.loadAsync(readFileSync(xlsmPath));
    const entry = zip.file('xl/vbaProject.bin');
    if (!entry) {
        console.error('xl/vbaProject.bin not found: this file has no VBA project yet.');
        console.error('Saving a workbook as .xlsm alone does not create one; add a VBA module in Excel first,');
        console.error('or use the bundled Build-Xlsm.ps1 script on Windows to create the initial workbook.');
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

function workbookHostNames(workbookXml: string): string[] {
    const names = new Set<string>(['ThisWorkbook']);
    for (const m of workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/gi)) names.add(m[1]);
    for (const m of workbookXml.matchAll(/<workbookPr\b[^>]*\bcodeName="([^"]+)"/gi)) names.add(m[1]);
    return [...names];
}

function minimalHostSource(name: string, isWorkbook: boolean): string {
    const base = isWorkbook
        ? '0{00020819-0000-0000-C000-000000000046}'
        : '0{00020820-0000-0000-C000-000000000046}';
    return [
        `Attribute VB_Name = "${name}"`,
        `Attribute VB_Base = "${base}"`,
        'Attribute VB_GlobalNameSpace = False',
        'Attribute VB_Creatable = False',
        'Attribute VB_PredeclaredId = True',
        'Attribute VB_Exposed = True',
        'Attribute VB_TemplateDerived = False',
        'Attribute VB_Customizable = True',
        '',
    ].join('\r\n');
}

function buildProjectWm(names: string[]): Buffer {
    const chunks: Buffer[] = [];
    for (const name of names) {
        chunks.push(Buffer.from(`${name}\0`, 'latin1'));
        chunks.push(Buffer.from(name, 'utf16le'));
        chunks.push(Buffer.from([0, 0]));
    }
    chunks.push(Buffer.from([0, 0]));
    return Buffer.concat(chunks);
}

async function createVbaProject(
    zip: JSZip, sourceMap: Map<string, string>, sourceNames: Map<string, string>,
    sourceClasses: Map<string, boolean>, encoding: string,
): Promise<void> {
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string') ?? '';
    const hostNames = workbookHostNames(workbookXml);
    const modules: VbaModuleFull[] = [];
    const moduleSources = new Map<string, string>();
    for (const name of hostNames) {
        const key = name.toLowerCase();
        moduleSources.set(key, sourceMap.get(key) ?? minimalHostSource(name, key === 'thisworkbook'));
        modules.push({ name, streamName: name, offset: 0, isClass: true, rawBlock: Buffer.alloc(0) });
    }
    for (const [key, source] of sourceMap) {
        if (modules.some(m => m.name.toLowerCase() === key)) continue;
        const name = sourceNames.get(key) ?? source.match(/^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/im)?.[1] ?? key;
        const isClass = sourceClasses.get(key) ?? false;
        moduleSources.set(key, source);
        modules.push({ name, streamName: name, offset: 0, isClass, rawBlock: Buffer.alloc(0) });
    }
    const cfb = CFB.utils.cfb_new();
    for (const mod of modules) {
        const source = moduleSources.get(mod.name.toLowerCase()) ?? '';
        CFB.utils.cfb_add(cfb, `/VBA/${mod.streamName}`, compress(iconv.encode(source, encoding)));
    }
    const dir = buildMinimalDirStream(modules, encoding);
    CFB.utils.cfb_add(cfb, '/VBA/dir', compress(dir));
    CFB.utils.cfb_add(cfb, '/VBA/_VBA_PROJECT', Buffer.from([0xcc, 0x61, 0xff, 0xff, 0, 0, 0]));
    const projectLines = [
        `ID="{${randomUUID().toUpperCase()}}"`,
        ...modules.filter(m => hostNames.some(h => h.toLowerCase() === m.name.toLowerCase()))
            .map(m => `Document=${m.name}/&H00000000`),
        ...modules.filter(m => !hostNames.some(h => h.toLowerCase() === m.name.toLowerCase()))
            .map(m => `${m.isClass ? 'Class' : 'Module'}=${m.name}`),
        'Name="VBAProject"',
        'HelpContextID="0"',
        'HelpFile=""',
        'Description=""',
        'VersionCompatible32="393222000"',
        '',
        '[Host Extender Info]',
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
        '',
        '[Workspace]',
        ...modules.map(m => `${m.name}=0, 0, 2000, 1000, C`),
        '',
    ].join('\r\n');
    CFB.utils.cfb_add(cfb, '/PROJECT', iconv.encode(projectLines, encoding));
    CFB.utils.cfb_add(cfb, '/PROJECTwm', buildProjectWm(modules.map(m => m.name)));
    zip.file('xl/vbaProject.bin', Buffer.from(CFB.write(cfb, { type: 'buffer' }) as unknown as ArrayBuffer));

    // Bind generated Document modules to their OOXML host objects.  Without
    // these code names Excel creates duplicate ThisWorkbook/Sheet modules.
    let patchedWorkbook = workbookXml;
    if (/<workbookPr\b/i.test(patchedWorkbook)) {
        patchedWorkbook = patchedWorkbook.replace(/<workbookPr\b([^>]*?)(\/>|>)/i, (all, attrs, end) =>
            /\bcodeName\s*=/.test(attrs) ? all : `<workbookPr${attrs} codeName="ThisWorkbook"${end}`);
    } else {
        patchedWorkbook = patchedWorkbook.replace(/(<workbook\b[^>]*>)/i, '$1<workbookPr codeName="ThisWorkbook"/>');
    }
    if (patchedWorkbook !== workbookXml) zip.file('xl/workbook.xml', patchedWorkbook);
    const sheetNames = [...patchedWorkbook.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/gi)].map(m => m[1]);
    for (let i = 0; i < sheetNames.length; i++) {
        const path = `xl/worksheets/sheet${i + 1}.xml`;
        const file = zip.file(path);
        if (!file) continue;
        let xml = await file.async('string');
        const codeName = sheetNames[i].replace(/&/g, '&amp;').replace(/"/g, '&quot;');
        if (/<sheetPr\b/i.test(xml)) {
            xml = xml.replace(/<sheetPr\b([^>]*?)(\/>|>)/i, (all, attrs, end) =>
                /\bcodeName\s*=/.test(attrs) ? all : `<sheetPr${attrs} codeName="${codeName}"${end}`);
        } else {
            xml = xml.replace(/(<worksheet\b[^>]*>)/i, `$1<sheetPr codeName="${codeName}"/>`);
        }
        zip.file(path, xml);
    }

    const types = zip.file('[Content_Types].xml');
    if (types) {
        let xml = await types.async('string');
        xml = xml.replace(/(PartName="\/xl\/workbook\.xml"\s+ContentType=")([^"]+)/i,
            '$1application/vnd.ms-excel.sheet.macroEnabled.main+xml');
        if (!/PartName="\/xl\/vbaProject\.bin"/i.test(xml)) {
            xml = xml.replace('</Types>', '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
        }
        zip.file('[Content_Types].xml', xml);
    }
    const rels = zip.file('xl/_rels/workbook.xml.rels');
    if (rels) {
        let xml = await rels.async('string');
        if (!/relationships\/vbaProject/i.test(xml)) {
            const ids = [...xml.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]));
            const id = `rId${Math.max(0, ...ids) + 1}`;
            xml = xml.replace('</Relationships>', `<Relationship Id="${id}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>`);
            zip.file('xl/_rels/workbook.xml.rels', xml);
        }
    }
}

function resolveEncoding(encodingOverride: string | undefined, codePage: number | null): string {
    if (!encodingOverride && codePage === null) {
        console.error('Error: PROJECTCODEPAGE not found. Specify encoding with --encoding.');
        process.exit(1);
    }
    return encodingOverride ?? `cp${codePage}`;
}

// The VBA project's internal module source stream (what `export` writes and what
// `import` expects) starts directly with `Attribute VB_Name = "..."`. The VBE's
// own "Export File..." menu command, by contrast, prepends a class-header block
// (`VERSION 1.0 CLASS` / `BEGIN ... END`) that holds component properties like
// MultiUse — this block lives in the PROJECT stream's component records, not in
// the module source stream itself. Feeding a VBE-exported .cls straight into
// `import` would write that header text as if it were VBA code and corrupt the
// module. Excel can place this block either before or after the initial Attribute
// lines, so locate it as a group of whole lines rather than assuming it is at byte 0.
function classHeaderRange(source: string): { start: number; end: number } | null {
    const startMatch = /(?:^﻿?|\r?\n)VERSION\s+[\d.]+\s+CLASS\r?\nBEGIN\r?\n/im.exec(source);
    if (!startMatch || startMatch.index === undefined) return null;

    // Keep the preceding newline (when present) and remove from VERSION onward.
    const leadingBreakLength = startMatch[0].startsWith('\r\n') ? 2 : startMatch[0].startsWith('\n') ? 1 : 0;
    const start = startMatch.index + leadingBreakLength;
    const afterBegin = start + startMatch[0].length - leadingBreakLength;
    const terminator = /^END\r?\n/im.exec(source.slice(afterBegin));
    if (!terminator || terminator.index === undefined) return null;

    return { start, end: afterBegin + terminator.index + terminator[0].length };
}

function stripClassHeader(source: string): string {
    const range = classHeaderRange(source);
    return range ? source.slice(0, range.start) + source.slice(range.end) : source;
}

// A VBE-exported .cls omits several attributes that Excel writes into the
// decompressed stream for a newly created ordinary class module. Preserve the
// user's line-ending style while adding only attributes that are absent.
function completeNewClassAttributes(source: string): string {
    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const hadFinalEol = /\r?\n$/.test(source);
    const lines = source.split(/\r\n|\r|\n/);
    if (hadFinalEol) lines.pop();

    const has = (name: string) => lines.some(line => new RegExp(`^\\s*Attribute\\s+${name}\\s*=`, 'i').test(line));
    const insertAfter = (name: string, line: string) => {
        const index = lines.findIndex(item => new RegExp(`^\\s*Attribute\\s+${name}\\s*=`, 'i').test(item));
        if (index >= 0) lines.splice(index + 1, 0, line);
        else lines.unshift(line);
    };

    if (!has('VB_Base')) insertAfter('VB_Name', 'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"');
    if (!has('VB_TemplateDerived')) insertAfter('VB_Exposed', 'Attribute VB_TemplateDerived = False');
    if (!has('VB_Customizable')) insertAfter('VB_TemplateDerived', 'Attribute VB_Customizable = False');

    return lines.join(eol) + (hadFinalEol ? eol : '');
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
    const { encoding: encodingOverride, yes, rest } = parseEncoding(args);
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
    console.log('   Standard/class modules are synced to <source-dir>; omitted host Document');
    console.log('   modules for existing workbook/sheets are retained with an empty source body.');
    console.log(`   Input  : ${absXlsm}`);
    console.log(`   Source : ${absSrc}`);
    console.log(`   Output : ${outPath}`);
    if (willOverwrite) {
        console.log(`   Backup : ${backupPath}`);
    }
    console.log('');

    if (yes) {
        console.log('--yes specified, skipping confirmation prompt.');
    } else {
        const question = willOverwrite
            ? 'Create a backup and run import? [y/N]: '
            : 'Run import? [y/N]: ';
        const confirmed = await promptYesNo(question);

        if (!confirmed) {
            console.log('Cancelled.');
            process.exit(0);
        }
    }

    if (willOverwrite) {
        copyFileSync(absXlsm, backupPath);
        console.log(`Backup created: ${backupPath}`);
    }

    // Build a map of source files: lowercase name → source text
    const sourceFileNames = new Map<string, string>(); // lowercase name → original filename (no ext)
    const sourceMap = new Map<string, string>();        // lowercase name → source text
    const sourceClasses = new Map<string, boolean>();   // lowercase name → .cls source
    let headerStrippedCount = 0;
    for (const f of readdirSync(absSrc)) {
        const ext = extname(f).toLowerCase();
        if (ext !== '.bas' && ext !== '.cls') continue;
        const baseName = basename(f, ext);
        let text = readFileSync(`${absSrc}/${f}`, 'utf8');
        if (ext === '.cls' && classHeaderRange(text)) {
            text = stripClassHeader(text);
            headerStrippedCount++;
        }
        sourceMap.set(baseName.toLowerCase(), text);
        sourceFileNames.set(baseName.toLowerCase(), baseName);
        sourceClasses.set(baseName.toLowerCase(), ext === '.cls');
    }
    console.log(`Source files : ${sourceMap.size}`);
    if (headerStrippedCount > 0) {
        console.log(`  (stripped VBE-style class header from ${headerStrippedCount} .cls file(s))`);
    }

    const inputZip = await JSZip.loadAsync(readFileSync(absXlsm));
    if (!inputZip.file('xl/vbaProject.bin')) {
        console.log('No VBA project found; initializing a minimal source-only project.');
        // Bootstrap only the host Document modules, then run the normal import
        // path.  This keeps module classification, casing, synchronization,
        // and PROJECT/dir updates identical for new and existing projects.
        const bootstrapDir = mkdtempSync(join(tmpdir(), 'vba-extractor-bootstrap-'));
        const bootstrapPath = join(bootstrapDir, 'bootstrap.xlsm');
        try {
            await createVbaProject(inputZip, new Map(), new Map(), new Map(), encodingOverride ?? 'cp932');
            const bootstrap = await inputZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
            writeFileSync(bootstrapPath, bootstrap);
            await runImport([bootstrapPath, absSrc, outPath, '--yes']);
        } finally {
            rmSync(bootstrapDir, { recursive: true, force: true });
        }
        return;
    }
    const { zip, cfb, codePage, dirEntry, dirDecompressed } = await openXlsm(absXlsm);
    const encoding = resolveEncoding(encodingOverride, codePage);
    const projectEntry = CFB.find(cfb, '/PROJECT');
    const projectText = projectEntry
        ? iconv.decode(Buffer.from(projectEntry.content as unknown as ArrayBuffer), encoding)
        : '';
    const documentNames = new Set<string>();
    for (const line of projectText.split(/\r\n|\r|\n/)) {
        if (line.trim().toLowerCase().startsWith('document=')) {
            documentNames.add(line.slice(line.indexOf('=') + 1).split('/')[0].trim().toLowerCase());
        }
    }
    const hostNames = await hostDocumentNames(zip, documentNames);

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
    // crashes or corrupts Excel. VB_Base alone is not sufficient for detection:
    // Excel-generated ordinary class modules can have it too. Designer/document
    // modules additionally have a predeclared instance, so require both attributes.
    for (const k of newModuleNames) {
        const origName = sourceFileNames.get(k)!;
        const newSource = sourceMap.get(k)!;
        if (/^\s*Attribute\s+VB_Base\s*=/im.test(newSource)
            && /^\s*Attribute\s+VB_PredeclaredId\s*=\s*True\s*$/im.test(newSource)) {
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
            // Document modules are owned by the workbook/sheet host.  Omitting
            // their exported source must not remove the OLE stream while the
            // PROJECT Document= declaration remains.  Keep the attributes and
            // clear only the procedure body so Excel can recompile safely.
            if (hostNames.has(mod.name.toLowerCase())) {
                const entry = CFB.find(cfb, `/VBA/${mod.streamName}`);
                if (!entry) {
                    console.warn(`  [warn] ${mod.name}: host stream not found`);
                    continue;
                }
                const original = iconv.decode(
                    decompress(Buffer.from(entry.content as unknown as ArrayBuffer).subarray(mod.offset)), encoding,
                );
                const attributes = original.match(/^Attribute\s+[^\r\n]+/gim) ?? [];
                const minimalSource = attributes.length > 0
                    ? `${attributes.join('\r\n')}\r\n`
                    : `Attribute VB_Name = "${mod.name}"\r\n`;
                const compressed = compress(iconv.encode(minimalSource, encoding));
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (entry as any).content = compressed;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                (entry as any).size = compressed.length;
                updatedModules.push({ ...mod, offset: 0 });
                console.log(`  ~ ${mod.name}: preserved host module (cleared source body)`);
                updated++;
                continue;
            }
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
    let completedClassAttributeCount = 0;
    for (const k of newModuleNames) {
        const origName = sourceFileNames.get(k)!;

        // Determine if class module by checking file extension
        const isClass = sourceClasses.get(k) ?? false;
        const source = sourceMap.get(k)!;
        const src = isClass ? completeNewClassAttributes(source) : source;
        if (isClass && src !== source) completedClassAttributeCount++;
        const srcBytes = iconv.encode(src, encoding);
        const compressed = compress(srcBytes);

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
    if (completedClassAttributeCount > 0) {
        console.log(`  (added Excel class attributes to ${completedClassAttributeCount} new .cls file(s))`);
    }

    if (updated === 0 && deleted === 0) {
        console.error('Error: no modules were updated or removed.');
        process.exit(1);
    }

    // Patch PROJECT stream (module list + workspace entries)
    if (projectEntry) {
        const deletedLower = new Set(deletedNames.map(n => n.toLowerCase()));
        const addedMods = newModuleNames.map(k => {
            const f = readdirSync(absSrc).find(file => basename(file, extname(file)).toLowerCase() === k);
            return { name: sourceFileNames.get(k)!, isClass: f ? extname(f).toLowerCase() === '.cls' : false };
        });
        const newText = patchProjectStream(projectText, deletedLower, addedMods);
        const newBytes = iconv.encode(newText, encoding);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projectEntry as any).content = newBytes;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (projectEntry as any).size = newBytes.length;
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
