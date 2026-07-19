/**
 * VBA Project dir stream builder  [MS-OVBA] 2.3.4
 * Reconstructs a dir stream binary from parsed data + updated module list.
 */

import iconv from 'iconv-lite';
import { DirStreamData, VbaModuleFull } from './dir-parser.js';

function writeU16(buf: number[], v: number) {
    buf.push(v & 0xFF, (v >> 8) & 0xFF);
}

function writeU32(buf: number[], v: number) {
    buf.push(v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF);
}

function writeRecord(buf: number[], id: number, data: number[]) {
    writeU16(buf, id);
    writeU32(buf, data.length);
    for (const b of data) buf.push(b);
}

function toBytes(s: string, enc: string): number[] {
    return [...iconv.encode(s, enc)];
}

function toUtf16le(s: string): number[] {
    return [...Buffer.from(s, 'utf16le')];
}

/**
 * Patch MODULEOFFSET (id=0x0031) value inside a rawBlock to newOffset.
 * Returns a new Buffer with the value replaced, or the original if not found.
 */
export function patchModuleOffset(rawBlock: Buffer, newOffset: number): Buffer {
    // Pattern: [0x31, 0x00,  0x04, 0x00, 0x00, 0x00,  <4 bytes value>]
    for (let i = 0; i + 10 <= rawBlock.length; i++) {
        if (rawBlock[i] === 0x31 && rawBlock[i + 1] === 0x00 &&
            rawBlock[i + 2] === 0x04 && rawBlock[i + 3] === 0x00 &&
            rawBlock[i + 4] === 0x00 && rawBlock[i + 5] === 0x00) {
            const out = Buffer.from(rawBlock);
            out.writeUInt32LE(newOffset, i + 6);
            return out;
        }
    }
    return rawBlock;
}

/** Build a minimal module block for a brand-new module (no rawBlock). */
function buildNewModuleBlock(name: string, streamName: string, isClass: boolean, encoding: string): Buffer {
    const buf: number[] = [];

    // Record order must match a real module block exactly (see MS-OVBA 2.3.4.2):
    //   MODULENAME → MODULENAMEUNICODE → MODULESTREAMNAME (+Unicode reserved) →
    //   MODULEDOCSTRING (+Unicode reserved) → MODULEOFFSET → MODULEHELPCONTEXT →
    //   MODULECOOKIE → MODULETYPE → MODULETERMINATOR.
    // MODULENAME (0x0019): codepage name
    writeRecord(buf, 0x0019, toBytes(name, encoding));
    // MODULENAMEUNICODE (0x0047)
    writeRecord(buf, 0x0047, toUtf16le(name));
    // MODULESTREAMNAME (0x001A): codepage
    writeRecord(buf, 0x001A, toBytes(streamName, encoding));
    // MODULESTREAMNAME Unicode reserved (0x0032)
    writeRecord(buf, 0x0032, toUtf16le(streamName));
    // MODULEDOCSTRING (0x001C): empty
    writeRecord(buf, 0x001C, []);
    // MODULEDOCSTRING Unicode reserved (0x0048): empty
    writeRecord(buf, 0x0048, []);
    // MODULEOFFSET (0x0031): 0 = no preamble
    writeRecord(buf, 0x0031, [0, 0, 0, 0]);
    // MODULEHELPCONTEXT (0x001E): 0
    writeRecord(buf, 0x001E, [0, 0, 0, 0]);
    // MODULECOOKIE (0x002C): 0xFFFF (MUST be 0xFFFF on write)
    writeRecord(buf, 0x002C, [0xFF, 0xFF]);
    // MODULETYPE: 0x0021 = procedural, 0x0022 = class/document
    writeRecord(buf, isClass ? 0x0022 : 0x0021, []);
    // MODULEPRIVATE is required for a class module. Excel's VBE writes it after
    // MODULETYPE even when it has no payload; omitting it produces a dir stream
    // that opens but fails VBE validation when the workbook is saved.
    if (isClass) writeRecord(buf, 0x0028, []);
    // MODULETERMINATOR (0x002B)
    writeRecord(buf, 0x002B, []);

    return Buffer.from(buf);
}

/**
 * Rebuild the uncompressed dir stream binary.
 *
 * @param data      Original parsed dir stream data (projectRaw + cookie)
 * @param modules   Updated module list; empty rawBlock → synthesize new block
 * @param encoding  Codepage encoding string (e.g. 'cp932')
 */
export function buildDirStream(data: DirStreamData, modules: VbaModuleFull[], encoding: string): Buffer {
    const out: number[] = [];

    // Project-level records verbatim
    for (const b of data.projectRaw) out.push(b);

    // PROJECTMODULES (0x000F): 2-byte count
    writeRecord(out, 0x000F, [modules.length & 0xFF, (modules.length >> 8) & 0xFF]);

    // PROJECTCOOKIE (0x0013): 2-byte value
    writeRecord(out, 0x0013, [data.cookie & 0xFF, (data.cookie >> 8) & 0xFF]);

    // Module blocks
    for (const mod of modules) {
        let block: Buffer;
        if (mod.rawBlock.length === 0) {
            block = buildNewModuleBlock(mod.name, mod.streamName, mod.isClass, encoding);
        } else {
            block = patchModuleOffset(mod.rawBlock, mod.offset);
        }
        for (const b of block) out.push(b);
    }

    // PROJECTTERMINATOR (0x0010)
    writeRecord(out, 0x0010, []);

    return Buffer.from(out);
}
