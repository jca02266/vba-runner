/**
 * VBA Project dir stream parser  [MS-OVBA] 2.3.4
 *
 * parseDirStream(buf) returns:
 *   { codePage: number, modules: Array<{ name, streamName, offset, isClass }> }
 */

export function parseDirStream(buf) {
    let i = 0;

    function readU16() { const v = buf.readUInt16LE(i); i += 2; return v; }
    function readU32() { const v = buf.readUInt32LE(i); i += 4; return v; }
    function readStr(n) { const v = buf.slice(i, i + n).toString('latin1'); i += n; return v; }
    function skip(n)  { i += n; }

    // ── Project-level records ──────────────────────────────────────────────
    // Scan until PROJECTMODULES (0x000F).
    // Special case: PROJECTVERSION (0x0009) has a MinorVersion WORD that lies
    // OUTSIDE the size field (size=4 covers only MajorVersion DWORD).
    let moduleCount = 0;
    let codePage = null; // filled from PROJECTCODEPAGE (0x0003)
    while (i < buf.length) {
        const id = readU16();
        const sz = readU32();

        if (id === 0x0003) {
            // PROJECTCODEPAGE: WORD code page number
            codePage = buf.readUInt16LE(i);
            skip(sz); // sz=2
        } else if (id === 0x0009) {
            // PROJECTVERSION: skip sz(=4) bytes + 2 extra (MinorVersion)
            skip(sz + 2);
        } else if (id === 0x000F) {
            // PROJECTMODULES: data is a WORD = module count
            moduleCount = buf.readUInt16LE(i);
            skip(sz); // sz=2
            break;
        } else {
            skip(sz);
        }
    }

    // Skip PROJECTCOOKIE (0x0013)
    if (buf.readUInt16LE(i) === 0x0013) {
        i += 2;
        const csz = readU32();
        skip(csz);
    }

    // ── Module records ────────────────────────────────────────────────────
    const modules = [];

    for (let m = 0; m < moduleCount; m++) {
        let name = '', streamName = '', offset = 0, isClass = false;

        while (i < buf.length) {
            const id = readU16();
            const sz = readU32();

            switch (id) {
                case 0x0019: // MODULENAME
                    name = readStr(sz);
                    break;

                case 0x001A: // MODULESTREAMNAMERECORD (ANSI)
                    streamName = readStr(sz);
                    // The Unicode variant (0x0032) always follows immediately
                    if (i + 6 <= buf.length && buf.readUInt16LE(i) === 0x0032) {
                        i += 2;
                        const usz = readU32();
                        skip(usz);
                    }
                    break;

                case 0x0021: // MODULETYPE procedural
                    isClass = false;
                    // sz=0, no data
                    break;

                case 0x0022: // MODULETYPE class/document
                    isClass = true;
                    break;

                case 0x0031: // MODULEOFFSET
                    offset = buf.readUInt32LE(i);
                    skip(sz); // sz=4
                    break;

                case 0x002B: // MODULE_TERMINATOR (sz=0)
                    modules.push({ name, streamName: streamName || name, offset, isClass });
                    break;

                default:
                    skip(sz);
                    break;
            }

            if (id === 0x002B) break; // module done
        }
    }

    return { codePage, modules };
}
