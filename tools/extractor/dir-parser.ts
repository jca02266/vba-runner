/**
 * VBA Project dir stream parser  [MS-OVBA] 2.3.4
 */

export interface VbaModule {
    name: string;
    streamName: string;
    offset: number;
    isClass: boolean;
}

export interface DirStreamResult {
    codePage: number | null;
    modules: VbaModule[];
}

/** モジュール情報 + rawBlock（MODULENAME〜MODULETERMINATOR の raw bytes） */
export interface VbaModuleFull extends VbaModule {
    rawBlock: Buffer; // empty Buffer → 新規モジュール（dir-builder で合成）
}

/** parseDirStreamFull の返り値 */
export interface DirStreamData {
    codePage: number;
    cookie: number;     // PROJECTCOOKIE (0x0013) の値
    projectRaw: Buffer; // PROJECTMODULES (0x000F) より前のレコード群の raw bytes
    modules: VbaModuleFull[];
}

export function parseDirStream(buf: Buffer): DirStreamResult {
    let i = 0;

    function readU16() { const v = buf.readUInt16LE(i); i += 2; return v; }
    function readU32() { const v = buf.readUInt32LE(i); i += 4; return v; }
    function readStr(n: number) { const v = buf.subarray(i, i + n).toString('latin1'); i += n; return v; }
    function skip(n: number) { i += n; }

    let moduleCount = 0;
    let codePage: number | null = null;

    while (i < buf.length) {
        const id = readU16();
        const sz = readU32();

        if (id === 0x0003) {
            codePage = buf.readUInt16LE(i);
            skip(sz);
        } else if (id === 0x0009) {
            skip(sz + 2);
        } else if (id === 0x000F) {
            moduleCount = buf.readUInt16LE(i);
            skip(sz);
            break;
        } else {
            skip(sz);
        }
    }

    if (buf.readUInt16LE(i) === 0x0013) {
        i += 2;
        const csz = readU32();
        skip(csz);
    }

    const modules: VbaModule[] = [];

    for (let m = 0; m < moduleCount; m++) {
        let name = '', streamName = '', offset = 0, isClass = false;

        while (i < buf.length) {
            const id = readU16();
            const sz = readU32();

            switch (id) {
                case 0x0019:
                    name = readStr(sz);
                    break;
                case 0x0047:
                    // MODULENAMEUNICODE -- Unicode version of module name
                    name = buf.subarray(i, i + sz).toString('utf16le');
                    skip(sz);
                    break;
                case 0x001A:
                    skip(sz); // codepage version -- prefer Unicode below
                    if (i + 6 <= buf.length && buf.readUInt16LE(i) === 0x0032) {
                        i += 2;
                        const usz = readU32();
                        streamName = buf.subarray(i, i + usz).toString('utf16le');
                        skip(usz);
                    }
                    break;
                case 0x0021:
                    isClass = false;
                    break;
                case 0x0022:
                    isClass = true;
                    break;
                case 0x0031:
                    offset = buf.readUInt32LE(i);
                    skip(sz);
                    break;
                case 0x002B:
                    modules.push({ name, streamName: streamName || name, offset, isClass });
                    break;
                default:
                    skip(sz);
                    break;
            }

            if (id === 0x002B) break;
        }
    }

    return { codePage, modules };
}

export function parseDirStreamFull(buf: Buffer): DirStreamData {
    let i = 0;

    function readU16() { const v = buf.readUInt16LE(i); i += 2; return v; }
    function readU32() { const v = buf.readUInt32LE(i); i += 4; return v; }
    function readStr(n: number) { const v = buf.subarray(i, i + n).toString('latin1'); i += n; return v; }
    function skip(n: number) { i += n; }

    let moduleCount = 0;
    let codePage = 0;

    // Scan project-level records; capture raw bytes up to (but not including) PROJECTMODULES
    while (i < buf.length) {
        const recStart = i;
        const id = readU16();
        const sz = readU32();

        if (id === 0x0003) {
            codePage = buf.readUInt16LE(i);
            skip(sz);
        } else if (id === 0x0009) {
            // PROJECTVERSION: Size bytes + 2 extra MinorVersion bytes
            skip(sz + 2);
        } else if (id === 0x000F) {
            // PROJECTMODULES — projectRaw ends just before this record
            const projectRaw = buf.subarray(0, recStart);
            moduleCount = buf.readUInt16LE(i);
            skip(sz);

            // PROJECTCOOKIE must immediately follow
            let cookie = 0xFFFF;
            if (i + 6 <= buf.length && buf.readUInt16LE(i) === 0x0013) {
                i += 2;
                const csz = readU32();
                cookie = buf.readUInt16LE(i);
                skip(csz);
            }

            const modules: VbaModuleFull[] = [];

            for (let m = 0; m < moduleCount; m++) {
                let name = '', streamName = '', offset = 0, isClass = false;
                const blockStart = i;

                while (i < buf.length) {
                    const id2 = readU16();
                    const sz2 = readU32();

                    switch (id2) {
                        case 0x0019:
                            name = readStr(sz2);
                            break;
                        case 0x0047:
                            name = buf.subarray(i, i + sz2).toString('utf16le');
                            skip(sz2);
                            break;
                        case 0x001A:
                            skip(sz2);
                            if (i + 6 <= buf.length && buf.readUInt16LE(i) === 0x0032) {
                                i += 2;
                                const usz = readU32();
                                streamName = buf.subarray(i, i + usz).toString('utf16le');
                                skip(usz);
                            }
                            break;
                        case 0x0021:
                            isClass = false;
                            break;
                        case 0x0022:
                            isClass = true;
                            break;
                        case 0x0031:
                            offset = buf.readUInt32LE(i);
                            skip(sz2);
                            break;
                        case 0x002B:
                            // MODULETERMINATOR — no data
                            break;
                        default:
                            skip(sz2);
                            break;
                    }

                    if (id2 === 0x002B) break;
                }

                const rawBlock = buf.subarray(blockStart, i);
                modules.push({ name, streamName: streamName || name, offset, isClass, rawBlock: Buffer.from(rawBlock) });
            }

            return { codePage, cookie, projectRaw: Buffer.from(projectRaw), modules };
        } else {
            skip(sz);
        }
    }

    throw new Error('parseDirStreamFull: PROJECTMODULES (0x000F) not found');
}
