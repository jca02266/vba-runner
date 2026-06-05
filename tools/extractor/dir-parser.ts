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
                case 0x001A:
                    streamName = readStr(sz);
                    if (i + 6 <= buf.length && buf.readUInt16LE(i) === 0x0032) {
                        i += 2;
                        const usz = readU32();
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
