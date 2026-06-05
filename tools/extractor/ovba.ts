/**
 * MS-OVBA compression / decompression
 * Spec: [MS-OVBA] 2.4
 */

export function decompress(src: Buffer): Buffer {
    if (!Buffer.isBuffer(src)) src = Buffer.from(src as unknown as ArrayBuffer);
    if (src.length < 1 || src[0] !== 0x01)
        throw new Error(`Invalid OVBA signature: 0x${src[0]?.toString(16)}`);

    const out: number[] = [];
    let i = 1;

    while (i < src.length) {
        if (i + 1 >= src.length) break;
        const hdr = src.readUInt16LE(i);
        i += 2;

        const isCompressed = (hdr & 0x8000) !== 0;

        if (!isCompressed) {
            for (let j = 0; j < 4096 && i < src.length; j++) out.push(src[i++]);
        } else {
            const dataLen = (hdr & 0x0FFF) + 1;
            const chunkEnd = i + dataLen;
            const chunkBase = out.length;

            while (i < chunkEnd && i < src.length) {
                const flags = src[i++];
                for (let bit = 0; bit < 8; bit++) {
                    if (i >= chunkEnd || i >= src.length) break;

                    if ((flags >> bit) & 1) {
                        if (i + 1 >= src.length) break;
                        const token = src.readUInt16LE(i);
                        i += 2;

                        const lenBits = copyTokenLengthBits(out.length - chunkBase);
                        const lenMask = (1 << lenBits) - 1;
                        const length = (token & lenMask) + 3;
                        const offset = (token >> lenBits) + 1;

                        for (let j = 0; j < length; j++) {
                            out.push(out[out.length - offset]);
                        }
                    } else {
                        out.push(src[i++]);
                    }
                }
            }
            i = chunkEnd;
        }
    }

    return Buffer.from(out);
}

export function compress(src: Buffer): Buffer {
    if (!Buffer.isBuffer(src)) src = Buffer.from(src as unknown as ArrayBuffer);

    const parts: Buffer[] = [Buffer.from([0x01])];

    for (let base = 0; base < src.length; base += 4096) {
        const chunk = src.subarray(base, base + 4096);
        const compressed = compressChunk(chunk);
        const isLastChunk = base + 4096 >= src.length;

        if (!isLastChunk && compressed.length >= chunk.length) {
            const hdr = Buffer.alloc(2);
            hdr.writeUInt16LE(0x0FFF, 0);
            const raw = Buffer.alloc(4096);
            chunk.copy(raw);
            parts.push(hdr, raw);
        } else {
            const hdr = Buffer.alloc(2);
            hdr.writeUInt16LE(((compressed.length - 1) & 0x0FFF) | 0x8000, 0);
            parts.push(hdr, compressed);
        }
    }

    return Buffer.concat(parts);
}

function copyTokenLengthBits(decompressedChunkSize: number): number {
    if (decompressedChunkSize <= 16)   return 12;
    if (decompressedChunkSize <= 32)   return 11;
    if (decompressedChunkSize <= 64)   return 10;
    if (decompressedChunkSize <= 128)  return 9;
    if (decompressedChunkSize <= 256)  return 8;
    if (decompressedChunkSize <= 512)  return 7;
    if (decompressedChunkSize <= 1024) return 6;
    if (decompressedChunkSize <= 2048) return 5;
    return 4;
}

function compressChunk(src: Buffer): Buffer {
    const out: number[] = [];
    const window: number[] = [];
    let i = 0;

    while (i < src.length) {
        const flagPos = out.length;
        out.push(0);
        let flags = 0;

        for (let bit = 0; bit < 8 && i < src.length; bit++) {
            const lenBits = copyTokenLengthBits(window.length);
            const maxLen = (1 << lenBits) + 2;
            const maxOffset = 1 << (16 - lenBits);

            let bestLen = 0, bestOffset = 1;
            const winStart = Math.max(0, window.length - maxOffset);

            for (let j = winStart; j < window.length; j++) {
                let len = 0;
                while (len < maxLen && (i + len) < src.length &&
                       window[j + len] === src[i + len]) {
                    len++;
                }
                if (len > bestLen) { bestLen = len; bestOffset = window.length - j; }
            }

            if (bestLen >= 3) {
                flags |= (1 << bit);
                const lenBitsNow = copyTokenLengthBits(window.length);
                const token = ((bestOffset - 1) << lenBitsNow) | (bestLen - 3);
                out.push(token & 0xFF, (token >> 8) & 0xFF);
                for (let j = 0; j < bestLen; j++) window.push(src[i++]);
            } else {
                out.push(src[i]);
                window.push(src[i++]);
            }
        }

        out[flagPos] = flags;
    }

    return Buffer.from(out);
}
