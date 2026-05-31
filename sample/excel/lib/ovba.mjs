/**
 * MS-OVBA compression / decompression
 * Spec: [MS-OVBA] 2.4
 */

/**
 * Decompress a CompressedContainer buffer.
 * @param {Buffer} src
 * @returns {Buffer}
 */
export function decompress(src) {
    if (!Buffer.isBuffer(src)) src = Buffer.from(src);
    if (src.length < 1 || src[0] !== 0x01)
        throw new Error(`Invalid OVBA signature: 0x${src[0]?.toString(16)}`);

    const out = [];
    let i = 1;

    while (i < src.length) {
        if (i + 1 >= src.length) break;
        const hdr = src.readUInt16LE(i);
        i += 2;

        const isCompressed = (hdr & 0x8000) !== 0;

        if (!isCompressed) {
            // Raw chunk: always exactly 4096 bytes
            for (let j = 0; j < 4096 && i < src.length; j++) out.push(src[i++]);
        } else {
            // CompressedChunkSize (bits 0-11) = total chunk byte count (2-byte
            // header + data) minus 3. Therefore data length = size + 3 - 2 = size + 1.
            const dataLen = (hdr & 0x0FFF) + 1;
            const chunkEnd = i + dataLen;
            const chunkBase = out.length;

            while (i < chunkEnd && i < src.length) {
                const flags = src[i++];
                for (let bit = 0; bit < 8; bit++) {
                    if (i >= chunkEnd || i >= src.length) break;

                    if ((flags >> bit) & 1) {
                        // CopyToken
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
                        // LiteralToken
                        out.push(src[i++]);
                    }
                }
            }
            i = chunkEnd; // ensure we advance past any padding
        }
    }

    return Buffer.from(out);
}

/**
 * Compress a buffer into a CompressedContainer.
 * @param {Buffer} src
 * @returns {Buffer}
 */
export function compress(src) {
    if (!Buffer.isBuffer(src)) src = Buffer.from(src);

    const parts = [Buffer.from([0x01])]; // SignatureByte

    for (let base = 0; base < src.length; base += 4096) {
        const chunk = src.slice(base, base + 4096);
        const compressed = compressChunk(chunk);
        const isLastChunk = base + 4096 >= src.length;

        // Use raw only for full (4096-byte) non-last chunks when compression doesn't help.
        // For the last (partial) chunk always use compression: a raw last chunk would
        // require padding to 4096 bytes, adding unwanted trailing nulls to the source.
        if (!isLastChunk && compressed.length >= chunk.length) {
            // Raw chunk (exactly 4096 bytes)
            const hdr = Buffer.alloc(2);
            hdr.writeUInt16LE(0x0FFF, 0);
            const raw = Buffer.alloc(4096);
            chunk.copy(raw);
            parts.push(hdr, raw);
        } else {
            // Compressed chunk.
            // size field = total chunk size (2-byte header + data) - 3 = data length - 1
            const hdr = Buffer.alloc(2);
            hdr.writeUInt16LE(((compressed.length - 1) & 0x0FFF) | 0x8000, 0);
            parts.push(hdr, compressed);
        }
    }

    return Buffer.concat(parts);
}

// --- helpers -----------------------------------------------------------------

/** Number of bits used for the length field in a CopyToken, given current chunk output size. */
function copyTokenLengthBits(decompressedChunkSize) {
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

/** Compress a single chunk (up to 4096 bytes). Returns CompressedChunkData bytes. */
function compressChunk(src) {
    const out = [];
    const window = []; // decompressed bytes output so far (within this chunk)
    let i = 0;

    while (i < src.length) {
        const flagPos = out.length;
        out.push(0); // placeholder for FlagByte
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
                const lenMask = (1 << lenBitsNow) - 1;
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
