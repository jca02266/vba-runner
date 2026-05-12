import * as path from 'path';

export interface FileSystem {
    existsSync(path: string): boolean;
    readFileSync(path: string, encoding: 'utf-8' | 'utf8'): string;
    writeFileSync(path: string, content: string): void;
    appendFileSync?(path: string, content: string): void;
    mkdirSync(path: string, options?: { recursive?: boolean }): void;
    rmdirSync?(path: string): void;
    rmSync?(path: string, options?: { recursive?: boolean, force?: boolean }): void;
    copyFileSync?(src: string, dest: string): void;
    unlinkSync(path: string): void;
    readdirSync(path: string): string[];
    statSync(path: string): {
        size: number;
        isFile(): boolean;
        isDirectory(): boolean;
        mtime: Date;
        birthtime?: Date;
        mode?: number;
    };
    openSync(path: string, flags: string): number;
    closeSync(fd: number): void;
    readSync(fd: number, buffer: Uint8Array, offset: number, length: number, position: number | null): number;
    writeSync(fd: number, buffer: string | Uint8Array, offset?: number, length?: number, position?: number | null): number;
}

/**
 * Memory-based file system for browser environments.
 */
export class MemoryFileSystem implements FileSystem {
    private files: Map<string, { data: Uint8Array | string, birthtime: Date, mtime: Date }> = new Map();
    private dirs: Map<string, { birthtime: Date, mtime: Date }> = new Map([['/', { birthtime: new Date(), mtime: new Date() }]]);
    private fileHandles: Map<number, { path: string, flags: string, pos: number }> = new Map();
    private nextFd = 1;

    existsSync(p: string): boolean {
        const norm = this.normalize(p);
        return this.files.has(norm) || this.dirs.has(norm);
    }

    readFileSync(p: string, _encoding: 'utf-8' | 'utf8'): string {
        const norm = this.normalize(p);
        const entry = this.files.get(norm);
        if (entry === undefined) throw new Error(`File not found: ${p}`);
        const data = entry.data;
        if (typeof data === 'string') return data;
        return new TextDecoder().decode(data);
    }

    writeFileSync(p: string, content: string | Uint8Array): void {
        const norm = this.normalize(p);
        this.ensureDir(path.dirname(norm));
        const now = new Date();
        const existing = this.files.get(norm);
        this.files.set(norm, { 
            data: content, 
            birthtime: existing ? existing.birthtime : now,
            mtime: now
        });
    }

    mkdirSync(p: string, options?: { recursive?: boolean }): void {
        const norm = this.normalize(p);
        const now = new Date();
        if (options?.recursive) {
            const parts = norm.split('/').filter(Boolean);
            let current = '';
            for (const part of parts) {
                current += '/' + part;
                if (!this.dirs.has(current)) {
                    this.dirs.set(current, { birthtime: now, mtime: now });
                }
            }
        } else {
            this.dirs.set(norm, { birthtime: now, mtime: now });
        }
    }

    unlinkSync(p: string): void {
        const norm = this.normalize(p);
        if (!this.files.has(norm)) {
            throw new Error(`ENOENT: no such file or directory, unlink '${p}'`);
        }
        this.files.delete(norm);
    }

    copyFileSync(src: string, dest: string): void {
        const s = this.normalize(src);
        const d = this.normalize(dest);
        const entry = this.files.get(s);
        if (!entry) throw new Error(`ENOENT: no such file or directory, copyFileSync '${src}'`);
        this.writeFileSync(d, entry.data);
    }

    rmSync(p: string, options?: { recursive?: boolean, force?: boolean }): void {
        const norm = this.normalize(p);
        if (!this.existsSync(norm)) {
            if (options?.force) return;
            throw new Error(`ENOENT: no such file or directory, rm '${p}'`);
        }
        if (this.files.has(norm)) {
            this.files.delete(norm);
        } else if (this.dirs.has(norm)) {
            if (options?.recursive) {
                const prefix = norm === '/' ? '/' : norm + '/';
                for (const f of Array.from(this.files.keys())) {
                    if (f.startsWith(prefix)) this.files.delete(f);
                }
                for (const d of Array.from(this.dirs.keys())) {
                    if (d.startsWith(prefix)) this.dirs.delete(d);
                }
                this.dirs.delete(norm);
            } else {
                if (this.readdirSync(norm).length > 0) throw new Error(`Directory not empty: ${p}`);
                this.dirs.delete(norm);
            }
        }
    }

    readdirSync(p: string): string[] {
        const norm = this.normalize(p);
        const prefix = norm === '/' ? '/' : norm + '/';
        const results = new Set<string>();
        for (const f of this.files.keys()) {
            if (f.startsWith(prefix)) {
                const rest = f.substring(prefix.length);
                const nextSlash = rest.indexOf('/');
                results.add(nextSlash === -1 ? rest : rest.substring(0, nextSlash));
            }
        }
        for (const d of this.dirs.keys()) {
            if (d !== norm && d.startsWith(prefix)) {
                const rest = d.substring(prefix.length);
                const nextSlash = rest.indexOf('/');
                results.add(nextSlash === -1 ? rest : rest.substring(0, nextSlash));
            }
        }
        return Array.from(results);
    }

    statSync(p: string) {
        const norm = this.normalize(p);
        const fileEntry = this.files.get(norm);
        const dirEntry = this.dirs.get(norm);
        if (!fileEntry && !dirEntry) throw new Error(`Not found: ${p}`);
        
        const entry = fileEntry || dirEntry!;
        const data = fileEntry?.data;
        const size = typeof data === 'string' ? new TextEncoder().encode(data).length : (data?.length || 0);

        return {
            size,
            isFile: () => !!fileEntry,
            isDirectory: () => !!dirEntry,
            mtime: entry.mtime,
            birthtime: entry.birthtime,
            mode: !!dirEntry ? 0o755 : 0o644
        };
    }

    openSync(p: string, flags: string): number {
        const norm = this.normalize(p);
        if (flags === 'r' && !this.existsSync(norm)) throw new Error(`File not found: ${p}`);
        if (flags === 'w' || flags === 'a') {
            if (!this.files.has(norm)) this.writeFileSync(norm, "");
        }
        const fd = this.nextFd++;
        this.fileHandles.set(fd, { path: norm, flags, pos: 0 });
        return fd;
    }

    closeSync(fd: number): void {
        this.fileHandles.delete(fd);
    }

    readSync(fd: number, buffer: Uint8Array, offset: number, length: number, position: number | null): number {
        const h = this.fileHandles.get(fd);
        if (!h) throw new Error("Invalid FD");
        const entry = this.files.get(h.path);
        if (!entry) return 0;
        
        const bin = typeof entry.data === 'string' ? new TextEncoder().encode(entry.data) : entry.data;
        const start = position !== null ? position : h.pos;
        const available = bin.length - start;
        const toRead = Math.max(0, Math.min(length, available));
        
        if (toRead <= 0) return 0;
        buffer.set(bin.subarray(start, start + toRead), offset);
        if (position === null) h.pos += toRead;
        return toRead;
    }

    writeSync(fd: number, buffer: string | Uint8Array, offset?: number, length?: number, position?: number | null): number {
        const h = this.fileHandles.get(fd);
        if (!h) throw new Error("Invalid FD");
        
        const newData = typeof buffer === 'string' ? new TextEncoder().encode(buffer) : buffer.subarray(offset || 0, (offset || 0) + (length || buffer.length));
        const entry = this.files.get(h.path);
        const oldBin = entry ? (typeof entry.data === 'string' ? new TextEncoder().encode(entry.data) : entry.data) : new Uint8Array(0);
        
        const start = position !== null && position !== undefined ? position : h.pos;
        const totalSize = Math.max(oldBin.length, start + newData.length);
        const combined = new Uint8Array(totalSize);
        combined.set(oldBin);
        combined.set(newData, start);
        
        const now = new Date();
        this.files.set(h.path, {
            data: combined,
            birthtime: entry ? entry.birthtime : now,
            mtime: now
        });
        if (position === null) h.pos += newData.length;
        return newData.length;
    }

    private normalize(p: string): string {
        return p.replace(/\\/g, '/').replace(/\/+$/, '') || '/';
    }

    private ensureDir(p: string) {
        this.mkdirSync(p, { recursive: true });
    }
}
