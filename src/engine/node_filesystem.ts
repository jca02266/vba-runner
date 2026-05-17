import * as fs from 'fs';
import { FileSystem } from './filesystem';

/**
 * Node.js based file system.
 */
export class NodeFileSystem implements FileSystem {
    existsSync(p: string) { return fs.existsSync(p); }
    readFileSync(p: string, encoding: 'utf-8' | 'utf8') { return fs.readFileSync(p, encoding); }
    writeFileSync(p: string, content: string) { fs.writeFileSync(p, content); }
    mkdirSync(p: string, options?: { recursive?: boolean }) { fs.mkdirSync(p, options); }
    rmdirSync(p: string) { fs.rmdirSync(p); }
    rmSync(p: string, options?: { recursive?: boolean, force?: boolean }) { fs.rmSync(p, options); }
    copyFileSync(src: string, dest: string) { fs.copyFileSync(src, dest); }
    unlinkSync(p: string) { fs.unlinkSync(p); }
    readdirSync(p: string) { return fs.readdirSync(p); }
    statSync(p: string) {
        const s = fs.statSync(p);
        return {
            size: s.size,
            isFile: () => s.isFile(),
            isDirectory: () => s.isDirectory(),
            mtime: s.mtime
        };
    }
    openSync(p: string, flags: string) { return fs.openSync(p, flags); }
    closeSync(fd: number) { fs.closeSync(fd); }
    readSync(fd: number, buffer: Uint8Array, offset: number, length: number, position: number | null) {
        return fs.readSync(fd, buffer, offset, length, position);
    }
    writeSync(fd: number, buffer: string | Uint8Array, offset?: number, length?: number, position?: number | null) {
        // @ts-ignore
        return fs.writeSync(fd, buffer, offset, length, position);
    }
}
