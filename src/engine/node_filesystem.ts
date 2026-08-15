import * as fs from 'fs';
import * as path from 'path';
import { FileSystem } from './filesystem';
import { VBA_FILE_ATTRIBUTE } from './filesystem';

/**
 * Node.js based file system.
 */
export class NodeFileSystem implements FileSystem {
    readonly hostBacked = true;
    private readonly attributeOverrides = new Map<string, number>();
    existsSync(p: string) { return fs.existsSync(p); }
    readFileSync(p: string, encoding: 'utf-8' | 'utf8') { return fs.readFileSync(p, encoding); }
    writeFileSync(p: string, content: string) { fs.writeFileSync(p, content); }
    mkdirSync(p: string, options?: { recursive?: boolean }) { fs.mkdirSync(p, options); }
    rmdirSync(p: string) { fs.rmdirSync(p); }
    rmSync(p: string, options?: { recursive?: boolean, force?: boolean }) { fs.rmSync(p, options); }
    copyFileSync(src: string, dest: string, options?: { overwrite?: boolean }) {
        fs.copyFileSync(src, dest, options?.overwrite ? 0 : fs.constants.COPYFILE_EXCL);
    }
    moveFileSync(src: string, dest: string) { fs.renameSync(src, dest); }
    copyDirectorySync(src: string, dest: string, options?: { overwrite?: boolean }) {
        const source = path.resolve(src);
        const target = path.resolve(dest);
        if (target === source || target.startsWith(`${source}${path.sep}`)) {
            throw new Error(`EINVAL: destination is inside source '${dest}'`);
        }
        if (!options?.overwrite) {
            this.assertNoDirectoryFileCollisions(source, target);
        }
        fs.cpSync(src, dest, { recursive: true, force: options?.overwrite === true, errorOnExist: options?.overwrite !== true });
    }

    /**
     * Match MemoryFileSystem's overwrite=false preflight. Node's cpSync can
     * copy earlier entries before reporting a later EEXIST, which would leave
     * the two FileSystem backends with different failure states.
     */
    private assertNoDirectoryFileCollisions(source: string, target: string): void {
        if (fs.existsSync(target) && !fs.statSync(target).isDirectory()) {
            throw new Error(`EEXIST: destination exists '${target}'`);
        }
        const visit = (current: string, relative: string): void => {
            for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
                const childRelative = path.join(relative, entry.name);
                const destination = path.join(target, childRelative);
                if (entry.isDirectory()) {
                    visit(path.join(current, entry.name), childRelative);
                } else if (fs.existsSync(destination)) {
                    throw new Error(`EEXIST: destination exists '${destination}'`);
                }
            }
        };
        visit(source, '');
    }
    moveDirectorySync(src: string, dest: string) {
        const source = path.resolve(src);
        const target = path.resolve(dest);
        if (target === source || target.startsWith(`${source}${path.sep}`)) {
            throw new Error(`EINVAL: destination is inside source '${dest}'`);
        }
        fs.renameSync(src, dest);
    }
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
    getAttributes(p: string): number {
        const normalized = path.resolve(p);
        const s = fs.statSync(normalized);
        const override = this.attributeOverrides.get(normalized);
        if (override !== undefined) {
            return s.isDirectory()
                ? override | VBA_FILE_ATTRIBUTE.DIRECTORY
                : override & ~VBA_FILE_ATTRIBUTE.DIRECTORY;
        }
        let attributes = s.isDirectory() ? VBA_FILE_ATTRIBUTE.DIRECTORY : VBA_FILE_ATTRIBUTE.NORMAL;
        if ((s.mode & 0o222) === 0) attributes |= VBA_FILE_ATTRIBUTE.READ_ONLY;
        if (path.basename(normalized).startsWith('.')) attributes |= VBA_FILE_ATTRIBUTE.HIDDEN;
        return attributes;
    }
    setAttributes(p: string, attributes: number): void {
        const normalized = path.resolve(p);
        const s = fs.statSync(normalized);
        if (!Number.isInteger(attributes) || attributes < 0 || attributes > 127) {
            throw new Error(`Invalid file attributes: ${attributes}`);
        }
        const normalizedAttributes = s.isDirectory()
            ? attributes | VBA_FILE_ATTRIBUTE.DIRECTORY
            : attributes & ~VBA_FILE_ATTRIBUTE.DIRECTORY;
        this.attributeOverrides.set(normalized, normalizedAttributes);
        // Read-only is the only VBA attribute with a portable POSIX analogue.
        const mode = s.mode & 0o777;
        fs.chmodSync(normalized, (normalizedAttributes & VBA_FILE_ATTRIBUTE.READ_ONLY)
            ? mode & ~0o222
            : mode | 0o200);
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
