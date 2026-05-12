import * as path from 'path';

/**
 * VBA File System Sandbox Path Resolver
 * Handles conversion between VBA virtual paths (e.g. C:\foo) and real OS paths.
 */
export class SandboxPath {
    private root: string;
    private env: Map<string, string> = new Map();

    constructor(root: string = '/sandbox', initialEnv: Record<string, string> = {}) {
        // In browser environment, we use a virtual root.
        // We ensure it's absolute within the virtual space.
        this.root = root.startsWith('/') ? root : '/' + root;
        
        // Initialize environment variables from memory
        for (const [k, v] of Object.entries(initialEnv)) {
            this.env.set(k.toLowerCase(), v);
        }
    }

    /**
     * Converts a VBA virtual path to a real OS path within the sandbox.
     * @param vbaPath VBA path (absolute C:\... or relative)
     * @returns Real absolute OS path
     */
    public toRealPath(vbaPath: string): string {
        if (!vbaPath) return this.root;

        // 1. Normalize separators
        let normalized = vbaPath.replace(/\\/g, '/');

        // 2. Handle Windows absolute paths: C:\foo -> /c/foo
        if (normalized.match(/^[a-zA-Z]:/)) {
            const drive = normalized[0].toLowerCase();
            normalized = '/' + drive + normalized.substring(2);
        }

        // 3. Resolve path relative to root
        // Remove leading slash if it exists after drive conversion to ensure it's relative to root
        const relativePart = normalized.startsWith('/') ? normalized.substring(1) : normalized;
        
        // Using path.join instead of path.resolve to avoid process.cwd dependency
        const resolved = path.join(this.root, relativePart);

        // 4. Traversal check
        if (!resolved.startsWith(this.root)) {
            throw new Error(`Execution error: Access denied outside of sandbox root (${vbaPath})`);
        }

        return resolved;
    }

    /**
     * Converts a real OS path to a VBA virtual path.
     * @param realPath Real absolute OS path
     * @returns VBA path (Windows format if it maps to a drive)
     */
    public toVirtualPath(realPath: string): string {
        const relative = path.relative(this.root, realPath);
        
        if (relative.startsWith('..')) {
            throw new Error(`Security Error: Real path is outside of sandbox root (${realPath})`);
        }

        if (relative === '' || relative === '.') return '\\';

        const parts = relative.split(/[/\\]/);
        // If first part is a single letter, treat as drive letter: c/foo -> C:\foo
        if (parts[0].length === 1 && /^[a-z]$/i.test(parts[0])) {
            const drive = parts[0].toUpperCase();
            const rest = parts.slice(1).join('\\');
            return drive + ':\\' + rest;
        }

        return relative.replace(/\//g, '\\');
    }

    public getEnv(key: string | number): string {
        if (typeof key === 'string') {
            return this.env.get(key.toLowerCase()) || "";
        } else if (typeof key === 'number') {
            // Index-based access (1-based)
            const entries = Array.from(this.env.entries());
            const idx = Math.floor(key) - 1;
            if (idx >= 0 && idx < entries.length) {
                const [k, v] = entries[idx];
                return `${k.toUpperCase()}=${v}`;
            }
            return "";
        }
        return "";
    }

    public getRoot(): string {
        return this.root;
    }
}
