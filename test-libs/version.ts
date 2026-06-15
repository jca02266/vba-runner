import { readFileSync } from 'fs';
import { join } from 'path';

function readVersion(): string {
    // tsx from test-libs/: ../build/runner/package.json
    // built CJS in build/runner/dist/bin/: ../../package.json
    for (const rel of ['../build/runner/package.json', '../../package.json']) {
        try {
            return JSON.parse(readFileSync(join(import.meta.dirname, rel), 'utf8')).version;
        } catch { /* try next */ }
    }
    return 'unknown';
}

export const VERSION = readVersion();
