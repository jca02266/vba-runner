import { assert, VBARunner } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const shared = new MemoryFileSystem();
const first = new VBARunner(null, { fs: shared });
const second = new VBARunner(null, { fs: shared });

first.fs.writeFileSync('/sandbox/shared.txt', 'shared state');
assert.strictEqual(second.fs.readFileSync('/sandbox/shared.txt', 'utf8'), 'shared state',
    '複数 VBARunner が同一 MemoryFileSystem を共有する');

second.fs.writeFileSync('/sandbox/second.txt', 'from second');
assert.strictEqual(first.fs.readFileSync('/sandbox/second.txt', 'utf8'), 'from second',
    '共有VFSへの後続書き込みも可視になる');

console.log('[PASS] VBARunner shared MemoryFileSystem');
