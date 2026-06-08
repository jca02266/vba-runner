import { evalVBASingle, assert } from '../../test-libs/test-runner';
import * as fs from 'fs';

console.log("Running MkDir/RmDir tests...");

const testDir = "vba_test_dir";
if (fs.existsSync(testDir)) fs.rmdirSync(testDir);

const code = `
    MkDir "vba_test_dir"
    Debug.Print Dir("vba_test_dir")
    RmDir "vba_test_dir"
`;

let output = '';
evalVBASingle(code, { onPrint: (o) => { output = o; } });

assert.strictEqual(output.trim(), "vba_test_dir", 'Dir should return directory name');
assert.ok(!fs.existsSync(testDir), 'Directory should be removed after RmDir');
console.log('✅ MkDir/RmDir: 全テスト通過');
