import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Shell tests...");

let output = '';
evalVBASingle('Debug.Print Shell("echo hello") > 0', { onPrint: (o) => { output = o; } });

assert.strictEqual(output.trim(), 'True', 'Shell returns True');
console.log('✅ Shell: テスト通過');
