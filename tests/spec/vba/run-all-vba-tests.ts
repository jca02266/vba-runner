/**
 * VBA Test Runner - すべての _Test.{cls,vba} ファイルを自動実行
 * tests/spec/vba/ ディレクトリ内の *_Test.cls および *_Test.vba ファイルを
 * 自動検出し、VBATest ランナーで実行する汎用テストランナー
 */

import * as fs from 'fs';
import * as path from 'path';
import { VBATest } from '../../../test-libs/test-runner';

const vbaDir = path.join(__dirname);

// VBA テストディレクトリ内のすべてのファイルをロード
console.log('=== VBA Test Runner ===\n');
console.log(`Loading VBA test files from: ${vbaDir}\n`);

let fileCount = 0;
let testCount = 0;
let passCount = 0;
const failedTests: string[] = [];

// ディレクトリ内のファイルをリスト
const files = fs.readdirSync(vbaDir);
const testFiles = files.filter(f => f.endsWith('_Test.cls') || f.endsWith('_Test.vba'));

console.log(`Found ${testFiles.length} test file(s):`);
testFiles.forEach(f => console.log(`  - ${f}`));
console.log('');

if (testFiles.length === 0) {
  console.log('No test files found.\n');
  process.exit(0);
}

// VBATest にすべてのファイルをロード
const vbaTest = new VBATest(vbaDir);
fileCount = testFiles.length;

// テスト関数を探して実行
// VBA での Sub テスト（パラメータなし、戻り値なし）を識別して実行
const evaluator = (vbaTest as any).evaluator;
const env = evaluator.env;

// プロシージャマップから Test_ で始まる Sub を検出
const procedures = new Map();

// 環境内のすべてのプロシージャを取得
function collectProcedures(environment: any): Map<string, any> {
  const procs = new Map();

  if (environment && environment.procedures) {
    for (const [name, proc] of environment.procedures.entries()) {
      if (typeof name === 'string') {
        procs.set(name, proc);
      }
    }
  }

  // 親環境も走査
  if (environment && environment.enclosing) {
    const parentProcs = collectProcedures(environment.enclosing);
    for (const [name, proc] of parentProcs) {
      if (!procs.has(name)) {
        procs.set(name, proc);
      }
    }
  }

  return procs;
}

const allProcs = collectProcedures(env);

// Test_ または test_ で始まるプロシージャを実行
for (const [procName] of allProcs) {
  if (typeof procName === 'string') {
    const lower = procName.toLowerCase();
    // Module-qualified names look like "module:procedure"
    // Extract the procedure name after the colon, or use the whole name if no colon
    const baseProcName = lower.includes(':') ? lower.split(':')[1] : lower;
    const moduleName = lower.includes(':') ? lower.split(':')[0] : '';

    // VBA での Sub テスト（Test_* で始まる）
    if (baseProcName.startsWith('test_')) {
      testCount++;
      console.log(`[Test] ${procName}`);

      try {
        // Call SetUp if it exists in the same module
        if (moduleName && allProcs.has(`${moduleName}:setup`)) {
          vbaTest.run(`${moduleName}:setup`, []);
        }

        // Sub テストを実行（戻り値なし）
        vbaTest.run(procName, []);
        passCount++;
        console.log(`  ✅ Pass\n`);

        // Call TearDown if it exists in the same module
        if (moduleName && allProcs.has(`${moduleName}:teardown`)) {
          vbaTest.run(`${moduleName}:teardown`, []);
        }
      } catch (e: any) {
        console.error(`  ❌ FAILED: ${e.message}\n`);
        failedTests.push(procName);
      }
    }
  }
}

// 結果をサマリー
console.log('=== Summary ===');
console.log(`Files loaded: ${fileCount}`);
console.log(`Tests run: ${testCount}`);
console.log(`Tests passed: ${passCount}`);

if (failedTests.length > 0) {
  console.log(`Tests failed: ${failedTests.length}`);
  console.log('\nFailed tests:');
  failedTests.forEach(name => console.log(`  - ${name}`));
  process.exit(1);
} else if (testCount > 0) {
  console.log(`\n✅ All ${testCount} test(s) passed!`);
  process.exit(0);
} else {
  console.log('\n⚠️  No tests found to run');
  process.exit(0);
}
