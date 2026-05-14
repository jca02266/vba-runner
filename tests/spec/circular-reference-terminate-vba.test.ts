import * as path from 'path';
import { VBATest } from '../../test-libs/test-runner';

// Load all VBA files from the vba/ directory (includes .cls and .vba files)
// VBA test files use _Test suffix for grouping related test artifacts
const vbaDir = path.join(__dirname, 'vba');
const vbaTest = new VBATest(vbaDir);

// Test 1: Basic Class_Terminate invocation
{
    try {
        vbaTest.run('Test_BasicTerminate', []);
        console.log('[PASS] Test 1: Basic Class_Terminate invocation');
    } catch (e) {
        console.error('[FAIL] Test 1:', (e as any).message);
        throw e;
    }
}

// Test 2: Mutual references with Set = Nothing
{
    try {
        vbaTest.run('Test_MutualReferences', []);
        console.log('[PASS] Test 2: Mutual references with Set = Nothing');
    } catch (e) {
        console.error('[FAIL] Test 2:', (e as any).message);
        throw e;
    }
}

// Test 3: Class_Terminate executes at most once
{
    try {
        vbaTest.run('Test_TerminateNotCalledTwice', []);
        console.log('[PASS] Test 3: Class_Terminate executes at most once');
    } catch (e) {
        console.error('[FAIL] Test 3:', (e as any).message);
        throw e;
    }
}

// Test 4: Multiple independent objects
{
    try {
        vbaTest.run('Test_MultipleObjects', []);
        console.log('[PASS] Test 4: Multiple independent objects');
    } catch (e) {
        console.error('[FAIL] Test 4:', (e as any).message);
        throw e;
    }
}

// Test 5: Circular chain cleanup
{
    try {
        vbaTest.run('Test_CircularChainCleanup', []);
        console.log('[PASS] Test 5: Circular chain cleanup');
    } catch (e) {
        console.error('[FAIL] Test 5:', (e as any).message);
        throw e;
    }
}

console.log('\n✅ Circular Reference & Class_Terminate (VBA Source): 全テスト通過');
