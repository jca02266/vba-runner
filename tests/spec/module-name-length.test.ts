import { Evaluator } from '../../src/engine/evaluator';
import { MemoryFileSystem } from '../../src/engine/filesystem';

// Test 1: Valid module name (exactly 31 characters)
const maxLengthName = 'A'.repeat(31); // 31 characters
try {
    const fs = new MemoryFileSystem();
    const evaluator = new Evaluator(console.log, { fs });
    evaluator.setSourceModule(maxLengthName);
    console.log('[PASS] Module name with exactly 31 characters');
} catch (e: any) {
    console.error(`[FAIL] Module name with 31 characters should be valid: ${e.message}`);
    throw e;
}

// Test 2: Valid module name (under 31 characters)
try {
    const fs = new MemoryFileSystem();
    const evaluator = new Evaluator(console.log, { fs });
    evaluator.setSourceModule('MyModule');
    console.log('[PASS] Module name under 31 characters');
} catch (e: any) {
    console.error(`[FAIL] Module name under 31 characters should be valid: ${e.message}`);
    throw e;
}

// Test 3: Invalid module name (32 characters)
const tooLongName = 'A'.repeat(32); // 32 characters
try {
    const fs = new MemoryFileSystem();
    const evaluator = new Evaluator(console.log, { fs });
    evaluator.setSourceModule(tooLongName);
    console.error('[FAIL] Module name with 32 characters should raise an error');
    throw new Error('Expected error was not raised');
} catch (e: any) {
    if (e.message.includes('exceeds the maximum length of 31 characters')) {
        console.log('[PASS] Module name with 32 characters raises appropriate error');
    } else {
        console.error(`[FAIL] Unexpected error message: ${e.message}`);
        throw e;
    }
}

// Test 4: Invalid module name (very long)
const veryLongName = 'ThisIsAVeryLongModuleNameThatExceedsTheLimit';
try {
    const fs = new MemoryFileSystem();
    const evaluator = new Evaluator(console.log, { fs });
    evaluator.setSourceModule(veryLongName);
    console.error('[FAIL] Very long module name should raise an error');
    throw new Error('Expected error was not raised');
} catch (e: any) {
    if (e.message.includes('exceeds the maximum length of 31 characters')) {
        console.log('[PASS] Very long module name raises appropriate error');
    } else {
        console.error(`[FAIL] Unexpected error message: ${e.message}`);
        throw e;
    }
}

console.log('\n✅ Module name length validation: 全テスト通過');
