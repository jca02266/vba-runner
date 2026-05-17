import * as fs from 'fs';
import * as path from 'path';
import { VBARunner } from '../../test-libs/test-runner';

const tempDir = '/tmp/vba-module-name-test';
const fixtureDir = path.join(tempDir, 'fixtures');

// Create temporary directory for test fixtures
if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
}
if (!fs.existsSync(fixtureDir)) {
    fs.mkdirSync(fixtureDir, { recursive: true });
}

// Test 1: Load VBA file with valid name (< 31 characters)
const validName = 'ValidModule';
const validFilePath = path.join(fixtureDir, validName + '.vba');
fs.writeFileSync(validFilePath, 'Sub Test()\n  Debug.Print "OK"\nEnd Sub\n', 'utf-8');

try {
    const vbaRunner = new VBARunner(validFilePath);
    console.log('[PASS] VBA file with valid module name loads successfully');
} catch (e: any) {
    console.error(`[FAIL] VBA file with valid name should load: ${e.message}`);
    throw e;
}

// Test 2: Load VBA file with name exactly at the limit (31 characters)
const limitName = 'A'.repeat(31);
const limitFilePath = path.join(fixtureDir, limitName + '.vba');
fs.writeFileSync(limitFilePath, 'Sub Test()\n  Debug.Print "OK"\nEnd Sub\n', 'utf-8');

try {
    const vbaRunner = new VBARunner(limitFilePath);
    console.log('[PASS] VBA file with module name exactly 31 characters loads successfully');
} catch (e: any) {
    console.error(`[FAIL] VBA file with 31-character name should load: ${e.message}`);
    throw e;
}

// Test 3: Load VBA file with name exceeding limit (32 characters)
const tooLongName = 'A'.repeat(32);
const tooLongFilePath = path.join(fixtureDir, tooLongName + '.vba');
fs.writeFileSync(tooLongFilePath, 'Sub Test()\n  Debug.Print "OK"\nEnd Sub\n', 'utf-8');

try {
    const vbaRunner = new VBARunner(tooLongFilePath);
    console.error('[FAIL] VBA file with 32-character name should raise an error');
    throw new Error('Expected error was not raised');
} catch (e: any) {
    if (e.message.includes('exceeds the maximum length of 31 characters')) {
        console.log('[PASS] VBA file with 32-character module name raises appropriate error');
    } else {
        console.error(`[FAIL] Unexpected error message: ${e.message}`);
        throw e;
    }
}

// Test 4: Load directory with a file that has a name exceeding the limit
const testDirName = 'VBAFilesToLoad';
const testDir = path.join(fixtureDir, testDirName);
if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir, { recursive: true });
}

// Create a file with valid name
const validFile = path.join(testDir, 'ValidModule.vba');
fs.writeFileSync(validFile, 'Sub Helper()\n  Debug.Print "Helper"\nEnd Sub\n', 'utf-8');

// Create a file with name that's too long
const tooLongFile = path.join(testDir, 'B'.repeat(32) + '.vba');
fs.writeFileSync(tooLongFile, 'Sub AnotherSub()\n  Debug.Print "Too Long"\nEnd Sub\n', 'utf-8');

try {
    const vbaRunner = new VBARunner(testDir);
    console.error('[FAIL] Loading directory with a file with 32-character name should raise an error');
    throw new Error('Expected error was not raised');
} catch (e: any) {
    if (e.message.includes('exceeds the maximum length of 31 characters')) {
        console.log('[PASS] Loading directory with file exceeding module name limit raises appropriate error');
    } else {
        console.error(`[FAIL] Unexpected error message: ${e.message}`);
        throw e;
    }
}

// Cleanup
fs.rmSync(tempDir, { recursive: true, force: true });

console.log('\n✅ Module name length integration test: 全テスト通過');
