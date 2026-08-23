/**
 * Excel queue host-global mock coverage.
 *
 * Keep this focused on the members used by the queue's VBA modules so the
 * mock remains a small host boundary rather than an Excel implementation.
 */
import path from 'node:path';
import { VBARunner, assert } from '../../test-libs/test-runner';

const queueDirectory = path.resolve('tests/excel/queue');
const runner = new VBARunner(
    path.join(queueDirectory, 'FormatMatrixVerification.bas'),
    { quiet: true },
);

assert.strictEqual(runner.eval('ThisWorkbook.Path'), queueDirectory,
    'ThisWorkbook.Path points to the queue directory');
assert.strictEqual(
    runner.eval('ThisWorkbook.Path & Application.PathSeparator & "FormatMatrix.result"'),
    path.join(queueDirectory, 'FormatMatrix.result'),
    'ThisWorkbook.Path composes the queue result path',
);

console.log('[PASS] Excel queue host mock preserves ThisWorkbook.Path');
