/**
 * Application host-global mock fixture.
 *
 * The fixture intentionally covers only the members used by the Excel queue:
 * PathSeparator and DisplayAlerts.
 */
import { VBARunner, assert } from '../../test-libs/test-runner';
import { spawnSync } from 'node:child_process';
import path from 'node:path';

const mockFile = path.resolve('tests/fixtures/application-mock/__mocks__/Application.js');
const nativeRequire = spawnSync(process.execPath, ['-e', `require(${JSON.stringify(mockFile)})`], {
    encoding: 'utf8',
});
assert.notStrictEqual(nativeRequire.status, 0,
    'fixture must exercise the package type=module/CommonJS compatibility boundary');
assert.ok(/ERR_REQUIRE_ESM/.test(nativeRequire.stderr),
    'plain Node require should demonstrate why the mock loader must force CommonJS semantics');

const suite = new VBARunner('tests/fixtures/application-mock');

assert.strictEqual(suite.run('ReadPathSeparator', []), '/',
    'Application.PathSeparator is provided by the mock');
assert.strictEqual(suite.run('ToggleDisplayAlerts', []), false,
    'Application.DisplayAlerts accepts assignment through the mock');

console.log('[PASS] Application host-global mock preserves queue members');
