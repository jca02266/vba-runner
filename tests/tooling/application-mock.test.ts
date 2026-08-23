/**
 * Application host-global mock fixture.
 *
 * The fixture intentionally covers only the members used by the Excel queue:
 * PathSeparator and DisplayAlerts.
 */
import { VBARunner, assert } from '../../test-libs/test-runner';

const suite = new VBARunner('tests/fixtures/application-mock');

assert.strictEqual(suite.run('ReadPathSeparator', []), '/',
    'Application.PathSeparator is provided by the mock');
assert.strictEqual(suite.run('ToggleDisplayAlerts', []), false,
    'Application.DisplayAlerts accepts assignment through the mock');

console.log('[PASS] Application host-global mock preserves queue members');
