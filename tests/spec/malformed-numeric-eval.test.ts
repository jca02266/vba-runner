import { VBARunner, assert } from '../../test-libs/test-runner';

function evalBare(source: string): any {
    return new VBARunner(null, { quiet: true }).eval(source);
}

for (const source of ['1.2.3', '1..2', '1 2', '1.2 .3']) {
    assert.throwsMatch(
        () => evalBare(source),
        /Parse error|unexpected token|invalid numeric label/,
        `${source} は不正な数値列として拒否される`,
    );
}

assert.strictEqual(evalBare('.3'), 0.3, '.3 は単一の数値式として維持する');
assert.strictEqual(evalBare('1.'), 1, '1. は単一の数値式として維持する');
assert.strictEqual(evalBare('1.2'), 1.2, '1.2 は単一の数値式として維持する');

// The runner intentionally supports a colonless integer line label.
assert.strictEqual(evalBare('10'), 10, '整数リテラルのbare eval動作');
assert.strictEqual(evalBare('10:'), undefined, 'コロン付き整数行ラベルの互換動作');

console.log('✅ malformed numeric eval boundaries: all tests passed');
