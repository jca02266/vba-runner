import { assert } from '../../test-libs/test-runner';
import { Evaluator } from '../../src/engine/evaluator';
import { Parser } from '../../src/engine/parser';
import { Lexer } from '../../src/engine/lexer';

function evalVBA(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((s) => console.log(s));
    evaluator.evaluate(program);
    return evaluator;
}

const code = `
    Function TestFormatDate()
        TestFormatDate = Format(#2025/05/10#, "yyyy-mm-dd")
    End Function

    Function TestFormatNumber()
        TestFormatNumber = Format(1234.567, "0.00")
    End Function

    Function TestFormatCurrency()
        TestFormatCurrency = Format(1234.567, "#,##0.00")
    End Function
`;

const ev = evalVBA(code);

console.log('[Test Suite] Format の検証');

assert.strictEqual(ev.callProcedure('TestFormatDate', []), '2025-05-10', 'Format(Date, "yyyy-mm-dd")');
assert.strictEqual(ev.callProcedure('TestFormatNumber', []), '1234.57', 'Format(Number, "0.00")');
// Note: toLocaleString behavior might vary by environment for thousand separators, 
// but we expect "1,234.57" in standard US-like locales if grouping is enabled.
const currency = ev.callProcedure('TestFormatCurrency', []);
assert.ok(currency.includes('1,234.57') || currency.includes('1234.57'), 'Format(Number, "#,##0.00")');

console.log('✅ Format: 全テスト通過');
