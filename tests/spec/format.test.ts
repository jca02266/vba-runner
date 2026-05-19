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

// --- 時刻フォーマットのバグ修正確認 ---
// Bug 1: HH (大文字) が時刻として認識されずリテラル出力されていた
// Bug 2: "hh:mm:ss" の mm が月として解釈されていた（文脈判定なし）
// Bug 3: formatDate が getHours() (ローカル時刻) を使っており UTC 保存の Date とずれていた
const timeCode = `
    Function TestHH()
        TestHH = Format(#2023-06-15 09:30:45#, "HH:mm:ss")
    End Function
    Function TestHhNn()
        TestHhNn = Format(#2023-06-15 09:30:45#, "hh:nn:ss")
    End Function
    Function TestIsoFormat()
        TestIsoFormat = Format(#2023-01-15 12:34:56#, "yyyy-mm-ddTHH:mm:ss")
    End Function
    Function TestMmIsMonth()
        TestMmIsMonth = Format(#2023-06-15#, "yyyy-mm-dd")
    End Function
    Function TestMmIsMinute()
        TestMmIsMinute = Format(#2023-06-15 09:07:05#, "hh:mm:ss")
    End Function
`;
const evTime = evalVBA(timeCode);

// HH (大文字) = 時刻
assert.strictEqual(evTime.callProcedure('TestHH', []), '09:30:45',
    'Format HH:mm:ss — HH(大文字)が時刻、mm が文脈上「分」になる');
// hh:nn:ss — nn は明示的な「分」
assert.strictEqual(evTime.callProcedure('TestHhNn', []), '09:30:45',
    'Format hh:nn:ss — nn は常に分');
// ISO 8601 形式: yyyy-mm-ddTHH:mm:ss
assert.strictEqual(evTime.callProcedure('TestIsoFormat', []), '2023-01-15T12:34:56',
    'Format "yyyy-mm-ddTHH:mm:ss" — 最初の mm は月、HH 後の mm は分');
// mm 単独（時刻トークンなし）→ 月
assert.strictEqual(evTime.callProcedure('TestMmIsMonth', []), '2023-06-15',
    'Format yyyy-mm-dd — mm は月');
// hh: 直後の mm → 分（ゼロ埋め）
assert.strictEqual(evTime.callProcedure('TestMmIsMinute', []), '09:07:05',
    'Format hh:mm:ss — mm は分（ゼロ埋め）');

console.log('[PASS] Format 時刻: HH大文字・文脈依存mm・UTC整合性');

console.log('\n✅ Format: 全テスト通過');
