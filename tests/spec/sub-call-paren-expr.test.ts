/**
 * Sub 呼び出しにおけるスペース前 `(` の式継続テスト
 *
 * VBA の仕様: `Sub (expr)` の形式でスペースの後に `(` が来る場合、
 * `(` はサブルーチンの引数式の一部であり、関数呼び出しの括弧ではない。
 * したがって `Debug.Print (1+2)*3` は `Debug.Print` に引数 `9` を渡す。
 *
 * 修正前は ParseError になっていた。
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function run(code: string): string[] {
    const out: string[] = [];
    const ast = new Parser(new Lexer(code).tokenize()).parse();
    const ev = new Evaluator((s: string) => out.push(String(s)));
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    ev.callProcedure('T', []);
    return out;
}

// ── 基本ケース ──────────────────────────────────────────────────────────────

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (1+2)*3
End Sub`), ['9'], '(1+2)*3 = 9');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (1+2)/3
End Sub`), ['1'], '(1+2)/3 = 1');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (2)^3
End Sub`), ['8'], '(2)^3 = 8');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (3) Mod 2
End Sub`), ['1'], '(3) Mod 2 = 1');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (10)\\3
End Sub`), ['3'], '(10)\\\\3 = 3');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print ("a") & "b"
End Sub`), ['ab'], '("a") & "b" = "ab"');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (1+2)+3
End Sub`), ['6'], '(1+2)+3 = 6');

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (1+2)-1
End Sub`), ['2'], '(1+2)-1 = 2');

// ── 既存動作の保護 ──────────────────────────────────────────────────────────

assert.deepStrictEqual(run(`
Sub T()
  Debug.Print (1+2)
End Sub`), ['3'], '(1+2) single paren arg = 3');

assert.deepStrictEqual(run(`
Sub T()
  Dim arr(5) As Integer
  arr(0) = 5
  Debug.Print arr(0)
End Sub`), ['5'], 'arr(0)=5 no-space unchanged');

assert.deepStrictEqual(run(`
Sub T()
  Dim arr(5) As Integer
  arr (0) = 5
  Debug.Print arr(0)
End Sub`), ['5'], 'arr (0)=5 with-space assignment unchanged');

console.log('[PASS] Sub call paren expression continuation - all tests passed');
