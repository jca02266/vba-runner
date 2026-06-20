/**
 * vba-runner 拡張: モジュールレベル文（Dim・代入・For/Next 等）はプロシージャの後に
 * 書いても allowTopLevelStatements（デフォルト true）が有効な間はコンパイルエラーにならない。
 *
 * 標準 VBA では、プロシージャの後（End Sub/End Function/End Property 以降）には
 * コメントしか書けず、Dim 等の宣言文や実行文を書くと構文エラーになる
 * （tests/vba/CompileError.bas の module_level_*_after_procedure 系ケースを参照）。
 * vba-runner は evalVBASingle/evalVBAModules を使った REPL・テストスクリプト用に、
 * 先頭にない Dim もモジュールレベル実行文と同様に許容する拡張を入れているため、
 * 両方の挙動（拡張あり / 標準 VBA 相当）を確認する。
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

const toplevelStmtCode = `
Sub Foo()
End Sub
For idx = 0 To 10
Next idx
`;

const dimAfterProcCode = `
Sub Foo()
End Sub
Dim v As Integer
`;

for (const [label, code] of [['For/Next', toplevelStmtCode], ['Dim', dimAfterProcCode]] as const) {
    assert.doesNotThrow(
        () => evalVBASingle(code, { onPrint: () => {} }),
        `allowTopLevelStatements (default true) ではプロシージャ後の ${label} がエラーにならないはず`,
    );
    console.log(`[PASS] default allowTopLevelStatements: no compile error after procedure (${label})`);

    assert.throwsMatch(
        () => evalVBASingle(code, { onPrint: () => {}, allowTopLevelStatements: false }),
        /only comments may appear after end sub/i,
        `allowTopLevelStatements: false では標準 VBA 同様コンパイルエラーになるはず (${label})`,
    );
    console.log(`[PASS] allowTopLevelStatements: false matches standard VBA compile error (${label})`);
}
