import { assertCompileErrorPreproc, assertCompileErrorResolve } from '../../test-libs/test-runner';

// VBA の宣言名は大文字小文字を区別しない。同一スコープでケースだけが
// 異なる宣言を別名として登録せず、コンパイルエラーにする。
assertCompileErrorResolve(`
    Dim Value As Long
    Sub value()
    End Sub
`, 3, /duplicate declaration|duplicate.*name/i, 'module variable/procedure collision');

assertCompileErrorPreproc(`
    Sub Check(arg As Long)
        Dim ARG As Long
    End Sub
`, 'Check', 3, /duplicate declaration/i, 'parameter/local collision');

assertCompileErrorPreproc(`
    Sub Check()
        Const Limit = 1
        Dim limit As Long
    End Sub
`, 'Check', 4, /duplicate declaration/i, 'constant/local collision');

console.log('[PASS] case-insensitive declaration collisions');
