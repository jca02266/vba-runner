import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const overwriteFs = new MemoryFileSystem();
const overwrite = evalVBASingle(`
    Function ProbeCreateNoOverwrite() As Long
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("C:\\existing.txt", True, False)
        stream.Write "original"
        stream.Close
        On Error Resume Next
        Set stream = fso.CreateTextFile("C:\\existing.txt", False, False)
        ProbeCreateNoOverwrite = Err.Number
    End Function
`, { fs: overwriteFs });
assert.strictEqual(overwrite.callProcedure('ProbeCreateNoOverwrite', []), 58,
    'CreateTextFile honors a VBA False overwrite flag');

const missingFs = new MemoryFileSystem();
const missing = evalVBASingle(`
    Function ProbeOpenMissing() As Long
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Set stream = fso.OpenTextFile("C:\\missing.txt", 1, False, 0)
        ProbeOpenMissing = Err.Number
    End Function
`, { fs: missingFs });
assert.strictEqual(missing.callProcedure('ProbeOpenMissing', []), 53,
    'OpenTextFile maps a missing path to File not found');

const copyFs = new MemoryFileSystem();
const copy = evalVBASingle(`
    Function ProbeCopyNoOverwrite() As Long
        Dim fso As Object, stream As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("C:\\source.txt", True, False)
        stream.Write "source"
        stream.Close
        Set stream = fso.CreateTextFile("C:\\target.txt", True, False)
        stream.Write "target"
        stream.Close
        On Error Resume Next
        fso.CopyFile "C:\\source.txt", "C:\\target.txt", False
        ProbeCopyNoOverwrite = Err.Number
    End Function
`, { fs: copyFs });
assert.strictEqual(copy.callProcedure('ProbeCopyNoOverwrite', []), 58,
    'CopyFile honors a VBA False overwrite flag');

console.log('[PASS] FSO Boolean flags and missing-file errors');

const pathProbe = (body: string, name: string, expected: number) => {
    const fs = new MemoryFileSystem();
    fs.mkdirSync('/sandbox', { recursive: true });
    const result = evalVBASingle(`
        Function ${name}() As Long
            Dim fso As Object, stream As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            ${body}
            ${name} = Err.Number
        End Function
    `, { fs });
    assert.strictEqual(result.callProcedure(name, []), expected, `${name} maps FSO errors`);
};
pathProbe('fso.CreateFolder "missing\\child"', 'ProbeMissingParent', 76);
pathProbe('fso.CreateFolder "existing": fso.CreateFolder "existing"', 'ProbeExistingFolder', 58);
pathProbe('fso.DeleteFolder "absent"', 'ProbeMissingFolder', 76);
pathProbe('Set stream = fso.GetFile("absent.txt")', 'ProbeMissingFile', 53);
pathProbe('Set stream = fso.GetFolder("absent")', 'ProbeMissingGetFolder', 76);

const variantFlagFs = new MemoryFileSystem();
const variantFlags = evalVBASingle(`
    Function ProbeVariantFlags() As String
        Dim fso As Object, stream As Object, result As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.CreateTextFile("C:\\variant.txt", True, False)
        stream.Write "old": stream.Close
        On Error Resume Next
        Set stream = fso.CreateTextFile("C:\\variant.txt", 1, False)
        result = result & Err.Number & ","
        Err.Clear
        Set stream = fso.OpenTextFile("C:\\created.txt", 1, 2, 0)
        result = result & Err.Number & ","
        stream.Close
        Set stream = fso.CreateTextFile("C:\\unicode.txt", True, 1)
        stream.Write "Aあ": stream.Close
        ProbeVariantFlags = result
    End Function
`, { fs: variantFlagFs });
assert.strictEqual(variantFlags.callProcedure('ProbeVariantFlags', []), '0,0,',
    'FSO Boolean parameters coerce nonzero numeric Variants to True');
