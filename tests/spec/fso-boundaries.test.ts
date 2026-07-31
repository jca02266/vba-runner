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
