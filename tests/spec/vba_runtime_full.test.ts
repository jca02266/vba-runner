import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

console.log("Running Full Runtime tests...");

const vfs = new MemoryFileSystem();

const vbaCode = `
    Sub Main()
        ' 1. Optional Default Values
        Debug.Print "Opt1: " & GetVal()
        Debug.Print "Opt2: " & GetVal(999)

        ' 2. File I/O
        Dim f
        f = FreeFile()
        Open "runtime_test.txt" For Output As #f
        Print #f, "Line1"
        Print #f, "Line2"; "Cont"
        Write #f, "Data", 123
        Debug.Print "LOF: " & LOF(f)
        Close #f

        Open "runtime_test.txt" For Input As #f
        Dim s, n
        Line Input #f, s
        Debug.Print "Read1: " & s
        Line Input #f, s
        Debug.Print "Read2: " & s
        Debug.Print "EOF before end: " & EOF(f)
        Input #f, s, n
        Debug.Print "Input: " & s & ", " & n
        Debug.Print "EOF after end: " & EOF(f)
        Close #f

        ' 3. Conversion & Math
        Debug.Print "Hex: " & Hex(255)
        Debug.Print "Oct: " & Oct(8)
        Debug.Print "Int: " & Int(-1.1)
        Debug.Print "Fix: " & Fix(-1.1)

        ' 4. Utility
        Debug.Print "Choose: " & Choose(2, "Apple", "Orange", "Grape")
        Debug.Print "Switch: " & Switch(1=2, "No", 1=1, "Yes")

        ' 5. Date/Time
        Dim d
        d = #2023-01-01#
        Debug.Print "DateAdd: " & DateAdd("m", 1, d)
        Debug.Print "DateDiff: " & DateDiff("d", d, #2023-01-10#)

        ' 6. String functions
        Debug.Print "LTrim: " & LTrim("  abc")
        Debug.Print "RTrim: " & RTrim("abc  ")

        ' 7. Arrays
        Dim arr(1 To 2, 1 To 2) As Integer
        arr(1, 1) = 100
        ReDim Preserve arr(1 To 2, 1 To 3)
        arr(1, 3) = 200
        Debug.Print "arr(1,1): " & arr(1, 1)
        Debug.Print "arr(1,3): " & arr(1, 3)
        Debug.Print "UBound(arr, 2): " & UBound(arr, 2)

        ' 8. Error Object
        On Error Resume Next
        Err.Raise 100, "Source", "Desc", "HelpFile", 1000
        Debug.Print "Err.Number: " & Err.Number
        Debug.Print "Err.Source: " & Err.Source
        Debug.Print "Err.Description: " & Err.Description
        Debug.Print "Err.HelpFile: " & Err.HelpFile
        Debug.Print "Err.HelpContext: " & Err.HelpContext
        Debug.Print "Err.LastDllError: " & Err.LastDllError
        On Error GoTo 0

        ' 9. File System Stubs
        Debug.Print "FileAttr: " & FileAttr(1)
    End Sub

    Function GetVal(Optional x As Integer = 123)
        GetVal = x
    End Function

    Main
`;

const lines: string[] = [];
evalVBASingle(vbaCode, {
    onPrint: (o) => { if (!o.startsWith('[STUB]')) lines.push(o.trim()); },
    fs: vfs,
});

console.log("Output lines:", lines);

const expected = [
    "Opt1: 123",
    "Opt2: 999",
    "LOF: 29",
    "Read1: Line1",
    "Read2: Line2Cont",
    "EOF before end: False",
    "Input: Data, 123",
    "EOF after end: True",
    "Hex: FF",
    "Oct: 10",
    "Int: -2",
    "Fix: -1",
    "Choose: Orange",
    "Switch: Yes",
    null, // DateAdd: partial match
    "DateDiff: 9",
    "LTrim: abc",
    "RTrim: abc",
    "arr(1,1): 100",
    "arr(1,3): 200",
    "UBound(arr, 2): 3",
    "Err.Number: 100",
    "Err.Source: Source",
    "Err.Description: Desc",
    "Err.HelpFile: HelpFile",
    "Err.HelpContext: 1000",
    "Err.LastDllError: 0",
    "FileAttr: 1",
];

expected.forEach((exp, i) => {
    const line = lines[i] ?? "";
    if (exp === null) {
        // DateAdd: Accept various date formats
        assert.ok(
            line.includes("2/1/2023") || line.includes("2023-02-01") || line.includes("2023/02/01") || line.includes("44958"),
            `DateAdd should match 2023-02-01 (got: "${line}")`
        );
    } else {
        assert.strictEqual(line, exp, `Line ${i + 1}`);
    }
});

console.log('✅ VBA Runtime Full: 全テスト通過');
