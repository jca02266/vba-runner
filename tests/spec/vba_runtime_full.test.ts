import { Evaluator, vbaMissing } from '../../src/compiler/evaluator';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import * as fs from 'fs';
import * as path from 'path';

function testFullRuntime() {
    console.log("Running Full Runtime tests...");
    
    // Clean workspace
    const workspace = path.join(process.cwd(), 'workspace');
    if (fs.existsSync(workspace)) {
        fs.rmSync(workspace, { recursive: true, force: true });
    }
    fs.mkdirSync(workspace);

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
        End Sub

        Function GetVal(Optional x As Integer = 123)
            GetVal = x
        End Function

        Main
    `;

    let output = "";
    const lexer = new Lexer(vbaCode);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((o) => { output += o + "\n"; }, { sandboxRoot: workspace });
    
    evaluator.evaluate(program);

    const lines = output.trim().split("\n").map(l => l.trim());
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
        "DateAdd: ", // Partial match
        "DateDiff: 9"
    ];

    let passed = 0;
    for (let i = 0; i < expected.length; i++) {
        const line = lines[i] || "";
        const exp = expected[i];
        
        if (exp === "DateAdd: ") {
            // Accept either formatted string or serial number
            if (line.includes("2/1/2023") || line.includes("2023-02-01") || line.includes("44958")) {
                passed++;
            } else {
                console.log(`[FAIL] Line ${i + 1}: Expected DateAdd to match 2023-02-01 or 44958, got "${line}"`);
            }
            continue;
        }

        if (line === exp) {
            passed++;
        } else {
            console.log(`[FAIL] Line ${i + 1}: Expected "${exp}", got "${line}"`);
        }
    }

    if (passed === expected.length) {
        console.log("✅ VBA Runtime Full: 全テスト通過");
    } else {
        console.log(`[FAIL] Passed ${passed}/${expected.length} tests`);
        process.exit(1);
    }
}

testFullRuntime();
