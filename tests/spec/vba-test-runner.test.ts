import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import * as fs from 'fs';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Load and test the VBA file
const vbaCode = fs.readFileSync(__dirname + '/vba/Test_CurrencyOperations.vba', 'utf8');

try {
    const result = runFunc(vbaCode, 'RunAllTests');
    if (result === true) {
        console.log('\n✅ VBA Test Suite: All tests passed');
    } else {
        console.log('\n❌ VBA Test Suite: Some tests failed');
    }
} catch (e: any) {
    console.log('Error running VBA tests:', e.message);
}
