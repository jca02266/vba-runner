import { assert } from '../../test-libs/test-runner';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

const ast = new Parser(new Lexer('Dim value As A.B.C').tokenize()).parse();
const declaration = (ast.body[0] as any).declarations[0];

assert.strictEqual(declaration.objectType, 'A.B.C');

console.log('✅ Dim accepts multi-qualified type names');
