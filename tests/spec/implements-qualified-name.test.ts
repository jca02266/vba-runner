import { assert } from '../../test-libs/test-runner';
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';

const ast = new Parser(new Lexer('Implements ADODB.ICommand').tokenize()).parse();

assert.strictEqual(ast.body.length, 1);
assert.strictEqual((ast.body[0] as any).type, 'ImplementsDirective');
assert.strictEqual((ast.body[0] as any).interfaceName, 'ADODB.ICommand');

console.log('✅ Implements accepts qualified interface names');
