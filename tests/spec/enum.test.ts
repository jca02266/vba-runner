import { VBATest, assert } from '../ts/test-runner';

async function main() {
    console.log("--- Starting Enum Declaration Tests ---");
    const vbaTest = new VBATest('sample/src/vba/Sample.vba');

    const vbaSource = `
        Enum Color
            Red = 1
            Green
            Blue = 10
            Yellow
        End Enum
    `;

    // We can't use eval for declarations because eval only handles expressions.
    // But VBATest constructor parses the file.
    // I'll use evaluate directly if needed, or just a temporary file.
    
    // For this test, I'll use a string source with the compiler
    const program = vbaTest['evaluator']['evaluate'](
        new (require('../../src/compiler/parser').Parser)(
            new (require('../../src/compiler/lexer').Lexer)(vbaSource).tokenize()
        ).parse()
    );

    console.log("[Test] Enum member direct access");
    assert.strictEqual(vbaTest.eval("Red"), 1);
    assert.strictEqual(vbaTest.eval("Green"), 2); // Incremented
    assert.strictEqual(vbaTest.eval("Blue"), 10);
    assert.strictEqual(vbaTest.eval("Yellow"), 11); // Incremented

    console.log("[Test] Enum member access via Enum name");
    assert.strictEqual(vbaTest.eval("Color.Red"), 1);
    assert.strictEqual(vbaTest.eval("Color.Green"), 2);

    console.log("--- All Enum Declaration Tests passed! ---");
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
