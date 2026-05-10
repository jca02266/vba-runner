import { VBATest, assert } from '../ts/test-runner';

async function main() {
    console.log("--- Starting Built-in String Functions Tests ---");
    const vbaTest = new VBATest('sample/src/vba/Sample.vba');

    console.log("[Test] Split & Join");
    const parts = vbaTest.eval(`Split("a,b,c", ",")`);
    assert.strictEqual(Array.isArray(parts), true);
    assert.strictEqual(parts.length, 3);
    assert.strictEqual(parts[0], "a");
    
    assert.strictEqual(vbaTest.eval(`Join(Split("a,b,c", ","), "-")`), "a-b-c");
    assert.strictEqual(vbaTest.eval(`Join(Split("hello world", " "), "|")`), "hello|world");

    console.log("[Test] Asc & Chr");
    assert.strictEqual(vbaTest.eval(`Asc("A")`), 65);
    assert.strictEqual(vbaTest.eval(`Chr(65)`), "A");
    assert.strictEqual(vbaTest.eval(`Chr(Asc("Z"))`), "Z");

    console.log("[Test] Space & String");
    assert.strictEqual(vbaTest.eval(`Space(3)`), "   ");
    assert.strictEqual(vbaTest.eval(`String(5, "*")`), "*****");
    assert.strictEqual(vbaTest.eval(`String(3, "ABC")`), "AAA"); // Only uses first char

    console.log("--- All Built-in String Functions Tests passed! ---");
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
