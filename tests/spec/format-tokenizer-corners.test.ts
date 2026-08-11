import { evalVBASingle, assert } from '../../test-libs/test-runner';

const source = String.raw`
Function FSemicolonEscape(): FSemicolonEscape = Format(12, "0\;0"): End Function
Function FSemicolonQuote(): FSemicolonQuote = Format(12, "0"";""0"): End Function
Function FBracketEscape(): FBracketEscape = Format(1000, "\[<=1000\]"): End Function
Function FBracketQuote(): FBracketQuote = Format(1000, """[abc]"""): End Function
Function FUnclosedBracket(): FUnclosedBracket = Format(1000, "[abc"): End Function
Function FColor(): FColor = Format(-12, "[Red]0"): End Function
Function FScientificLower(): FScientificLower = Format(12.3, "0.0e-00"): End Function
Function FPercentLiteral(): FPercentLiteral = Format(12, "%"): End Function
Function FStringLeft(): FStringLeft = Format("Ab9", "!@@@"): End Function
Function FStringRight(): FStringRight = Format("Ab9", "@@@"): End Function
Function FStringLiteralPrefix(): FStringLiteralPrefix = Format("Ab9", "pre @@@"): End Function
Function FStringQuotedLiteral(): FStringQuotedLiteral = Format("Ab9", """[x]""@@@"): End Function
Function FStringEscapedLiteral(): FStringEscapedLiteral = Format("Ab9", "\[x\]@@@"): End Function
Function FStringCaseLiteral(): FStringCaseLiteral = Format("Ab9", ">pre @@@"): End Function
Function FStringNumericPattern(): FStringNumericPattern = Format("Ab9", "0.00"): End Function
Function FStringDatePattern(): FStringDatePattern = Format("Ab9", "yyyy-mm-dd"): End Function
Function FDateMinutes(): FDateMinutes = Format(CDate("2026-03-15 12:34:56"), "hh:mm"): End Function
Function FDateMonth(): FDateMonth = Format(CDate("2026-03-15"), "m"): End Function
Function FDateMinute(): FDateMinute = Format(CDate("2026-03-15 12:34:56"), "n"): End Function
Function FEra(): FEra = Format(CDate("2026-03-15"), "ggee"): End Function
`;

const ev = evalVBASingle(source);
assert.strictEqual(ev.callProcedure('FSemicolonEscape', []), '12', 'escaped semicolon is literal');
assert.strictEqual(ev.callProcedure('FSemicolonQuote', []), '12', 'quoted semicolon is literal');
assert.strictEqual(ev.callProcedure('FBracketEscape', []), '[=1000]1000', 'escaped comparison brackets follow Excel tokenization');
assert.strictEqual(ev.callProcedure('FBracketQuote', []), '[abc]', 'quoted brackets remain literal');
assert.strictEqual(ev.callProcedure('FUnclosedBracket', []), '', 'unclosed bracket follows numeric error boundary');
assert.strictEqual(ev.callProcedure('FColor', []), '-12', 'color directive is removed from value output');
assert.strictEqual(ev.callProcedure('FScientificLower', []), '1.2e01', 'lowercase scientific token preserves E case');
assert.strictEqual(ev.callProcedure('FPercentLiteral', []), '%', 'percent without numeric token is literal');
assert.strictEqual(ev.callProcedure('FStringLeft', []), 'Ab9', '! fills string placeholders left to right');
assert.strictEqual(ev.callProcedure('FStringRight', []), 'Ab9', 'string placeholders fill right to left');
assert.strictEqual(ev.callProcedure('FStringLiteralPrefix', []), 'pre Ab9', 'string format preserves plain literals');
assert.strictEqual(ev.callProcedure('FStringQuotedLiteral', []), '[x]Ab9', 'string format preserves quoted literals');
assert.strictEqual(ev.callProcedure('FStringEscapedLiteral', []), '[x]Ab9', 'string format preserves escaped literals');
assert.strictEqual(ev.callProcedure('FStringCaseLiteral', []), 'PRE AB9', 'string format applies case control after rendering literals');
assert.strictEqual(ev.callProcedure('FStringNumericPattern', []), 'Ab9', 'numeric format does not format String input');
assert.strictEqual(ev.callProcedure('FStringDatePattern', []), 'Ab9', 'date format does not format non-date String input');
assert.strictEqual(ev.callProcedure('FDateMinutes', []), '12:34', 'm after h is minutes');
assert.strictEqual(ev.callProcedure('FDateMonth', []), '3', 'm without hour is month');
assert.strictEqual(ev.callProcedure('FDateMinute', []), '34', 'n is minutes');
assert.strictEqual(ev.callProcedure('FEra', []), 'R08', 'Japanese era tokens are composed independently');

console.log('[PASS] Format tokenizer corner boundaries');
