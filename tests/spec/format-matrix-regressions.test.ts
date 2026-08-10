import { evalVBASingle, assert } from '../../test-libs/test-runner';

const source = String.raw`
Function FBracket(): FBracket = Format(1000, "[abc]"): End Function
Function FBracketString(): FBracketString = Format("Ab9", "[<=1000]"): End Function
Function FBracketDate(): FBracketDate = Format(CDate("2026-03-15 12:34:56"), "[abc]"): End Function
Function FDateNumber(): FDateNumber = Format(CDate("2026-03-15 12:34:56"), "0.00"): End Function
Function FLiteralSections(): FLiteralSections = Format(-1000, "Lo;Hi"): End Function
Function FNumericSection(): FNumericSection = Format(-1000, "Lo0;Hi"): End Function
Function FPercentLiteral(): FPercentLiteral = Format(1000, "%"): End Function
Function FSingleAt(): FSingleAt = Format("Ab9", "@"): End Function
Function FExclamation(): FExclamation = Format(1000, "!"): End Function
`;

const ev = evalVBASingle(source);
assert.strictEqual(ev.callProcedure('FBracket', []), '1000', 'unescaped bracket directive is ignored');
assert.strictEqual(ev.callProcedure('FBracketString', []), 'Ab9', 'bracket directive does not case-convert strings');
assert.strictEqual(ev.callProcedure('FBracketDate', []), '2026/03/15 12:34:56', 'bracket-only Date uses General Date');
assert.strictEqual(ev.callProcedure('FDateNumber', []), '46096.52', 'Date numeric format uses serial value');
assert.strictEqual(ev.callProcedure('FLiteralSections', []), 'Lo', 'literal-only first section is reused');
assert.strictEqual(ev.callProcedure('FNumericSection', []), 'Hi', 'numeric first section enables negative selection');
assert.strictEqual(ev.callProcedure('FPercentLiteral', []), '%', 'percent without a numeric placeholder is literal');
assert.strictEqual(ev.callProcedure('FSingleAt', []), 'Ab9', 'single string placeholder preserves the string');
assert.strictEqual(ev.callProcedure('FExclamation', []), '', 'exclamation-only format emits no value');

console.log('[PASS] Format matrix regression boundaries');
