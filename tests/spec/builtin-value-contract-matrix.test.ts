import { assert, assertClose, evalVBASingle } from '../../test-libs/test-runner';
import { vbaNull } from '../../src/engine/evaluator';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as yaml from 'js-yaml';

type ContractGroup = { name: string; rule: string; functions: string[] };
type ParameterContract = { name: string; categories: string[]; nullBehavior: string; optional?: boolean };
type DetailedContract = { name: string; parameters: ParameterContract[]; overloads?: { parameters: string[] }[] };
type GeneratedCase = { functionName: string; arguments: string[]; omitted: string[]; oracle: string[] };

/** The YAML file is the canonical contract table; this test only executes it. */
const contractPath = path.join(process.cwd(), 'tests/spec/data/builtin-value-contracts.yml');
const contractGroups = yaml.load(fs.readFileSync(contractPath, 'utf8')) as ContractGroup[];
const detailPath = path.join(process.cwd(), 'tests/spec/data/builtin-value-contract-details.yml');
const detailedContracts = yaml.load(fs.readFileSync(detailPath, 'utf8')) as DetailedContract[];
const examplePath = path.join(process.cwd(), 'tests/spec/data/builtin-value-contract-examples.yml');
const exactNormalContracts = yaml.load(fs.readFileSync(examplePath, 'utf8')) as {
    function: string;
    expression: string;
    expected?: any;
    error?: number;
    tolerance?: { absolute?: number; relative?: number };
}[];
// YAML 1.2 parses an unquoted `null` category as JavaScript null. Normalize
// categories at the boundary so the contract vocabulary remains string-based.
for (const contract of detailedContracts) {
    for (const parameter of contract.parameters) parameter.categories = parameter.categories.map(String);
}

const registeredNames = new Set(contractGroups.flatMap((group) => group.functions));
const detailedNames = new Set(detailedContracts.map((contract) => contract.name));
for (const example of exactNormalContracts) {
    assert.strictEqual(registeredNames.has(example.function), true, `${example.function}: example must be registered`);
    assert.strictEqual(detailedNames.has(example.function), true, `${example.function}: example must have a detailed contract`);
    assert.strictEqual((example.expected === undefined) !== (example.error === undefined), true,
        `${example.function}: example must specify exactly one of expected or error`);
    if (example.tolerance) {
        assert.strictEqual(example.error === undefined, true, `${example.function}: tolerance cannot accompany error`);
        assert.strictEqual(typeof example.expected === 'number', true, `${example.function}: tolerance requires numeric expected value`);
        assert.strictEqual((example.tolerance.absolute ?? 0) >= 0 && (example.tolerance.relative ?? 0) >= 0, true,
            `${example.function}: tolerance must be non-negative`);
    }
}
assert.strictEqual(registeredNames.size, 134, 'contract inventory must cover all registered built-ins');
assert.strictEqual(contractGroups.reduce((n, group) => n + group.functions.length, 0), 134, 'contract inventory must not contain duplicates');
assert.strictEqual(new Set(detailedContracts.map((contract) => contract.name)).size, detailedContracts.length,
    'detailed contract names must be unique');
assert.strictEqual(detailedNames.size, registeredNames.size,
    'every registered built-in must have a detailed contract');
for (const name of registeredNames) {
    assert.strictEqual(detailedNames.has(name), true, `${name}: detailed contract is required`);
}
for (const contract of detailedContracts) {
    assert.strictEqual(registeredNames.has(contract.name), true, `${contract.name}: must be registered`);
    assert.strictEqual(contract.parameters.length >= 0, true, `${contract.name}: parameters must be an array`);
    for (const parameter of contract.parameters) {
        assert.strictEqual(parameter.categories.length > 0, true, `${contract.name}.${parameter.name}: categories are required`);
        assert.strictEqual(parameter.nullBehavior.length > 0, true, `${contract.name}.${parameter.name}: nullBehavior is required`);
    }
}

/** Generate symbolic boundary combinations from the YAML contract table. */
function generateContractCases(contract: DetailedContract): GeneratedCase[] {
    if (contract.parameters.length === 0) return [{ functionName: contract.name, arguments: [], omitted: [], oracle: [] }];
    const cases: GeneratedCase[] = [];
    const recurse = (index: number, args: string[], omitted: string[]): void => {
        if (index === contract.parameters.length) {
            const firstOmitted = args.indexOf('omitted');
            if (firstOmitted >= 0 && args.slice(firstOmitted).some((category) => category !== 'omitted')) return;
            const parameterOracle = contract.parameters.map((parameter, parameterIndex) => {
                const category = args[parameterIndex];
                if (category !== 'null') return 'normal';
                if (/propagate-null/.test(parameter.nullBehavior)) return 'return-null';
                if (/or-null/.test(parameter.nullBehavior)) return 'function-specific';
                if (/error-(?:5|13|94)|invalid-argument|object-error|type-mismatch/.test(parameter.nullBehavior)) return 'raise-error';
                return 'function-specific';
            });
            const overallOracle = parameterOracle.includes('raise-error')
                ? 'raise-error'
                : parameterOracle.includes('return-null') ? 'return-null' : 'normal';
            const oracle = parameterOracle.map(() => overallOracle);
            cases.push({ functionName: contract.name, arguments: args, omitted, oracle });
            return;
        }
        const parameter = contract.parameters[index];
        for (const category of parameter.categories) {
            recurse(index + 1, [...args, category], omitted);
        }
        if (parameter.optional) recurse(index + 1, [...args, 'omitted'], [...omitted, parameter.name]);
    };
    recurse(0, [], []);
    return cases;
}

const generatedCases = detailedContracts.flatMap(generateContractCases);
assert.strictEqual(generatedCases.length > detailedContracts.length, true,
    'contract table must generate boundary combinations');
assert.strictEqual(new Set(generatedCases.map((testCase) => testCase.functionName)).size, detailedContracts.length,
    'generated combinations must cover every detailed contract');
assert.strictEqual(generatedCases.every((testCase) => testCase.oracle.length === testCase.arguments.length), true,
    'every generated combination must have an oracle classification');

const categoryValues: Record<string, string> = {
    'all-variant-types': 'Empty', array: 'Array(1)', 'array-element': '1', binary: '0', boolean: 'False',
    'boolean-value-pairs': 'False, "no", True, "yes"', byte: '1', custom: '"0"', currency: '1',
    'currency-array': 'Array(1)', date: '#2020-01-01#', 'date-format-constant': '1', 'date-part': '"d"',
    'date-string': '"2020-01-01"', double: '1', 'double-array': 'Array(1)', empty: 'Empty', 'empty-array': 'Array()',
    enum: '1', 'error-value': 'CVErr(5)', fraction: '1.5', integer: '1', 'integer-array': 'Array(1)', invalid: '"invalid"',
    negative: '-1', null: 'Null', 'null-element': 'Null', number: '1', nothing: 'Nothing', object: 'Nothing',
    omitted: 'Empty', one: '1', 'out-of-range': '999', overflow: '1E+309', positive: '1',
    'radix-string': '"&H10"', 'numeric-string': '"123"', scalar: '1', standard: '"General Number"', string: '"abc"', 'string-array': 'Array("a")',
    suffix: '"1!"', text: '1', 'time-string': '"12:00:00"', unicode: '"あ"', variable: '1', 'variant-array': 'Array(1)',
    'week-constant': '1', 'weekday-constant': '1', zero: '0',
};
for (const testCase of generatedCases) {
    for (const category of testCase.arguments) {
        assert.strictEqual(categoryValues[category] !== undefined, true,
            `${testCase.functionName}: no representative VBA value for ${category}`);
    }
}
const generatedSources = generatedCases.map((testCase) => {
    const firstOmitted = testCase.arguments.findIndex((category, index) =>
        category === 'omitted' && testCase.arguments.slice(index).every((rest) => rest === 'omitted'));
    const included = firstOmitted < 0 ? testCase.arguments : testCase.arguments.slice(0, firstOmitted);
    return `${testCase.functionName}(${included.map((category) => categoryValues[category]).join(', ')})`;
});
assert.strictEqual(generatedSources.length, generatedCases.length, 'every generated combination must render to VBA');
const oracleCounts = generatedCases.flatMap((testCase) => testCase.oracle).reduce<Record<string, number>>((counts, kind) => {
    counts[kind] = (counts[kind] ?? 0) + 1;
    return counts;
}, {});

const executableNames = new Set([
    'abs', 'atn', 'cbool', 'cbyte', 'ccur', 'cdate', 'cdec', 'cdbl', 'cint', 'clng', 'clnglng', 'csng', 'cstr',
    'cos', 'dateadd', 'datediff', 'datepart', 'datevalue', 'day', 'exp', 'fix', 'format', 'formatcurrency',
    'formatdatetime', 'formatnumber', 'formatpercent', 'hour', 'instrrev', 'int', 'join', 'lcase',
    'left', 'leftb', 'len', 'lenb', 'log', 'ltrim', 'mid', 'midb', 'mirr', 'minute', 'month', 'npv', 'replace',
    'right', 'rightb', 'round', 'rtrim', 'sgn', 'sin', 'split', 'sqr', 'strcomp', 'strconv',
    'strreverse', 'tan', 'trim', 'ucase', 'ubound', 'lbound', 'weekday', 'year',
]);
const oracleProbes: { source: string; expected: 'return-null' | 'raise-error' }[] = [];
const oracleProbeCounts = new Map<string, number>();
for (let index = 0; index < generatedCases.length; index++) {
    const testCase = generatedCases[index];
    if (!executableNames.has(testCase.functionName)) continue;
    if (testCase.arguments.filter((category) => category === 'null').length > 1) continue;
    const expected = testCase.oracle.includes('return-null') ? 'return-null' :
        testCase.oracle.includes('raise-error') ? 'raise-error' : undefined;
    const key = `${testCase.functionName}:${expected}`;
    if (expected !== undefined && (oracleProbeCounts.get(key) ?? 0) < 8) {
        oracleProbeCounts.set(key, (oracleProbeCounts.get(key) ?? 0) + 1);
        oracleProbes.push({ source: generatedSources[index], expected });
    }
}
for (const probe of oracleProbes) {
    const source = String.raw`Function Probe() As Variant
    Probe = ${probe.source}
End Function`;
    const ev = evalVBASingle(source);
    let value: any;
    let error: any;
    try { value = ev.callProcedure('Probe', []); } catch (caught) { error = caught; }
    if (probe.expected === 'return-null') {
        assert.strictEqual(error, undefined, `${probe.source}: unexpected generated error`);
        assert.strictEqual(value, vbaNull, `${probe.source}: generated Null oracle`);
    } else {
        assert.strictEqual(error !== undefined, true, `${probe.source}: generated error oracle`);
    }
}

const normalReturnTypes: Record<string, 'number' | 'string' | 'boolean' | 'object'> = {
    abs: 'number', atn: 'number', cbool: 'object', cbyte: 'number', cdbl: 'number', cint: 'number', clng: 'number',
    csng: 'number', exp: 'number', fix: 'number', int: 'number', len: 'number', lenb: 'number', log: 'number',
    lcase: 'string', left: 'string', leftb: 'string', mid: 'string', midb: 'string', right: 'string', rightb: 'string',
    rtrim: 'string', ltrim: 'string', trim: 'string', ucase: 'string', strreverse: 'string', strcomp: 'number',
    sgn: 'number', sin: 'number', sqr: 'number', tan: 'number', val: 'number',
};
for (const contract of exactNormalContracts) {
    const source = String.raw`Function Probe() As Variant
    Probe = ${contract.expression}
End Function`;
    const ev = evalVBASingle(source);
    let actual: any;
    let error: any;
    try { actual = ev.callProcedure('Probe', []); } catch (caught) { error = caught; }
    if (contract.error !== undefined) {
        assert.strictEqual(error?.number, contract.error, `${contract.function}: exact error contract`);
    } else {
        assert.strictEqual(error, undefined, `${contract.function}: unexpected exact contract error`);
        if (contract.tolerance && typeof contract.expected === 'number') {
            assertClose(actual as number, contract.expected, { ...contract.tolerance, message: `${contract.function}: exact normal contract` });
        } else {
            assert.strictEqual(actual, contract.expected, `${contract.function}: exact normal contract`);
        }
    }
}
const normalProbes = new Map<string, { source: string; expected: string }>();
for (let index = 0; index < generatedCases.length; index++) {
    const testCase = generatedCases[index];
    const expectedType = normalReturnTypes[testCase.functionName];
    if (!expectedType || !testCase.oracle.includes('normal') || normalProbes.has(testCase.functionName)) continue;
    normalProbes.set(testCase.functionName, { source: generatedSources[index], expected: expectedType });
}
for (const probe of normalProbes.values()) {
    const source = String.raw`Function Probe() As Variant
    Probe = ${probe.source}
End Function`;
    const ev = evalVBASingle(source);
    let value: any;
    let error: any;
    try { value = ev.callProcedure('Probe', []); } catch (caught) { error = caught; }
    assert.strictEqual(error, undefined, `${probe.source}: unexpected normal-case error`);
    const actualType = typeof value === 'bigint' ? 'number' : typeof value;
    assert.strictEqual(actualType, probe.expected, `${probe.source}: generated normal return type`);
}

type Case = { name: string; expression: string; expected?: any; error?: number };
const cases: Case[] = [
    { name: 'IsNull(Null)', expression: 'IsNull(Null)', expected: -1 },
    { name: 'IsNull(Empty)', expression: 'IsNull(Empty)', expected: 0 },
    { name: 'IsEmpty(Empty)', expression: 'IsEmpty(Empty)', expected: -1 },
    { name: 'IsEmpty(Null)', expression: 'IsEmpty(Null)', expected: 0 },
    { name: 'IsNumeric(Empty)', expression: 'IsNumeric(Empty)', expected: -1 },
    { name: 'IsNumeric(Null)', expression: 'IsNumeric(Null)', expected: 0 },
    { name: 'Len(Empty)', expression: 'Len(Empty)', expected: 0 },
    { name: 'Len(Null)', expression: 'Len(Null)', expected: vbaNull },
    { name: 'LCase(Empty)', expression: 'LCase(Empty)', expected: '' },
    { name: 'LCase(Null)', expression: 'LCase(Null)', expected: vbaNull },
    { name: 'Trim(Empty)', expression: 'Trim(Empty)', expected: '' },
    { name: 'Trim(Null)', expression: 'Trim(Null)', expected: vbaNull },
    { name: 'CStr(Empty)', expression: 'CStr(Empty)', expected: '' },
    { name: 'CStr(Null)', expression: 'CStr(Null)', error: 94 },
    { name: 'CBool(Empty)', expression: 'CBool(Empty)', expected: 0 },
    { name: 'CByte(Empty)', expression: 'CByte(Empty)', expected: 0 },
    { name: 'CDbl(Empty)', expression: 'CDbl(Empty)', expected: 0 },
    { name: 'CStr(CCur(Empty))', expression: 'CStr(CCur(Empty))', expected: '0' },
    { name: 'CStr(CDec(Empty))', expression: 'CStr(CDec(Empty))', expected: '0' },
    { name: 'TypeName(CCur(Empty))', expression: 'TypeName(CCur(Empty))', expected: 'Currency' },
    { name: 'TypeName(CDec(Empty))', expression: 'TypeName(CDec(Empty))', expected: 'Decimal' },
    { name: 'CInt(Empty)', expression: 'CInt(Empty)', expected: 0 },
    { name: 'CInt(Null)', expression: 'CInt(Null)', error: 94 },
    { name: 'CLng(Empty)', expression: 'CLng(Empty)', expected: 0 },
    { name: 'CSng(Empty)', expression: 'CSng(Empty)', expected: 0 },
    { name: 'CVar(Null)', expression: 'CVar(Null)', expected: vbaNull },
    { name: 'TypeName(CVDate(Empty))', expression: 'TypeName(CVDate(Empty))', expected: 'Date' },
    { name: 'TypeName(CVErr(5))', expression: 'TypeName(CVErr(5))', expected: 'Error' },
    { name: 'CLngLng(Empty)', expression: 'CLngLng(Empty)', error: 13 },
    { name: 'Hex(255)', expression: 'Hex(255)', expected: 'FF' },
    { name: 'Oct(255)', expression: 'Oct(255)', expected: '377' },
    { name: 'Val numeric text', expression: 'Val("12.5")', expected: 12.5 },
    { name: 'Asc A', expression: 'Asc("A")', expected: 65 },
    { name: 'AscB A', expression: 'AscB("A")', expected: 65 },
    { name: 'AscW A', expression: 'AscW("A")', expected: 65 },
    { name: 'Chr 65', expression: 'Chr(65)', expected: 'A' },
    { name: 'ChrB 65', expression: 'ChrB(65)', expected: 'A' },
    { name: 'ChrW 65', expression: 'ChrW(65)', expected: 'A' },
    { name: 'Split Null delimiter', expression: 'Split("a b", Null)', error: 94 },
    { name: 'Join Empty delimiter', expression: 'Join(Array("a", "b"), Empty)', expected: 'ab' },
    { name: 'Join Null delimiter', expression: 'Join(Array("a", "b"), Null)', expected: vbaNull },
    { name: 'InStr Null text', expression: 'InStr(Null, "a")', expected: vbaNull },
    { name: 'InStrRev Null text', expression: 'InStrRev(Null, "a")', expected: vbaNull },
    { name: 'InStr basic overload', expression: 'InStr("abc", "b")', expected: 2 },
    { name: 'Mid omitted length', expression: 'Mid("abc", 2)', expected: 'bc' },
    { name: 'Left zero length', expression: 'Left("abc", 0)', expected: '' },
    { name: 'Right zero length', expression: 'Right("abc", 0)', expected: '' },
    { name: 'LTrim spaces', expression: 'LTrim("  abc")', expected: 'abc' },
    { name: 'RTrim spaces', expression: 'RTrim("abc  ")', expected: 'abc' },
    { name: 'LenB empty', expression: 'LenB(Empty)', expected: 0 },
    { name: 'LeftB zero length', expression: 'LeftB("abc", 0)', expected: '' },
    { name: 'RightB zero length', expression: 'RightB("abc", 0)', expected: '' },
    { name: 'MidB omitted length', expression: 'MidB("abc", 2)', expected: 'abc' },
    { name: 'Space zero', expression: 'Space(0)', expected: '' },
    { name: 'String repeat', expression: 'String(2, "x")', expected: 'xx' },
    { name: 'StrComp equal', expression: 'StrComp("a", "a")', expected: 0 },
    { name: 'StrConv upper', expression: 'StrConv("abc", 1)', expected: 'ABC' },
    { name: 'StrReverse text', expression: 'StrReverse("abc")', expected: 'cba' },
    { name: 'Filter string array', expression: 'TypeName(Filter(Array("a"), "a"))', expected: 'Variant()' },
    { name: 'Replace Null find', expression: 'Replace("abc", Null, "x")', expected: vbaNull },
    { name: 'Round omitted optional', expression: 'Round(2.5)', expected: 2 },
    { name: 'Round Empty optional', expression: 'Round(2.5, Empty)', expected: 2 },
    { name: 'Round Null optional', expression: 'Round(2.5, Null)', error: 94 },
    { name: 'Abs(Null)', expression: 'Abs(Null)', expected: vbaNull },
    { name: 'Abs(-2)', expression: 'Abs(-2)', expected: 2 },
    { name: 'Int(-1.2)', expression: 'Int(-1.2)', expected: -2 },
    { name: 'Fix(-1.2)', expression: 'Fix(-1.2)', expected: -1 },
    { name: 'Sgn(-2)', expression: 'Sgn(-2)', expected: -1 },
    { name: 'Sqr(9)', expression: 'Sqr(9)', expected: 3 },
    { name: 'Exp(0)', expression: 'Exp(0)', expected: 1 },
    { name: 'Log(1)', expression: 'Log(1)', expected: 0 },
    { name: 'Sin(0)', expression: 'Sin(0)', expected: 0 },
    { name: 'Cos(0)', expression: 'Cos(0)', expected: 1 },
    { name: 'Tan(0)', expression: 'Tan(0)', expected: 0 },
    { name: 'Atn(0)', expression: 'Atn(0)', expected: 0 },
    { name: 'Sqr(Empty)', expression: 'Sqr(Empty)', expected: 0 },
    { name: 'Year(DateValue(Empty))', expression: 'Year(DateValue(Empty))', expected: 1899 },
    { name: 'Day(DateSerial)', expression: 'Day(DateSerial(2020, 1, 2))', expected: 2 },
    { name: 'Hour(DateSerial)', expression: 'Hour(DateSerial(2020, 1, 2))', expected: 0 },
    { name: 'Minute(DateSerial)', expression: 'Minute(DateSerial(2020, 1, 2))', expected: 0 },
    { name: 'Month(DateSerial)', expression: 'Month(DateSerial(2020, 1, 2))', expected: 1 },
    { name: 'Second(DateSerial)', expression: 'Second(DateSerial(2020, 1, 2))', expected: 0 },
    { name: 'DateAdd day', expression: 'Day(DateAdd("d", 1, DateSerial(2020, 1, 1)))', expected: 2 },
    { name: 'DateDiff day', expression: 'DateDiff("d", DateSerial(2020, 1, 1), DateSerial(2020, 1, 2))', expected: 1 },
    { name: 'DatePart year', expression: 'DatePart("yyyy", DateSerial(2020, 1, 2))', expected: 2020 },
    { name: 'TypeName(TimeSerial)', expression: 'TypeName(TimeSerial(1, 2, 3))', expected: 'Date' },
    { name: 'Hour(TimeValue)', expression: 'Hour(TimeValue("12:34:56"))', expected: 12 },
    { name: 'MonthName January', expression: 'MonthName(1)', expected: 'January' },
    { name: 'Weekday(DateSerial)', expression: 'Weekday(DateSerial(2020, 1, 5))', expected: 1 },
    { name: 'WeekdayName Sunday', expression: 'WeekdayName(1)', expected: 'Sunday' },
    { name: 'DateValue(Null)', expression: 'DateValue(Null)', expected: vbaNull },
    { name: 'MIRR Null finance rate', expression: 'MIRR(Array(-100, 100), Null, 0.1)', error: 13 },
    { name: 'NPV scalar value argument', expression: 'NPV(0.1, 100)', error: 13 },
    { name: 'FV Null type', expression: 'FV(0.01, 12, -100, 0, Null)', error: 94 },
    { name: 'PV Null periods', expression: 'PV(0.01, Null, -100)', error: 94 },
    { name: 'Pmt invalid periods', expression: 'Pmt(0.01, 0, 1000)', error: 5 },
    { name: 'NPer Null payment', expression: 'NPer(0.01, Null, 1000, 0)', error: 94 },
    { name: 'Rate Null future value', expression: 'Rate(12, -100, 100, Null)', error: 94 },
    { name: 'IRR scalar value argument', expression: 'IRR(100)', error: 13 },
    { name: 'Format numeric', expression: 'Format(0.5, "0")', expected: '1' },
    { name: 'TypeName(FormatCurrency)', expression: 'TypeName(FormatCurrency(1))', expected: 'String' },
    { name: 'TypeName(FormatNumber)', expression: 'TypeName(FormatNumber(1))', expected: 'String' },
    { name: 'TypeName(FormatPercent)', expression: 'TypeName(FormatPercent(0.5))', expected: 'String' },
    { name: 'TypeName(FormatDateTime)', expression: 'TypeName(FormatDateTime(DateSerial(2020, 1, 1)))', expected: 'String' },
    { name: 'Choose Null index', expression: 'Choose(Null, "a", "b")', expected: vbaNull },
    { name: 'Switch first true pair', expression: 'Switch(False, "no", True, "yes")', expected: 'yes' },
    { name: 'Array no arguments', expression: 'TypeName(Array())', expected: 'Variant()' },
    { name: 'LBound scalar argument', expression: 'LBound(1)', error: 9 },
    { name: 'UBound scalar argument', expression: 'UBound(1)', error: 9 },
    { name: 'LBound(array)', expression: 'LBound(Array(1, 2))', expected: 0 },
    { name: 'UBound(array)', expression: 'UBound(Array(1, 2))', expected: 1 },
    { name: 'IsObject(Nothing)', expression: 'IsObject(value)', expected: -1 },
];

// The executable matrix covers representative pure/value functions. Host-facing
// contracts remain declarative because executing them would open dialogs or
// mutate host state; their dedicated suites are the verification source.

for (const testCase of cases) {
    const source = String.raw`Function Probe() As Variant
    Dim value As Object
    Probe = ${testCase.expression}
End Function`;
    const ev = evalVBASingle(source);
    let value: any;
    let error: any;
    try { value = ev.callProcedure('Probe', []); } catch (caught) { error = caught; }
    if (testCase.error !== undefined) {
        assert.strictEqual(error?.number, testCase.error, `${testCase.name}: Error ${testCase.error}`);
    } else {
        assert.strictEqual(error, undefined, `${testCase.name}: unexpected error`);
        assert.strictEqual(value, testCase.expected, `${testCase.name}: value`);
    }
}

console.log(`[PASS] Built-in value contract inventory (${registeredNames.size} functions, ${cases.length} executable boundary cases, ${generatedCases.length} generated combinations, ${oracleProbes.length} boundary oracle probes, ${normalProbes.size} normal return-type probes, oracle=${JSON.stringify(oracleCounts)})`);
