import { Statement, ProcedureDeclaration } from '../engine/parser';
import { buildScopedSymbolTable } from './symbol-table';

export interface SignatureInfo {
    label: string;
    parameters: string[];
}

export interface SignatureHelpResult {
    signature: SignatureInfo;
    activeParameter: number;
}

/** `detail` 文字列からパラメーター名リストを抽出する。
 *  例: "MsgBox(prompt, [buttons], [title]) As Long" → ["prompt", "[buttons]", "[title]"]
 *  ネストした [] 内のカンマ（例: InStr([start,] string1, ...）は無視する。
 */
function parseParamsFromDetail(detail: string): string[] {
    const match = detail.match(/\(([^)]*)\)/);
    if (!match || match[1].trim() === '') return [];

    const inner = match[1];
    const params: string[] = [];
    let depth = 0;
    let start = 0;

    for (let i = 0; i < inner.length; i++) {
        if (inner[i] === '[') depth++;
        else if (inner[i] === ']') depth--;
        else if (inner[i] === ',' && depth === 0) {
            params.push(inner.slice(start, i).trim());
            start = i + 1;
        }
    }
    const last = inner.slice(start).trim();
    if (last) params.push(last);
    return params;
}

/** VBA 組み込み関数のシグネチャ一覧（completion-provider.ts の standardFunctions と同じ関数群）。
 *  detail 文字列をそのまま label として使い、パラメーターは parseParamsFromDetail で導出する。
 */
const BUILTIN_DETAILS: string[] = [
    'Len(string) As Long',
    'LCase(string) As String',
    'UCase(string) As String',
    'Trim(string) As String',
    'LTrim(string) As String',
    'RTrim(string) As String',
    'Left(string, length) As String',
    'Right(string, length) As String',
    'Mid(string, start, [length]) As String',
    'Str(number) As String',
    'Val(string) As Double',
    'InStr([start,] string1, string2) As Long',
    'InStrRev(string1, string2, [start]) As Long',
    'Replace(expression, find, replace) As String',
    'Split(expression, [delimiter]) As String()',
    'Join(sourceArray, [delimiter]) As String',
    'String(number, character) As String',
    'Space(number) As String',
    'Int(number) As Long',
    'Fix(number) As Long',
    'Round(expression, [decimals]) As Double',
    'Abs(number) As Variant',
    'Sqr(number) As Double',
    'CDbl(expression) As Double',
    'CLng(expression) As Long',
    'CInt(expression) As Integer',
    'CStr(expression) As String',
    'CBool(expression) As Boolean',
    'CCur(expression) As Currency',
    'CSng(expression) As Single',
    'CDate(expression) As Date',
    'IsEmpty(expression) As Boolean',
    'IsNumeric(expression) As Boolean',
    'IsNull(expression) As Boolean',
    'IsDate(expression) As Boolean',
    'IsObject(expression) As Boolean',
    'IsArray(expression) As Boolean',
    'TypeName(varname) As String',
    'VarType(varname) As Integer',
    'UBound(array, [dimension]) As Long',
    'LBound(array, [dimension]) As Long',
    'Array(arglist) As Variant',
    'MsgBox(prompt, [buttons], [title]) As Long',
    'InputBox(prompt, [title], [default]) As String',
    'CreateObject(class) As Object',
    'GetObject([pathname], [class]) As Object',
    'Format(expression, [format]) As String',
    'Now() As Date',
    'Date() As Date',
    'Time() As Date',
    'DateSerial(year, month, day) As Date',
    'DateAdd(interval, number, date) As Date',
    'DateDiff(interval, date1, date2) As Long',
    'DateValue(date) As Date',
    'Year(date) As Integer',
    'Month(date) As Integer',
    'Day(date) As Integer',
    'Hour(date) As Integer',
    'Minute(date) As Integer',
    'Second(date) As Integer',
    'Weekday(date, [firstdayofweek]) As Integer',
    'Timer() As Single',
    'Shell(pathname, [windowstyle]) As Long',
    'Environ(expression) As String',
    'Chr(charcode) As String',
    'Asc(string) As Integer',
    'Hex(number) As String',
    'Oct(number) As String',
    'Error([errornumber]) As String',
    'IIf(expr, truepart, falsepart) As Variant',
    'Choose(index, choice1, [choice2]) As Variant',
    'Switch(expr1, value1, [expr2, value2]) As Variant',
];

function buildBuiltinSignatures(): Map<string, SignatureInfo> {
    const map = new Map<string, SignatureInfo>();
    for (const detail of BUILTIN_DETAILS) {
        const parenIdx = detail.indexOf('(');
        if (parenIdx < 0) continue;
        const name = detail.slice(0, parenIdx).trim().toLowerCase();
        map.set(name, {
            label: detail,
            parameters: parseParamsFromDetail(detail),
        });
    }
    return map;
}

const BUILTIN_SIGNATURES = buildBuiltinSignatures();

/** カーソル左のテキストを走査し、現在の関数呼び出しの名前と activeParameter を返す。
 *  ネストした括弧を考慮し、最も外側の未閉じ `(` を対象の呼び出しとみなす。
 *  文字列リテラル内の `(`, `,` は無視する。
 */
export function findCallContext(lineText: string, character: number): { name: string; activeParameter: number } | null {
    const text = lineText.slice(0, character);

    let depth = 0;
    let activeParameter = 0;
    let inString = false;

    for (let i = text.length - 1; i >= 0; i--) {
        const ch = text[i];

        // 文字列リテラルの境界（簡易判定: " を数えて偶奇で判定）
        if (ch === '"') {
            // 直前のバックスラッシュエスケープは VBA にないので単純に反転
            inString = !inString;
            continue;
        }
        if (inString) continue;

        if (ch === ')') {
            depth++;
        } else if (ch === '(') {
            if (depth === 0) {
                // この `(` が対象の呼び出し
                const name = extractFunctionName(text, i);
                if (!name) return null;
                return { name, activeParameter };
            }
            depth--;
        } else if (ch === ',' && depth === 0) {
            activeParameter++;
        }
    }

    return null;
}

/** `(` の直前にある識別子を取り出す */
function extractFunctionName(text: string, parenIdx: number): string | null {
    let end = parenIdx;
    // 空白をスキップ
    while (end > 0 && text[end - 1] === ' ') end--;
    if (end === 0) return null;

    let start = end;
    while (start > 0 && /[a-zA-Z0-9_]/.test(text[start - 1])) start--;

    if (start === end) return null;
    return text.slice(start, end);
}

export class SignatureHelpProvider {
    getSignatureHelp(
        statements: Statement[],
        source: string,
        line: number,
        character: number,
    ): SignatureHelpResult | null {
        const lines = source.split('\n');
        const lineText = lines[line] ?? '';

        // コメント行は対象外
        const trimmed = lineText.trimStart();
        if (trimmed.startsWith("'")) return null;

        const ctx = findCallContext(lineText, character);
        if (!ctx) return null;

        const nameLower = ctx.name.toLowerCase();

        // 1. ユーザー定義シンボルを優先
        const table = buildScopedSymbolTable(statements);
        const entry = table.moduleSymbols.get(nameLower);
        if (entry?.displayText) {
            const sig: SignatureInfo = {
                label: entry.displayText,
                parameters: this.paramsFromProc(statements, nameLower),
            };
            return { signature: sig, activeParameter: ctx.activeParameter };
        }

        // 2. 組み込み関数にフォールバック
        const builtin = BUILTIN_SIGNATURES.get(nameLower);
        if (builtin) {
            return { signature: builtin, activeParameter: ctx.activeParameter };
        }

        return null;
    }

    private paramsFromProc(statements: Statement[], nameLower: string): string[] {
        for (const stmt of statements) {
            if (stmt.type !== 'ProcedureDeclaration') continue;
            const proc = stmt as ProcedureDeclaration;
            if (proc.name.name.toLowerCase() !== nameLower) continue;
            return proc.parameters.map(p => {
                const pname = p.name as unknown as string;
                const ptype = p.paramType || 'Variant';
                const byref = p.isByVal ? 'ByVal ' : '';
                return `${byref}${pname} As ${ptype}`;
            });
        }
        return [];
    }
}
