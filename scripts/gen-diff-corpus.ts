/**
 * 実 VBA 差分テスト用コーパスジェネレーター。
 *
 * 「同一の .bas を実 Excel と vba-runner の両方で実行して出力を diff する」ための
 * 自己完結コーパス（diff-vba/DiffMain.bas + DiffCorpus_NNN.bas + manifest.json）を生成する。
 * 式はコードに焼き込まれ、実行結果は locale 非依存の正規化形式（Canon）で
 * vba-diff-results.txt に 1 式 1 行で書き出される。
 *
 * 実行: npx tsx scripts/gen-diff-corpus.ts
 *
 * 生成時に各式を vba-runner のパーサーで事前スクリーニングし、パースできない式は
 * モジュールから除外して manifest に SKIPPED_PARSE として記録する
 * （= vba-runner 側のパーサー欠落そのものの記録になる）。
 *
 * 設計上の除外: 非決定要素（Now/Timer/Rnd）、ロケール依存（名前付き書式・曖昧な
 * CDate 文字列・mmm/ddd パターン）、サンドボックス依存（ファイル I/O・MsgBox 等）。
 * 浮動小数の実装依存関数（Log/Exp/三角・財務）は式側で Format による桁丸めを掛ける。
 */
import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';

const OUT_DIR = 'diff-vba';
const CASES_PER_MODULE = 200;

interface Case { id: string; expr: string }
const cases: Case[] = [];
const skippedParse: Case[] = [];
const counters = new Map<string, number>();

function add(cat: string, expr: string): void {
    const n = (counters.get(cat) ?? 0) + 1;
    counters.set(cat, n);
    const id = `${cat}_${String(n).padStart(4, '0')}`;
    // vba-runner のパーサーで事前スクリーニング
    const snippet = `Sub T()\nDim v As Variant\nv = ${expr}\nEnd Sub`;
    try {
        new Parser(new Lexer(snippet).tokenize(), { sourceLines: snippet.split('\n') }).parse();
        cases.push({ id, expr });
    } catch {
        skippedParse.push({ id, expr });
    }
}

// ---------------------------------------------------------------------------
// 1. 演算子 × 型の行列
// ---------------------------------------------------------------------------
{
    const vals = ['2', '3.5', '-1', '0', '32767', '2147483647', '"7"', '"3.5"', '"abc"',
        'True', 'False', '#2024/03/15#', 'Null', 'Empty'];
    const ops = ['+', '-', '*', '/', '\\', 'Mod', '&', '=', '<'];
    for (const op of ops) {
        for (const a of vals) {
            for (const b of vals) {
                add('OP', `(${a}) ${op} (${b})`);
            }
        }
    }
    // ^ は組み合わせ爆発が激しいので代表のみ
    for (const [a, b] of [['2', '10'], ['2', '0.5'], ['-2', '3'], ['-2', '0.5'], ['0', '0'],
        ['2', '-2'], ['"2"', '"3"'], ['True', '2'], ['Null', '2'], ['2', 'Null'], ['Empty', '2']]) {
        add('OP', `(${a}) ^ (${b})`);
    }
}

// ---------------------------------------------------------------------------
// 2. 変換関数 × 境界値グリッド
// ---------------------------------------------------------------------------
{
    const fns = ['CInt', 'CLng', 'CDbl', 'CSng', 'CCur', 'CByte', 'CBool', 'CDate', 'CStr', 'Int', 'Fix'];
    const vals = ['2.5', '3.5', '-2.5', '-3.5', '0.5', '1.5', '32767.4', '32767.6', '255.5',
        '2147483647', '32768', '-32769', '256',
        '"2.5"', '"  12  "', '"&H10"', '"&HFF"', '"&O17"', '"1E3"', '"1,234"', '"abc"', '""',
        'True', 'False', '#2024/03/15#', '#2024/03/15 13:45:30#', 'Null', 'Empty'];
    for (const fn of fns) {
        for (const v of vals) {
            add('CONV', `${fn}(${v})`);
        }
    }
}

// ---------------------------------------------------------------------------
// 3. 文字列関数グリッド
// ---------------------------------------------------------------------------
{
    const strCases = [
        'Left("abcde", 2)', 'Left("abcde", 0)', 'Left("abcde", 10)', 'Left("", 3)',
        'Right("abcde", 2)', 'Right("abcde", 0)', 'Right("abcde", 10)',
        'Mid("abcde", 2)', 'Mid("abcde", 2, 2)', 'Mid("abcde", 5, 10)', 'Mid("abcde", 6)',
        'InStr("abcabc", "b")', 'InStr(3, "abcabc", "b")', 'InStr("abc", "")', 'InStr("", "a")',
        'InStr(1, "ABC", "b", 1)', 'InStr(1, "ABC", "b", 0)',
        'InStrRev("abcabc", "b")', 'InStrRev("abcabc", "b", 3)',
        'Replace("aaa", "a", "bb")', 'Replace("aaa", "a", "b", 2)', 'Replace("aaa", "a", "b", 1, 2)',
        'Replace("abc", "", "x")',
        'StrComp("a", "B", 0)', 'StrComp("a", "B", 1)', 'StrComp("a", "a", 0)',
        'String(3, "ab")', 'String(0, "x")', 'String(3, 65)', 'String(3, 321)',
        'Space(0)', 'Space(3) & "|"',
        'Len("")', 'Len("abc")', 'LenB("abc")',
        'Asc("Ab")', 'AscW(ChrW(12354))', 'Asc(ChrW(12354))',
        'Chr(65)', 'Chr(255)', 'ChrW(12354)', 'ChrW(65535)',
        'UCase("aBc")', 'LCase("AbC")', 'UCase(ChrW(228))',
        'Trim("  a  b  ")', 'LTrim("  ab")', 'RTrim("ab  ")', 'Trim(Chr(9) & "a")',
        'StrReverse("abc")', 'StrReverse(ChrW(12354) & ChrW(12356))',
        'Len(ChrW(12354) & ChrW(12356))', 'Mid(ChrW(12354) & ChrW(12356), 2, 1)',
        'StrConv("abc", 1)', 'StrConv("ABC", 2)', 'StrConv("hello world", 3)',
        'StrConv(ChrW(12354), 16)', 'StrConv(ChrW(12450), 32)',
        'StrConv("abc", 4)', 'StrConv(ChrW(65345), 8)',
        'Str(1.5)', 'Str(-1.5)', 'Str(0)',
        'Val("1.5abc")', 'Val("  1 2 3  ")', 'Val("&H10")', 'Val("&O17")', 'Val("1E3")', 'Val(".")', 'Val("abc")',
        'Hex(255)', 'Hex(-1)', 'Hex(0)', 'Oct(8)', 'Oct(-1)',
        '"abc" Like "a*"', '"abc" Like "A*"', '"abc" Like "?b?"', '"a1c" Like "a#c"',
        '"abc" Like "[a-c][a-c][a-c]"', '"abc" Like "[!x-z]bc"', '"[x]" Like "[[]x]"', '"" Like "*"', '"" Like ""',
    ];
    for (const e of strCases) add('STR', e);

    const fmtPatterns = ['"0"', '"0.00"', '"#,##0.00"', '"0%"', '"0.00E+00"', '"000"', '"#.##"'];
    const fmtVals = ['0', '0.5', '-0.5', '1234.567', '0.045', '2.675'];
    for (const p of fmtPatterns) {
        for (const v of fmtVals) {
            add('STR', `Format(${v}, ${p})`);
        }
    }
    const dateFmtPatterns = ['"yyyy-mm-dd"', '"yyyy/mm/dd hh:nn:ss"', '"hh:nn"', '"m/d/yyyy"', '"q"', '"y"', '"ww"'];
    for (const p of dateFmtPatterns) {
        add('STR', `Format(#2024/03/15 13:45:30#, ${p})`);
    }
}

// ---------------------------------------------------------------------------
// 4. 日付関数グリッド（カバレッジ暗部: 2 要素リテラル・DateAdd "ww"・firstdayofweek）
// ---------------------------------------------------------------------------
{
    const dateCases = [
        'DateSerial(2024, 13, 1)', 'DateSerial(2024, 0, 15)', 'DateSerial(2024, 2, 30)',
        'DateSerial(2024, 1, 60)', 'DateSerial(1899, 12, 31)', 'DateSerial(2024, -1, 1)',
        'TimeSerial(25, 0, 0)', 'TimeSerial(0, 90, 0)', 'TimeSerial(12, -30, 0)',
        'Year(#3/15#)', 'Month(#3/15#)', 'Day(#3/15#)',
        'Year(#12/31#)', 'Month(#1/2#)', 'Day(#1/2#)',
        'Hour(#13:45#)', 'Minute(#1:2:3#)', 'Second(#1:2:3#)', 'Hour(#3:15:00 PM#)',
        'Weekday(#2024/03/15#)', 'Weekday(#2024/03/15#, 2)', 'Weekday(#2024/03/15#, 7)',
        'DateDiff("ww", #2024/01/01#, #2024/03/15#)', 'DateDiff("ww", #2024/01/01#, #2024/03/15#, 2)',
        'DateDiff("w", #2024/01/01#, #2024/03/15#)',
        'DateDiff("q", #2024/01/01#, #2024/12/31#)', 'DateDiff("y", #2024/01/01#, #2024/03/15#)',
        'DatePart("ww", #2024/01/06#)', 'DatePart("ww", #2024/01/06#, 2)', 'DatePart("ww", #2024/01/06#, 1, 2)',
        'DatePart("q", #2024/03/15#)', 'DatePart("y", #2024/03/15#)',
        'CStr(#2024/03/15# + 1)', 'CStr(#2024/03/15# - #2024/03/10#)',
        'Format(DateSerial(2024, 2, 29), "yyyy-mm-dd")',
    ];
    for (const e of dateCases) add('DATE', e);
    const intervals = ['"yyyy"', '"q"', '"m"', '"ww"', '"w"', '"d"', '"h"', '"n"', '"s"'];
    for (const iv of intervals) {
        add('DATE', `Format(DateAdd(${iv}, 1, #2024/01/31#), "yyyy-mm-dd hh:nn:ss")`);
        add('DATE', `Format(DateAdd(${iv}, -1, #2024/03/31 12:00:00#), "yyyy-mm-dd hh:nn:ss")`);
    }
}

// ---------------------------------------------------------------------------
// 5. 数学関数（実装依存の超越関数は桁丸めを掛ける）
// ---------------------------------------------------------------------------
{
    const exact = [
        'Sqr(2)', 'Sqr(0)', 'Abs(-2.5)', 'Sgn(-0.1)', 'Sgn(0)',
        'Int(-2.7)', 'Fix(-2.7)', 'Int(2.7)', 'Fix(2.7)',
        'Round(2.5)', 'Round(3.5)', 'Round(-2.5)', 'Round(2.675, 2)', 'Round(0.125, 2)',
        '7 \\ 2', '-7 \\ 2', '7 \\ -2', '7.5 \\ 2', '7 Mod 3', '-7 Mod 3', '7 Mod -3', '7.5 Mod 2',
    ];
    for (const e of exact) add('MATH', e);
    const rounded = [
        'Log(2)', 'Exp(1)', 'Sin(1)', 'Cos(1)', 'Tan(1)', 'Atn(1)', 'Sqr(3)', '2 ^ 0.5',
    ];
    for (const e of rounded) add('MATH', `Format(${e}, "0.0000000000")`);
}

// ---------------------------------------------------------------------------
// 6. 財務関数（浮動小数の実装差を吸収するため桁丸め）
// ---------------------------------------------------------------------------
{
    const finCases = [
        'Pmt(0.05 / 12, 360, -100000)', 'Pmt(0, 12, -1200)',
        'FV(0.05 / 12, 120, -100)', 'PV(0.05 / 12, 120, -100)',
        'NPer(0.05 / 12, -100, 10000)', 'IPmt(0.05 / 12, 1, 360, -100000)',
        'PPmt(0.05 / 12, 1, 360, -100000)', 'SLN(10000, 1000, 5)',
        'SYD(10000, 1000, 5, 1)', 'DDB(10000, 1000, 5, 1)',
    ];
    for (const e of finCases) add('FIN', `Format(${e}, "0.00000")`);
}

// ---------------------------------------------------------------------------
// 7. エラーコード互換性（意図的に失敗する式）
// ---------------------------------------------------------------------------
{
    const errCases = [
        '1 / 0', '1 \\ 0', '1 Mod 0', '0 / 0',
        'CInt("x")', 'CByte(300)', 'CByte(-1)', 'CInt(40000)', 'CDate("xyz")',
        'CInt(Null)', 'CStr(Null)', 'Sqr(-1)', 'Log(0)', 'Log(-1)',
        'Left("a", -1)', 'Mid("abc", 0)', 'Mid("abc", -1)', 'String(-1, "x")', 'Space(-1)',
        'Chr(-1)', 'Chr(256)', 'ChrW(65536)', 'Asc("")',
        'StrComp("a", "b", 99)', 'Weekday(#2024/03/15#, 99)',
        '"abc" + 1', '"abc" - 1', 'True + "x"',
    ];
    for (const e of errCases) add('ERRP', e);
}

// ---------------------------------------------------------------------------
// 8. 深掘り第2弾: \ / Mod の Boolean 保持規則の境界（第1回で「左Boolean・右文字列」
//    規則を発見。右が文字列以外のとき・他の演算子でも起きるかを検証する）
// ---------------------------------------------------------------------------
{
    const boolVals = ['True', 'False'];
    const others = ['7', '"7"', '"3.5"', 'Empty', 'Null', '#2024/03/15#', '3.5', '2'];
    for (const b of boolVals) {
        for (const o of others) {
            add('BOOLPRES', `(${b}) \\ (${o})`);
            add('BOOLPRES', `(${o}) \\ (${b})`);
            add('BOOLPRES', `(${b}) Mod (${o})`);
            add('BOOLPRES', `(${o}) Mod (${b})`);
        }
    }
    // 他の算術演算子でも同じ非対称規則が起きるか（第1回は \/Mod でしか確認していない）
    const arithOps = ['+', '-', '*', '/'];
    for (const op of arithOps) {
        for (const b of boolVals) {
            add('BOOLPRES', `(${b}) ${op} ("7")`);
            add('BOOLPRES', `("7") ${op} (${b})`);
        }
    }
}

// ---------------------------------------------------------------------------
// 9. 深掘り第2弾: CDate の "H.N" 時刻解釈の境界値
//    （第1回は "2.5"/"3.5" のみ確認。無効値・ゼロ埋め・3桁以上・区切りなしを検証）
// ---------------------------------------------------------------------------
{
    const hnCases = [
        '"2.5"', '"3.5"', '"0.0"', '"23.59"', '"24.0"', '"23.60"',
        '"2.50"', '"02.05"', '"12.345"', '"5"', '"1.1"', '"9.9"',
        '"100.5"', '".5"', '"5."',
    ];
    for (const s of hnCases) {
        add('HNTIME', `CDate(${s})`);
        add('HNTIME', `(#2024/03/15#) + (${s})`);
        add('HNTIME', `(${s}) + (#2024/03/15#)`);
        // + 以外の演算子では時刻解釈されないはず（第1回で確認済みだが念のため回帰確認）
        add('HNTIME', `(#2024/03/15#) - (${s})`);
    }
}

// ---------------------------------------------------------------------------
// 10. 深掘り第2弾: CDate の "M,Y" 形式（第1回はタイムゾーン問題で検証不能だった
//     ため、現代に近い年で再検証する）
// ---------------------------------------------------------------------------
{
    const myCases = [
        ['3', '2024'], ['1', '2020'], ['12', '1999'], ['6', '2050'],
        ['1', '234'],  // 第1回の元ケース（タイムゾーン問題ありの参考用）
    ];
    for (const [m, y] of myCases) {
        add('MYDATE', `CDate("${m},${y}")`);
    }
}

// ---------------------------------------------------------------------------
// 11. 深掘り第2弾: "+" 演算子の Date 型時刻解釈の他バリエーション
//     （Date + Date、時刻風文字列同士の + 等）
// ---------------------------------------------------------------------------
{
    add('PLUSDATE', '(#2024/03/15#) + (#2024/03/15#)');       // Date + Date（数値どうしの加算になるはず）
    add('PLUSDATE', '("3.5") + ("2.5")');                       // 文字列どうし（数値どうしの加算 6 になるはず、Date が絡まないので時刻解釈は起きない）
    add('PLUSDATE', '(#2024/03/15 10:00:00#) + ("3.5")');       // 時刻付き Date + 時刻文字列
    add('PLUSDATE', '(CDate("2.5")) + (#2024/03/15#)');         // CDate() 結果（既に Date 型）+ Date
}

// ---------------------------------------------------------------------------
// 12. 深掘り第3弾: 論理演算子（And/Or/Xor/Eqv/Imp）の型強制
//     \/Mod で発見した「Boolean が絡む二項演算は型変換の優先順位が特殊」という
//     パターンが論理演算子でも再現するかを確認する
// ---------------------------------------------------------------------------
{
    const logicOps = ['And', 'Or', 'Xor', 'Eqv', 'Imp'];
    const logicVals = ['True', 'False', '"7"', '"0"', '"-1"', '5', '0', 'Null', 'Empty'];
    for (const op of logicOps) {
        for (const a of logicVals) {
            for (const b of logicVals) {
                // 全組み合わせは爆発するので、片方は必ず True/False literal に絞る
                if (a === 'True' || a === 'False' || b === 'True' || b === 'False') {
                    add('LOGIC', `(${a}) ${op} (${b})`);
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// 13. 深掘り第3弾: StrConv の日本語変換モードの往復・複合フラグ
// ---------------------------------------------------------------------------
{
    // ChrW でひらがな・カタカナ・全角英数・半角カナを組み立てる
    const hiragana = 'ChrW(12354) & ChrW(12356)';           // "あい"
    const katakanaFull = 'ChrW(12450) & ChrW(12452)';        // "アイ"
    const katakanaHalf = 'ChrW(65393) & ChrW(65394)';        // "ｱｲ"（半角カナ）
    const zenkakuAlnum = 'ChrW(65313) & ChrW(65314)';        // "ＡＢ"（全角英字）
    const hankakuAlnum = '"AB"';

    // vbHiragana=32, vbKatakana=16, vbWide=4, vbNarrow=8,
    // vbUpperCase=1, vbLowerCase=2, vbProperCase=3
    add('STRCONV', `StrConv(${hiragana}, 16)`);              // ひらがな→カタカナ
    add('STRCONV', `StrConv(${katakanaFull}, 32)`);          // カタカナ→ひらがな
    add('STRCONV', `StrConv(${katakanaHalf}, 4)`);           // 半角カナ→全角
    add('STRCONV', `StrConv(${zenkakuAlnum}, 8)`);           // 全角英字→半角
    add('STRCONV', `StrConv(${hankakuAlnum}, 4)`);           // 半角英字→全角（英字も対象か確認）
    add('STRCONV', `StrConv(${katakanaFull}, 8)`);           // 全角カタカナ→半角（カナも半角化されるか）
    // 複合フラグ: vbKatakana(16) + vbNarrow(8) = 24 → ひらがなを半角カタカナに
    add('STRCONV', `StrConv(${hiragana}, 24)`);
    // 複合フラグ: vbWide(4) + vbUpperCase(1) = 5
    add('STRCONV', `StrConv("ab", 5)`);
    // 往復: ひらがな→カタカナ→ひらがな で元に戻るか
    add('STRCONV', `StrConv(StrConv(${hiragana}, 16), 32)`);
    // 記号・数字はひらがな/カタカナ変換の対象外のはず
    add('STRCONV', `StrConv("123" & ${hiragana} & "abc", 16)`);
}

// ---------------------------------------------------------------------------
// .bas 生成
// ---------------------------------------------------------------------------
fs.rmSync(OUT_DIR, { recursive: true, force: true });
fs.mkdirSync(OUT_DIR, { recursive: true });

const moduleNames: string[] = [];
for (let m = 0; m * CASES_PER_MODULE < cases.length; m++) {
    const chunk = cases.slice(m * CASES_PER_MODULE, (m + 1) * CASES_PER_MODULE);
    const name = `DiffCorpus_${String(m + 1).padStart(3, '0')}`;
    moduleNames.push(name);
    const subs = chunk.map(c => `Private Sub Case_${c.id}()
    Dim v As Variant
    Err.Clear
    On Error Resume Next
    v = ${c.expr}
    If Err.Number <> 0 Then
        DiffMain.Rec "${c.id}", "ERR", CStr(Err.Number)
    Else
        DiffMain.Rec "${c.id}", "OK", DiffMain.Canon(v)
    End If
End Sub`).join('\n\n');
    const calls = chunk.map(c => `    Case_${c.id}`).join('\n');
    fs.writeFileSync(path.join(OUT_DIR, `${name}.bas`), `Attribute VB_Name = "${name}"
Option Explicit

' 自動生成: scripts/gen-diff-corpus.ts（手動編集しない）

Public Sub RunModule()
${calls}
End Sub

${subs}
`);
}

const runCalls = moduleNames.map(n => `    ${n}.RunModule`).join('\n');
fs.writeFileSync(path.join(OUT_DIR, 'DiffMain.bas'), `Attribute VB_Name = "DiffMain"
Option Explicit

' 自動生成: scripts/gen-diff-corpus.ts（手動編集しない）
'
' 実 Excel での実行手順:
'   1. 空の .xlsm を用意し、vba-extractor import book.xlsm diff-vba/ で全モジュールを注入
'   2. Excel で開き、DiffMain.RunAll を実行（Alt+F8）
'   3. カレントディレクトリー（通常は Documents）に vba-diff-results.txt が生成される
'   4. npx tsx scripts/diff-results.ts diff-vba/results-runner.txt vba-diff-results.txt

Private g_f As Integer

Public Sub RunAll()
    g_f = FreeFile
    Open "vba-diff-results.txt" For Output As #g_f
${runCalls}
    Close #g_f
    Debug.Print "diff corpus done"
End Sub

Public Sub Rec(ByVal id As String, ByVal status As String, ByVal payload As String)
    Print #g_f, id & Chr(9) & status & Chr(9) & payload
End Sub

' locale 非依存の正規化シリアライザー
Public Function Canon(ByVal v As Variant) As String
    On Error GoTo Fallback
    Dim t As String
    t = TypeName(v)
    Select Case t
        Case "String"
            Canon = "String:" & HexStr(v)
        Case "Date"
            Canon = "Date:" & Format(v, "yyyy-mm-dd hh:nn:ss")
        Case "Null"
            Canon = "Null"
        Case "Empty"
            Canon = "Empty"
        Case "Boolean"
            Canon = "Boolean:" & CStr(v)
        Case "Decimal"
            Canon = "Decimal:" & CStr(v)
        Case Else
            ' 数値系: Str は locale 非依存（先頭スペースは Trim で除去）
            Canon = t & ":" & Trim(Str(v))
    End Select
    Exit Function
Fallback:
    Canon = "CANONERR:" & Err.Number
End Function

Public Function HexStr(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & Right("0000" & Hex(AscW(Mid(s, i, 1))), 4)
    Next i
    HexStr = r
End Function
`);

// マニフェスト
const manifest = {
    generated: new Date().toISOString().slice(0, 10),
    total: cases.length,
    modules: moduleNames,
    cases: Object.fromEntries(cases.map(c => [c.id, c.expr])),
    skippedParse: Object.fromEntries(skippedParse.map(c => [c.id, c.expr])),
};
fs.writeFileSync(path.join(OUT_DIR, 'manifest.json'), JSON.stringify(manifest, null, 1));

console.log(`生成完了: ${cases.length} 式 / ${moduleNames.length} モジュール → ${OUT_DIR}/`);
for (const [cat, n] of counters) console.log(`  ${cat}: ${n}`);
if (skippedParse.length > 0) {
    console.log(`\n⚠️ vba-runner でパースできず除外した式 ${skippedParse.length} 件（manifest.json の skippedParse 参照）:`);
    for (const c of skippedParse.slice(0, 10)) console.log(`  ${c.id}: ${c.expr}`);
}
