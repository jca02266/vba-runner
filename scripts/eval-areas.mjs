/** Canonical evaluation-area vocabulary and legacy focus inference. */
export const EVALUATION_AREAS = [
  ['Select Case', /Select Case/],
  ['Format', /Format/],
  ['数値精度', /LongLong|Decimal|Currency|数値String|算術|整数除算|指数|桁区切り|Scientific/],
  ['引数・ByRef・Property', /ByRef|ByVal|Property|引数|CallByName/],
  ['ファイルI/O・FSO', /FSO|ファイル|Append|Input|EOF|VFS|OpenTextFile|FileCopy|MkDir|RmDir|Kill/],
  ['配列・UDT', /配列|UDT|Binary|ReDim|Erase/],
  ['Option Compare・Like', /Option Compare|Like/],
  ['Date', /日付|Date|Weekday|MonthName|TimeSerial/],
  ['Err・On Error', /Err|エラー|On Error/],
  ['LSP・拡張機能', /LSP|拡張機能/],
  ['Parser・Lexer', /Lexer|文法|識別子|宣言重複/],
  ['評価基盤', /EVAL_LOG|カバレッジ|評価|テスト|Namespace|名前空間|リファクタリング|ミューテーション/],
  ['制御フロー', /GoSub|Return|For制御|DoEvents|Sleep|Case Is/],
  ['コレクション・Dictionary', /Collection|Dictionary|列挙変更|キー|インデックス/],
  ['オブジェクト・COM', /COM|Object|オブジェクト|既定Value|既定メンバー|SAFEARRAY|As New|With/],
  ['宣言・型・リテラル', /リテラル|サフィックス|基数|型強制|型変換|型境界|CDec|IsNumeric/],
  ['組み込み関数', /\bVal\b|StrConv|StrReverse|StrComp|Chr|LeftRightMid|RGB|Round|Join|InStr|Choose|Rate|MIRR|SYD|DDB|SLN|情報関数|金融関数|ダイアログ/],
];

export const VALID_EVALUATION_AREAS = new Set([
  ...EVALUATION_AREAS.map(([area]) => area),
  'その他',
]);

export function inferEvaluationArea(focus) {
  const text = String(focus ?? '');
  return EVALUATION_AREAS.find(([, pattern]) => pattern.test(text))?.[0] ?? 'その他';
}

export function evaluationArea(record) {
  return record.area ?? inferEvaluationArea(record.focus);
}
