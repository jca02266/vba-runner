# VBA Runner Internal Knowledge Bundle

## Specifications

* [MS-VBAL BNF Grammar](BNF.md) - VBA ソース構造の文法
* [リテラル・型・演算子](LITERALS_AND_OPERATORS.md) - 言語仕様と実装対応
* [名前解決](NAME_RESOLUTION.md) - 名前空間と識別子解決
* [型システム厳密化](TYPE_SYSTEM_SPEC.md) - 型メタデータと変換規則

## Design notes

* [構文的曖昧性](AMBIGUOUS_SYNTAX.md) - 文と式の解釈
* [評価パイプライン](EVALUATION.md) - パースから実行まで
* [関数コールと引数処理](FUNCTION_CALL.md) - 呼び出しと引数束縛
* [識別子分類](IDENTIFIER_CLASSIFICATION.md) - Lexer / Parser の分類
* [Currency / Decimal 実装](DECIMAL_IMPLEMENTATION.md) - BigInt 固定小数点
* [VBA エクスポート・インポート](VBA_EXPORT_IMPORT.md) - Office バイナリ処理
* [評価記録とバグ探索状態](EVALUATION_STORAGE.md) - 評価履歴・探索キュー・横展開調査の永続化

## Plans

* [データフロー解析 実装計画](DATAFLOW_ANALYSIS_TODO.md) - 静的解析のロードマップ
