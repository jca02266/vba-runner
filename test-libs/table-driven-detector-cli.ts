/**
 * テーブル駆動検出器 CLI
 *
 * 使い方:
 *   node table-driven-detector-cli.cjs <file.bas> [options]
 *
 * ビルド:
 *   ./node_modules/.bin/esbuild test-libs/table-driven-detector-cli.ts \
 *     --bundle --platform=node \
 *     --outfile=test-libs/table-driven-detector-cli.cjs
 */

import * as fs from 'fs';
import * as path from 'path';
import { TableDrivenDetector } from './table-driven-detector';

const HELP = `
テーブル駆動リファクタリング候補検出ツール

使い方:
  node table-driven-detector-cli.cjs <file.bas> [options]

引数:
  <file.bas>         解析対象の VBA ファイル（.bas）

オプション:
  --decision-table   デシジョンテーブルを ASCII 形式で出力する（人間向け確認用）
  --json             結果を JSON 形式で出力する（AI・他ツール連携向け）
  --help             このヘルプを表示する

JSON 出力フィールド（--json）:
  候補ごとに以下のフィールドを含む配列を出力します。

  基本情報:
    functionName          関数名
    startLine / endLine   if 文の開始・終了行（0 始まり）
    confidenceScore       テーブル駆動化の信頼度スコア（0-100）
    riskLevel             実装リスク（"low" | "medium" | "high"）
    recommendation        推奨テキスト

  分岐統計:
    nestingDepth          ネスト深度
    outerBranchCount      外側分岐数（Map のキー数）
    innerBranchCount      内側分岐数（しきい値段階数）
    totalBranchCombinations  外側 × 内側の組み合わせ数

  削減見積もり:
    currentLines          現在の推定行数
    estimatedAfterLines   リファクタリング後の推定行数
    linesToSave           削減見込み行数
    reductionPercent      削減率（%）

  判定根拠:
    reasons.isRepeatingStructure   繰り返し構造か
    reasons.isSimpleAssignment     単純代入パターンか
    reasons.isSimpleCondition      単純条件式か
    reasons.branchCountThreshold   分岐数が閾値以上か
    reasons.hasNoSideEffects       副作用なしか

  診断情報:
    diagnostics.detectedPatterns   検出されたパターンの説明リスト
    diagnostics.warnings           問題点の警告リスト

  デシジョンテーブル（decisionTable）:
    keyVariable           外側キーの変数名（例: "department"）
    valueVariable         しきい値の変数名（例: "amount"）
    rows[]                各キーの組み合わせ
      keyPath[]           キーのパス（例: ["Engineering","Capex","Emergency"]）
      rules[]             しきい値ルール
        condition         条件式（例: "amount < 100,000"）
        result            結果値（例: "TeamLead"）
      defaultResult       else 節の結果値

  副作用分析（sideEffects）:
    hasUniformStructure   全ブランチで副作用の構造が同一か
    sideEffectCount       副作用文数（結果代入を除く）
    structureSummary      副作用の AST 形状（例: "log=(log&STR), count=(count+NUM)"）
    mismatchedBranches[]  構造が不一致だったブランチのキーパス（不均一時のみ）

例:
  # 候補の検出（テキスト形式）
  node table-driven-detector-cli.cjs sample/src/vba/ApprovalRules_Before.bas

  # デシジョンテーブルも一緒に出力（人間向け）
  node table-driven-detector-cli.cjs sample/src/vba/ApprovalRules_Before.bas --decision-table

  # JSON 形式で出力（AI・他ツール連携向け、decisionTable と sideEffects を含む）
  node table-driven-detector-cli.cjs sample/src/vba/ApprovalRules_Before.bas --json
`.trimStart();

function main() {
    const args = process.argv.slice(2);

    if (args.length === 0 || args.includes('--help')) {
        process.stdout.write(HELP);
        process.exit(0);
    }

    const filePath = args.find((a) => !a.startsWith('--'));
    if (!filePath) {
        process.stderr.write('エラー: ファイルパスを指定してください。\n');
        process.exit(1);
    }

    const showDecisionTable = args.includes('--decision-table');
    const jsonOutput = args.includes('--json');

    let code: string;
    try {
        code = fs.readFileSync(path.resolve(filePath), 'utf8');
    } catch {
        process.stderr.write(`エラー: ファイルを読み込めません: ${filePath}\n`);
        process.exit(1);
    }

    const detector = new TableDrivenDetector();
    const candidates = detector.detectFromCode(code);

    if (jsonOutput) {
        process.stdout.write(JSON.stringify(candidates, null, 2) + '\n');
        return;
    }

    if (candidates.length === 0) {
        process.stdout.write(`${filePath}: テーブル駆動化の候補は見つかりませんでした。\n`);
        return;
    }

    for (const c of candidates) {
        process.stdout.write([
            `\n関数: ${c.functionName}  (行 ${c.startLine + 1}-${c.endLine + 1})`,
            `スコア: ${c.confidenceScore}/100  リスク: ${c.riskLevel}  削減効果: ${c.reductionPercent}%（${c.linesToSave}行）`,
            `外側分岐: ${c.outerBranchCount}  内側分岐: ${c.innerBranchCount}  組み合わせ: ${c.totalBranchCombinations}`,
            `推奨: ${c.recommendation}`,
        ].join('\n') + '\n');

        if (c.diagnostics.warnings.length > 0) {
            process.stdout.write('警告:\n');
            for (const w of c.diagnostics.warnings) {
                process.stdout.write(`  - ${w}\n`);
            }
        }

        if (c.conditionVariables && c.conditionVariables.length > 0) {
            process.stdout.write('条件変数:\n');
            for (const v of c.conditionVariables) {
                const decl = v.declaredAt != null ? `Dim 行${v.declaredAt}` : '引数/外部';
                const assigns = v.assignedLines.length > 0
                    ? `代入 行${v.assignedLines.join(', ')}`
                    : '代入なし';
                process.stdout.write(`  ${v.name}: ${decl}  ${assigns}\n`);
            }
        }

        if (showDecisionTable && c.decisionTable) {
            process.stdout.write('\nデシジョンテーブル:\n');
            process.stdout.write(detector.renderDecisionTable(c.decisionTable) + '\n');
        }
    }
}

main();
