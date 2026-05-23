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
  --decision-table   デシジョンテーブルを ASCII 形式で出力する
  --json             結果を JSON 形式で出力する（デフォルト: テキスト）
  --help             このヘルプを表示する

例:
  # 候補の検出（テキスト形式）
  node table-driven-detector-cli.cjs sample/src/vba/ApprovalRules_Before.bas

  # デシジョンテーブルも一緒に出力
  node table-driven-detector-cli.cjs sample/src/vba/ApprovalRules_Before.bas --decision-table

  # JSON 形式で出力（他ツールへのパイプに便利）
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

        if (showDecisionTable && c.decisionTable) {
            process.stdout.write('\nデシジョンテーブル:\n');
            process.stdout.write(detector.renderDecisionTable(c.decisionTable) + '\n');
        }
    }
}

main();
