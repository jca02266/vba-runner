/**
 * exceljs アダプターを使って .xlsx に VBA マクロを適用する CLI
 *
 * Usage:
 *   npx tsx tools/exceljs-adapter/run.ts <input.xlsx> <macro.bas> <ProcName> [output.xlsx]
 *
 * Example:
 *   npx tsx tools/exceljs-adapter/run.ts \
 *     tools/exceljs-adapter/input.xlsx \
 *     tools/exceljs-adapter/sample.bas \
 *     Main \
 *     tools/exceljs-adapter/output.xlsx
 */
import ExcelJS from 'exceljs';
import * as path from 'path';
import { VBARunner } from '../../test-libs/test-runner';
import { ExcelJsApplication } from './ExcelJsApplication';

async function main() {
    const args = process.argv.slice(2);
    if (args.length < 3) {
        console.error('Usage: run.ts <input.xlsx> <macro.bas> <ProcName> [output.xlsx]');
        process.exit(1);
    }

    const [inputPath, macroPath, procName, outputPath] = args;
    const absInput  = path.resolve(inputPath);
    const absMacro  = path.resolve(macroPath);
    const absOutput = outputPath
        ? path.resolve(outputPath)
        : absInput.replace(/\.xlsx$/i, '-out.xlsx');

    // 1. Excel ファイルを読み込む
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(absInput);
    console.log(`Loaded:  ${absInput}`);
    console.log(`Sheets:  ${workbook.worksheets.map(w => w.name).join(', ')}`);

    // 2. exceljs アダプターを作成し VBARunner に注入
    const app = new ExcelJsApplication(workbook);
    const runner = new VBARunner([absMacro], {
        excelStub: app as any,  // duck-typed — ExcelJsApplication は MockApplication と同じインターフェース
        quiet: false,
        onPrint: (msg) => console.log('[VBA]', msg),
    });

    // 3. 指定プロシージャを実行
    console.log(`Running: ${procName}()`);
    runner.run(procName);

    // 4. 結果を保存
    await workbook.xlsx.writeFile(absOutput);
    console.log(`Saved:   ${absOutput}`);
}

main().catch(e => {
    console.error(e instanceof Error ? e.message : e);
    process.exit(1);
});
