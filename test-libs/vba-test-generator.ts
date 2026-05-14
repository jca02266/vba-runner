import * as fs from 'fs';
import * as path from 'path';

/**
 * VBA テストファイルから Test* パターンのプロシージャを検出し、
 * 実VBA環境で動作するテストランナー Sub を自動生成するツール
 */
export class VBATestGenerator {
    /**
     * VBA ファイルから Test* で始まるプロシージャを検出
     * @param vbaSource VBA ソースコード
     * @returns テストプロシージャ名の配列（検出順）
     */
    static extractTestProcedures(vbaSource: string): string[] {
        const testProcedures: string[] = [];
        // Sub Test*() または Function Test*() パターンを検出
        // 先頭の空白を許容、括弧内の空白も許容
        const procedurePattern = /^\s*(?:Sub|Function)\s+(Test\w+)\s*\(\s*\)/gim;

        let match;
        while ((match = procedurePattern.exec(vbaSource)) !== null) {
            testProcedures.push(match[1]);
        }

        return testProcedures;
    }

    /**
     * 実VBA環境で動作するテストランナー Sub を生成
     * @param testProcedures テストプロシージャ名の配列
     * @param runnerName ランナーの Sub 名（デフォルト: RunAllTests）
     * @returns VBA Sub ソースコード
     */
    static generateTestRunner(testProcedures: string[], runnerName: string = 'RunAllTests'): string {
        if (testProcedures.length === 0) {
            return `' No test procedures found\nSub ${runnerName}()\n    MsgBox "No test procedures found"\nEnd Sub\n`;
        }

        // テストランナー Sub の生成（実VBA環境で動作）
        let runner = `' Auto-generated test runner from vba-test-generator\n`;
        runner += `' Run this Sub in Excel VBA environment to execute all tests\n`;
        runner += `Sub ${runnerName}()\n`;
        runner += `    Dim allPass As Boolean\n`;
        runner += `    Dim passCount As Integer\n`;
        runner += `    Dim failCount As Integer\n`;
        runner += `    Dim testResults As String\n`;
        runner += `    Dim i As Integer\n\n`;
        runner += `    allPass = True\n`;
        runner += `    passCount = 0\n`;
        runner += `    failCount = 0\n`;
        runner += `    testResults = "=== Test Results ===" & vbCrLf & vbCrLf\n\n`;

        // 各テストプロシージャの実行コードを生成
        for (const testProc of testProcedures) {
            runner += `    ' Execute ${testProc}\n`;
            runner += `    On Error Resume Next\n`;
            runner += `    If ${testProc}() Then\n`;
            runner += `        testResults = testResults & "[PASS] ${testProc}" & vbCrLf\n`;
            runner += `        passCount = passCount + 1\n`;
            runner += `    Else\n`;
            runner += `        testResults = testResults & "[FAIL] ${testProc}" & vbCrLf\n`;
            runner += `        failCount = failCount + 1\n`;
            runner += `        allPass = False\n`;
            runner += `    End If\n`;
            runner += `    On Error GoTo 0\n\n`;
        }

        runner += `    testResults = testResults & vbCrLf\n`;
        runner += `    testResults = testResults & "=== Test Summary ===" & vbCrLf\n`;
        runner += `    testResults = testResults & "Total: " & (passCount + failCount) & vbCrLf\n`;
        runner += `    testResults = testResults & "Passed: " & passCount & vbCrLf\n`;
        runner += `    testResults = testResults & "Failed: " & failCount & vbCrLf\n`;
        runner += `    testResults = testResults & "=== Test Complete ===" & vbCrLf\n\n`;
        runner += `    Debug.Print testResults\n`;
        runner += `    MsgBox testResults, IIf(allPass, vbInformation, vbCritical), "Test Results"\n`;
        runner += `End Sub\n`;

        return runner;
    }

    /**
     * VBA ファイルを読み込み、テストランナーを生成して出力
     * @param inputFilePath 入力 VBA ファイルパス
     * @param outputFilePath 出力ファイルパス（指定しない場合は stdout）
     * @param runnerName ランナー Sub の名前
     */
    static generateFromFile(inputFilePath: string, outputFilePath?: string, runnerName: string = 'RunAllTests'): void {
        if (!fs.existsSync(inputFilePath)) {
            throw new Error(`File not found: ${inputFilePath}`);
        }

        const vbaSource = fs.readFileSync(inputFilePath, 'utf-8');
        const testProcedures = this.extractTestProcedures(vbaSource);

        if (testProcedures.length === 0) {
            console.warn(`No test procedures found in ${inputFilePath}`);
        } else {
            console.error(`Found ${testProcedures.length} test procedure(s): ${testProcedures.join(', ')}`);
        }

        const runner = this.generateTestRunner(testProcedures, runnerName);

        if (outputFilePath) {
            fs.writeFileSync(outputFilePath, runner, 'utf-8');
            console.error(`✓ Generated test runner: ${outputFilePath}`);
        } else {
            console.log(runner);
        }
    }

    /**
     * ディレクトリ内のすべての VBA テストファイルに対してランナーを生成
     * @param dirPath ディレクトリパス
     * @param outputDir 出力ディレクトリ（指定しない場合は元のファイルの隣に _runner.vba で出力）
     */
    static generateForDirectory(dirPath: string, outputDir?: string): void {
        if (!fs.existsSync(dirPath)) {
            throw new Error(`Directory not found: ${dirPath}`);
        }

        const files = fs.readdirSync(dirPath)
            .filter(f => f.toLowerCase().endsWith('.vba'))
            .sort();

        if (files.length === 0) {
            console.error(`No .vba files found in ${dirPath}`);
            return;
        }

        console.error(`Processing ${files.length} VBA file(s)...`);

        for (const file of files) {
            const inputPath = path.join(dirPath, file);
            const baseName = path.basename(file, '.vba');
            const runnerName = 'RunAllTests';  // 統一したランナー名

            let outputPath: string | undefined;
            if (outputDir) {
                if (!fs.existsSync(outputDir)) {
                    fs.mkdirSync(outputDir, { recursive: true });
                }
                outputPath = path.join(outputDir, baseName + '_runner.vba');
            } else {
                outputPath = path.join(dirPath, baseName + '_runner.vba');
            }

            try {
                this.generateFromFile(inputPath, outputPath, runnerName);
            } catch (e: any) {
                console.error(`✗ Error processing ${file}: ${e.message}`);
            }
        }
    }
}

// CLI 実行（Node.js で直接実行された場合のみ）
// esbuild でバンドルされた場合は実行されないようにする
const isMainModule = typeof require !== 'undefined' && require.main === module && process.argv.length > 2 && !process.argv[1].includes('esbuild');
if (isMainModule || (typeof process !== 'undefined' && process.argv[1]?.includes('vba-test-generator'))) {
    const args = process.argv.slice(2);

    if (args.length === 0) {
        console.error('VBA Test Generator - Generate test runners from VBA test files');
        console.error('');
        console.error('Usage:');
        console.error('  vba-test-generator.ts <input.vba> [output.vba]');
        console.error('  vba-test-generator.ts --dir <directory> [output-directory]');
        console.error('');
        console.error('Examples:');
        console.error('  # Generate and print to stdout');
        console.error('  npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_CurrencyOperations.vba');
        console.error('');
        console.error('  # Generate and save to file');
        console.error('  npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_CurrencyOperations.vba tests/spec/vba/Test_CurrencyOperations_runner.vba');
        console.error('');
        console.error('  # Generate runners for all VBA files in a directory');
        console.error('  npx ts-node test-libs/vba-test-generator.ts --dir tests/spec/vba');
        console.error('');
        console.error('  # Generate runners in a separate output directory');
        console.error('  npx ts-node test-libs/vba-test-generator.ts --dir tests/spec/vba tests/spec/runners');
        process.exit(1);
    }

    try {
        if (args[0] === '--dir') {
            const dirPath = args[1];
            const outputDir = args[2];
            if (!dirPath) {
                throw new Error('--dir requires a directory path');
            }
            VBATestGenerator.generateForDirectory(dirPath, outputDir);
        } else {
            const inputPath = args[0];
            const outputPath = args[1];
            VBATestGenerator.generateFromFile(inputPath, outputPath, 'RunAllTests');
        }
    } catch (e: any) {
        console.error(`Error: ${e.message}`);
        process.exit(1);
    }
}

export default VBATestGenerator;
