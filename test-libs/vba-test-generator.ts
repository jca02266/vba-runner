import * as fs from 'fs';
import * as path from 'path';

/**
 * VBA テストファイルから Test* パターンのプロシージャを検出し、
 * 実VBA環境で動作するテストランナー Sub を自動生成するツール
 */
export class VBATestGenerator {
    /**
     * VBA ファイルからテスト関連プロシージャを検出
     * @param vbaSource VBA ソースコード
     * @returns { procedures: テストプロシージャ名の配列, hasSetUp: SetUp があるか, hasTearDown: TearDown があるか }
     */
    static extractTestProcedures(vbaSource: string): { procedures: string[], hasSetUp: boolean, hasTearDown: boolean } {
        const testProcedures: string[] = [];
        let hasSetUp = false;
        let hasTearDown = false;

        // Sub Test*() または Function Test*() パターンを検出
        // 先頭の空白を許容、括弧内の空白も許容
        const procedurePattern = /^\s*(?:Sub|Function)\s+(\w+)\s*\(\s*\)/gim;

        let match;
        while ((match = procedurePattern.exec(vbaSource)) !== null) {
            const name = match[1];
            if (name.match(/^Test/i)) {
                testProcedures.push(name);
            } else if (name.match(/^SetUp$/i)) {
                hasSetUp = true;
            } else if (name.match(/^TearDown$/i)) {
                hasTearDown = true;
            }
        }

        return { procedures: testProcedures, hasSetUp, hasTearDown };
    }

    /**
     * モジュール名を抽出（ファイル名から .vba 拡張子を除いたもの）
     * @param filePath ファイルパス
     * @returns モジュール名
     */
    static extractModuleName(filePath: string): string {
        return path.basename(filePath, '.vba');
    }

    /**
     * 実VBA環境で動作するテストランナー Sub を生成
     * @param testProcedures テストプロシージャ名の配列
     * @param hasSetUp SetUp Sub があるか
     * @param hasTearDown TearDown Sub があるか
     * @param runnerName ランナーの Sub 名（デフォルト: RunAllTests）
     * @param moduleName モジュール名（省略可）- 指定されない場合は修飾なしで生成
     * @returns VBA Sub ソースコード
     */
    static generateTestRunner(testProcedures: string[], hasSetUp: boolean = false, hasTearDown: boolean = false, runnerName: string = 'RunAllTests', moduleName?: string): string {
        if (testProcedures.length === 0) {
            return `' No test procedures found\nSub ${runnerName}()\n    MsgBox "No test procedures found"\nEnd Sub\n`;
        }

        // 修飾プレフィックス（モジュール名が指定されている場合のみ使用）
        const modulePrefix = moduleName ? `${moduleName}.` : '';

        // テストランナー Sub の生成（実VBA環境で動作）
        let runner = `' Auto-generated test runner from vba-test-generator\n`;
        runner += `' Run this Sub in Excel VBA environment to execute all tests\n`;
        if (moduleName) {
            runner += `' Module-qualified calls: ${modulePrefix}TestProcedures\n`;
        }
        if (hasSetUp || hasTearDown) {
            runner += `' Note: SetUp/TearDown will be called before/after each test\n`;
        }
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

            // SetUp を呼び出す
            if (hasSetUp) {
                runner += `    On Error Resume Next\n`;
                runner += `    ${modulePrefix}SetUp\n`;
                runner += `    On Error GoTo 0\n`;
            }

            // テストを実行
            runner += `    On Error Resume Next\n`;
            runner += `    If ${modulePrefix}${testProc}() Then\n`;
            runner += `        testResults = testResults & "[PASS] ${testProc}" & vbCrLf\n`;
            runner += `        passCount = passCount + 1\n`;
            runner += `    Else\n`;
            runner += `        testResults = testResults & "[FAIL] ${testProc}" & vbCrLf\n`;
            runner += `        failCount = failCount + 1\n`;
            runner += `        allPass = False\n`;
            runner += `    End If\n`;
            runner += `    On Error GoTo 0\n`;

            // TearDown を呼び出す
            if (hasTearDown) {
                runner += `    On Error Resume Next\n`;
                runner += `    ${modulePrefix}TearDown\n`;
                runner += `    On Error GoTo 0\n`;
            }

            runner += `\n`;
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
     * @param useModuleQualifier モジュール修飾を使用するか（デフォルト: true）
     */
    static generateFromFile(inputFilePath: string, outputFilePath?: string, runnerName: string = 'RunAllTests', useModuleQualifier: boolean = true): void {
        if (!fs.existsSync(inputFilePath)) {
            throw new Error(`File not found: ${inputFilePath}`);
        }

        const vbaSource = fs.readFileSync(inputFilePath, 'utf-8');
        const { procedures, hasSetUp, hasTearDown } = this.extractTestProcedures(vbaSource);

        if (procedures.length === 0) {
            console.warn(`No test procedures found in ${inputFilePath}`);
        } else {
            console.error(`Found ${procedures.length} test procedure(s): ${procedures.join(', ')}`);
            if (hasSetUp) console.error(`  - SetUp Sub detected`);
            if (hasTearDown) console.error(`  - TearDown Sub detected`);
        }

        // モジュール名を抽出
        const moduleName = useModuleQualifier ? this.extractModuleName(inputFilePath) : undefined;
        if (moduleName) {
            console.error(`  - Module name: ${moduleName}`);
        }

        const runner = this.generateTestRunner(procedures, hasSetUp, hasTearDown, runnerName, moduleName);

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
     * @param useModuleQualifier モジュール修飾を使用するか（デフォルト: true）
     */
    static generateForDirectory(dirPath: string, outputDir?: string, useModuleQualifier: boolean = true): void {
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
            const runnerName = 'RunAllTests';

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
                this.generateFromFile(inputPath, outputPath, runnerName, useModuleQualifier);
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
        console.error('  vba-test-generator.ts <input.vba> [output.vba] [options]');
        console.error('  vba-test-generator.ts --dir <directory> [output-directory] [options]');
        console.error('');
        console.error('Options:');
        console.error('  --no-module-qualifier  Do not qualify function calls with module name');
        console.error('                         (default: module qualifier is enabled for real VBA environments)');
        console.error('');
        console.error('Examples:');
        console.error('  # Generate with module qualifiers (for real VBA environments)');
        console.error('  npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_CurrencyOperations.vba');
        console.error('');
        console.error('  # Generate without module qualifiers (for VBA interpreter)');
        console.error('  npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_CurrencyOperations.vba --no-module-qualifier');
        console.error('');
        console.error('  # Generate and save to file');
        console.error('  npx ts-node test-libs/vba-test-generator.ts tests/spec/vba/Test_CurrencyOperations.vba output.vba');
        console.error('');
        console.error('  # Generate runners for all VBA files in a directory');
        console.error('  npx ts-node test-libs/vba-test-generator.ts --dir tests/spec/vba');
        console.error('');
        console.error('  # Generate runners in a separate output directory');
        console.error('  npx ts-node test-libs/vba-test-generator.ts --dir tests/spec/vba tests/spec/runners');
        process.exit(1);
    }

    try {
        // Parse options
        let useModuleQualifier = true;
        const optionArgs = args.filter(arg => {
            if (arg === '--no-module-qualifier') {
                useModuleQualifier = false;
                return false;
            }
            return true;
        });

        if (optionArgs[0] === '--dir') {
            const dirPath = optionArgs[1];
            const outputDir = optionArgs[2];
            if (!dirPath) {
                throw new Error('--dir requires a directory path');
            }
            VBATestGenerator.generateForDirectory(dirPath, outputDir, useModuleQualifier);
        } else {
            const inputPath = optionArgs[0];
            const outputPath = optionArgs[1];
            VBATestGenerator.generateFromFile(inputPath, outputPath, 'RunAllTests', useModuleQualifier);
        }
    } catch (e: any) {
        console.error(`Error: ${e.message}`);
        process.exit(1);
    }
}

export default VBATestGenerator;
