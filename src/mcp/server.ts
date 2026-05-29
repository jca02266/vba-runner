import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import * as fs from 'fs';
import * as path from 'path';
import { execSync } from 'child_process';
import { Lexer } from '../engine/lexer.js';
import { Parser } from '../engine/parser.js';
import { Evaluator } from '../engine/evaluator.js';
import { lintProgram } from '../engine/vba-lint.js';

import {
    collectVbaFilesForMcp,
    analyzeWorkspaceForMcp,
    formatWorkspaceReportForMcp,
    formatWorkspaceOutlineForMcp,
} from './analyzer-bridge.js';

const server = new McpServer({
    name: 'vba-runner',
    version: '1.0.0',
});

// ------------------------------------------------------------------
// vba_analyze: コードスメルレポート
// ------------------------------------------------------------------
server.tool(
    'vba_analyze',
    'VBA ソースの品質指標（行数・ネスト深さ・Excel 依存度など）を分析して報告する',
    {
        path: z.string().describe('分析対象の .bas/.cls/.frm ファイルまたはディレクトリのパス'),
        format: z.enum(['text', 'json']).optional().default('text').describe('出力形式'),
    },
    async ({ path: targetPath, format }) => {
        const abs = path.resolve(targetPath);
        if (!fs.existsSync(abs)) {
            return { content: [{ type: 'text', text: `パスが見つかりません: ${abs}` }], isError: true };
        }
        try {
            const files = collectVbaFilesForMcp(abs);
            if (files.length === 0) {
                return { content: [{ type: 'text', text: 'VBA ファイルが見つかりません' }], isError: true };
            }
            const workspace = analyzeWorkspaceForMcp(files);
            const text = format === 'json'
                ? JSON.stringify(workspace, null, 2)
                : formatWorkspaceReportForMcp(workspace);
            return { content: [{ type: 'text', text }] };
        } catch (e: any) {
            return { content: [{ type: 'text', text: `エラー: ${e.message}` }], isError: true };
        }
    }
);

// ------------------------------------------------------------------
// vba_get_outline: AI 向けコンテキスト圧縮アウトライン
// ------------------------------------------------------------------
server.tool(
    'vba_get_outline',
    'VBA ワークスペースの構造をコンパクトに要約する（AI コンテキスト節約向け）',
    {
        path: z.string().describe('対象ディレクトリまたはファイルのパス'),
    },
    async ({ path: targetPath }) => {
        const abs = path.resolve(targetPath);
        if (!fs.existsSync(abs)) {
            return { content: [{ type: 'text', text: `パスが見つかりません: ${abs}` }], isError: true };
        }
        try {
            const files = collectVbaFilesForMcp(abs);
            if (files.length === 0) {
                return { content: [{ type: 'text', text: 'VBA ファイルが見つかりません' }], isError: true };
            }
            const workspace = analyzeWorkspaceForMcp(files);
            const text = formatWorkspaceOutlineForMcp(workspace);
            return { content: [{ type: 'text', text }] };
        } catch (e: any) {
            return { content: [{ type: 'text', text: `エラー: ${e.message}` }], isError: true };
        }
    }
);

// ------------------------------------------------------------------
// vba_run_tests: VBA ネイティブテストの実行
// ------------------------------------------------------------------
server.tool(
    'vba_run_tests',
    'VBA ソース内の Test_* プロシージャを実行し、pass/fail 件数と失敗詳細を返す',
    {
        path: z.string().describe('テスト対象の .bas/.cls/.frm ファイルまたはディレクトリのパス'),
    },
    async ({ path: targetPath }) => {
        const abs = path.resolve(targetPath);
        if (!fs.existsSync(abs)) {
            return { content: [{ type: 'text', text: `パスが見つかりません: ${abs}` }], isError: true };
        }
        try {
            const files = collectVbaFilesForMcp(abs);
            if (files.length === 0) {
                return { content: [{ type: 'text', text: 'VBA ファイルが見つかりません' }], isError: true };
            }

            const results: string[] = [];
            let pass = 0;
            let fail = 0;

            for (const filePath of files) {
                const src = fs.readFileSync(filePath, 'utf-8');
                const tokens = new Lexer(src).tokenize();
                const ast = new Parser(tokens).parse();
                const evaluator = new Evaluator(() => {});
                evaluator.evaluate(ast);

                const testProcs = ast.body
                    .filter((n: any) => n.type === 'SubDeclaration' && /^test_/i.test(n.name?.name ?? ''))
                    .map((n: any) => n.name?.name as string);

                for (const procName of testProcs) {
                    try {
                        evaluator.callProcedure(procName, []);
                        results.push(`  ✅ PASS  ${procName}`);
                        pass++;
                    } catch (e: any) {
                        results.push(`  ❌ FAIL  ${procName}: ${e.message}`);
                        fail++;
                    }
                }
            }

            const summary = `テスト結果: ${pass + fail} 件中 ${pass} PASS / ${fail} FAIL`;
            const text = [summary, '', ...results].join('\n');
            return { content: [{ type: 'text', text }] };
        } catch (e: any) {
            return { content: [{ type: 'text', text: `エラー: ${e.message}` }], isError: true };
        }
    }
);

// ------------------------------------------------------------------
// vba_get_diagnostics: lint 診断結果
// ------------------------------------------------------------------
server.tool(
    'vba_get_diagnostics',
    'VBA ソースに対して VBA001〜VBA011 の lint 診断を実行して報告する',
    {
        path: z.string().describe('診断対象の .bas/.cls/.frm ファイルのパス'),
    },
    async ({ path: targetPath }) => {
        const abs = path.resolve(targetPath);
        if (!fs.existsSync(abs)) {
            return { content: [{ type: 'text', text: `パスが見つかりません: ${abs}` }], isError: true };
        }
        try {
            const src = fs.readFileSync(abs, 'utf-8');
            const tokens = new Lexer(src).tokenize();
            const ast = new Parser(tokens, { errorRecovery: true }).parse();
            const diags = lintProgram(ast);
            if (diags.length === 0) {
                return { content: [{ type: 'text', text: '診断なし（問題は見つかりませんでした）' }] };
            }
            const lines = diags.map((d: any) =>
                `  [${d.code}] L${d.loc.start.line}:${d.loc.start.column}  ${d.message}`
            );
            return { content: [{ type: 'text', text: lines.join('\n') }] };
        } catch (e: any) {
            return { content: [{ type: 'text', text: `エラー: ${e.message}` }], isError: true };
        }
    }
);

// ------------------------------------------------------------------
// vba_run_ts_tests: TypeScript テストの実行
// ------------------------------------------------------------------
server.tool(
    'vba_run_ts_tests',
    'TypeScript テストファイル (.test.ts) を esbuild でバンドルして Node.js で実行し、結果を返す',
    {
        path: z.string().describe('テスト対象の .test.ts ファイルまたはディレクトリのパス'),
    },
    async ({ path: targetPath }) => {
        const projectRoot = path.resolve(__dirname, '..');
        const esbuild = path.join(projectRoot, 'node_modules/.bin/esbuild');
        const abs = path.resolve(targetPath);
        if (!fs.existsSync(abs)) {
            return { content: [{ type: 'text', text: `パスが見つかりません: ${abs}` }], isError: true };
        }

        const testFiles: string[] = [];
        const stat = fs.statSync(abs);
        if (stat.isDirectory()) {
            const walk = (dir: string) => {
                for (const entry of fs.readdirSync(dir)) {
                    const full = path.join(dir, entry);
                    if (fs.statSync(full).isDirectory()) walk(full);
                    else if (entry.endsWith('.test.ts')) testFiles.push(full);
                }
            };
            walk(abs);
        } else if (abs.endsWith('.test.ts')) {
            testFiles.push(abs);
        }

        if (testFiles.length === 0) {
            return { content: [{ type: 'text', text: '.test.ts ファイルが見つかりません' }], isError: true };
        }

        const results: string[] = [];
        for (const tsFile of testFiles) {
            const cjsFile = tsFile.replace(/\.ts$/, '.cjs');
            try {
                execSync(`${esbuild} ${tsFile} --bundle --outfile=${cjsFile} --platform=node`, { cwd: projectRoot });
                const output = execSync(`node ${cjsFile}`, { cwd: projectRoot, encoding: 'utf-8' });
                results.push(`=== ${path.basename(tsFile)} ===\n${output}`);
            } catch (e: any) {
                const out = (e.stdout ?? '') + (e.stderr ?? '') + (e.message ?? '');
                results.push(`=== ${path.basename(tsFile)} [FAILED] ===\n${out}`);
            }
        }

        return { content: [{ type: 'text', text: results.join('\n') }] };
    }
);

// ------------------------------------------------------------------
// サーバー起動
// ------------------------------------------------------------------
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
}

main().catch(console.error);
