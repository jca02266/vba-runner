import { ProcedureDeclaration, Statement } from '../engine/parser';
import { findAllReferences } from './references-provider';

export interface CodeLensItem {
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
    command: {
        title: string;
        command: string;
        arguments?: any[];
    };
}

export interface ProcInfo {
    name: string;
    line: number;           // 0-based (declaration line)
    endLine: number;        // 0-based (End Sub/Function line)
    isPrivate: boolean;
    hasRequiredParams: boolean;
    isTestProc: boolean;    // Test_* with exactly 1 param (the assert helper)
    refCount: number;       // call sites from outside this procedure
    isTested: boolean;      // any Test_* proc references this one
}

export class CodeLensProvider {
    getCodeLens(statements: Statement[], sourceText: string, uri: string): CodeLensItem[] {
        const procs = this.collectProcs(statements, sourceText, uri);
        const items: CodeLensItem[] = [];

        for (const proc of procs) {
            const range = {
                start: { line: proc.line, character: 0 },
                end:   { line: proc.line, character: 0 },
            };

            // ▶ Run: parameterless、または Test_* の assert 1引数プロシージャ
            if (!proc.hasRequiredParams || proc.isTestProc) {
                items.push({
                    range,
                    command: {
                        title: '▶ Run',
                        command: 'vba-runner.runProcedure',
                        arguments: [uri, proc.name, proc.isTestProc],
                    },
                });
            }

            // N references
            const refLabel = proc.refCount === 0
                ? (proc.isPrivate ? '⚠ 0 references' : '0 references')
                : `${proc.refCount} reference${proc.refCount !== 1 ? 's' : ''}`;
            items.push({
                range,
                command: {
                    title: refLabel,
                    command: 'vba-runner.findReferences',
                    arguments: [uri, proc.line, 0],
                },
            });

            // テスト状態
            items.push({
                range,
                command: {
                    title: proc.isTested ? '✓ テスト済み' : '未テスト',
                    command: proc.isTested ? 'vba-runner.goToTest' : 'vba-runner.generateTest',
                    arguments: [uri, proc.name],
                },
            });
        }

        return items;
    }

    getDeadCodeWarnings(statements: Statement[], sourceText: string, uri: string): any[] {
        const procs = this.collectProcs(statements, sourceText, uri);
        return procs
            .filter(p => p.isPrivate && p.refCount === 0)
            .map(p => ({
                range: {
                    start: { line: p.line, character: 0 },
                    end:   { line: p.line, character: 0 },
                },
                severity: 2, // warning
                message: `'${p.name}' は参照されていません (Private Dead code 候補)`,
                source: 'vba-runner',
            }));
    }

    private collectProcs(statements: Statement[], sourceText: string, uri: string): ProcInfo[] {
        const result: ProcInfo[] = [];

        const testProcNames = new Set<string>();
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                if (proc.name.name.toLowerCase().startsWith('test_')) {
                    testProcNames.add(proc.name.name.toLowerCase());
                }
            }
        }

        for (const stmt of statements) {
            if (stmt.type !== 'ProcedureDeclaration') continue;
            const proc = stmt as ProcedureDeclaration;
            if (!proc.loc) continue;

            const name = proc.name.name;
            const line = proc.loc.start.line - 1;
            const endLine = proc.loc.end.line - 1;
            const isPrivate = proc.scope === 'private';

            const hasRequiredParams = proc.parameters.some(
                p => !p.isOptional && !p.isParamArray && p.defaultValue == null
            );
            const isTestProc = name.toLowerCase().startsWith('test_') && proc.parameters.length === 1;

            // Count call sites from outside this procedure's own body
            const refs = findAllReferences(sourceText, name, uri, statements, true);
            // externalRefs excludes both the declaration and intra-proc references
            const externalRefs = refs.filter(r =>
                r.range.start.line < line || r.range.start.line > endLine
            );
            const refCount = externalRefs.length;

            // テスト済み: Test_* プロシージャのソーステキスト中に name が登場するか
            const isTested = [...testProcNames].some(testName =>
                sourceText.toLowerCase().includes(name.toLowerCase()) &&
                this.testProcReferences(statements, testName, name.toLowerCase(), sourceText)
            );

            result.push({ name, line, endLine, isPrivate, hasRequiredParams, isTestProc, refCount, isTested });
        }

        return result;
    }

    private testProcReferences(statements: Statement[], testProcName: string, targetLower: string, sourceText: string): boolean {
        const testProc = statements.find(s =>
            s.type === 'ProcedureDeclaration' &&
            (s as ProcedureDeclaration).name.name.toLowerCase() === testProcName
        ) as ProcedureDeclaration | undefined;

        if (!testProc?.loc) return false;

        const lines = sourceText.split('\n');
        const start = testProc.loc.start.line - 1;
        const end = testProc.loc.end.line - 1;
        const procBody = lines.slice(start, end + 1).join('\n');
        return new RegExp(`(?<![a-zA-Z0-9_])${targetLower}(?![a-zA-Z0-9_])`, 'i').test(procBody);
    }
}
