import { ProcedureDeclaration, VariableDeclaration, Statement } from '../engine/parser';
import { findAllReferences } from './references-provider';

/**
 * VBA/Excel/Access/MSForms の組み込みイベント名（すべて小文字）のホワイトリスト。
 * プロシージャ名が "<prefix>_<suffix>" の形式で suffix がこのセットに一致する場合、
 * イベントハンドラーとみなし Dead Code 警告を抑制する。
 */
const KNOWN_VBA_EVENT_NAMES = new Set([
    // --- Worksheet ---
    'activate', 'beforedoubleclick', 'beforerightclick', 'calculate', 'change',
    'deactivate', 'followhyperlink', 'pivottableaftervaluechange',
    'pivottablebeforeallocatechanges', 'pivottablebeforecommitchanges',
    'pivottablebeforediscardchanges', 'pivottablechangesync', 'pivottableupdate',
    'selectionchange', 'tableupdate',
    // --- Workbook ---
    'addininstall', 'addinuninstall', 'afterremotechange', 'aftersave',
    'afterxmlexport', 'afterxmlimport', 'beforeclose', 'beforeprint',
    'beforeremotechange', 'beforesave', 'beforexmlexport', 'beforexmlimport',
    'modelchange', 'newchart', 'newsheet', 'open', 'pivottablecloseconnection',
    'pivottableopenconnection', 'rowsetcomplete', 'sheetactivate', 'sheetbeforedelete',
    'sheetbeforedoubleclick', 'sheetbeforerightclick', 'sheetcalculate', 'sheetchange',
    'sheetdeactivate', 'sheetfollowhyperlink', 'sheetlensgalleryrendercomplete',
    'sheetpivottableaftervaluechange', 'sheetpivottablebeforeallocatechanges',
    'sheetpivottablebeforecommitchanges', 'sheetpivottablebeforediscardchanges',
    'sheetpivottablechangesync', 'sheetpivottableupdate', 'sheetselectionchange',
    'sheettableupdate', 'syncevent', 'windowactivate', 'windowdeactivate', 'windowresize',
    // --- Chart ---
    'dragover', 'dragplot', 'mousedown', 'mousemove', 'mouseup', 'resize', 'select',
    'serieschange',
    // --- Application ---
    'aftercalculate', 'newworkbook', 'protectedviewwindowactivate',
    'protectedviewwindowbeforeclose', 'protectedviewwindowbeforeedit',
    'protectedviewwindowdeactivate', 'protectedviewwindowopen', 'protectedviewwindowresize',
    'workbookactivate', 'workbookaddininstall', 'workbookaddinuninstall',
    'workbookafterremotechange', 'workbookaftersave', 'workbookafterxmlexport',
    'workbookafterxmlimport', 'workbookbeforeclose', 'workbookbeforeprint',
    'workbookbeforeremotechange', 'workbookbeforesave', 'workbookbeforexmlexport',
    'workbookbeforexmlimport', 'workbookdeactivate', 'workbookmodelchange',
    'workbooknewchart', 'workbooknewsheet', 'workbookopen',
    'workbookpivottablecloseconnection', 'workbookpivottableopenconnection',
    'workbookrowsetcomplete', 'workbooksync',
    // --- UserForm ---
    'addcontrol', 'beforedragover', 'beforedroportpaste', 'click', 'dblclick', 'error',
    'initialize', 'keydown', 'keypress', 'keyup', 'layout', 'queryclose', 'removecontrol',
    'scroll', 'spindown', 'spinup', 'terminate', 'zoom',
    // --- MSForms controls (UserForm 上のコントロールイベント) ---
    'afterupdate', 'beforeupdate', 'dropbuttonclick', 'additem', 'removeitem',
    'enter', 'exit', 'gotfocus', 'lostfocus',
    // --- Access Form/Report ---
    'applyfilter', 'beforedelconfirm', 'afterdelconfirm', 'beforeinsert', 'afterinsert',
    'current', 'delete', 'dirty', 'filter', 'load', 'page', 'print', 'recordexit',
    'timer', 'undo', 'unload',
]);

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
    nameChar: number;       // 0-based column of the procedure name identifier
    isPrivate: boolean;
    hasRequiredParams: boolean;
    isTestProc: boolean;    // Test_* with exactly 1 param (the assert helper)
    refCount: number;       // call sites from outside this procedure
    isTested: boolean;      // any Test_* proc references this one
    isEventHandler: boolean; // WithEvents 変数・クラスライフサイクル・既知イベント名サフィックスに該当するイベントハンドラー
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

            // ▶ Run / 🐛 Debug: parameterless、または Test_* の assert 1引数プロシージャ
            if (!proc.hasRequiredParams || proc.isTestProc) {
                items.push({
                    range,
                    command: {
                        title: '▶ Run',
                        command: 'vba-runner.runProcedure',
                        arguments: [uri, proc.name, proc.isTestProc],
                    },
                });
                items.push({
                    range,
                    command: {
                        title: '🐛 Debug',
                        command: 'vba-runner.debugProcedure',
                        arguments: [uri, proc.name],
                    },
                });
            }

            // N references
            const refLabel = proc.isEventHandler
                ? '🔔 Event Handler'
                : proc.refCount === 0
                    ? (proc.isPrivate ? '⚠ 0 references' : '0 references')
                    : `${proc.refCount} reference${proc.refCount !== 1 ? 's' : ''}`;
            items.push({
                range,
                command: {
                    title: refLabel,
                    command: 'vba-runner.findReferences',
                    arguments: [uri, proc.line, proc.nameChar],
                },
            });

            // テスト状態
            items.push({
                range,
                command: {
                    title: proc.isTested ? '✓ Tested' : 'Untested',
                    command: proc.isTested ? 'vba-runner.goToTest' : 'vba-runner.generateTest',
                    arguments: [uri, proc.name],
                },
            });

            // Show in Call Graph
            items.push({
                range,
                command: {
                    title: '📊 Show in Call Graph',
                    command: 'vba-runner.showInCallGraph',
                    arguments: [uri, proc.name],
                },
            });

        }

        return items;
    }

    getDeadCodeWarnings(statements: Statement[], sourceText: string, uri: string): any[] {
        const procs = this.collectProcs(statements, sourceText, uri);
        return procs
            .filter(p => p.isPrivate && p.refCount === 0 && !p.isEventHandler)
            .map(p => ({
                range: {
                    start: { line: p.line, character: 0 },
                    end:   { line: p.line, character: 0 },
                },
                severity: 2, // warning
                message: `'${p.name}' is not referenced (Private dead code candidate)`,
                l10nKey: "'{0}' is not referenced (Private dead code candidate)",
                l10nArgs: [p.name],
                source: 'vba-runner',
            }));
    }

    /** モジュール内の WithEvents 変数名（小文字）を収集する */
    private collectWithEventsVarNames(statements: Statement[]): Set<string> {
        const names = new Set<string>();
        for (const stmt of statements) {
            if (stmt.type !== 'VariableDeclaration') continue;
            for (const decl of (stmt as VariableDeclaration).declarations) {
                if (decl.isWithEvents) names.add(decl.name.name.toLowerCase());
            }
        }
        return names;
    }

    /** プロシージャ名がイベントハンドラーかどうかを判定する。
     *  1. Class_Initialize / Class_Terminate はクラスライフサイクルとして無条件に該当
     *  2. <varName>_<anything> で varName が WithEvents 宣言済みの変数名と一致する場合
     *  3. <anything>_<eventName> で eventName が KNOWN_VBA_EVENT_NAMES に一致する場合 */
    private isEventHandlerProc(nameLower: string, withEventsVars: Set<string>): boolean {
        if (nameLower === 'class_initialize' || nameLower === 'class_terminate') return true;
        const sep = nameLower.indexOf('_');
        if (sep <= 0) return false;
        const prefix = nameLower.slice(0, sep);
        const suffix = nameLower.slice(sep + 1);
        if (withEventsVars.has(prefix)) return true;
        return KNOWN_VBA_EVENT_NAMES.has(suffix);
    }

    private collectProcs(statements: Statement[], sourceText: string, uri: string): ProcInfo[] {
        const result: ProcInfo[] = [];

        const withEventsVars = this.collectWithEventsVarNames(statements);

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
            const nameChar = (proc.name.loc?.start.column ?? 1) - 1;
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

            const isEventHandler = this.isEventHandlerProc(name.toLowerCase(), withEventsVars);

            // テスト済み: Test_* プロシージャのソーステキスト中に name が登場するか
            const isTested = [...testProcNames].some(testName =>
                sourceText.toLowerCase().includes(name.toLowerCase()) &&
                this.testProcReferences(statements, testName, name.toLowerCase(), sourceText)
            );

            result.push({ name, line, endLine, nameChar, isPrivate, hasRequiredParams, isTestProc, refCount, isTested, isEventHandler });
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
