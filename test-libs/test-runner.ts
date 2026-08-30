import * as fs from 'fs';
import * as path from 'path';
import { globSync, isDynamicPattern } from 'tinyglobby';
import { Lexer } from '../src/engine/lexer';
import { Parser, ParseError, TypeDeclaration, Program } from '../src/engine/parser';
import { Evaluator, SpyRecord, vbaTrue, vbaFalse, VbaBoolean, vbaNull, vbaEmpty, vbaNothing } from '../src/engine/evaluator';
import { VbaCurrency, VbaDecimal } from '../src/engine/vba-types';
import type { VbaType, VbaDefaultProperty, VbaIterable, VbaComObject } from '../src/engine/vba-types';
import { FileSystem, MemoryFileSystem } from '../src/engine/filesystem';
import { preprocess, stripVBAFileHeader, CompilerConstants } from '../src/engine/preprocessor';
import { loadMocks } from './mock-loader';
import type { MockLoadContext } from './mock-loader';
import { injectExcelStub } from './excel-stub';
import { MockApplication, MockWorksheet, MockRange, MockRows, MockColumns, MockWorkbook } from '../src/engine/mock/MockExcel';
import { LSPServer } from '../src/lsp/server';
import { publishDiagnostics, DiagnosticApi, RawDiagnostic } from '../src/lsp/diagnostic-publisher';
export { vbaTrue, vbaFalse, vbaNull, vbaEmpty };
export { MockApplication, MockWorksheet, MockRange, MockRows, MockColumns, MockWorkbook };
export type { VbaType, VbaDefaultProperty, VbaIterable, VbaComObject };

const VBA_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

/**
 * VBA 内部値を JS ネイティブ値に正規化する。
 * - VbaBoolean → boolean
 * - vbaNothing → null
 * - VbaCurrency / VbaDecimal → そのまま返す（精度損失を避けるため number に変換しない）
 * それ以外はそのまま返す。
 */
function normalizeVbaValue(raw: any): any {
    if (raw instanceof VbaBoolean) return raw.value !== 0;
    if (raw === vbaNothing) return null;
    return raw;
}

/**
 * run() のログ出力用に引数を文字列化する。
 * VbaCurrency / VbaDecimal が BigInt を含むため JSON.stringify が失敗するケースを回避する。
 */
function formatVbaArg(a: any): string {
    if (a === null) return 'Nothing';
    if (a instanceof VbaCurrency) return `CCur(${a.toString()})`;
    if (a instanceof VbaDecimal) return `CDec(${a.toString()})`;
    if (typeof a === 'object') {
        try {
            return JSON.stringify(a);
        } catch {
            return String(a);
        }
    }
    return String(a);
}

export class VBARunner {
    public evaluator: Evaluator;
    private _asts: Program[] = [];
    private _moduleNames: string[] = [];
    private _resolved = false;
    private readonly _quiet: boolean;

    /** `excelStub: true` のとき注入された MockApplication。セル初期値の設定に使う。 */
    public readonly excelStub: MockApplication | null = null;

    /** VBA コードが使う仮想ファイルシステム。テスト用入力ファイルの事前配置などに使う。 */
    get fs(): FileSystem { return this.evaluator.fs; }

    constructor(pathOrDir: string | string[] | null = null, config: {
        sandboxRoot?: string,
        env?: Record<string, string>,
        compilerConstants?: CompilerConstants,
        /** 複数 VBARunner インスタンスで共有する仮想ファイルシステム。 */
        fs?: FileSystem,
        excelStub?: boolean | MockApplication,
        /** true にすると run() の `[PASS] ...` ログを抑制する */
        quiet?: boolean,
        /** Debug.Print の出力先（既定: console.log） */
        onPrint?: (s: string) => void,
    } = {}) {
        this._quiet = config.quiet ?? false;
        this.evaluator = new Evaluator(config.onPrint ?? console.log, {
            ...config,
            fs: config.fs ?? new MemoryFileSystem(),
            allowTopLevelStatements: false,
        });
        let mockContext: MockLoadContext = {};
        if (config.excelStub) {
            const app = config.excelStub === true ? undefined : config.excelStub;
            (this as any).excelStub = injectExcelStub(this.evaluator, app);
            mockContext = {
                excel: { Application: (this as any).excelStub },
            };
        }

        if (!pathOrDir) return;

        // glob パターンまたはファイルパス配列の場合: mock スキャンなしで順番にロード
        if (Array.isArray(pathOrDir) || isDynamicPattern(pathOrDir)) {
            const patterns = Array.isArray(pathOrDir) ? pathOrDir : [pathOrDir];
            const files = patterns.flatMap(p =>
                isDynamicPattern(p)
                    ? globSync(p, { cwd: process.cwd(), absolute: true }).sort()
                    : [path.resolve(p)]
            );
            for (const file of files) {
                this._loadSingleFile(file, config.compilerConstants);
            }
            return;
        }

        const stat = fs.statSync(pathOrDir);
        const dir = stat.isDirectory() ? pathOrDir : path.dirname(pathOrDir);
        mockContext.sourceDirectory = dir;
        const files = stat.isDirectory()
            ? fs.readdirSync(pathOrDir)
                  .filter(f => VBA_EXTENSIONS.has(path.extname(f).toLowerCase()))
                  .sort()
                  .map(f => path.join(pathOrDir, f))
            : [pathOrDir];

        // モックを先にロード（標準ライブラリより後、本番コードより前に注入）
        const mockModules = loadMocks(dir, this.evaluator, mockContext);
        for (const { ast, moduleName } of mockModules) {
            this._asts.push(ast);
            this._moduleNames.push(moduleName);
        }

        // Pass 1: 全モジュールをロードして手続き・変数を登録する
        for (const file of files) {
            try {
                this._loadSingleFile(file, config.compilerConstants);
            } catch (e: any) {
                // ディレクトリ読み込み時: 意図的に構文エラーを含む仕様ファイル
                // （例: CompileError.bas）はスキップしてログだけ残す。
                // 単一ファイル指定時はエラーをそのまま上げる。
                if (stat.isDirectory() && e.message && /^\[.*\] Parse error:/i.test(e.message)) {
                    console.warn(`  (skip: parse errors in spec) ${path.basename(file)}`);
                    continue;
                }
                throw e;
            }
        }

        // Pass 2（resolveIdentifiers）は set() による定数注入を可能にするため
        // 初回 run()/eval() 呼び出し時まで遅延する。
    }

    private _loadSingleFile(file: string, compilerConstants?: CompilerConstants): void {
        let source = fs.readFileSync(file, 'utf-8');
        try {
            const moduleName = path.basename(file, path.extname(file));
            this.evaluator.setSourceModule(moduleName);
            const ext = path.extname(file).toLowerCase();
            source = stripVBAFileHeader(source);
            source = preprocess(source, compilerConstants);
            const isRawCls = ext === '.cls'
                && !source.trim().toLowerCase().startsWith('class ')
                && !source.toLowerCase().includes('end class');
            const parseOpts = isRawCls ? { parseAsClass: moduleName, sourceLines: source.split('\n') } : { sourceLines: source.split('\n') };
            const ast = new Parser(new Lexer(source).tokenize(), parseOpts).parse();
            this._asts.push(ast);
            this._moduleNames.push(moduleName);
            this.evaluator.evaluateModule(ast);
        } catch (e: any) {
            throw new Error(`[${path.basename(file)}] ${e.message}`);
        }
    }

    private _ensureResolved(): void {
        if (this._resolved) return;
        this._resolved = true;
        this.evaluator.resolveIdentifiers(
            this._asts.map((ast, i) => ({ ast, moduleName: this._moduleNames[i] }))
        );
    }

    run(procedureName: string, args: any[] = [], type?: 'get' | 'let' | 'set'): any {
        this._ensureResolved();
        const start = Date.now();
        const raw = this.evaluator.callProcedure(procedureName, args, type);
        const result = normalizeVbaValue(raw);
        const duration = Date.now() - start;
        const formatArgs = args.map(a => formatVbaArg(a)).join(', ');
        const typeStr = type ? `:${type}` : '';
        if (!this._quiet) {
            console.log(`[PASS] ${procedureName}${typeStr}(${formatArgs}) -> ${result} (${duration}ms)`);
        }
        return result;
    }

    eval(exprString: string): any {
        this._ensureResolved();
        const raw = this.evaluator.evalExpression(exprString);
        return normalizeVbaValue(raw);
    }

    set(name: string, value: any): void {
        this.evaluator.set(name, value);
    }

    /**
     * 複数の定数・変数を一括で VBA 環境に注入する。
     * vba-analyzer の `--gen-test-dir` で生成した `allConstants` を渡す典型用途を想定。
     *
     * @example
     *   import { allConstants } from './generated/const';
     *   vbaRunner.setConstants(allConstants);
     */
    setConstants(constants: Record<string, any>): void {
        for (const [name, value] of Object.entries(constants)) {
            this.evaluator.setConstant(name, value);
        }
    }

    /**
     * COM オブジェクトのファクトリを登録する。factory() が返すオブジェクトは
     * `VbaComObject`（`__progId__: string` を持つ）である必要がある。
     * 既存の組み込みスタブよりも優先される。
     *
     * @example
     *   import { createRegExpMock } from '../../test-libs/regexp-mock';
     *   const vbaRunner = new VBARunner('source.bas');
     *   vbaRunner.registerComObject(createRegExpMock);
     */
    registerComObject(factory: () => VbaComObject): void {
        this.evaluator.registerComObject(factory);
    }

    /**
     * VBA 関数をスパイでラップし、呼び出し記録を返す。
     * returnFn を指定すると戻り値をオーバーライドできる。
     *
     * @example
     *   const spy = vbaRunner.spy('MsgBox');
     *   vbaRunner.run('MyProc', []);
     *   console.log(spy.callCount);           // 呼び出し回数
     *   console.log(spy.lastCall);            // 最後の引数配列
     *   console.log(spy.calledWith('Error!')); // 引数一致チェック
     *
     * @example MsgBox の戻り値をモックする（vbYes=6）
     *   const spy = vbaRunner.spy('MsgBox', () => 6);
     */
    spy(name: string, returnFn?: (...args: any[]) => any): SpyRecord {
        return this.evaluator.spy(name, returnFn);
    }

    /**
     * ロード済みの VBA ソースに含まれる `Type` 宣言を TypeScript の型情報として返す。
     * 各フィールドの VBA 型は TypeScript 型文字列にマッピングされる。
     *
     * @returns `{ [TypeName]: { [fieldName]: tsType } }` の形式
     *
     * @example
     *   const vbaRunner = new VBARunner('src/vba/inventory.bas');
     *   const types = vbaRunner.getTypeDefinitions();
     *   // => { InventoryParams: { CurrentStock: 'number', SoldUnits: 'number', ... } }
     *
     *   // TypeScript interface として文字列出力する場合:
     *   for (const [name, fields] of Object.entries(types)) {
     *       const body = Object.entries(fields).map(([f, t]) => `  ${f}: ${t};`).join('\n');
     *       console.log(`interface ${name} {\n${body}\n}`);
     *   }
     */
    getTypeDefinitions(): Record<string, Record<string, string>> {
        const result: Record<string, Record<string, string>> = {};
        for (const ast of this._asts) {
            for (const stmt of ast.body) {
                if (stmt.type === 'TypeDeclaration') {
                    const decl = stmt as TypeDeclaration;
                    result[decl.name] = {};
                    for (const member of decl.members) {
                        result[decl.name][member.name] = vbaTypeToTs(member.memberType);
                    }
                }
            }
        }
        return result;
    }

    /**
     * Fix the date/time returned by `Now`, `Date`, `Time`, and `Timer`.
     * Accepts any value parseable by `new Date(...)`, e.g. `'2024-12-31T09:00:00'`.
     * Pass `null` to restore real system time.
     */
    mockDate(dateStr: string | null): void {
        if (dateStr === null) {
            this.evaluator.setNowFn(null);
        } else {
            const fixed = new Date(dateStr);
            this.evaluator.setNowFn(() => fixed);
        }
    }
}

function vbaTypeToTs(vbaType: string): string {
    switch (vbaType.toLowerCase()) {
        case 'integer': case 'long': case 'single': case 'double':
        case 'currency': case 'decimal': case 'byte': case 'date':
            return 'number';
        case 'string':    return 'string';
        case 'boolean':   return 'boolean';
        case 'object':    return 'object';
        default:          return 'any';
    }
}

// Keep backward compatibility
export function runVBARunner(filePath: string, procedureName: string, args: any[], config: { sandboxRoot?: string, env?: Record<string, string> } = {}): any {
    const vbaRunner = new VBARunner(filePath, config);
    return vbaRunner.run(procedureName, args);
}

/** evalVBASingle / evalVBAModules 共通オプション */
export interface EvalOptions {
    /** Debug.Print の出力先（デフォルト: console.log） */
    onPrint?: (s: string) => void;
    /** 仮想ファイルシステム（ファイル I/O テスト用） */
    fs?: FileSystem;
    /** Sandbox ルートパス */
    sandboxRoot?: string;
    /** Environ() が参照する環境変数 */
    env?: Record<string, string>;
    /** §5.6.10 Tier 6 オブジェクト（Excel Application 等）。
     *  Option Explicit 違反チェックは Pass 2 で即時実行されるため、
     *  defaultBindingObject を使う場合はここで渡す必要がある。 */
    defaultBindingObject?: any;
    /** Evaluator 生成直後・evaluateModule 呼び出し前に実行されるコールバック。
     *  setNowFn / set / setConstant 等の前処理に使う。 */
    setup?: (ev: Evaluator) => void;
    /** パース直後・Evaluator 生成前に実行されるコールバック。
     *  ast.diagnostics のチェック等に使う。throw すれば評価を中断できる。 */
    afterParse?: (ast: Program) => void;
    /** vba-runner 拡張: モジュールレベル実行文をプロシージャの後にも書けるようにするか（デフォルト true）。
     *  false にすると標準 VBA 仕様どおりコンパイルエラーになる。 */
    allowTopLevelStatements?: boolean;
    /**
     * Run the same inline sources through the LSP diagnostic publication path.
     * `clean` requires that no published engine Error diagnostic is present.
     * `lsp-error` explicitly expects a published LSP Error while leaving the
     * engine's own load and call timing unchanged. Failures are deferred until
     * process shutdown so they cannot replace an engine exception.
     */
    diagnostics?: DiagnosticConsistencyOptions;
}

export type DiagnosticConsistencyOptions =
    | { expectation: 'clean' }
    | { expectation: 'lsp-error'; message?: RegExp }
    | { expectation: 'skip'; reason: string };

type InlineEvalModule = { name: string; code: string; parseAsClass?: string };

type PublishedDiagnostic = {
    range: RawDiagnostic['range'];
    severity: number;
    message: string;
    source?: string;
    code?: string | number;
};

const diagnosticApi: DiagnosticApi<unknown, PublishedDiagnostic> = {
    range: (startLine, startCharacter, endLine, endCharacter) => ({
        start: { line: startLine, character: startCharacter },
        end: { line: endLine, character: endCharacter },
    }),
    diagnostic: (range, message, severity) => ({
        range: range as PublishedDiagnostic['range'], message, severity,
    }),
    setSource: (diagnostic, source) => { diagnostic.source = source; },
    setCode: (diagnostic, code) => { if (code != null) diagnostic.code = code; },
};

function inlineModuleUri(module: InlineEvalModule): string {
    const extension = module.parseAsClass ? '.cls' : '.bas';
    const name = encodeURIComponent(module.name.replace(/\.[^.]+$/, ''));
    return `file:///vba-runner-consistency/${name}${extension}`;
}

type DiagnosticContext = {
    server: LSPServer;
    cache: Map<string, PublishedDiagnostic[]>;
};

function collectBindingObjectNames(obj: any): Set<string> {
    const names = new Set<string>();
    let current = obj;
    while (current && current !== Object.prototype && current !== Function.prototype) {
        for (const name of Object.getOwnPropertyNames(current)) names.add(name.toLowerCase());
        current = Object.getPrototypeOf(current);
    }
    return names;
}

function createDiagnosticContext(
    modules: InlineEvalModule[],
    knownExternalNames: Iterable<string> = [],
): DiagnosticContext {
    const server = new LSPServer();
    server.setMockedHostIdentifiers(knownExternalNames);
    for (const module of modules) {
        server.didOpen(inlineModuleUri(module), module.code);
    }
    return { server, cache: new Map() };
}

function publishedDiagnosticsFor(
    context: DiagnosticContext,
    uri: string,
): PublishedDiagnostic[] {
    const cached = context.cache.get(uri);
    if (cached) return cached;
    let result: PublishedDiagnostic[] = [];
    const collection = {
        set: (_uri: string, diagnostics: PublishedDiagnostic[]) => { result = diagnostics; },
    };
    publishDiagnostics(uri, context.server.getDiagnostics(uri) as RawDiagnostic[], collection, diagnosticApi,
        (key, args) => `${key}${args.length ? ` ${args.join(',')}` : ''}`);
    context.cache.set(uri, result);
    return result;
}

const deferredDiagnosticFailures: string[] = [];
let diagnosticFailureHandlerInstalled = false;

function deferDiagnosticFailure(message: string): void {
    deferredDiagnosticFailures.push(message);
    if (diagnosticFailureHandlerInstalled) return;
    diagnosticFailureHandlerInstalled = true;
    process.once('beforeExit', () => {
        if (deferredDiagnosticFailures.length === 0) return;
        console.error('[FAIL] LSP diagnostic consistency checks:');
        for (const failure of deferredDiagnosticFailures) console.error(`  ${failure}`);
        process.exitCode = 1;
    });
}

type ProcedureSpan = {
    uri: string;
    moduleName: string;
    name: string;
    startLine: number;
    endLine: number;
};

function procedureSpans(modules: InlineEvalModule[]): ProcedureSpan[] {
    const spans: ProcedureSpan[] = [];
    for (const module of modules) {
        let ast: Program;
        try {
            ast = new Parser(new Lexer(module.code).tokenize(), {
                parseAsClass: module.parseAsClass,
                errorRecovery: true,
                sourceLines: module.code.split('\n'),
            } as any).parse();
        } catch {
            continue;
        }
        const visit = (statements: any[]): void => {
            for (const statement of statements) {
                if (statement.type === 'ProcedureDeclaration' && statement.loc) {
                    spans.push({
                        uri: inlineModuleUri(module),
                        moduleName: module.name,
                        name: statement.name.name.toLowerCase(),
                        startLine: statement.loc.start.line - 1,
                        endLine: statement.loc.end.line - 1,
                    });
                }
                if (statement.type === 'ClassDeclaration' && Array.isArray(statement.body)) {
                    visit(statement.body);
                }
            }
        };
        visit(ast.body);
    }
    return spans;
}

function findProcedureSpan(
    spans: ProcedureSpan[],
    name: string,
    moduleName?: string,
): ProcedureSpan | undefined {
    const parts = name.toLowerCase().split('.');
    const procName = parts.length === 2 ? parts[1] : parts[0];
    const qualifier = (moduleName ?? (parts.length === 2 ? parts[0] : undefined))?.toLowerCase();
    return spans.find(span => span.name === procName
        && (!qualifier || span.moduleName.toLowerCase() === qualifier));
}

function errorsInProcedure(
    context: DiagnosticContext,
    span: ProcedureSpan,
): PublishedDiagnostic[] {
    return publishedDiagnosticsFor(context, span.uri).filter(diagnostic =>
        diagnostic.severity === 0
        && diagnostic.source === 'vba-runner'
        && diagnostic.range.end.line >= span.startLine
        && diagnostic.range.start.line <= span.endLine);
}

function isPreprocError(error: unknown): boolean {
    return typeof (error as any)?.message === 'string'
        && (error as any).message.startsWith('Compile error:');
}

/**
 * A procedure wrapper can observe an error rethrown by a nested call.  The
 * engine records the originating frame first, so only that frame should be
 * compared with the LSP diagnostics.  Otherwise a compile error in `Inner`
 * is incorrectly reported again while the `Outer` wrapper unwinds.
 */
function errorOriginatesInProcedure(
    error: unknown,
    span: ProcedureSpan,
): boolean {
    const vbaError = error as any;
    const origin = Array.isArray(vbaError?.vbaStack) ? vbaError.vbaStack[0] : undefined;
    if (!origin) return true;
    const originName = typeof origin.name === 'string' ? origin.name.toLowerCase() : undefined;
    const originModule = typeof origin.moduleName === 'string' ? origin.moduleName.toLowerCase() : undefined;
    if (originName && originName !== span.name) return false;
    if (originModule && originModule !== span.moduleName.toLowerCase()) return false;
    return true;
}

function checkCalledProcedureDiagnostics(
    context: DiagnosticContext,
    span: ProcedureSpan | undefined,
    error: unknown,
    options: DiagnosticConsistencyOptions,
): void {
    if (options.expectation === 'skip') return;
    if (!span) return;
    // A nested call may rethrow its original error through this wrapper.  The
    // inner wrapper already checked the originating procedure; do not compare
    // the parent's source range with that error's diagnostics.
    if (error && !errorOriginatesInProcedure(error, span)) return;
    const errors = errorsInProcedure(context, span);
    const compileError = isPreprocError(error);
    if (options.expectation === 'clean') {
        if (!compileError && errors.length > 0) {
            deferDiagnosticFailure(`expected clean diagnostics for ${span.moduleName}.${span.name}, got: ${errors.map(d => d.message).join('; ')}`);
        }
        if (compileError && errors.length === 0) {
            deferDiagnosticFailure(`engine compile error in ${span.moduleName}.${span.name} has no LSP Error`);
        }
        return;
    }
    if (errors.length === 0) {
        deferDiagnosticFailure(`expected an LSP Error in ${span.moduleName}.${span.name}, but none was published`);
    } else if (options.message && !errors.some(d => options.message!.test(d.message))) {
        deferDiagnosticFailure(`LSP Error in ${span.moduleName}.${span.name} did not match ${options.message}`);
    }
}

function attachDiagnosticConsistency(
    ev: Evaluator,
    modules: InlineEvalModule[],
    options: DiagnosticConsistencyOptions,
    defaultBindingObject?: any,
): void {
    if (options.expectation === 'skip') {
        if (options.reason.trim().length === 0) {
            throw new Error('diagnostics skip requires a non-empty reason');
        }
        return;
    }
    let externalNamesKey = '';
    let context: DiagnosticContext;
    const currentContext = (): DiagnosticContext => {
        const names = new Set(ev.getDiagnosticExternalNames());
        if (defaultBindingObject !== undefined) {
            for (const name of collectBindingObjectNames(defaultBindingObject)) names.add(name);
        }
        const sortedNames = [...names].sort();
        const key = sortedNames.join('\0');
        if (!context || key !== externalNamesKey) {
            context = createDiagnosticContext(modules, sortedNames);
            externalNamesKey = key;
        }
        return context;
    };
    const spans = procedureSpans(modules);
    const originalCall = ev.callProcedure;
    (ev as any).callProcedure = function(this: Evaluator, name: string, args: any[], ...rest: any[]): any {
        const moduleName = rest[1] as string | undefined;
        const span = findProcedureSpan(spans, name, moduleName);
        try {
            const result = originalCall.call(this, name, args, ...rest);
            checkCalledProcedureDiagnostics(currentContext(), span, undefined, options);
            return result;
        } catch (error) {
            checkCalledProcedureDiagnostics(currentContext(), span, error, options);
            throw error;
        }
    };
    const originalCheck = ev.checkProcedure;
    (ev as any).checkProcedure = function(this: Evaluator, name: string): void {
        const span = findProcedureSpan(spans, name);
        try {
            originalCheck.call(this, name);
            checkCalledProcedureDiagnostics(currentContext(), span, undefined, options);
        } catch (error) {
            checkCalledProcedureDiagnostics(currentContext(), span, error, options);
            throw error;
        }
    };
}

/**
 * 単一モジュールのインライン VBA コードを評価して Evaluator を返す。
 * Pass 1（登録）と Pass 2（定数確定）を両方実行する。
 * tests/spec/ のテストで evalVBASingle をローカル定義する代わりに使う。
 */
export function evalVBASingle(code: string, options?: EvalOptions): Evaluator {
    const diagnosticOptions = options?.diagnostics ?? { expectation: 'clean' as const };
    const ast = new Parser(new Lexer(code).tokenize(), { sourceLines: code.split('\n') }).parse();
    options?.afterParse?.(ast);
    const ev = new Evaluator(options?.onPrint ?? console.log, {
        fs: options?.fs,
        sandboxRoot: options?.sandboxRoot,
        env: options?.env,
        allowTopLevelStatements: options?.allowTopLevelStatements,
    });
    if (options?.defaultBindingObject !== undefined) {
        ev.setDefaultBindingObject(options.defaultBindingObject);
    }
    options?.setup?.(ev);
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    attachDiagnosticConsistency(ev, [{ name: 'InlineModule', code }], diagnosticOptions,
        options?.defaultBindingObject);
    return ev;
}

/**
 * 複数モジュールのインライン VBA コードを評価して Evaluator を返す。
 * 各モジュールを Pass 1（evaluateModule）でロードした後、全モジュール分まとめて
 * Pass 2（resolveIdentifiers）を実行する。
 * クロスモジュール定数参照（ModA.X = ModB.Y + 1 など）を正しく解決するため、
 * resolveIdentifiers は全モジュールのロード後に1回だけ呼ぶ。
 */
export function evalVBAModules(
    modules: InlineEvalModule[],
    options?: EvalOptions,
): Evaluator {
    const diagnosticOptions = options?.diagnostics ?? { expectation: 'clean' as const };
    const ev = new Evaluator(options?.onPrint ?? console.log, {
        fs: options?.fs,
        sandboxRoot: options?.sandboxRoot,
        env: options?.env,
        allowTopLevelStatements: options?.allowTopLevelStatements,
    });
    if (options?.defaultBindingObject !== undefined) {
        ev.setDefaultBindingObject(options.defaultBindingObject);
    }
    options?.setup?.(ev);

    const asts = modules.map(({ name, code, parseAsClass }) => {
        const ast = new Parser(new Lexer(code).tokenize(), { parseAsClass, sourceLines: code.split('\n') }).parse();
        ev.setSourceModule(name);
        ev.evaluateModule(ast);
        return { ast, moduleName: name };
    });
    ev.resolveIdentifiers(asts);
    attachDiagnosticConsistency(ev, modules, diagnosticOptions, options?.defaultBindingObject);
    return ev;
}

// Minimal assert framework
export const assert = {
    strictEqual: (actual: any, expected: any, message?: string) => {
        const a = (actual && (actual instanceof VbaBoolean || (actual as any).__isVbaBoolean__)) ? actual.valueOf() : actual;
        const e = (expected && (expected instanceof VbaBoolean || (expected as any).__isVbaBoolean__)) ? expected.valueOf() : expected;
        if (a !== e) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected ${String(expected)} but got ${String(actual)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    deepStrictEqual: (actual: any, expected: any, message?: string) => {
        const normalize = (v: any) =>
            (v && (v instanceof VbaBoolean || (v as any).__isVbaBoolean__)) ? v.valueOf() : v;
        const seen = new WeakMap<object, WeakSet<object>>();
        const deepEqual = (a: any, e: any): boolean => {
            a = normalize(a);
            e = normalize(e);
            if (a === e) return true;
            if (a === null || e === null || typeof a !== 'object' || typeof e !== 'object') return false;
            if (!seen.has(a)) seen.set(a, new WeakSet());
            if (seen.get(a)!.has(e)) return true;
            seen.get(a)!.add(e);
            if (Array.isArray(a) !== Array.isArray(e)) return false;
            const aKeys = Object.keys(a);
            const eKeys = Object.keys(e);
            if (aKeys.length !== eKeys.length) return false;
            return aKeys.every(k => Object.prototype.hasOwnProperty.call(e, k) && deepEqual(a[k], e[k]));
        };
        if (!deepEqual(actual, expected)) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected ${JSON.stringify(expected)} but got ${JSON.stringify(actual)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    notStrictEqual: (actual: any, expected: any, message?: string) => {
        const a = (actual && (actual instanceof VbaBoolean || (actual as any).__isVbaBoolean__)) ? actual.valueOf() : actual;
        const e = (expected && (expected instanceof VbaBoolean || (expected as any).__isVbaBoolean__)) ? expected.valueOf() : expected;
        if (a === e) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected values to differ but both were ${String(actual)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    ok: (value: any, message?: string) => {
        if (!value) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected truthy value but got ${String(value)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    fail: (message?: string) => {
        console.error(`[FAIL] ${message || 'Assertion failed'}`);
        throw new Error(`Assertion Failed`);
    },
    doesNotThrow: (fn: () => void, message?: string) => {
        try {
            fn();
        } catch (e: any) {
            const detail = e?.message ?? String(e);
            console.error(`[FAIL] ${message || 'Expected no exception'} - threw: ${detail}`);
            throw new Error(`Assertion Failed`);
        }
    },
    throws: (fn: () => void, patternOrMessage?: string | RegExp, message?: string) => {
        let threw = false;
        let detail = '';
        try {
            fn();
        } catch (e: any) {
            threw = true;
            detail = e?.message ?? String(e);
        }
        if (!threw) {
            console.error(`[FAIL] ${message || (typeof patternOrMessage === 'string' ? patternOrMessage : 'Expected an exception but none was thrown')}`);
            throw new Error(`Assertion Failed`);
        }
        if (patternOrMessage instanceof RegExp && !patternOrMessage.test(detail)) {
            console.error(`[FAIL] ${message || 'Exception message did not match'} - pattern: ${patternOrMessage}, got: "${detail}"`);
            throw new Error(`Assertion Failed`);
        }
    },
    throwsMatch: (fn: () => void, pattern: string | RegExp, message?: string) => {
        let threw = false;
        let detail = '';
        try {
            fn();
        } catch (e: any) {
            threw = true;
            detail = e?.message ?? String(e);
        }
        if (!threw) {
            console.error(`[FAIL] ${message || 'Expected an exception but none was thrown'}`);
            throw new Error(`Assertion Failed`);
        }
        const matched = typeof pattern === 'string' ? detail.includes(pattern) : pattern.test(detail);
        if (!matched) {
            console.error(`[FAIL] ${message || 'Exception message did not match'} - pattern: ${pattern}, got: "${detail}"`);
            throw new Error(`Assertion Failed`);
        }
    },
};

/**
 * Compare floating-point results using both absolute and relative tolerance.
 *
 * The default is a few machine epsilons: it accepts harmless operation-order
 * differences while still rejecting ordinary numerical regressions. Callers
 * can tighten or widen either bound for a domain-specific comparison.
 */
export function assertClose(
    actual: number,
    expected: number,
    options: {
        absolute?: number;
        relative?: number;
        message?: string;
    } = {},
): void {
    const absolute = options.absolute ?? Number.EPSILON * 4;
    const relative = options.relative ?? Number.EPSILON * 4;
    const difference = Math.abs(actual - expected);
    const scale = Math.max(1, Math.abs(actual), Math.abs(expected));
    const tolerance = Math.max(absolute, relative * scale);
    if (!(difference <= tolerance)) {
        const label = options.message ?? 'Floating-point values are not close';
        console.error(`[FAIL] ${label} - Expected ${expected} but got ${actual} (difference ${difference}, tolerance ${tolerance})`);
        throw new Error('Assertion Failed');
    }
}

/** Pass 1 (parse) コンパイルエラー専用アサーション。Parser.parse() のみを try で囲み ParseError を検証する。 */
function assertCompileErrorImpl(opts: {
    label: string;
    phaseName: string;
    pattern: RegExp;
    expectedLine: number | undefined;
    okAction?: () => void;
    throwAction: () => void;
    extraCheck?: (e: unknown) => string | null;
}): void {
    const { label, phaseName, pattern, expectedLine, okAction, throwAction, extraCheck } = opts;
    if (okAction) {
        try {
            okAction();
        } catch (e: any) {
            console.error(`[FAIL] ${label}: Unexpected error before ${phaseName}: ${e?.message ?? e}`);
            throw new Error('Assertion Failed');
        }
    }
    let threw = false, msg = '';
    let caughtError: unknown;
    try {
        throwAction();
    } catch (e: any) {
        threw = true;
        caughtError = e;
        msg = e?.message ?? String(e);
    }
    if (!threw) {
        console.error(`[FAIL] ${label}: Expected ${phaseName} error but none was thrown`);
        throw new Error('Assertion Failed');
    }
    if (extraCheck) {
        const extraMsg = extraCheck(caughtError);
        if (extraMsg) {
            console.error(`[FAIL] ${label}: ${extraMsg}`);
            throw new Error('Assertion Failed');
        }
    }
    if (!pattern.test(msg)) {
        console.error(`[FAIL] ${label}: Message mismatch - pattern: ${pattern}, got: "${msg}"`);
        throw new Error('Assertion Failed');
    }
    if (expectedLine !== undefined && !new RegExp(`\\bline ${expectedLine}\\b`).test(msg)) {
        console.error(`[FAIL] ${label}: Line mismatch - expected line ${expectedLine}, got: "${msg}"`);
        throw new Error('Assertion Failed');
    }
}

export function assertCompileErrorPass1(src: string, expectedLine: number, pattern: RegExp, label: string): void {
    assertCompileErrorImpl({
        label, phaseName: 'parse',
        pattern, expectedLine,
        throwAction: () => new Parser(new Lexer(src).tokenize()).parse(),
        extraCheck: e => e instanceof ParseError ? null : `Expected ParseError (parse) but got: "${(e as any)?.message ?? e}"`,
    });
}

/** Pass 2 (resolve) コンパイルエラー専用アサーション。evaluateModule + resolveIdentifiers で throw を検証。 */
export function assertCompileErrorPass2(src: string, expectedLine: number, pattern: RegExp, label: string): void {
    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const ev = new Evaluator(console.log);
    assertCompileErrorImpl({
        label, phaseName: 'resolve (pass 2)',
        pattern, expectedLine,
        throwAction: () => { ev.evaluateModule(ast); ev.resolveIdentifiers([{ ast, moduleName: '' }]); },
    });
}

/** [resolve] evaluateModule OK → resolveIdentifiers で throw を検証。
 *  evalOptions: Evaluator のコンストラクター設定（allowTopLevelStatements 等）を上書きする場合に指定。 */
export function assertCompileErrorResolve(
    src: string,
    expectedLine: number | undefined,
    pattern: RegExp,
    label: string,
    evalOptions?: { allowTopLevelStatements?: boolean },
): void {
    const ast = new Parser(new Lexer(src).tokenize()).parse();
    const ev = new Evaluator(console.log, evalOptions);
    assertCompileErrorImpl({
        label, phaseName: 'resolve',
        pattern, expectedLine,
        okAction: () => ev.evaluateModule(ast),
        throwAction: () => ev.resolveIdentifiers([{ ast, moduleName: '' }]),
    });
}


/** [preproc] evalVBASingle OK → precheckProc のみ（本体実行なし）で throw を検証。 */
export function assertCompileErrorPreproc(src: string, procName: string, expectedLine: number | undefined, pattern: RegExp, label: string): void {
    let ev!: Evaluator;
    assertCompileErrorImpl({
        label, phaseName: 'preproc',
        pattern, expectedLine,
        okAction: () => { ev = evalVBASingle(src); },
        throwAction: () => ev.checkProcedure(procName),
    });
}

/** [exec] evalVBASingle OK → precheckProc 後の実行中の throw を検証。 */
export function assertCompileErrorExec(src: string, procName: string, expectedLine: number | undefined, pattern: RegExp, label: string): void {
    let ev!: Evaluator;
    assertCompileErrorImpl({
        label, phaseName: 'exec',
        pattern, expectedLine,
        okAction: () => {
            ev = evalVBASingle(src);
            // An exec case must survive the static precheck. Without this
            // guard, a preproc failure could be misreported as an execution
            // failure because callProcedure performs the same precheck.
            ev.checkProcedure(procName);
        },
        throwAction: () => ev.callProcedure(procName, []),
    });
}

/**
 * RUNNER: TBD ケース調査用: 各フェーズで実際にスローされるエラーメッセージをそのまま返す。
 * assertCompileError* は pattern にマッチした時点で正常終了してしまうため、
 * 実際のメッセージを観察する目的には使えない（pattern: /.+/i は常にマッチしてしまう）。
 */
export function captureCompileErrorMessage(
    phase: 'parse' | 'resolve' | 'preproc' | 'exec',
    src: string,
    procName?: string,
    evalOptions?: { allowTopLevelStatements?: boolean },
): string {
    try {
        if (phase === 'parse') {
            new Parser(new Lexer(src).tokenize()).parse();
            return '(no error thrown)';
        }
        const ast = new Parser(new Lexer(src).tokenize()).parse();
        const ev = new Evaluator(console.log, evalOptions);
        ev.evaluateModule(ast);
        if (phase === 'resolve') {
            ev.resolveIdentifiers([{ ast, moduleName: '' }]);
            return '(no error thrown)';
        }
        ev.resolveIdentifiers([{ ast, moduleName: '' }]);
        ev.callProcedure(procName ?? '__test__', []);
        return '(no error thrown)';
    } catch (e: any) {
        return e?.message ?? String(e);
    }
}
