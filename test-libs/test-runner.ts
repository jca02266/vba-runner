import * as fs from 'fs';
import * as path from 'path';
import { globSync, isDynamicPattern } from 'tinyglobby';
import { Lexer } from '../src/engine/lexer';
import { Parser, ParseError, TypeDeclaration, Program } from '../src/engine/parser';
import { Evaluator, SpyRecord, vbaTrue, vbaFalse, VbaBoolean, vbaNull, vbaEmpty, vbaNothing } from '../src/engine/evaluator';
import type { VbaType, VbaDefaultProperty, VbaIterable, VbaComObject } from '../src/engine/vba-types';
import { FileSystem, MemoryFileSystem } from '../src/engine/filesystem';
import { preprocess, stripVBAFileHeader, CompilerConstants } from '../src/engine/preprocessor';
import { loadMocks } from './mock-loader';
import { injectExcelStub } from './excel-stub';
import { MockApplication, MockWorksheet, MockRange, MockRows, MockColumns, MockWorkbook } from '../src/engine/mock/MockExcel';
export { vbaTrue, vbaFalse, vbaNull, vbaEmpty };
export { MockApplication, MockWorksheet, MockRange, MockRows, MockColumns, MockWorkbook };
export type { VbaType, VbaDefaultProperty, VbaIterable, VbaComObject };

const VBA_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

export class VBARunner {
    public evaluator: Evaluator;
    private _asts: Program[] = [];
    private _moduleNames: string[] = [];
    private _resolved = false;
    private readonly _quiet: boolean;

    /** `excelStub: true` のとき注入された MockApplication。セル初期値の設定に使う。 */
    public readonly excelStub: MockApplication | null = null;

    constructor(pathOrDir: string | string[] | null = null, config: {
        sandboxRoot?: string,
        env?: Record<string, string>,
        compilerConstants?: CompilerConstants,
        excelStub?: boolean | MockApplication,
        /** true にすると run() の `[PASS] ...` ログを抑制する */
        quiet?: boolean,
        /** Debug.Print の出力先（既定: console.log） */
        onPrint?: (s: string) => void,
    } = {}) {
        this._quiet = config.quiet ?? false;
        this.evaluator = new Evaluator(config.onPrint ?? console.log, { ...config, fs: new MemoryFileSystem(), allowTopLevelStatements: false });
        if (config.excelStub) {
            const app = config.excelStub === true ? undefined : config.excelStub;
            (this as any).excelStub = injectExcelStub(this.evaluator, app);
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
        const files = stat.isDirectory()
            ? fs.readdirSync(pathOrDir)
                  .filter(f => VBA_EXTENSIONS.has(path.extname(f).toLowerCase()))
                  .sort()
                  .map(f => path.join(pathOrDir, f))
            : [pathOrDir];

        // モックを先にロード（標準ライブラリより後、本番コードより前に注入）
        const mockModules = loadMocks(dir, this.evaluator);
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

    run(procedureName: string, args: any[], type?: 'get' | 'let' | 'set'): any {
        this._ensureResolved();
        const start = Date.now();
        const raw = this.evaluator.callProcedure(procedureName, args, type);
        const result = raw instanceof VbaBoolean ? raw.value !== 0 : raw === vbaNothing ? null : raw;
        const duration = Date.now() - start;
        const formatArgs = args.map(a => typeof a === 'object' ? (a === null ? 'Nothing' : JSON.stringify(a)) : String(a)).join(', ');
        const typeStr = type ? `:${type}` : '';
        if (!this._quiet) {
            console.log(`[PASS] ${procedureName}${typeStr}(${formatArgs}) -> ${result} (${duration}ms)`);
        }
        return result;
    }

    eval(exprString: string): any {
        this._ensureResolved();
        const raw = this.evaluator.evalExpression(exprString);
        return raw instanceof VbaBoolean ? raw.value !== 0 : raw === vbaNothing ? null : raw;
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
}

/**
 * 単一モジュールのインライン VBA コードを評価して Evaluator を返す。
 * Pass 1（登録）と Pass 2（定数確定）を両方実行する。
 * tests/spec/ のテストで evalVBASingle をローカル定義する代わりに使う。
 */
export function evalVBASingle(code: string, options?: EvalOptions): Evaluator {
    const ast = new Parser(new Lexer(code).tokenize(), { sourceLines: code.split('\n') }).parse();
    options?.afterParse?.(ast);
    const ev = new Evaluator(options?.onPrint ?? console.log, {
        fs: options?.fs,
        sandboxRoot: options?.sandboxRoot,
        env: options?.env,
        allowTopLevelStatements: options?.allowTopLevelStatements,
    });
    options?.setup?.(ev);
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
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
    modules: Array<{ name: string; code: string; parseAsClass?: string }>,
    options?: EvalOptions,
): Evaluator {
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
    throws: (fn: () => void, message?: string) => {
        let threw = false;
        try {
            fn();
        } catch {
            threw = true;
        }
        if (!threw) {
            console.error(`[FAIL] ${message || 'Expected an exception but none was thrown'}`);
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
        okAction: () => { ev = evalVBASingle(src); },
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
