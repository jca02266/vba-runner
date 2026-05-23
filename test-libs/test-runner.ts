import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../src/engine/lexer';
import { Parser, TypeDeclaration, Program } from '../src/engine/parser';
import { Evaluator, SpyRecord, vbaTrue, vbaFalse, VbaBoolean, vbaNull, vbaEmpty } from '../src/engine/evaluator';
import { NodeFileSystem } from '../src/engine/node_filesystem';
import { MemoryFileSystem } from '../src/engine/filesystem';
import { preprocess, CompilerConstants } from '../src/engine/preprocessor';
export { vbaTrue, vbaFalse, vbaNull, vbaEmpty };

const VBA_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

export class VBARunner {
    public evaluator: Evaluator;
    private _asts: Program[] = [];

    constructor(pathOrDir: string | null = null, config: { sandboxRoot?: string, env?: Record<string, string>, useVirtualFS?: boolean, compilerConstants?: CompilerConstants } = {}) {
        const useVFS = config.useVirtualFS ?? (typeof process !== 'undefined' && process.env.USE_VFS === '1');
        const fileSystem = useVFS ? new MemoryFileSystem() : new NodeFileSystem();
        this.evaluator = new Evaluator(console.log, { ...config, fs: fileSystem });

        if (!pathOrDir) return;

        const stat = fs.statSync(pathOrDir);
        const files = stat.isDirectory()
            ? fs.readdirSync(pathOrDir)
                  .filter(f => VBA_EXTENSIONS.has(path.extname(f).toLowerCase()))
                  .sort()
                  .map(f => path.join(pathOrDir, f))
            : [pathOrDir];

        for (const file of files) {
            let source = fs.readFileSync(file, 'utf-8');
            try {
                // Set module name without file extension
                const moduleName = path.basename(file, path.extname(file));
                this.evaluator.setSourceModule(moduleName);

                source = preprocess(source, config.compilerConstants);

                // .cls files are always class modules in VBA (file name = class name).
                // Use parseAsClass option instead of string-wrapping the source.
                const ext = path.extname(file).toLowerCase();
                const isRawCls = ext === '.cls'
                    && !source.trim().toLowerCase().startsWith('class ')
                    && !source.toLowerCase().includes('end class');
                const parseOpts = isRawCls ? { parseAsClass: moduleName } : {};

                const ast = new Parser(new Lexer(source).tokenize(), parseOpts).parse();
                this._asts.push(ast);
                this.evaluator.evaluate(ast);
            } catch (e: any) {
                throw new Error(`[${path.basename(file)}] ${e.message}`);
            }
        }
    }

    run(procedureName: string, args: any[], type?: 'get' | 'let' | 'set'): any {
        const start = Date.now();
        const result = this.evaluator.callProcedure(procedureName, args, type);
        const duration = Date.now() - start;
        const formatArgs = args.map(a => typeof a === 'object' ? (a === null ? 'Nothing' : '[Object]') : String(a)).join(', ');
        const typeStr = type ? `:${type}` : '';
        console.log(`[PASS] ${procedureName}${typeStr}(${formatArgs}) -> ${result} (${duration}ms)`);
        return result;
    }

    eval(exprString: string): any {
        return this.evaluator.evalExpression(exprString);
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
     * VBA の `CreateObject(progId)` が返すオブジェクトを差し替えるためのファクトリを登録する。
     * 既存の組み込みスタブよりも優先される。
     *
     * @example
     *   import { createRegExpMock } from '../../test-libs/regexp-mock';
     *   const vbaRunner = new VBARunner('source.bas');
     *   vbaRunner.registerExternalObject('VBScript.RegExp', createRegExpMock);
     */
    registerExternalObject(progId: string, factory: () => any): void {
        this.evaluator.registerExternalObject(progId, factory);
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
export function runVBARunner(filePath: string, procedureName: string, args: any[], config: { sandboxRoot?: string, env?: Record<string, string>, useVirtualFS?: boolean } = {}): any {
    const vbaRunner = new VBARunner(filePath, config);
    return vbaRunner.run(procedureName, args);
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
    ok: (value: any, message?: string) => {
        if (!value) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected truthy value but got ${String(value)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    isTrue: (actual: VbaBoolean, message?: string) => {
        if (actual !== vbaTrue) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected vbaTrue (${String(vbaTrue)}) but got ${String(actual)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    isFalse: (actual: VbaBoolean, message?: string) => {
        if (actual !== vbaFalse) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected vbaFalse (${String(vbaFalse)}) but got ${String(actual)}`);
            throw new Error(`Assertion Failed`);
        }
    },
    fail: (message?: string) => {
        console.error(`[FAIL] ${message || 'Assertion failed'}`);
        throw new Error(`Assertion Failed`);
    },
};
