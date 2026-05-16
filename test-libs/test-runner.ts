import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../src/compiler/lexer';
import { Parser } from '../src/compiler/parser';
import { Evaluator, vbaTrue, vbaFalse, VbaBoolean, vbaNull, vbaEmpty } from '../src/compiler/evaluator';
import { NodeFileSystem } from '../src/compiler/node_filesystem';
import { MemoryFileSystem } from '../src/compiler/filesystem';
export { vbaTrue, vbaFalse, vbaNull, vbaEmpty };

const VBA_EXTENSIONS = new Set(['.vba', '.bas', '.cls', '.frm']);

export class VBATest {
    public evaluator: Evaluator;

    constructor(pathOrDir: string, config: { sandboxRoot?: string, env?: Record<string, string>, useVirtualFS?: boolean } = {}) {
        const useVFS = config.useVirtualFS ?? (typeof process !== 'undefined' && process.env.USE_VFS === '1');
        const fileSystem = useVFS ? new MemoryFileSystem() : new NodeFileSystem();
        this.evaluator = new Evaluator(console.log, { ...config, fs: fileSystem });

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

                // .cls files are always class modules in VBA (file name = class name).
                // Use parseAsClass option instead of string-wrapping the source.
                const ext = path.extname(file).toLowerCase();
                const isRawCls = ext === '.cls'
                    && !source.trim().toLowerCase().startsWith('class ')
                    && !source.toLowerCase().includes('end class');
                const parseOpts = isRawCls ? { parseAsClass: moduleName } : {};

                const ast = new Parser(new Lexer(source).tokenize(), parseOpts).parse();
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
     * VBA の `CreateObject(progId)` が返すオブジェクトを差し替えるためのファクトリを登録する。
     * 既存の組み込みスタブよりも優先される。
     *
     * @example
     *   import { createRegExpMock } from '../../test-libs/regexp-mock';
     *   const vbaTest = new VBATest('source.vba');
     *   vbaTest.registerExternalObject('VBScript.RegExp', createRegExpMock);
     */
    registerExternalObject(progId: string, factory: () => any): void {
        this.evaluator.registerExternalObject(progId, factory);
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

// Keep backward compatibility
export function runVBATest(filePath: string, procedureName: string, args: any[], config: { sandboxRoot?: string, env?: Record<string, string>, useVirtualFS?: boolean } = {}): any {
    const vbaTest = new VBATest(filePath, config);
    return vbaTest.run(procedureName, args);
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
};
