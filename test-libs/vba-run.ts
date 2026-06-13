import * as fs from 'fs';
import * as path from 'path';
import { Lexer } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';
import { Evaluator } from '../src/engine/evaluator';
import { NodeFileSystem } from '../src/engine/node_filesystem';
import { preprocess } from '../src/engine/preprocessor';
import { VERSION } from './version';

const VBA_EXTENSIONS = new Set(['.bas', '.cls', '.frm']);

function parseArg(s: string): any {
    if (s === 'True' || s === 'true')  return true;
    if (s === 'False' || s === 'false') return false;
    if (s === 'Nothing' || s === 'nothing') return null;
    const n = Number(s);
    if (!isNaN(n) && s.trim() !== '') return n;
    return s;
}

function usage(): never {
    process.stderr.write([
        'Usage: vba-runner run <file> <ProcedureName> [args...]',
        '',
        '  file            VBA ソースファイル（同じディレクトリ内の全 .bas/.cls/.frm をロード）',
        '  ProcedureName   呼び出す Sub/Function 名',
        '  args            引数（数値・True/False/Nothing は自動変換、それ以外は文字列）',
        '',
        'Examples:',
        '  vba-runner run sample/src/vba/closure/ClosureTest.bas RunSubtotalTest',
        '  vba-runner run sample/src/refactoring/TaskScheduler_Core.bas InitCalendarConfig',
    ].join('\n') + '\n');
    process.exit(1);
}

export function main(args: string[]): void {
    if (args.includes('--version') || args.includes('-v')) {
        console.log(VERSION);
        process.exit(0);
    }

    const [filePath, procName, ...rawArgs] = args;

    if (!filePath || !procName) usage();

    const absFile = path.resolve(filePath);
    const dir     = path.dirname(absFile);
    const callArgs = rawArgs.map(parseArg);

    // Debug.Print → stderr、戻り値のみ stdout に出す
    const ev = new Evaluator(
        (s: string) => process.stderr.write(s + '\n'),
        { fs: new NodeFileSystem() }
    );

    const files = fs.readdirSync(dir)
        .filter(f => VBA_EXTENSIONS.has(path.extname(f).toLowerCase()))
        .sort()
        .map(f => path.join(dir, f));

    const asts: Array<{ ast: ReturnType<Parser['parse']>; moduleName: string }> = [];

    for (const file of files) {
        const moduleName = path.basename(file, path.extname(file));
        ev.setSourceModule(moduleName);
        const src = preprocess(fs.readFileSync(file, 'utf-8'));
        const ext = path.extname(file).toLowerCase();
        const isRawCls = ext === '.cls'
            && !src.trim().toLowerCase().startsWith('class ')
            && !src.toLowerCase().includes('end class');
        const ast = new Parser(new Lexer(src).tokenize(), isRawCls ? { parseAsClass: moduleName } : {}).parse();
        ev.evaluateModule(ast);
        asts.push({ ast, moduleName });
    }

    ev.resolveIdentifiers(asts);

    const result = ev.callProcedure(procName, callArgs);

    if (result !== undefined && result !== null) {
        process.stdout.write(String(result) + '\n');
    }
}

const _scriptBase = path.basename(process.argv[1] ?? '');
if (_scriptBase === 'vba-run' || _scriptBase === 'vba-run.ts' || _scriptBase === 'vba-run.cjs') {
    main(process.argv.slice(2));
}
