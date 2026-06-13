import { Lexer, LexError } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';
import { stripClsFileHeader } from '../src/engine/preprocessor';
import * as fs from 'fs';
import * as path from 'path';
import { VERSION } from './version';

interface Diagnostic {
    file: string;
    line: number;
    column: number;
    severity: 'error' | 'warning';
    source: 'lexer' | 'parser';
    message: string;
}

function checkFile(filePath: string): Diagnostic[] {
    const ext = path.extname(filePath).toLowerCase();
    const raw = fs.readFileSync(filePath, 'utf8');
    const src = ext === '.cls' ? stripClsFileHeader(raw) : raw;
    const diags: Diagnostic[] = [];

    const lexer = new Lexer(src);
    let tokens;
    try {
        tokens = lexer.tokenize();
    } catch (e) {
        if (e instanceof LexError) {
            return [{
                file: filePath,
                line: e.line,
                column: e.column,
                severity: 'error',
                source: 'lexer',
                message: e.message.replace(/ \(line \d+\)$/, ''),
            }];
        }
        throw e;
    }

    for (const d of lexer.diagnostics) {
        diags.push({
            file: filePath,
            line: d.line,
            column: d.column,
            severity: 'error',
            source: 'lexer',
            message: d.message,
        });
    }

    const ast = new Parser(tokens, { errorRecovery: true }).parse();
    for (const d of ast.diagnostics) {
        diags.push({
            file: filePath,
            line: d.loc.start.line,
            column: d.loc.start.column,
            severity: d.severity === 'error' ? 'error' : 'warning',
            source: 'parser',
            message: d.message,
        });
    }

    return diags;
}

function collectFiles(target: string): string[] {
    const stat = fs.statSync(target);
    if (stat.isFile()) return [target];
    return fs.readdirSync(target)
        .filter(f => /\.(bas|cls|frm)$/i.test(f))
        .map(f => path.join(target, f));
}

export function main(args: string[]): void {
    if (args.includes('--version') || args.includes('-v')) {
        console.log(VERSION);
        process.exit(0);
    }

    const jsonMode = args.includes('--json');
    const targets = args.filter(a => !a.startsWith('--'));

    if (targets.length === 0) {
        process.stderr.write('Usage: vba-runner parse-check <file|dir> [--json] [--version]\n');
        process.exit(1);
    }

    const allDiags: Diagnostic[] = [];
    for (const target of targets) {
        for (const file of collectFiles(target)) {
            allDiags.push(...checkFile(file));
        }
    }

    if (jsonMode) {
        process.stdout.write(JSON.stringify(allDiags, null, 2) + '\n');
    } else {
        if (allDiags.length === 0) {
            process.stdout.write('No errors found.\n');
        } else {
            for (const d of allDiags) {
                process.stdout.write(
                    `${d.file}:${d.line}:${d.column}: ${d.severity} [${d.source}] ${d.message}\n`
                );
            }
            process.exit(1);
        }
    }
}

if (process.argv[1]?.includes('vba-parse-check')) main(process.argv.slice(2));
