import * as fs from 'fs';
import * as path from 'path';
import { format, applyEdits, FormatterOptions } from '../src/lsp/formatter';
import { VERSION } from './version';

const VBA_EXTENSIONS = /\.(bas|cls|frm)$/i;

function collectFiles(target: string): string[] {
    const stat = fs.statSync(target);
    if (stat.isFile()) return [target];
    if (stat.isDirectory()) {
        return fs.readdirSync(target)
            .flatMap(entry => {
                const full = path.join(target, entry);
                const s = fs.statSync(full);
                if (s.isDirectory()) return collectFiles(full);
                if (VBA_EXTENSIONS.test(entry)) return [full];
                return [];
            });
    }
    return [];
}

function buildDiff(original: string, formatted: string, filePath: string): string {
    const origLines = original.split('\n');
    const fmtLines = formatted.split('\n');
    const maxLen = Math.max(origLines.length, fmtLines.length);
    const diffLines: string[] = [`--- ${filePath} (original)`, `+++ ${filePath} (formatted)`];
    let hasDiff = false;

    for (let i = 0; i < maxLen; i++) {
        const orig = origLines[i] ?? '';
        const fmtd = fmtLines[i] ?? '';
        if (orig !== fmtd) {
            hasDiff = true;
            diffLines.push(`@@ line ${i + 1} @@`);
            diffLines.push(`-${orig}`);
            diffLines.push(`+${fmtd}`);
        }
    }

    return hasDiff ? diffLines.join('\n') : '';
}

export function main(args: string[]): void {
    if (args.includes('--version') || args.includes('-v')) {
        console.log(VERSION);
        process.exit(0);
    }
    if (args.length === 0 || args[0] === '--help' || args[0] === '-h') {
        console.log(`Usage: vba-runner format <file-or-dir> [--check | --write] [options]

Options:
  --check            差分があれば表示して exit code 1 で終了
  --write            ファイルを上書き（指定しない場合は stdout に出力）
  --indent-size=N    インデント幅（デフォルト: 4）
  --no-keyword-case  キーワードの大文字化を無効化
  --version          バージョンを表示

Example:
  vba-runner format src/vba/Main.bas
  vba-runner format src/vba/ --check
  vba-runner format src/vba/Main.bas --write --indent-size=2`);
        process.exit(0);
    }

    // First non-option argument is the target file/directory
    const target = args.find(a => !a.startsWith('-'));
    const isCheck = args.includes('--check');
    const isWrite = args.includes('--write');
    const noKeywordCase = args.includes('--no-keyword-case');

    const indentSizeArg = args.find(a => a.startsWith('--indent-size='));
    const indentSize = indentSizeArg ? parseInt(indentSizeArg.split('=')[1], 10) : 4;

    if (!target || !fs.existsSync(target)) {
        console.error(`Error: "${target ?? '(指定なし)'}" が見つかりません`);
        process.exit(1);
    }

    const options: FormatterOptions = {
        indentSize,
        keywordCase: noKeywordCase ? false : 'pascal',
    };

    const files = collectFiles(target);
    if (files.length === 0) {
        console.error(`Error: VBAファイル（.bas/.cls/.frm）が見つかりません: ${target}`);
        process.exit(1);
    }

    let hasAnyDiff = false;

    for (const filePath of files) {
        const original = fs.readFileSync(filePath, 'utf8');
        const edits = format(original, options);
        const formatted = applyEdits(original, edits);

        const hasDiff = formatted !== original;

        if (isCheck) {
            if (hasDiff) {
                hasAnyDiff = true;
                const diff = buildDiff(original, formatted, filePath);
                console.log(diff);
            }
        } else if (isWrite) {
            if (hasDiff) {
                fs.writeFileSync(filePath, formatted, 'utf8');
                console.log(`フォーマット済み: ${filePath}`);
            }
        } else {
            // stdout モード: 単一ファイルの場合のみ対応
            if (files.length > 1) {
                console.error('Error: stdout モードは単一ファイルのみ対応しています。ディレクトリには --check または --write を使用してください。');
                process.exit(1);
            }
            process.stdout.write(formatted);
        }
    }

    if (isCheck && hasAnyDiff) {
        process.exit(1);
    }
}

if (process.argv[1]?.includes('vba-formatter')) main(process.argv.slice(2));
