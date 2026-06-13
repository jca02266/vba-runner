import { main as runMain } from './vba-run';
import { main as analyzeMain } from './vba-analyzer';
import { main as formatMain } from './vba-formatter';
import { main as parseCheckMain } from './vba-parse-check';
import { VERSION } from './version';

const [,, subcmd, ...rest] = process.argv;

function showHelp(exitCode: number): never {
    process.stdout.write([
        `vba-runner v${VERSION}`,
        '',
        'Usage: vba-runner <subcommand> [options]',
        '',
        'Subcommands:',
        '  run          VBA プロシージャを実行する',
        '  analyze      リファクタリング候補を静的解析する',
        '  format       VBA ソースを整形する',
        '  parse-check  字句・構文エラーをチェックする',
        '',
        "Run 'vba-runner <subcommand> --help' for subcommand-specific options.",
        '',
        'Options:',
        '  --version    バージョンを表示',
        '  --help       このヘルプを表示',
    ].join('\n') + '\n');
    process.exit(exitCode);
}

if (subcmd === '--version' || subcmd === '-v') {
    console.log(VERSION);
    process.exit(0);
}

if (!subcmd || subcmd === '--help' || subcmd === '-h') {
    showHelp(subcmd ? 0 : 1);
}

switch (subcmd) {
    case 'run':          runMain(rest); break;
    case 'analyze':      analyzeMain(rest); break;
    case 'format':       formatMain(rest); break;
    case 'parse-check':  parseCheckMain(rest); break;
    default:
        process.stderr.write(`vba-runner: unknown subcommand '${subcmd}'\n`);
        showHelp(1);
}
