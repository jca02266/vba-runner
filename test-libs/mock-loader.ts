/**
 * __mocks__ ディレクトリ / ファイルからモックをロードして Evaluator に注入する。
 *
 * ## ロード順と優先順位
 *
 * 収集対象:
 *   - `__mocks__/` ディレクトリ内のファイル（ソートキー = ファイル名のみ、例: `ExcelObjects.bas`）
 *   - `__mocks__.<ext>` 単一ファイル（ソートキー = ファイル名、例: `__mocks__.bas`）
 *
 * ソートアルゴリズム:
 *   ソートキー（basename）を ASCII コードで辞書順比較（ロケール非依存）。
 *   小文字・大文字・記号の ASCII 値は以下の通り:
 *     A–Z = 65–90、_ = 95、a–z = 97–122
 *   ファイル名の先頭文字が A–Z（大文字）のファイルは `_` より前に来るため、
 *   `__mocks__.*` 系は通常のファイル名より後にソートされる。
 *
 * 優先順位ルール:
 *   **後にロードされたファイルが先にロードされたファイルを上書きする（後勝ち）。**
 *   ソートが早い（先頭に近い）ファイルほど優先度が低く、
 *   ソートが遅い（末尾に近い）ファイルほど優先度が高い。
 *
 * 典型的なソート順と優先度（低 → 高）:
 *   1. `ExcelObjects.bas`  (E=69、最も早い → 最低優先)
 *   2. `MsgBox.js`         (M=77)
 *   3. `RegExp.ts`         (R=82)
 *   4. `__mocks__.bas`     (_=95、b=98)
 *   5. `__mocks__.js`      (_=95、j=106)
 *   6. `__mocks__.ts`      (_=95、t=116 → 最高優先)
 *
 * 同一関数名やクラス名が複数ファイルに存在する場合、最終的に上位のファイルの定義が使われる。
 *
 * ## 各形式の処理
 *
 * `.js` / `.ts`: createRequire で読み込み、エクスポートを evaluator に適用:
 *   - 通常の関数エクスポート → `setBuiltinOverride` (env を上書き)
 *   - `__addCreateObject__` キー → `registerExternalObject` (CreateObject ファクトリ登録)
 *
 * `.bas` / `.cls`: VBA として評価:
 *   - トップレベル関数 → 評価後に env へ自動登録（後から定義したもので上書き）
 *   - Class 定義 → `promoteMockVbaClasses` で `externalObjectFactories` へ昇格
 *     （`instantiateClass` では classDefinitions より externalObjectFactories が優先）
 */
import * as fs from 'fs';
import * as path from 'path';
import { createRequire } from 'node:module';
import { Lexer } from '../src/engine/lexer';
import { Parser, Program } from '../src/engine/parser';
import { Evaluator } from '../src/engine/evaluator';
import { preprocess } from '../src/engine/preprocessor';

// ESM 環境でも require() を使えるようにする
const _require = createRequire(import.meta.url);

export interface MockModule {
    ast: Program;
    moduleName: string;
}

/** 収集したモックファイルの内部表現 */
interface MockEntry {
    /** ファイルの絶対パス */
    fullPath: string;
    /** ソートキー（basename）。ASCII 辞書順で比較する */
    sortKey: string;
    /** 小文字の拡張子 */
    ext: string;
}

/**
 * dir 配下の __mocks__/ と __mocks__.* を検出してロードする。
 *
 * VBA モックのみ AST を返す（resolveIdentifiers に含めるため）。
 * JS/TS モックは require() でロードして evaluator に直接適用する。
 *
 * ファイルは basename の ASCII 辞書順にソートされ、後にロードされたものが勝つ。
 */
export function loadMocks(dir: string, evaluator: Evaluator): MockModule[] {
    const entries: MockEntry[] = [];

    // --- 収集: __mocks__/ ディレクトリ内のファイル ---
    const mocksDir = path.join(dir, '__mocks__');
    if (fs.existsSync(mocksDir) && fs.statSync(mocksDir).isDirectory()) {
        for (const entry of fs.readdirSync(mocksDir)) {
            const fullPath = path.join(mocksDir, entry);
            if (!fs.statSync(fullPath).isFile()) continue;
            const ext = path.extname(entry).toLowerCase();
            if (ext === '.ts' || ext === '.js' || ext === '.bas' || ext === '.cls') {
                entries.push({ fullPath, sortKey: entry, ext });
            }
        }
    }

    // --- 収集: __mocks__.<ext> 単一ファイル ---
    for (const ext of ['.ts', '.js', '.bas']) {
        const candidate = path.join(dir, `__mocks__${ext}`);
        if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
            entries.push({ fullPath: candidate, sortKey: `__mocks__${ext}`, ext });
        }
    }

    // --- ソート: basename を ASCII コードで辞書順比較（ロケール非依存） ---
    entries.sort((a, b) => {
        if (a.sortKey < b.sortKey) return -1;
        if (a.sortKey > b.sortKey) return 1;
        return 0;
    });

    // --- ロード: ソート順に処理（後にロードしたものが勝つ） ---
    const vbaMockModules: MockModule[] = [];

    for (const entry of entries) {
        if (entry.ext === '.js' || entry.ext === '.ts') {
            loadJsMock(entry.fullPath, evaluator);
        } else {
            const mod = loadVbaMock(entry.fullPath, evaluator);
            if (mod) vbaMockModules.push(mod);
        }
    }

    return vbaMockModules;
}

/** JS / TS モックを require() してエクスポートを evaluator に適用する */
function loadJsMock(file: string, evaluator: Evaluator): void {
    try {
        const mod = _require(path.resolve(file));
        applyJsMockExports(mod, evaluator, file);
    } catch (e: any) {
        console.warn(`[mock-loader] Failed to load JS/TS mock "${file}": ${e.message}`);
    }
}

/** VBA モック (.bas / .cls) を評価し、クラスを externalObjectFactories へ昇格する */
function loadVbaMock(file: string, evaluator: Evaluator): MockModule | null {
    try {
        const source = fs.readFileSync(file, 'utf-8');
        const ext = path.extname(file).toLowerCase();
        const moduleName = path.basename(file, ext);

        const processed = preprocess(source);

        const isRawCls = ext === '.cls'
            && !processed.trim().toLowerCase().startsWith('class ')
            && !processed.toLowerCase().includes('end class');
        const parseOpts = isRawCls ? { parseAsClass: moduleName } : {};

        const classesBefore = evaluator.getRegisteredClassNames();
        evaluator.setSourceModule(moduleName);
        const ast = new Parser(new Lexer(processed).tokenize(), parseOpts).parse();
        evaluator.evaluateModule(ast);
        evaluator.promoteMockVbaClasses(classesBefore);

        return { ast, moduleName };
    } catch (e: any) {
        console.warn(`[mock-loader] Failed to load VBA mock "${file}": ${e.message}`);
        return null;
    }
}

/**
 * JS/TS モジュールのエクスポートを evaluator に適用する。
 *
 * - TS default export か CJS module.exports を正規化して走査する
 * - `__addCreateObject__` キー → registerComObject に登録
 *   factory が __progId__ を持たない場合はキー名を __progId__ として補完する
 * - それ以外の関数 → setBuiltinOverride で env に登録
 */
function applyJsMockExports(mod: any, evaluator: Evaluator, file: string): void {
    const exports = mod?.default ?? mod;
    if (!exports || typeof exports !== 'object') {
        console.warn(`[mock-loader] "${file}" exports nothing usable (expected object)`);
        return;
    }

    for (const [key, value] of Object.entries(exports)) {
        if (key === '__addCreateObject__') {
            if (typeof value === 'object' && value !== null) {
                for (const [progId, factory] of Object.entries(value as Record<string, () => any>)) {
                    // factory が __progId__ を持たない場合はキーで補完して VbaComObject 型に合わせる
                    const wrappedFactory = () => {
                        const obj = (factory as () => any)();
                        if (obj && typeof obj === 'object' && !('__progId__' in obj)) {
                            obj.__progId__ = progId;
                        }
                        return obj;
                    };
                    evaluator.registerComObject(wrappedFactory);
                }
            }
        } else if (typeof value === 'function') {
            evaluator.setBuiltinOverride(key, value);
        }
    }
}
