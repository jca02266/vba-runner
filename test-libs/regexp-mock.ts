/**
 * VBScript.RegExp のモック実装
 *
 * VBA の `CreateObject("VBScript.RegExp")` が返すオブジェクトを、
 * JavaScript の RegExp をベースにエミュレートする。
 *
 * 使い方:
 *   import { createRegExpMock } from '../../test-libs/regexp-mock';
 *
 *   const vbaRunner = new VBARunner('source.vba');
 *   vbaRunner.registerExternalObject('VBScript.RegExp', createRegExpMock);
 *
 * 対応している API:
 *   プロパティ:
 *     - Pattern      (string)
 *     - IgnoreCase   (Boolean: vbaTrue/vbaFalse)
 *     - Global       (Boolean: vbaTrue/vbaFalse)
 *     - MultiLine    (Boolean: vbaTrue/vbaFalse)
 *   メソッド:
 *     - Test(string)         -> Boolean
 *     - Execute(string)      -> MatchCollection
 *     - Replace(s, repl)     -> string
 *   MatchCollection:
 *     - Count, Item(i), For Each 反復
 *   Match:
 *     - Value, FirstIndex (0-based), Length, SubMatches(i) (0-based)
 */
import { vbaTrue, vbaFalse, VbaBoolean } from '../src/engine/evaluator';

function isTruthy(v: any): boolean {
    if (v instanceof VbaBoolean) return v.valueOf() !== 0;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'boolean') return v;
    return Boolean(v);
}

export function createRegExpMock(): any {
    const state: any = {
        __isVbaRegExp__: true,
        // 参照設定相当: VBA で `Dim re As RegExp` / `Set re = New RegExp` と
        // 書く際の class 名。Evaluator.registerExternalObject はこの値も
        // 別名としてファクトリに登録する。
        __className__: 'RegExp',
        pattern: '',
        ignorecase: vbaFalse,
        global: vbaFalse,
        multiline: vbaFalse,

        test(s: any) {
            const re = new RegExp(state.pattern, buildFlags());
            return re.test(String(s ?? '')) ? vbaTrue : vbaFalse;
        },

        execute(s: any) {
            // Execute 内では常に global 検索（VBA Global=False では先頭 1 件のみ返す）
            const flags = buildFlags();
            const reFlags = flags.includes('g') ? flags : 'g' + flags;
            const re = new RegExp(state.pattern, reFlags);
            const items: any[] = [];
            const str = String(s ?? '');
            const isGlobal = isTruthy(state.global);
            let m: RegExpExecArray | null;
            while ((m = re.exec(str)) !== null) {
                items.push(buildMatch(m));
                if (!isGlobal) break;
                if (m[0].length === 0) re.lastIndex++;
            }
            return buildMatchCollection(items);
        },

        replace(s: any, replacement: any) {
            const re = new RegExp(state.pattern, buildFlags());
            return String(s ?? '').replace(re, String(replacement ?? ''));
        },
    };

    function buildFlags(): string {
        let f = '';
        if (isTruthy(state.global)) f += 'g';
        if (isTruthy(state.ignorecase)) f += 'i';
        if (isTruthy(state.multiline)) f += 'm';
        return f;
    }

    function buildMatch(m: RegExpExecArray): any {
        return {
            value: m[0],
            firstindex: m.index,
            length: m[0].length,
            submatches: buildSubMatches(m.slice(1).map(s => s ?? '')),
        };
    }

    function buildSubMatches(values: string[]): any {
        const items = values.slice();
        // SubMatches(i) は VBA のデフォルトプロパティ（Item）アクセス。
        // Evaluator は obj(i) を関数呼び出しとして扱うので、本体は関数にする。
        const fn: any = (i: number) => items[i];
        fn.__isVbaCollection__ = true;
        fn.count = items.length;
        fn.items = items;
        fn.item = (i: number) => items[i];
        fn[Symbol.iterator] = function* () { for (const v of items) yield v; };
        return fn;
    }

    function buildMatchCollection(items: any[]): any {
        return {
            __isVbaCollection__: true,
            count: items.length,
            items,
            item: (i: number) => items[i],
            [Symbol.iterator]: function* () { for (const it of items) yield it; },
        };
    }

    return state;
}
