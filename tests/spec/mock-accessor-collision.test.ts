/**
 * モック注入オブジェクトの「裏フィールド vs 書き込み可能アクセサ」case 衝突ガード
 *
 * 背景:
 *   Evaluator は VBA の大文字小文字無視仕様に合わせ、注入された JS オブジェクトの
 *   メンバを case-insensitive に解決する（resolveObjectMemberKey）。
 *   このとき own データプロパティ（裏フィールド）が prototype 上のアクセサより
 *   先に拾われるため、「小文字フィールド + 大文字アクセサ」のように case だけ違う
 *   名前で衝突していると、setter を素通りしてフィールドへ直接書き込んでしまう。
 *   （例: かつての MockRange の private `value` フィールドと `Value` setter）
 *
 * このテストは、モッククラスのインスタンス（プロトタイプチェーン含む）を走査し、
 * setter を持つアクセサと case 衝突する own データプロパティが無いことを保証する。
 *
 * ハーム境界について:
 *   - 害が出るのは「書き込み（setter）が裏フィールドに食われる」場合。よって判定対象は
 *     setter を持つアクセサに限定する。
 *   - getter のみ（read-only）の単純パススルー（例: MockWorksheet.Name）はフィールドを
 *     拾っても返る値が同じため無害。よってフラグしない。
 *     ※ ただし将来そのプロパティに setter を足すと、本ガードが衝突を検出するようになる。
 */
import { MockRange, MockWorksheet, MockApplication } from '../../src/engine/mock/MockWorksheet';
import { assert } from '../../test-libs/test-runner';

/**
 * obj のプロトタイプチェーン（Object/Function.prototype は除く）を走査し、
 * データプロパティ名と「setter を持つアクセサ名」を収集する。
 */
function collectProps(obj: any): { data: Set<string>; settable: Set<string> } {
    const data = new Set<string>();
    const settable = new Set<string>();
    let cur = obj;
    while (cur && cur !== Object.prototype && cur !== Function.prototype) {
        for (const name of Object.getOwnPropertyNames(cur)) {
            const desc = Object.getOwnPropertyDescriptor(cur, name)!;
            if (typeof desc.set === 'function') {
                settable.add(name);          // 書き込み可能アクセサ
            } else if (!desc.get) {
                data.add(name);              // 純データ（getter-only は data でも settable でもない）
            }
        }
        cur = Object.getPrototypeOf(cur);
    }
    return { data, settable };
}

function checkInstance(label: string, obj: any): string[] {
    const { data, settable } = collectProps(obj);
    const collisions: string[] = [];
    for (const a of settable) {
        for (const d of data) {
            if (d.toLowerCase() === a.toLowerCase()) {
                collisions.push(`${label}: データフィールド '${d}' が書き込み可能アクセサ '${a}' と case 衝突`);
            }
        }
    }
    return collisions;
}

// 検査対象: クラス直のインスタンスに加え、defineProperty で Value を持つ Range も網羅する
const ws = new MockWorksheet('Sheet1');
const targets: Array<[string, any]> = [
    ['MockRange', new MockRange(0)],
    ['MockRange(single via Range)', ws.Range('A1')],
    ['MockRange(multi via Range)', ws.Range('A1:B2')],
    ['MockRange(union via Range)', ws.Range('A1:A2,C1:C2')],
    ['MockWorksheet', ws],
    ['MockApplication', new MockApplication()],
];

let all: string[] = [];
for (const [label, obj] of targets) {
    all = all.concat(checkInstance(label, obj));
}

if (all.length > 0) {
    console.error('[FAIL] case 衝突を検出:');
    for (const c of all) console.error('  - ' + c);
}
assert.strictEqual(all.length, 0, 'モックに裏フィールド vs 書き込み可能アクセサの case 衝突が存在する');

console.log('[PASS] モック注入オブジェクトに case 衝突なし（検査 ' + targets.length + ' インスタンス）');
console.log('\n✅ mock-accessor-collision: 全テスト通過');
