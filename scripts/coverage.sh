#!/bin/bash
# テストスイート全体を V8 カバレッジ付きで実行し、src/engine の正確な行カバレッジを出す。
#
# 使い方:
#   ./scripts/coverage.sh           # 全テスト実行 + 変換 + ユニオン集計（全体で 30 分〜1 時間）
#   ./scripts/coverage.sh --report  # 既存の coverage-v8/ から集計だけやり直す
#
# ⚠️ c8 の複数プロセスマージ（--merge-async 含む）は tsx 子プロセス構成で大幅に
# 過少報告する（例: evaluator.ts 実際 92% → c8 マージ 51%）。そのため本スクリプトは
# プロセス（V8 カバレッジ JSON）ごとに個別に istanbul 形式へ変換し、行カバレッジの
# ユニオンを自前で取る。変換は npx c8 の起動コストが支配的（約 3 秒 × プロセス数）。
set -e
cd "$(dirname "$0")/.."

COV_DIR="coverage-v8"
CHUNK_DIR="coverage-chunks"
INCLUDES=(--include 'src/engine/evaluator.ts' --include 'src/engine/builtins.ts'
          --include 'src/engine/coerce.ts' --include 'src/engine/parser.ts'
          --include 'src/engine/lexer.ts' --include 'src/engine/vba-types.ts')

if [ "$1" != "--report" ]; then
    rm -rf "$COV_DIR"
    echo "⚡ カバレッジ計測付きで全テスト実行中..."
    NODE_V8_COVERAGE="$COV_DIR" ./run_all_tests.sh
fi

echo "🔄 プロセスごとに istanbul 形式へ変換中（$(ls "$COV_DIR"/coverage-*.json | wc -l | tr -d ' ') プロセス）..."
rm -rf "$CHUNK_DIR"
mkdir -p "$CHUNK_DIR"
i=0
for f in "$COV_DIR"/coverage-*.json; do
    i=$((i+1))
    d="$CHUNK_DIR/tmp$i"
    mkdir -p "$d"
    cp "$f" "$d/"
    npx c8 report --temp-directory "$d" "${INCLUDES[@]}" \
        --reporter=json --report-dir "$CHUNK_DIR/r$i" > /dev/null 2>&1 || true
    rm -rf "$d"
    if [ $((i % 100)) -eq 0 ]; then echo "  ... $i 件変換済み"; fi
done

echo ""
echo "--- 行カバレッジ（全プロセスのユニオン） ---"
node --input-type=module - "$CHUNK_DIR" <<'EOF'
import * as fs from 'fs';
import * as path from 'path';

const chunkDir = process.argv[2];
const covered = new Map();   // path -> Set(実行された行)
const seen = new Map();      // path -> Set(計測対象の行)

for (const r of fs.readdirSync(chunkDir)) {
    const f = path.join(chunkDir, r, 'coverage-final.json');
    if (!fs.existsSync(f)) continue;
    let data;
    try { data = JSON.parse(fs.readFileSync(f, 'utf8')); } catch { continue; }
    for (const [p, d] of Object.entries(data)) {
        if (!covered.has(p)) { covered.set(p, new Set()); seen.set(p, new Set()); }
        const cov = covered.get(p), ins = seen.get(p);
        for (const [sid, loc] of Object.entries(d.statementMap)) {
            for (let ln = loc.start.line; ln <= loc.end.line; ln++) {
                ins.add(ln);
                if ((d.s[sid] ?? 0) > 0) cov.add(ln);
            }
        }
    }
}

for (const p of [...seen.keys()].sort()) {
    const ins = seen.get(p), cov = covered.get(p);
    const unc = [...ins].filter(l => !cov.has(l)).sort((a, b) => a - b);
    const pct = (100 * cov.size / ins.size).toFixed(1);
    // 未到達行を連続範囲に圧縮して表示
    const ranges = [];
    for (const ln of unc) {
        const last = ranges[ranges.length - 1];
        if (last && ln <= last[1] + 1) last[1] = ln;
        else ranges.push([ln, ln]);
    }
    const shortName = p.split('/src/')[1] ?? p;
    console.log(`${shortName}: ${pct}%  未到達 ${unc.length} 行`);
    const big = ranges.filter(([s, e]) => e - s >= 4).slice(0, 20);
    for (const [s, e] of big) console.log(`    ${s}-${e}`);
}
EOF
