#!/bin/bash
# bg-test.sh — 現在の HEAD からテストブランチ＋worktree を作成してテスト実行
#
# 使い方（ターミナルからバックグラウンド実行）:
#   ./scripts/bg-test.sh &    # バックグラウンドで実行、PIDを記録しておく
#
# Claude Code から使う場合:
#   Bash ツールで run_in_background: true を指定して呼ぶ。
#   スクリプト自体は同期実行（exit code がテスト結果を表す）。
#
# テスト失敗時の修正手順:
#   git stash                                   # 現在の作業を退避
#   git checkout -b fix-branch <COMMIT>         # 失敗したコミットからブランチ作成
#   # ... 修正 ...
#   git add . && git commit                     # 修正をコミット
#   git checkout main
#   git rebase fix-branch                       # main を fix-branch の先端に積み直す
#   git branch -d fix-branch
#   git stash pop                               # 作業を復元

set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")/.." rev-parse --show-toplevel)"

# ワークツリーが clean であることを確認（コミット漏れ・未追跡ファイルを防ぐ）
DIRTY=$(git -C "${REPO_ROOT}" status --porcelain)
if [ -n "${DIRTY}" ]; then
    echo "❌ ワークツリーが clean ではありません。コミットまたは stash してから実行してください。"
    echo ""
    git -C "${REPO_ROOT}" status --short
    exit 1
fi

TIMESTAMP=$(date +%Y%m%d-%H%M%S)
BRANCH="test/${TIMESTAMP}"
# /tmp (=/private/tmp) は Claude Code の作業領域と競合して溢れることがある。
# ホームディレクトリ配下の専用ディレクトリを使う。
WORKTREE_BASE="${HOME}/tmp/vba-runner-tests"
mkdir -p "${WORKTREE_BASE}"
WORKTREE="${WORKTREE_BASE}/test-${TIMESTAMP}"
LOG="${WORKTREE}/test-output.log"

echo "📋 テストブランチ作成: ${BRANCH}"
git -C "${REPO_ROOT}" branch "${BRANCH}"

echo "🗂  worktree 追加: ${WORKTREE}"
git -C "${REPO_ROOT}" worktree add "${WORKTREE}" "${BRANCH}"

# node_modules はインストール済みのものをシンボリックリンクで共有
echo "🔗 node_modules をリンク..."
ln -sf "${REPO_ROOT}/node_modules" "${WORKTREE}/node_modules"

echo "⚡ テスト実行中... (ログ: ${LOG})"

set +e
(cd "${WORKTREE}" && npm run typecheck && npm test) > "${LOG}" 2>&1
EXIT_CODE=$?
set -e

# テスト結果サマリーをログから抜粋して表示
echo "--- テスト結果 ---"
grep -E "passed|failed|PASS|FAIL|Tests passed|All [0-9]+|error TS" "${LOG}" | tail -10 || true
echo "------------------"

if [ ${EXIT_CODE} -eq 0 ]; then
    echo "✅ テスト成功 — worktree / ブランチを削除します"
    git -C "${REPO_ROOT}" worktree remove "${WORKTREE}" --force
    git -C "${REPO_ROOT}" branch -d "${BRANCH}"
else
    COMMIT=$(git -C "${WORKTREE}" rev-parse HEAD)
    git -C "${REPO_ROOT}" worktree remove "${WORKTREE}" --force
    git -C "${REPO_ROOT}" branch -D "${BRANCH}"
    echo "❌ テスト失敗 (exit: ${EXIT_CODE})"
    echo "   ログ: ${LOG}"
    echo "   コミット: ${COMMIT}"
    echo ""
    echo "   修正手順:"
    echo "     git stash"
    echo "     git checkout -b fix-branch ${COMMIT}"
    echo "     # ... 修正 ..."
    echo "     git add . && git commit"
    echo "     git checkout main"
    echo "     git rebase fix-branch"
    echo "     git branch -d fix-branch"
    echo "     git stash pop"
    exit ${EXIT_CODE}
fi
