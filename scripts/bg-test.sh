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
#   git stash                          # 現在の作業を退避
#   cd <WORKTREE>                      # ログに表示されたパスへ移動
#   # ... 修正 ...
#   git add . && git commit --amend    # または新規コミット
#   cd <REPO_ROOT>
#   git cherry-pick <BRANCH>           # 修正をメインブランチに適用
#   git worktree remove <WORKTREE> --force
#   git branch -d <BRANCH>
#   git stash pop                      # 作業を復元

set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")/.." rev-parse --show-toplevel)"
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
(cd "${WORKTREE}" && npm test) > "${LOG}" 2>&1
EXIT_CODE=$?
set -e

if [ ${EXIT_CODE} -eq 0 ]; then
    echo "✅ テスト成功 — worktree / ブランチを削除します"
    git -C "${REPO_ROOT}" worktree remove "${WORKTREE}" --force
    git -C "${REPO_ROOT}" branch -d "${BRANCH}"
else
    echo "❌ テスト失敗 (exit: ${EXIT_CODE})"
    echo "   ログ: ${LOG}"
    echo "   worktree: ${WORKTREE}  ブランチ: ${BRANCH}"
    echo ""
    echo "   修正手順:"
    echo "     git stash"
    echo "     cd ${WORKTREE}"
    echo "     # ... 修正 ..."
    echo "     cd ${REPO_ROOT}"
    echo "     git cherry-pick ${BRANCH}"
    echo "     git worktree remove ${WORKTREE} --force"
    echo "     git branch -d ${BRANCH}"
    echo "     git stash pop"
    tail -20 "${LOG}"
    exit ${EXIT_CODE}
fi
