#!/bin/bash
# bg-test.sh — 現在の HEAD からテストブランチ＋worktree を作成してバックグラウンドでテスト実行
#
# 使い方:
#   ./scripts/bg-test.sh          # テスト実行（完了後にブランチ・worktree を自動削除）
#   ./scripts/bg-test.sh --keep   # テスト失敗時もworktreeを保持（デフォルト動作と同じ）
#
# テスト失敗時:
#   1. worktree と test ブランチは保持される
#   2. カレントブランチで作業中の変更を stash してからworktreeへ移動して修正
#      git stash && cd <worktree_path> && ...修正... && cd - && git stash pop

set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")/.." rev-parse --show-toplevel)"
TIMESTAMP=$(date +%Y%m%d-%H%M%S)
BRANCH="test/${TIMESTAMP}"
WORKTREE="/tmp/vba-runner-test-${TIMESTAMP}"
LOG="${WORKTREE}/test-output.log"

echo "📋 テストブランチ作成: ${BRANCH}"
git -C "${REPO_ROOT}" branch "${BRANCH}"

echo "🗂  worktree 追加: ${WORKTREE}"
git -C "${REPO_ROOT}" worktree add "${WORKTREE}" "${BRANCH}"

echo "⚡ バックグラウンドでテスト開始..."
echo "   ログ: ${LOG}"
echo "   進捗確認: tail -f ${LOG}"

(
  cd "${WORKTREE}"
  npm test > "${LOG}" 2>&1
  EXIT_CODE=$?

  if [ ${EXIT_CODE} -eq 0 ]; then
    echo "" >> "${LOG}"
    echo "✅ TESTS PASSED — worktree / ブランチを削除します" >> "${LOG}"
    echo "✅ [bg-test] テスト成功 (${BRANCH})"
    git -C "${REPO_ROOT}" worktree remove "${WORKTREE}" --force
    git -C "${REPO_ROOT}" branch -d "${BRANCH}"
  else
    echo "" >> "${LOG}"
    echo "❌ TESTS FAILED — worktree を保持: ${WORKTREE}" >> "${LOG}"
    echo "❌ [bg-test] テスト失敗 (${BRANCH})"
    echo "   修正手順:"
    echo "     git stash                  # 現在の作業を退避"
    echo "     cd ${WORKTREE}             # テストブランチへ移動"
    echo "     # ... 修正 ..."
    echo "     git add . && git commit    # 修正をコミット"
    echo "     cd ${REPO_ROOT}"
    echo "     git cherry-pick ${BRANCH}  # 修正をメインブランチに適用"
    echo "     git worktree remove ${WORKTREE} --force"
    echo "     git branch -d ${BRANCH}"
    echo "     git stash pop              # 作業を復元"
    echo "   ログ全文: cat ${LOG}"
    exit ${EXIT_CODE}
  fi
) &

BG_PID=$!
echo "🔄 バックグラウンドPID: ${BG_PID}"
echo "   完了まで作業を続けられます"
