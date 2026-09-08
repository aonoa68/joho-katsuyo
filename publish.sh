#!/usr/bin/env bash
# 公開前チェック → 合格時のみ公開（gh-deploy）するラッパー。
#
# 使い方:
#   ./publish.sh            # チェック→（任意でpush）→gh-deploy
#   ./publish.sh --no-push  # main への git push をしない
#
# 「git push しただけでは公開されない」「ナビ/トップ一覧に入れ忘れる」
# という2つの定番ミスを、まとめて防ぐためのもの。

set -euo pipefail
cd "$(dirname "$0")"

echo "▶ 1/3 整合チェック（ナビ・トップ一覧・実ファイル）"
if ! python3 check_consistency.py; then
  echo ""
  echo "🛑 整合チェックで重大エラー（❌）。公開を中止しました。"
  echo "   上の指摘を直してから、もう一度実行してください。"
  exit 1
fi

echo ""
echo "▶ 2/3 ソースを GitHub(main) にバックアップ push"
if [[ "${1:-}" == "--no-push" ]]; then
  echo "   （--no-push 指定のためスキップ）"
elif [[ -n "$(git status --porcelain)" ]]; then
  echo "   未コミットの変更があります。コミットしてから実行してください："
  git status --short
  exit 1
else
  git push origin main
fi

echo ""
echo "▶ 3/3 公開（mkdocs gh-deploy → gh-pages へ反映）"
mkdocs gh-deploy --message "publish.sh による公開"

echo ""
echo "✅ 公開完了：https://aonoa68.github.io/joho-katsuyo/"
echo "   （CDN反映に1〜2分かかる場合があります。⌘+Shift+R で再読み込み）"
