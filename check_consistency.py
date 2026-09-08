#!/usr/bin/env python3
"""公開前チェック：ナビ・トップ一覧・実ファイルの整合を確認する。

3者を突き合わせて「どこかから辿れないページ」を洗い出す。

- ナビ      : mkdocs.yml の nav:
- トップ一覧: docs/index.md 内のリンク
- 実ファイル: docs/ 配下の *.md

使い方:
    python3 check_consistency.py
    # 問題があれば終了コード 1（公開前のガードに使える）
"""

from __future__ import annotations

import os
import re
import sys

import yaml  # mkdocs の依存なので入っている

BASE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(BASE, "docs")
MKDOCS_YML = os.path.join(BASE, "mkdocs.yml")
INDEX_MD = os.path.join(DOCS, "index.md")

# index.md は意図的にリンクしない補助ページ（一覧に出さなくてよいもの）
INDEX_EXEMPT = {"index.md"}
# ナビに出さなくてよいページ
NAV_EXEMPT = {"index.md"}


def collect_nav_targets(nav) -> set[str]:
    """nav 構造（list/dict のネスト）から .md / http のターゲットを集める。"""
    found: set[str] = set()

    def walk(node):
        if isinstance(node, list):
            for item in node:
                walk(item)
        elif isinstance(node, dict):
            for value in node.values():
                walk(value)
        elif isinstance(node, str):
            found.add(node)

    walk(nav)
    return found


def load_yaml_lenient(path: str):
    """mkdocs.yml の独自タグ（!!python/... 等）があっても nav だけ読めるようにする。"""
    class SafeLoaderIgnore(yaml.SafeLoader):
        pass

    def ignore_unknown(loader, suffix, node):
        return None

    SafeLoaderIgnore.add_multi_constructor("tag:yaml.org,2002:python/", ignore_unknown)
    SafeLoaderIgnore.add_multi_constructor("!", ignore_unknown)
    with open(path, encoding="utf-8") as f:
        return yaml.load(f, Loader=SafeLoaderIgnore)


def main() -> int:
    config = load_yaml_lenient(MKDOCS_YML)
    nav_targets = collect_nav_targets(config.get("nav", []))
    nav_md = {t for t in nav_targets if t.endswith(".md")}
    nav_ext = {t for t in nav_targets if t.startswith("http")}

    # index.md 内のリンク
    with open(INDEX_MD, encoding="utf-8") as f:
        index_text = f.read()
    index_links = set(re.findall(r"\]\(([^)]+)\)", index_text))
    index_md = {l for l in index_links if l.endswith(".md")}
    index_ext = {l for l in index_links if l.startswith("http")}

    # 実ファイル
    actual_md = {
        f for f in os.listdir(DOCS)
        if f.endswith(".md") and os.path.isfile(os.path.join(DOCS, f))
    }

    problems: list[str] = []

    # ① 実ファイルなのにナビに無い（左メニューから辿れない）
    missing_in_nav = actual_md - nav_md - NAV_EXEMPT
    for f in sorted(missing_in_nav):
        problems.append(f"❌ ナビに無い（左メニューから辿れない）: {f}")

    # ② 実ファイルなのにトップ一覧に無い（index.md から辿れない）
    missing_in_index = actual_md - index_md - INDEX_EXEMPT
    for f in sorted(missing_in_index):
        problems.append(f"⚠️  トップ一覧に無い（index.md にリンクなし）: {f}")

    # ③ ナビにあるのに実ファイルが無い（リンク切れ）
    nav_dangling = {t for t in nav_md if t not in actual_md}
    for f in sorted(nav_dangling):
        problems.append(f"❌ ナビのリンク切れ（ファイル不在）: {f}")

    # ④ index.md にあるのに実ファイルが無い（リンク切れ／files等は除外）
    for l in sorted(index_md):
        if l not in actual_md and not l.startswith("files/"):
            problems.append(f"❌ トップ一覧のリンク切れ（ファイル不在）: {l}")

    # 外部リンク（NotebookLM・Colab 等）は「ナビかトップ一覧のどちらか一方」で
    # よいため、両方に在るかは検査しない（件数だけサマリに出す）。

    # 出力
    print("=" * 60)
    print("  公開前チェック：ナビ / トップ一覧 / 実ファイルの整合")
    print("=" * 60)
    print(f"  実ファイル: {len(actual_md)} ページ")
    print(f"  ナビ掲載  : {len(nav_md)} ページ ＋ 外部 {len(nav_ext)} 件")
    print(f"  トップ掲載: {len(index_md)} ページ ＋ 外部 {len(index_ext)} 件")
    print("-" * 60)

    blocking = [p for p in problems if p.startswith("❌")]
    warnings = [p for p in problems if p.startswith(("⚠️", "ℹ️"))]

    if not problems:
        print("  ✅ 問題なし。すべて整合しています。")
        print("=" * 60)
        return 0

    for p in blocking + warnings:
        print("  " + p)
    print("-" * 60)
    print(f"  重大(❌): {len(blocking)} 件 ／ 注意(⚠️ℹ️): {len(warnings)} 件")
    print("=" * 60)

    # ❌ があれば終了コード1（公開を止めるガードに使える）
    return 1 if blocking else 0


if __name__ == "__main__":
    sys.exit(main())
