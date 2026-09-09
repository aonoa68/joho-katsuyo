#!/usr/bin/env python3
"""各回のページの「今日のリンク」欄を links.json から生成する。

学生が Moodle を探しにいかなくて済むように、提出先・小テスト・ノートブック・
配布データへのリンクを、各回のページの先頭に置く。

使い方:
    python3 build_links.py          # docs/weekNN.md のマーカー間を書き換える
マーカー:
    <!-- LINKS:START --> ... <!-- LINKS:END -->
"""
from __future__ import annotations
import json, pathlib, re, sys

HERE = pathlib.Path(__file__).parent
DOCS = HERE / "docs"
CFG = json.loads((HERE / "links.json").read_text(encoding="utf-8"))

LABELS = [
    ("submit",   "📝 課題を提出する（Moodle）"),
    ("quiz",     "✅ ふりかえりテスト（Moodle）"),
    ("notebook", "📓 ノートブックを開く（Colab）"),
    ("data",     "📊 配布データ（CSV）"),
    ("ai",       "🤖 今日使うAI"),
]

def block(week: str) -> str:
    w = CFG["weeks"].get(week, {})
    rows = []
    for key, label in LABELS:
        url = (w.get(key) or "").strip()
        if url:
            rows.append(f"- [{label}]({url})")
        elif key in ("submit",):          # 提出先は毎回あるので、無ければ準備中と出す
            rows.append(f"- {label} … 準備中")
    if not rows:
        return ""
    return "!!! info \"今日のリンク\"\n" + "\n".join("    " + r for r in rows)

def main() -> int:
    changed = 0
    for path in sorted(DOCS.glob("week*.md")):
        m = re.search(r"week(\d+)\.md$", path.name)
        if not m:
            continue
        week = str(int(m.group(1)))
        text = path.read_text(encoding="utf-8")
        if "<!-- LINKS:START -->" not in text:
            print(f"  マーカーなし: {path.name}")
            continue
        new = re.sub(
            r"<!-- LINKS:START -->.*?<!-- LINKS:END -->",
            "<!-- LINKS:START -->\n" + block(week) + "\n<!-- LINKS:END -->",
            text, flags=re.S)
        if new != text:
            path.write_text(new, encoding="utf-8")
            changed += 1
    print(f"✅ 「今日のリンク」を更新: {changed} ページ")
    return 0

if __name__ == "__main__":
    sys.exit(main())
