# 情報活用（2026年度後期）授業サイト

北星学園大学 共通科目「情報活用／情報活用Ⅰ」（グC 木3／社B 金3）の授業資料。

- 公開URL: https://aonoa68.github.io/joho-katsuyo/
- 生成: MkDocs + Material

## 公開のしかた

```bash
./publish.sh          # 整合チェック → main へ push → gh-deploy
./publish.sh --no-push
```

`check_consistency.py` が「ナビ・トップ一覧・実ファイル」の3者を突き合わせ、どこからも辿れないページを検出する。チェックに落ちると公開されない。

## この授業の道具立て

Forms（調査）／CSV（手入力しない）／Python・pandas・matplotlib／GitHub Codespaces（ブラウザ）／成果物はHTMLをGitHub Pagesで公開。**Word・Excel・PowerPointは使わない。ピボットも使わない**（計算過程が残らないため）。
