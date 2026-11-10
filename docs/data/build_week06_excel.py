#!/usr/bin/env python3
# coding: utf-8
"""第6回の課題を「Excelで、過程が残る形」でやった場合の模範解答を組み立てる。

    python3 build_week06_excel.py

出力:
    第6回_模範解答_Excel版.xlsx      … 提出できる形（数式が残った状態）
    week06_excel_counts.json         … 手数の集計（サイトの比較ページが使う）

ねらい:
    同じ課題を Excel とノートブックの両方でやって、「過程を残すのに何操作かかるか」
    を実測する。数えた中身は 00_手順ログ シートに全部入れてあるので、
    数え方そのものに文句をつけられる形にしてある。
"""

import json
import os
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "primates_raw.csv")
OUT_XLSX = os.path.join(HERE, "第6回_模範解答_Excel版.xlsx")
OUT_JSON = os.path.join(HERE, "week06_excel_counts.json")

RAW = "01_元データ"
MISS = "02_欠測の数"
CLEAN = "03_整形後"
STAT = "04_基本統計量"
CMP = "05_間違いとの比較"
CHECK = "06_3通り照合"
DIRTY = "07_どこが汚かったか"
COUNTS = "08_手数の集計"
LOG = "00_手順ログ"

# ---------------------------------------------------------------- 手順ログ
# (手順, 何をするか, Excelでの操作, 操作回数, ファイルに残るか, 行が増えたら, 列が増えたら)
# 「操作」= クリック1回・ドラッグ1回・ダイアログの選択1回・数式やことばの入力1回。
# 「行が増えたら」= 回答が1件増えたときに、もう一度やらないといけない操作の回数。
# 「列が増えたら」= 設問を1つ足したときに、もう一度やらないといけない操作の回数。
STEPS = [
    (1, "CSVを文字化けさせずに取り込む",
     "データ → テキストまたはCSVから → ファイルを選ぶ → 開く → 文字コード 65001:UTF-8 を選ぶ → 読み込み",
     6, "接続（クエリ）として残る", 1, 1),
    (2, "シートの名前を「01_元データ」にする",
     "シートタブを右クリック → 名前の変更 → 入力",
     3, "名前だけ残る", 0, 0),
    (3, "整形用のシートを作る（元データを壊さないため）",
     "＋ で新しいシート → 右クリック → 名前の変更 → 入力",
     4, "シートとして残る", 0, 0),
    (4, "見出し行を整形用シートに写す",
     "01_元データ の1行目をコピー → 03_整形後 に貼り付け",
     2, "値として残る（どこから来たかは残らない）", 0, 2),
    (5, "文字の3列（学名・科・属）はそのまま参照する",
     "A2に ='01_元データ'!A2 → C列まで右へフィル",
     2, "数式として残る", 0, 0),
    (6, "🔴 -999 を空欄に直す数式を1つ書く（ここが判断の中身）",
     "D2に =IF('01_元データ'!D2=-999,\"\",'01_元データ'!D2)",
     1, "数式として残る", 0, 0),
    (7, "右に17列ぶんフィルする",
     "D2のフィルハンドルを T2 までドラッグ",
     1, "数式として残る", 0, 2),
    (8, "最終行までフィルする（数式が7千個できる）",
     "A2:T2 を選択 → フィルハンドルを最終行までドラッグ",
     2, "数式として残る", 2, 0),
    (9, "欠測の数を数えるシートを作る",
     "＋ → 右クリック → 名前の変更 → 入力",
     4, "シートとして残る", 0, 0),
    (10, "列名を「横向き」に貼る（※縦向きにすると後で17回手入力になる）",
     "見出し行をコピー → 貼り付け",
     2, "値として残る", 0, 2),
    (11, "-999 が何個あるかを数える",
     "「=COUNTIF('01_元データ'!D:D,-999)」を入力 → 右へフィル",
     2, "数式として残る", 0, 2),
    (12, "整形後に残った有効件数を数える",
     "「=COUNT('03_整形後'!D:D)」を入力 → 右へフィル",
     2, "数式として残る", 0, 2),
    (13, "基本統計量のシートを作る",
     "＋ → 右クリック → 名前の変更 → 入力",
     4, "シートとして残る", 0, 0),
    (14, "列名を横向きに貼る",
     "見出し行をコピー → 貼り付け",
     2, "値として残る", 0, 2),
    (15, "件数・平均・中央値・標準偏差・最小・最大の6行を作る",
     "COUNT / AVERAGE / MEDIAN / STDEV / MIN / MAX を1つずつ書いて、それぞれ右へフィル（6×2）",
     12, "数式として残る", 0, 6),
    (16, "「-999 を入れたまま計算した場合」と比べるシートを作る",
     "＋ → 右クリック → 名前の変更 → 入力",
     4, "シートとして残る", 0, 0),
    (17, "間違った平均・正しい平均・ずれ% の3行を作る",
     "AVERAGE を2つと割り算を1つ書いて、それぞれ右へフィル（3×2）",
     6, "数式として残る", 0, 3),
    (18, "3通り照合のシートを作る",
     "＋ → 右クリック → 名前の変更 → 入力",
     4, "シートとして残る", 0, 0),
    (19, "🔴 ピボットテーブルで、同じ平均をもう1通り出す",
     "挿入 → ピボットテーブル → 範囲を確認 → OK → 行に「科」 → 値に「体重g」 → 値フィールドの設定 → 平均 → OK",
     8, "❌ 表は残るが、作った手順は残らない", 1, 8),
    (20, "AIに出させた数字を貼る",
     "貼り付け",
     1, "値として残る（何と聞いたかは残らない）", 0, 0),
    (21, "3つが一致するかの判定式を作る",
     "「=IF(ABS(…)<0.1,\"一致\",\"ずれ\")」を入力 → 下へフィル",
     2, "数式として残る", 0, 0),
    (22, "🔴 手順ログ（このシート）を手で書く",
     "＋ → 名前の変更（4操作）＋ 1手順ごとに「何をしたか／どこで／なぜ」の3つを入力（手順の数×3）",
     None, "✅ 残る（ただし手で書いたから残っている）", 3, 3),
    (23, "保存して提出する",
     "ファイル → 保存（※PDFにすると数式が消えるので、.xlsx のまま出す）",
     2, "数式が残った状態で残る", 1, 1),
]

# 7の列増が2、4の列増が2 … などは「右フィルの延長／見出しの追記」の分。

# ノートブック側（同じ成果物を出すのに必要な操作）
NB_STEPS = [
    ("ノートブックのリンクを押す", 1, 0, 0),
    ("CSVを左の📁にドラッグする", 1, 0, 0),
    ("raw = pd.read_csv(\"primates_raw.csv\")", 1, 0, 0),
    ("🔴 df = raw.replace(-999, pd.NA)   ← 17列まとめて1行", 1, 0, 0),
    ("(raw == -999).sum()   ← 欠測コードの数を全列ぶん", 1, 0, 0),
    ("df.describe().T   ← 件数・平均・標準偏差・最小・最大を全列ぶん", 1, 0, 0),
    ("df.median(numeric_only=True)   ← 中央値", 1, 0, 0),
    ("raw.mean(numeric_only=True)   ← -999 込みの平均（間違いとの比較）", 1, 0, 0),
    ("df.groupby(\"科\")[\"体重g\"].mean()   ← もう1通りの出し方", 1, 0, 0),
    ("▶ を押す（セル2つ）", 2, 1, 0),
    ("ファイル → ダウンロード → .ipynb をダウンロード", 3, 1, 0),
]
NB_LOG_COST = 0  # 実行すればコードと出力がそのまま残るので、残すための追加作業はない

# Excelの「マクロの記録」を使った場合に、過程を残すためにかかる操作
#   記録開始2（開発タブ→マクロの記録）＋名前をつけてOK 2＋記録終了1＋.xlsm で保存4
#   ※ 開発タブを初めて出すなら、さらに4操作（環境設定→リボンとツールバー→開発→OK）
#   ※ 記録開始を押し忘れると何も残らない。あとから記録することはできない
MACRO_LOG_COST = 9

# 手順ログを手で書く手数は、手順の数から出す（4 = シート追加と名前の変更、3 = 1手順あたりの記入欄）
LOG_COST = 4 + len(STEPS) * 3
STEPS = [(s[0], s[1], s[2], LOG_COST if s[3] is None else s[3], s[4], s[5], s[6]) for s in STEPS]


def check_formulas(wb):
    """テキストのつもりのセルが数式になっていないか検査する。

    openpyxl は "=" で始まる文字列を数式として保存する。日本語の説明文が数式に
    なると、Excel は「エラーが検出されました」と言ってそのセルを削除する
    （LibreOffice は黙って通すので、再計算の検証だけでは見つからない）。
    """
    bad = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type != "f" or not isinstance(c.value, str):
                    continue
                body = re.sub(r"'[^']*'!", "", c.value)       # シート名の参照を外す
                body = re.sub(r'"[^"]*"', "", body)            # 文字列リテラルを外す
                if re.search(r"[ぁ-んァ-ヶ一-龥→]", body):
                    bad.append(f"{ws.title}!{c.coordinate}: {c.value[:60]}")
    if bad:
        raise SystemExit(
            "数式になってはいけないセルが数式になっています（Excelが修復を求めます）:\n  "
            + "\n  ".join(bad))


def main():
    df_raw = pd.read_csv(SRC)
    cols = list(df_raw.columns)
    ncol = len(cols)
    nrow = len(df_raw)
    text_cols = ["学名", "科", "属"]
    num_cols = [c for c in cols if c not in text_cols]
    first_num = get_column_letter(len(text_cols) + 1)          # D
    last_col = get_column_letter(ncol)                          # T
    last_row = nrow + 1                                         # 見出し1行ぶん

    # ---- 答え合わせ用に、pandas でも同じ数字を出しておく
    clean = df_raw.replace(-999, pd.NA)
    truth = {}
    for c in num_cols:
        v = pd.to_numeric(clean[c], errors="coerce").dropna()
        rawv = pd.to_numeric(df_raw[c], errors="coerce").dropna()
        truth[c] = dict(
            n=int(len(v)), mean=float(v.mean()), median=float(v.median()),
            sd=float(v.std()), vmin=float(v.min()), vmax=float(v.max()),
            n_missing=int((df_raw[c] == -999).sum()),
            wrong_mean=float(rawv.mean()),
        )

    wb = Workbook()

    # ============================================================ 共通の見た目
    H = Font(bold=True, color="FFFFFF")
    HFILL = PatternFill("solid", fgColor="2A78D6")
    SUB = Font(bold=True)
    NOTE = Font(italic=True, color="52514E")
    WRAP = Alignment(wrap_text=True, vertical="top")
    THIN = Border(*[Side(style="thin", color="D5D5D0")] * 4)

    def header_row(ws, row, values, width=None):
        for j, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=j, value=v)
            c.font, c.fill, c.alignment, c.border = H, HFILL, WRAP, THIN
        ws.freeze_panes = ws.cell(row=row + 1, column=1)
        if width:
            for j, w in enumerate(width, start=1):
                ws.column_dimensions[get_column_letter(j)].width = w

    # ============================================================ 00_手順ログ
    ws = wb.active
    ws.title = LOG
    ws["A1"] = "第6回の課題を、Excelで「過程が残る形」でやった記録"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        "ノートブックなら、実行すればコードと出力がそのまま残る。"
        "Excelは数式は残るが、「どの順でやったか」「なぜそうしたか」は残らないので、"
        "この表を手で書かないと過程が残らない。その手間も下の「操作回数」に入れて数えてある。"
    )
    ws["A2"].font = NOTE
    ws["A3"] = "※ 数式は、Excelで開いたときに計算されます（このファイルを作った時点では値が入っていません）。"
    ws["A3"].font = NOTE
    header_row(ws, 5,
               ["手順", "何をするか", "Excelでの操作", "操作回数",
                "ファイルに残るか", "回答が1件増えたら、やり直す操作", "設問が1つ増えたら、やり直す操作"],
               [6, 46, 62, 8, 34, 16, 18])
    r = 6
    for no, what, how, n, keep, rowadd, coladd in STEPS:
        ws.cell(row=r, column=1, value=no)
        ws.cell(row=r, column=2, value=what)
        ws.cell(row=r, column=3, value=how)
        ws.cell(row=r, column=4, value=n)
        ws.cell(row=r, column=5, value=keep)
        ws.cell(row=r, column=6, value=rowadd)
        ws.cell(row=r, column=7, value=coladd)
        for j in range(1, 8):
            ws.cell(row=r, column=j).alignment = WRAP
            ws.cell(row=r, column=j).border = THIN
        r += 1
    ws.cell(row=r, column=2, value="合計").font = SUB
    ws.cell(row=r, column=4, value=f"=SUM(D6:D{r-1})").font = SUB
    ws.cell(row=r, column=6, value=f"=SUM(F6:F{r-1})").font = SUB
    ws.cell(row=r, column=7, value=f"=SUM(G6:G{r-1})").font = SUB

    # ============================================================ 01_元データ
    ws = wb.create_sheet(RAW)
    header_row(ws, 1, cols, [26, 20] + [12] * (ncol - 2))
    for i, row in enumerate(df_raw.itertuples(index=False), start=2):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=v)

    # ============================================================ 02_欠測の数
    ws = wb.create_sheet(MISS)
    ws["A1"] = "欠測コード -999 が、どの列にいくつ入っていたか"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "数えた元：01_元データ（取り込んだまま、何も直していない状態）"
    ws["A2"].font = NOTE
    header_row(ws, 4, ["", *cols], [22] + [14] * ncol)
    labels = [
        ("全部の行数", lambda L: f"=COUNTA('{RAW}'!{L}2:{L}{last_row})"),
        ("うち -999 の数", lambda L: f"=COUNTIF('{RAW}'!{L}2:{L}{last_row},-999)"),
        ("-999 を除いた有効件数", lambda L: f"=COUNT('{CLEAN}'!{L}2:{L}{last_row})"),
        ("有効の割合%", None),
    ]
    for k, (lab, f) in enumerate(labels):
        rr = 5 + k
        ws.cell(row=rr, column=1, value=lab).font = SUB
        for j in range(1, ncol + 1):
            L = get_column_letter(j)
            col_letter = get_column_letter(j + 1)
            if lab == "有効の割合%":
                ws.cell(row=rr, column=j + 1,
                        value=f"=ROUND({col_letter}7/{col_letter}5*100,1)")
            else:
                ws.cell(row=rr, column=j + 1, value=f(L))
    ws["A10"] = (
        "ノートブックなら (raw == -999).sum() の1行で、全列ぶんが一度に出る。"
        "Excelでは、数式を1つ書いて右にフィルする（＝列の向きを間違えると、同じことを17回手で書くことになる）。"
    )
    ws["A10"].font = NOTE

    # ============================================================ 03_整形後
    ws = wb.create_sheet(CLEAN)
    header_row(ws, 1, cols, [26, 20] + [12] * (ncol - 2))
    for i in range(2, last_row + 1):
        for j in range(1, ncol + 1):
            L = get_column_letter(j)
            if cols[j - 1] in text_cols:
                ws.cell(row=i, column=j, value=f"='{RAW}'!{L}{i}")
            else:
                ws.cell(row=i, column=j,
                        value=f"=IF('{RAW}'!{L}{i}=-999,\"\",'{RAW}'!{L}{i})")

    # ============================================================ 04_基本統計量
    ws = wb.create_sheet(STAT)
    ws["A1"] = "基本統計量（-999 を空欄に直した 03_整形後 から計算）"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "どの数字も、セルをクリックすれば数式が見える。それが「過程が残っている」ということ。"
    ws["A2"].font = NOTE
    header_row(ws, 4, ["", *cols], [22] + [14] * ncol)
    rows = [
        ("件数", "COUNT"), ("平均", "AVERAGE"), ("中央値", "MEDIAN"),
        ("標準偏差", "STDEV"), ("最小", "MIN"), ("最大", "MAX"),
    ]
    for k, (lab, fn) in enumerate(rows):
        rr = 5 + k
        ws.cell(row=rr, column=1, value=lab).font = SUB
        for j in range(1, ncol + 1):
            L = get_column_letter(j)
            ws.cell(row=rr, column=j + 1,
                    value=f"={fn}('{CLEAN}'!{L}2:{L}{last_row})")
    ws.cell(row=12, column=1, value="使った数式").font = SUB
    ws.cell(row=12, column=2,
            value=f"「=COUNT('{CLEAN}'!{first_num}2:{first_num}{last_row})」など。"
                  "文字の列（学名・科・属）は COUNT が 0 になる＝数えるものがない、で正しい。")
    ws.cell(row=13, column=1, value="注意").font = SUB
    ws.cell(row=13, column=2,
            value="STDEV は標本標準偏差（n-1で割る）。ノートブックの .std() と同じ。"
                  "STDEVP（n で割る）を使うと数字が変わるので、どちらを使ったかを書き残す。")

    # ============================================================ 05_間違いとの比較
    ws = wb.create_sheet(CMP)
    ws["A1"] = "-999 を見落としたまま平均を出すと、どれだけずれるか"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "エラーは出ない。グラフも描ける。だから気づかない。"
    ws["A2"].font = NOTE
    header_row(ws, 4, ["", *cols], [30] + [14] * ncol)
    for k, (lab, src) in enumerate([
        ("① -999 のまま計算した平均", RAW),
        ("② -999 を空欄にした平均", CLEAN),
    ]):
        rr = 5 + k
        ws.cell(row=rr, column=1, value=lab).font = SUB
        for j in range(1, ncol + 1):
            L = get_column_letter(j)
            ws.cell(row=rr, column=j + 1,
                    value=f"=IFERROR(AVERAGE('{src}'!{L}2:{L}{last_row}),\"\")")
    ws.cell(row=7, column=1, value="③ ①は②より何%低いか （①−②）÷②").font = SUB
    ws.cell(row=8, column=1, value="④ ②は①より何%高いか （②−①）÷①").font = SUB
    for j in range(1, ncol + 1):
        L = get_column_letter(j + 1)
        ws.cell(row=7, column=j + 1,
                value=f"=IFERROR(ROUND(({L}5-{L}6)/{L}6*100,1),\"\")")
        ws.cell(row=8, column=j + 1,
                value=f"=IFERROR(ROUND(({L}6-{L}5)/{L}5*100,1),\"\")")
    ws["A10"] = (
        "🔴 ③と④は、同じ1つのずれを、別の分母で言ったもの。体重gの列では ③ −34.5% ／ ④ +52.8%。"
        "資料の「平均が34%ずれる」は③（正しい値を分母にした言い方）、"
        "ノートブックが出す +52.8% は④（間違った値を分母にした言い方）。"
        "どちらでもよいが、\"何を分母にしたか\"を書かないと、読む人には別の話に見える。"
        "第7回のクロス集計で「どちらの方向で割ったか」を必ず書くのは、これと同じ理由。")
    ws["A10"].font = NOTE
    ws["A11"] = ("新しいデータを受け取ったら、計算の前に最小値と最大値を見る（04_基本統計量の下2行）。"
                 "それだけで、このずれは防げる。")
    ws["A11"].font = NOTE

    # ============================================================ 06_3通り照合
    ws = wb.create_sheet(CHECK)
    ws["A1"] = "同じ数字が、3通りの方法で出るか"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "ずれたら、どれかが間違っている。どれが間違っていたかを突き止めるまでが課題。"
    ws["A2"].font = NOTE
    header_row(ws, 4,
               ["出し方", "体重g の平均", "どこに過程が残るか", "やり直すとき"],
               [34, 18, 46, 30])
    wm = truth["体重g"]["mean"]
    rows6 = [
        ("① Excelの関数（04_基本統計量）",
         f"=AVERAGE('{CLEAN}'!{first_num}2:{first_num}{last_row})",
         "✅ セルに数式が残る", "元データを差し替えれば計算し直される"),
        ("② ピボットテーブル",
         "（ここに、ピボットで出した総平均を貼る）",
         "❌ 表は残るが、どう作ったかは残らない", "作り直す（手順を覚えていないと再現できない）"),
        ("③ AIに投げた結果",
         "（ここに、AIが答えた数字を貼る）",
         "△ やり取りを貼れば残る。貼らなければ残らない", "もう一度聞く（同じ答えが返るとは限らない）"),
        ("参考：ノートブックの答え",
         round(wm, 1),
         "✅ コードと出力がそのまま残る", "セルをもう一度実行するだけ"),
    ]
    for k, row in enumerate(rows6):
        rr = 5 + k
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=rr, column=j, value=v)
            c.alignment, c.border = WRAP, THIN
    ws.cell(row=10, column=1, value="①と④は一致したか").font = SUB
    ws.cell(row=10, column=2, value='=IF(ABS(B5-B8)<0.5,"一致","ずれている")')
    ws["A12"] = ("②のピボットは、同じ数字が出る。問題は数字ではなく、"
                 "「半年後に、どうやって出したかを言えるか」。"
                 "この行だけは、ファイルを開いても作り方が分からない。")
    ws["A12"].font = NOTE

    # ============================================================ 07_どこが汚かったか
    ws = wb.create_sheet(DIRTY)
    ws["A1"] = "元データのどこが汚かったか（課題6の4番）"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "この書き残しは、Excelでもノートブックでも同じだけ書く。道具の差はつかない。"
    ws["A2"].font = NOTE
    header_row(ws, 4,
               ["どの列", "何が汚かったか", "どう判断したか", "どうしたか"],
               [18, 40, 40, 40])
    dirty_rows = [
        ("体重g ほか17列",
         f"欠測が -999 という数値で入っていた（体重gだけで{truth['体重g']['n_missing']}件）",
         "最小値を見たら -999 だった。体重が負の霊長類はいない",
         "IF で空欄に直した（03_整形後）。元データは消していない"),
        ("体重g",
         f"最大 {truth['体重g']['vmax']:,.0f}g（平均の約25倍）",
         "ヒガシゴリラ。調べたら実際に149kgある",
         "外れ値だが正しい値なので、残した。除くと別の集団の話になる"),
        ("最長寿命月",
         f"最大 {truth['最長寿命月']['vmax']:,.0f}か月（約122歳）",
         "Homo sapiens の行。自分もこの表の1行として入っている",
         "正しい値なので残した"),
        ("出産間隔日 など",
         f"有効件数が{truth['出産間隔日']['n']}件しかない（全{nrow}種のうち）",
         "COUNT で数えた。平均の分母が列ごとに違う",
         "平均と一緒に件数も必ず書く（04_基本統計量の1行目）"),
        ("（列名）",
         "この配布データは大丈夫だが、フォームの回答CSVは列名に改行が入る",
         "見た目は「睡眠時間」でも、中身は「睡眠時間␊（半角数字で）」",
         "自分のデータでは、列名を全部出して確かめてから列を指定する"),
    ]
    for k, row in enumerate(dirty_rows):
        rr = 5 + k
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=rr, column=j, value=v)
            c.alignment, c.border = WRAP, THIN

    # ============================================================ 08_手数の集計
    excel_calc = sum(s[3] for s in STEPS if s[0] != 22)
    excel_log = next(s[3] for s in STEPS if s[0] == 22)
    excel_row = sum(s[5] for s in STEPS)
    excel_col = sum(s[6] for s in STEPS)
    nb_calc = sum(s[1] for s in NB_STEPS)
    nb_row = sum(s[2] for s in NB_STEPS)
    nb_col = sum(s[3] for s in NB_STEPS)

    ws = wb.create_sheet(COUNTS)
    ws["A1"] = "手数の集計（数え方は 00_手順ログ を見てください）"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("操作＝クリック1回・ドラッグ1回・ダイアログの選択1回・数式やことばの入力1回。"
                "同じ数え方を両方に当てている。")
    ws["A2"].font = NOTE
    header_row(ws, 4, ["数えたもの", "Excel", "ノートブック", "差"], [40, 12, 16, 10])
    table = [
        ("A 計算そのものに必要な操作", excel_calc, nb_calc),
        ("B 過程を残すための追加作業（手順ログを手で書く）", excel_log, NB_LOG_COST),
        ("B' 同上（マクロの記録を使う）", MACRO_LOG_COST, NB_LOG_COST),
        ("A+B 1回やり切るまでの操作", excel_calc + excel_log, nb_calc + NB_LOG_COST),
        ("C 回答が1件増えたときの操作", excel_row, nb_row),
        ("D 設問が1つ増えたときの操作", excel_col, nb_col),
    ]
    for k, (lab, a, b) in enumerate(table):
        rr = 5 + k
        ws.cell(row=rr, column=1, value=lab).font = SUB
        ws.cell(row=rr, column=2, value=a)
        ws.cell(row=rr, column=3, value=b)
        ws.cell(row=rr, column=4, value=f"=B{rr}-C{rr}")
    ws.cell(row=11, column=1, value=(
        "※ B の 73 は「手順ログを手で書く」場合。Excelの「マクロの記録」を使えば 9 操作で済む"
        "（記録開始2＋名前2＋記録終了1＋.xlsm で保存4）。ただし記録開始を押し忘れると何も残らず、"
        "あとから記録することはできない。ノートブックは実行すれば残る。")).alignment = WRAP
    ws.cell(row=13, column=1, value="ノートブック側の内訳").font = SUB
    for k, (what, n, _, _) in enumerate(NB_STEPS):
        ws.cell(row=14 + k, column=1, value=what).alignment = WRAP
        ws.cell(row=14 + k, column=2, value=n)
    ws.cell(row=14 + len(NB_STEPS), column=1, value="合計").font = SUB
    ws.cell(row=14 + len(NB_STEPS), column=2,
            value=f"=SUM(B14:B{13+len(NB_STEPS)})").font = SUB

    check_formulas(wb)
    wb.save(OUT_XLSX)

    counts = dict(
        rows=nrow, cols=ncol, num_cols=len(num_cols),
        formulas_in_clean=nrow * ncol,
        excel=dict(calc=excel_calc, log=excel_log, log_macro=MACRO_LOG_COST,
                   total=excel_calc + excel_log, row_add=excel_row, col_add=excel_col),
        notebook=dict(calc=nb_calc, log=NB_LOG_COST,
                      total=nb_calc + NB_LOG_COST, row_add=nb_row, col_add=nb_col),
        sheets=[LOG, RAW, MISS, CLEAN, STAT, CMP, CHECK, DIRTY, COUNTS],
        truth={k: v for k, v in truth.items() if k in
               ["体重g", "集団サイズ", "最長寿命月", "出産間隔日"]},
    )
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(counts, f, ensure_ascii=False, indent=2)

    print(json.dumps(counts, ensure_ascii=False, indent=2))
    print(f"\n書き出し: {OUT_XLSX}  ({os.path.getsize(OUT_XLSX)/1024:,.0f} KB)")


if __name__ == "__main__":
    main()
