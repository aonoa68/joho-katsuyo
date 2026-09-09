#!/usr/bin/env python3
"""霊長類376種の生活史・生態データを、授業用CSVに整形する。

出典（必ず教材に明記すること）:
  Jones, K.E. et al. (2009) PanTHERIA: a species-level database of life history,
  ecology, and geography of extant and recently extinct mammals.
  Ecology 90(9): 2648. Ecological Archives E090-184.
  https://esapubs.org/archive/ecol/E090/184/

書き出すファイル:
  primates_raw.csv  … 欠測が -999 のまま。第6回「データの整形」で使う
  primates.csv      … -999 を空欄に直したもの。それ以外の回で使う

※ 値そのものは加工していない。霊長類の行だけを抜き、列を選び、名前を日本語にし、
   気温の単位を 0.1℃ → ℃ に直しただけである。
   体重はg、寿命は月、日齢は日のまま（元データの単位）。
"""
from __future__ import annotations
import sys, pathlib
import numpy as np
import pandas as pd

SRC = "https://esapubs.org/archive/ecol/E090/184/PanTHERIA_1-0_WR05_Aug2008.txt"

FAMILY_JA = {
    "Cercopithecidae": "オナガザル科", "Cebidae": "オマキザル科",
    "Pitheciidae": "サキ科", "Atelidae": "クモザル科",
    "Cheirogaleidae": "コビトキツネザル科", "Lemuridae": "キツネザル科",
    "Galagidae": "ガラゴ科", "Hylobatidae": "テナガザル科",
    "Indriidae": "インドリ科", "Lorisidae": "ロリス科",
    "Lepilemuridae": "イタチキツネザル科", "Aotidae": "ヨザル科",
    "Hominidae": "ヒト科", "Tarsiidae": "メガネザル科",
    "Daubentoniidae": "アイアイ科",
}

# 元の列名 → 授業で使う列名
COLS = {
    "MSW05_Binomial":             "学名",
    "MSW05_Genus":                "属",
    "5-1_AdultBodyMass_g":        "体重g",
    "13-1_AdultHeadBodyLen_mm":   "頭胴長mm",
    "5-3_NeonateBodyMass_g":      "新生児体重g",
    "10-2_SocialGrpSize":         "集団サイズ",
    "9-1_GestationLen_d":         "妊娠期間日",
    "25-1_WeaningAge_d":          "離乳日齢",
    "3-1_AgeatFirstBirth_d":      "初産日齢",
    "14-1_InterbirthInterval_d":  "出産間隔日",
    "15-1_LitterSize":            "一腹産子数",
    "17-1_MaxLongevity_m":        "最長寿命月",
    "22-1_HomeRange_km2":         "行動圏km2",
    "21-1_PopulationDensity_n/km2": "個体群密度",
    "26-1_GR_Area_km2":           "分布域km2",
    "6-2_TrophicLevel":           "栄養段階",
    "12-1_HabitatBreadth":        "生息環境幅",
    "28-2_Temp_Mean_01degC":      "平均気温01degC",
    "28-1_Precip_Mean_mm":        "月降水量mm",
}

ORDER = ["学名", "科", "属", "体重g", "頭胴長mm", "新生児体重g", "集団サイズ",
         "妊娠期間日", "離乳日齢", "初産日齢", "出産間隔日", "一腹産子数",
         "最長寿命月", "行動圏km2", "個体群密度", "分布域km2",
         "栄養段階", "生息環境幅", "平均気温C", "月降水量mm"]


def build(src: str = SRC) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_csv(src, sep="\t")
    p = raw[raw["MSW05_Order"] == "Primates"].copy()

    out = p[list(COLS)].rename(columns=COLS)
    out.insert(1, "科", p["MSW05_Family"].map(FAMILY_JA))
    if out["科"].isna().any():
        missing = sorted(p.loc[out["科"].isna(), "MSW05_Family"].unique())
        raise SystemExit(f"和名の対応がない科があります: {missing}")

    # 0.1℃単位 → ℃（欠測コードは保ったまま換算しないよう、有効値だけ割る）
    t = out.pop("平均気温01degC")
    out["平均気温C"] = np.where(t == -999, -999, (t / 10).round(1))

    out = out[ORDER].sort_values("学名").reset_index(drop=True)

    raw_csv = out.copy()                      # -999 のまま
    clean = out.replace(-999, np.nan)         # 欠測として扱う
    return raw_csv, clean


def main() -> int:
    dest = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    dest.mkdir(parents=True, exist_ok=True)
    raw_csv, clean = build()

    raw_csv.to_csv(dest / "primates_raw.csv", index=False, encoding="utf-8")
    clean.to_csv(dest / "primates.csv", index=False, encoding="utf-8")

    print(f"種数: {len(clean)}　科数: {clean['科'].nunique()}")
    print(f"→ {dest/'primates_raw.csv'}（欠測 -999 のまま）")
    print(f"→ {dest/'primates.csv'}（欠測を空欄に）")
    print("\n列ごとの有効データ数:")
    for c in ORDER:
        n = clean[c].notna().sum()
        print(f"  {c:<12} {n:4d}種 ({n/len(clean)*100:5.1f}%)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
