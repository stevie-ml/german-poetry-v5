"""
04_merge.py — Merge poems_raw.json + surprisal + annotations into poems_v5.csv.

Matching strategy: poet+poem_title+n_lines key (NFC-normalized) rather than poem_id,
because multi-file poets (Klopstock, trakl) reset the counter per file, creating
duplicate poem_ids that make poem_id joins produce cartesian products.

v4 data covers 17 poets (all except George). George's data comes from fresh runs.
"""

import json, re, unicodedata
from pathlib import Path
import pandas as pd
import numpy as np

V4_OUT = Path.home() / "german-poetry-apr27" / "output"
V5_OUT = Path.home() / "german-poetry-v5" / "output"


def nfc(s):
    return unicodedata.normalize("NFC", str(s)) if pd.notna(s) and s else ""

def make_key(poet, title, n_lines):
    return nfc(poet) + "||" + nfc(title) + "||" + str(int(n_lines) if pd.notna(n_lines) else 0)

def parse_year(val):
    if not val:
        return np.nan
    m = re.search(r"\d{4}", str(val))
    return int(m.group()) if m else np.nan


# ── Load and key v5 raw poems ─────────────────────────────────────────────────
with open(V5_OUT / "poems_raw.json", encoding="utf-8") as f:
    raw = json.load(f)

df_raw = pd.DataFrame([{
    "poem_id":                r["poem_id"],
    "poet":                   r["poet"],
    "collection":             r["collection"],
    "poem_title":             r["poem_title"],
    "text":                   r["text"],
    "n_stanzas":              r["n_stanzas"],
    "n_lines":                r["n_lines"],
    "pub_place":              r.get("pub_place", ""),
    "edition_year":           r.get("edition_year", ""),
    "erstdruck_year":         r.get("erstdruck_year", ""),
    "comp_date_from":         r.get("comp_date_from", ""),
    "comp_date_to":           r.get("comp_date_to", ""),
    "george_normalized_text": r.get("george_normalized_text", ""),
} for r in raw])

df_raw["edition_year_int"]   = df_raw["edition_year"].apply(parse_year)
df_raw["erstdruck_year_int"] = df_raw["erstdruck_year"].apply(parse_year)
df_raw["comp_year"] = df_raw.apply(
    lambda r: np.nanmean([parse_year(r["comp_date_from"]), parse_year(r["comp_date_to"])]),
    axis=1
)

# date_uncertain: comp range > 20 years, or no erstdruck when collection has some
df_raw["comp_range"] = df_raw.apply(
    lambda r: abs((parse_year(r["comp_date_to"]) or 0) - (parse_year(r["comp_date_from"]) or 0)),
    axis=1
)
df_raw["date_uncertain"] = df_raw["comp_range"] > 40

df_raw["_key"] = df_raw.apply(lambda r: make_key(r["poet"], r["poem_title"], r["n_lines"]), axis=1)


# ── Load, key, deduplicate v4 surprisal ───────────────────────────────────────
surp_cols = ["poem_id", "poet", "poem_title", "n_lines", "n_words",
             "mean_surprisal", "mean_entropy", "s2_mean", "uid_sigma",
             "tension", "peak_pos", "ac1", "ttr"]

df_surp_v4     = pd.read_csv(V4_OUT / "poems_surprisal.csv")[surp_cols]
df_surp_george = pd.read_csv(V5_OUT / "george_surprisal.csv")[surp_cols]
df_surp_all    = pd.concat([df_surp_v4, df_surp_george], ignore_index=True)
df_surp_all["_key"] = df_surp_all.apply(
    lambda r: make_key(r["poet"], r["poem_title"], r["n_lines"]), axis=1)
df_surp_dedup = df_surp_all.drop_duplicates(subset=["_key"], keep="first")

surp_merge_cols = ["_key", "n_words", "mean_surprisal", "mean_entropy",
                   "s2_mean", "uid_sigma", "tension", "peak_pos", "ac1", "ttr"]


# ── Load, key, deduplicate annotations ────────────────────────────────────────
ann_cols = ["poem_id", "poet", "poem_title", "person", "has_addressee",
            "emotional_valence", "emotional_intensity", "dominant_theme",
            "imagery_density", "tone", "temporal_frame", "setting", "closure"]

df_ann_v4     = pd.read_csv(V4_OUT / "poems_annotated.csv")[ann_cols]
df_ann_george = pd.read_csv(V5_OUT / "george_annotated.csv")[ann_cols]

# annotated CSV may not have n_lines — we need it for the key; join back via poem_id match
# Instead, build key from poet + poem_title only for v4 (n_lines not in annotated CSV)
# Use two-step: match v4 annotated to v4 surprisal to get n_lines, then key
df_surp_v4_for_ann = df_surp_v4[["poem_id","n_lines"]].copy()
df_ann_v4 = df_ann_v4.merge(df_surp_v4_for_ann, on="poem_id", how="left")

df_surp_george_for_ann = df_surp_george[["poem_id","n_lines"]].copy()
df_ann_george = df_ann_george.merge(df_surp_george_for_ann, on="poem_id", how="left")

df_ann_all = pd.concat([df_ann_v4, df_ann_george], ignore_index=True)
df_ann_all["_key"] = df_ann_all.apply(
    lambda r: make_key(r["poet"], r["poem_title"], r["n_lines"]), axis=1)
df_ann_dedup = df_ann_all.drop_duplicates(subset=["_key"], keep="first")

ann_merge_cols = ["_key", "person", "has_addressee", "emotional_valence",
                  "emotional_intensity", "dominant_theme", "imagery_density",
                  "tone", "temporal_frame", "setting", "closure"]


# ── Merge ─────────────────────────────────────────────────────────────────────
df = df_raw.merge(df_surp_dedup[surp_merge_cols], on="_key", how="left")
df = df.merge(df_ann_dedup[ann_merge_cols],        on="_key", how="left")
df = df.drop(columns=["_key"])


# ── Movement mapping ──────────────────────────────────────────────────────────
MOVEMENT = {
    "Albrech von Haller":        "Aufklärung",
    "Barthold Heinrich Brockes": "Aufklärung",
    "Klopstock":                 "Empfindsamkeit",
    "goethe":                    "Klassik / Sturm und Drang",
    "Holderlin":                 "Klassik / Romantik",
    "Mörike":                    "Biedermeier",
    "Droste-Hülshoff":           "Biedermeier",
    "storm":                     "Realismus",
    "Schiller":                  "Klassik",
    "Rilke":                     "Moderne",
    "hvh":                       "Symbolismus",
    "george":                    "Symbolismus",
    "Felix Dörmann":             "Symbolismus",
    "heym":                      "Expressionismus",
    "stadler":                   "Expressionismus",
    "trakl":                     "Expressionismus",
    "wedekind":                  "Expressionismus",
    "Morgenstern":               "Moderne",
}
df["movement"] = df["poet"].map(MOVEMENT).fillna("Unbekannt")

# temporal_year: best available date for temporal analysis.
# Prefers erstdruck_year_int (actual first publication) when available.
# Falls back to comp_year only when date_uncertain=False (comp range ≤ 40 yrs).
# Lifespan proxies (date_uncertain=True, no erstdruck) are left as NaN.
df["temporal_year"] = df.apply(
    lambda r: r["erstdruck_year_int"] if pd.notna(r["erstdruck_year_int"])
              else (r["comp_year"] if not r["date_uncertain"] and pd.notna(r["comp_year"]) else np.nan),
    axis=1
)


# ── Reorder columns ───────────────────────────────────────────────────────────
ordered = [
    "poem_id", "poet", "movement", "collection", "poem_title",
    "pub_place",
    "comp_date_from", "comp_date_to", "comp_year",
    "erstdruck_year", "erstdruck_year_int",
    "edition_year", "edition_year_int",
    "date_uncertain", "temporal_year",
    "n_lines", "n_stanzas", "n_words",
    "mean_surprisal", "mean_entropy", "s2_mean", "uid_sigma",
    "tension", "peak_pos", "ac1", "ttr",
    "person", "has_addressee", "emotional_valence", "emotional_intensity",
    "dominant_theme", "imagery_density", "tone", "temporal_frame", "setting", "closure",
    "text", "george_normalized_text",
]
df = df[[c for c in ordered if c in df.columns]]

out_path = V5_OUT / "poems_v5.csv"
df.to_csv(out_path, index=False)
print(f"Merged: {len(df)} poems -> {out_path}")
print(f"Columns ({len(df.columns)}): {list(df.columns)}")
print()
print("Poet breakdown:")
for poet, grp in df.groupby("poet"):
    surp_ok  = grp["mean_surprisal"].notna().sum()
    comp_ok  = grp["comp_year"].notna().sum()
    erst_ok  = grp["erstdruck_year_int"].notna().sum()
    ed_ok    = grp["edition_year_int"].notna().sum()
    uncert   = grp["date_uncertain"].sum()
    print(f"  {poet:35s} {len(grp):4d}  surprisal={surp_ok}  comp={comp_ok}  erstdruck={erst_ok}  edition={ed_ok}  uncertain={uncert}")

print(f"\nDate coverage summary:")
print(f"  comp_year:      {df['comp_year'].notna().sum()}/{len(df)}")
print(f"  erstdruck_year: {df['erstdruck_year_int'].notna().sum()}/{len(df)}")
print(f"  edition_year:   {df['edition_year_int'].notna().sum()}/{len(df)}")
print(f"  date_uncertain: {df['date_uncertain'].sum()} poems (comp range >20 yrs)")
