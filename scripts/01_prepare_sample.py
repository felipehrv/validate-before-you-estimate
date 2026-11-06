"""
Build the lesson sample from the Comparative Agendas Project (CAP) file of
U.S. congressional bills.

Every bill in the CAP file carries a major-topic code assigned by trained
human coders. We keep that code as the gold standard, collapse the 21 CAP
topics into the lesson's seven labels, and draw a stratified random sample
(the same number of bills per Congress) so the time series is evenly
covered. A separate, disjoint pool of titles is drawn for the few-shot
examples and the hand-coding exercise, so nothing used inside a prompt is
ever scored.

Run once:  python3 scripts/01_prepare_sample.py
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd

PROJECT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT / "data" / "raw"
OUT_DIR = PROJECT / "data" / "processed"
OUT_DIR.mkdir(parents=True, exist_ok=True)

FIRST_CONGRESS, LAST_CONGRESS = 93, 114     # 1973 to 2016
PER_CONGRESS = 75
EXAMPLES_PER_LABEL = 6
SEED = 2026

# CAP major topic codes kept as their own label; everything else is Other.
CAP_TO_LABEL = {3: "Health", 6: "Education", 7: "Environment",
                9: "Immigration", 12: "Law and Crime", 13: "Social Welfare"}
LABELS = list(CAP_TO_LABEL.values()) + ["Other"]

# ---------------------------------------------------------------------------
# Load and inspect
# ---------------------------------------------------------------------------
files = sorted(RAW_DIR.glob("*.csv"))
if not files:
    sys.exit(f"No CSV found in {RAW_DIR}. Download the CAP bills file first.")
raw_file = files[0]
print(f"Reading {raw_file.name} ...")
raw = pd.read_csv(raw_file, low_memory=False, encoding_errors="replace")
print(f"{len(raw):,} rows, columns: {list(raw.columns)}")

def find(*candidates):
    """First column whose lower-cased name contains one of the candidates."""
    for cand in candidates:
        for col in raw.columns:
            if cand in col.lower():
                return col
    return None

col_title = find("title", "description")
col_topic = "majortopic" if "majortopic" in raw.columns else find("majortopic")
col_year = find("year")
col_cong = find("cong")
col_chamber = find("chamber")
col_party = find("party")
col_id = find("id")
print(f"Using title={col_title!r}, topic={col_topic!r}, year={col_year!r}, "
      f"congress={col_cong!r}, chamber={col_chamber!r}, party={col_party!r}, "
      f"id={col_id!r}")
if col_title is None or col_topic is None:
    print(raw.head(3).T)
    sys.exit("Could not identify the title or topic column; see columns above.")

# ---------------------------------------------------------------------------
# Clean
# ---------------------------------------------------------------------------
df = pd.DataFrame({
    "bill_id": raw[col_id] if col_id else np.arange(len(raw)),
    "title": raw[col_title].astype(str).str.strip(),
    "cap_majortopic": pd.to_numeric(raw[col_topic], errors="coerce"),
    "year": pd.to_numeric(raw[col_year], errors="coerce") if col_year else np.nan,
    "congress": pd.to_numeric(raw[col_cong], errors="coerce") if col_cong else np.nan,
    "chamber": raw[col_chamber] if col_chamber else np.nan,
    "party": raw[col_party] if col_party else np.nan,
})
if df["congress"].isna().all() and not df["year"].isna().all():
    # Congress number from the year: the 93rd Congress began in 1973.
    df["congress"] = ((df["year"] - 1789) // 2 + 1).astype("Int64")

df["title"] = df["title"].str.replace(r'"{2,}', '"', regex=True)
df = df.dropna(subset=["title", "cap_majortopic"])
df = df[df["title"].str.len().between(15, 400)]
# Private relief bills phrased "For the relief of ..." are dropped; the 62 phrased
# "A bill for the relief of ..." stay in (trivially Other for humans and model).
df = df[~df["title"].str.lower().str.startswith("for the relief of")]
df["cap_majortopic"] = df["cap_majortopic"].astype(int)
df["label_gold"] = df["cap_majortopic"].map(CAP_TO_LABEL).fillna("Other")
df = df[df["congress"].between(FIRST_CONGRESS, LAST_CONGRESS)]
df = df.drop_duplicates(subset="title")
print(f"\nAfter cleaning: {len(df):,} bills in Congresses "
      f"{FIRST_CONGRESS} to {LAST_CONGRESS}")
print("Gold label distribution in the population:")
print(df["label_gold"].value_counts(normalize=True).round(3))

# ---------------------------------------------------------------------------
# Draw the example pool first, then the evaluation sample from what remains
# ---------------------------------------------------------------------------
pool = (df.groupby("label_gold", group_keys=False)
          .apply(lambda g: g.sample(n=EXAMPLES_PER_LABEL, random_state=SEED)))
rest = df.drop(index=pool.index)

sample = (rest.groupby("congress", group_keys=False)
              .apply(lambda g: g.sample(n=min(PER_CONGRESS, len(g)),
                                        random_state=SEED)))
sample = sample.sample(frac=1, random_state=SEED).reset_index(drop=True)

sample.to_csv(OUT_DIR / "bills_sample.csv", index=False)
pool.to_csv(OUT_DIR / "example_pool.csv", index=False)
print(f"\nEvaluation sample: {len(sample):,} titles "
      f"({sample['congress'].nunique()} Congresses)")
print(sample["label_gold"].value_counts().to_string())
print(f"Example pool: {len(pool)} titles, {EXAMPLES_PER_LABEL} per label (disjoint from the sample)")
print("\nFive random titles from the sample:")
for _, r in sample.sample(5, random_state=1).iterrows():
    print(f"  [{r.label_gold}] {r.title[:110]}")
