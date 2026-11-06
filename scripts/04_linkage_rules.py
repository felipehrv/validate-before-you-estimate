"""
Record linkage, step 1: rules do the easy work, and we find the hard pairs.

Data: Febrl 4 (bundled with the recordlinkage package). Two files of 5,000
person records; every record in file B is a corrupted copy of one record in
file A, so the 5,000 true links are known. That is what lets us measure
precision and recall instead of guessing.

Pipeline
--------
1. Candidate pairs by blocking on several keys (union), so a typo in one
   field does not hide a true pair.
2. Field-by-field similarity (Jaro-Winkler for names, exact for date of
   birth and postcode). Only fields that two administrative files would
   realistically share; there is no shared identifier.
3. A simple additive score. Pairs above a high cutoff are declared matches,
   pairs below a low cutoff are discarded, and the band in between is the
   "ambiguous" set that a human reviewer, or an LLM, must adjudicate.
4. Precision and recall for the rules alone, under two policies for the
   ambiguous band (treat as match / treat as non-match).

Outputs
-------
outputs/linkage_scores.csv           all candidate pairs with score and truth
outputs/linkage_ambiguous_pairs.csv  the band, with both records side by side
"""

from pathlib import Path

import pandas as pd
import recordlinkage as rl
from recordlinkage.datasets import load_febrl4

PROJECT = Path(__file__).resolve().parents[1]
OUT = PROJECT / "outputs"
OUT.mkdir(exist_ok=True)

LOW, HIGH = 1.5, 2.5   # cutoffs on the additive score (0 to 4): score 2 is the band

dfA, dfB, true_links = load_febrl4(return_links=True)
print(f"File A: {len(dfA)} records, file B: {len(dfB)} records, "
      f"true links: {len(true_links)}")

# 1. Candidate pairs: union of three blocking keys ------------------------------
indexer = rl.Index()
indexer.sortedneighbourhood("surname", window=5)
indexer.block("date_of_birth")
indexer.block("postcode")
candidates = indexer.index(dfA, dfB)
truth_in_candidates = candidates.isin(true_links).sum()
print(f"Candidate pairs: {len(candidates):,}; true links among them: "
      f"{truth_in_candidates} of {len(true_links)} "
      f"(blocking recall {truth_in_candidates/len(true_links):.3f})")

# 2. Similarity per field -----------------------------------------------------
comparer = rl.Compare()
# Only the fields two administrative files typically share. No shared ID.
comparer.string("given_name", "given_name", method="jarowinkler",
                threshold=0.85, label="given_name")
comparer.string("surname", "surname", method="jarowinkler",
                threshold=0.85, label="surname")
comparer.exact("date_of_birth", "date_of_birth", label="date_of_birth")
comparer.exact("postcode", "postcode", label="postcode")
features = comparer.compute(candidates, dfA, dfB)

# 3. Additive score and three zones --------------------------------------------
features["score"] = features.sum(axis=1)
features["truth"] = features.index.isin(true_links).astype(int)
features["zone"] = pd.cut(features["score"], bins=[-1, LOW, HIGH, 5],
                          labels=["non-match", "ambiguous", "match"])
print("\nPairs by zone (rows) and truth (columns):")
print(pd.crosstab(features["zone"], features["truth"]))

# 4. Precision and recall of the rules alone ----------------------------------
def precision_recall(predicted, truth_index):
    predicted = set(predicted)
    truth = set(truth_index)
    tp = len(predicted & truth)
    precision = tp / len(predicted) if predicted else float("nan")
    recall = tp / len(truth)
    return precision, recall

sure = features.index[features["zone"] == "match"]
amb = features.index[features["zone"] == "ambiguous"]
for policy, predicted in [("ambiguous treated as non-match", sure),
                          ("ambiguous treated as match", sure.union(amb))]:
    p, r = precision_recall(predicted, true_links)
    print(f"Rules only, {policy}: precision {p:.3f}, recall {r:.3f}")

features.reset_index().to_csv(OUT / "linkage_scores.csv", index=False)

# The ambiguous band, laid out for a reviewer ----------------------------------
rows = []
for (a, b) in amb:
    ra, rb = dfA.loc[a], dfB.loc[b]
    rows.append({"pair_id": f"{a}|{b}", "id_a": a, "id_b": b,
                 **{f"{c}_a": ra[c] for c in dfA.columns},
                 **{f"{c}_b": rb[c] for c in dfB.columns},
                 "score": features.loc[(a, b), "score"],
                 "truth": int((a, b) in set(true_links))})
pd.DataFrame(rows).to_csv(OUT / "linkage_ambiguous_pairs.csv", index=False)
print(f"\nAmbiguous pairs written: {len(rows)}")


# 5. Why it matters downstream: a "moved" rate ---------------------------------
# Suppose the research question is residential mobility: the share of people
# whose postcode differs between the two files. Postcode is also one of the
# matching fields, so linkage errors are correlated with the outcome.
pairs = features.reset_index()
pairs["moved"] = (pairs["postcode"] == 0).astype(int)
truth_pairs = pairs[pairs["truth"] == 1]
print(f"\nTrue share with a different postcode: "
      f"{truth_pairs['moved'].mean():.3f} (n = {len(truth_pairs)})")
for policy, mask in [("ambiguous treated as match",
                      pairs["zone"].isin(["match", "ambiguous"])),
                     ("ambiguous treated as non-match",
                      pairs["zone"] == "match")]:
    est = pairs.loc[mask, "moved"].mean()
    print(f"Estimated share, {policy}: {est:.3f} (n = {mask.sum()})")
