"""
Build the lesson notebooks (lessons/ with scaffolded challenges, solutions/
filled in) following the D-Lab Jupyter style guide, then execute the solution
notebooks to make sure every cell runs on the cached data.

Run: python3 scripts/07_build_notebooks.py
"""

import subprocess
from pathlib import Path

import nbformat as nbf

PROJECT = Path(__file__).resolve().parents[1]
LESSONS = PROJECT / "lessons"
SOLUTIONS = PROJECT / "solutions"
LESSONS.mkdir(exist_ok=True)
SOLUTIONS.mkdir(exist_ok=True)


def build(cells, solution):
    nb = nbf.v4.new_notebook()
    nb.metadata["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
    out = []
    for kind, text, *extra in cells:
        if kind == "md":
            out.append(nbf.v4.new_markdown_cell(text))
        elif kind == "code":
            out.append(nbf.v4.new_code_cell(text))
        elif kind == "challenge":
            # text is the scaffold shown in lessons, extra[0] the solution
            out.append(nbf.v4.new_code_cell(extra[0] if solution else text))
    nb.cells = out
    return nb


# ===========================================================================
# Notebook 1: Labels as Measurements
# ===========================================================================
nb1 = [
("md", """# Validate Before You Estimate: Labels as Measurements

* * *

<div class="alert alert-success">

### Learning Objectives

After this notebook, you will be able to:

- Explain why a classification prompt is a measurement instrument, and what changes when the codebook goes into it.
- Compare AI labels with human codes using accuracy, kappa, precision and recall, by class and by group.
- Explain why high accuracy does not guarantee an unbiased estimate or a valid comparison between groups.
- Correct an estimate with a small hand-coded sample (prediction-powered inference).
- Decide how many documents to code by hand from an interval curve.
</div>

### Icons Used in This Notebook
🔔 **Question**: A quick question to help you understand what's going on.<br>
🥊 **Challenge**: Interactive exercise. We'll work through these in the workshop!<br>
💡 **Tip**: How to do something a bit more efficiently or effectively.<br>
⚠️ **Warning**: Heads-up about tricky stuff or common mistakes.<br>
🎬 **Demo**: Showing off something more advanced, so you know what's possible!

The in-room exercises (the handout's Challenges 1 to 3) are hand-coding, betting and reading a curve. The notebook challenges below are their coding counterparts and are numbered on their own.

### Sections
1. [The Data](#data)
2. [From Codebook to Prompt](#prompt)
3. [Validate](#validate)
4. [Estimate](#estimate)
5. [How Many to Hand-Code?](#howmany)
6. [Key Points](#keypoints)

Every model call used in this notebook was run beforehand and saved in `outputs/`, so you do not need an API key. The script that produced the labels is `scripts/02_classify_titles.py`."""),
("code", """import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# The agreement metrics live in scripts/lesson_tools.py; the estimators are written out below.
sys.path.insert(0, "../scripts")
from lesson_tools import accuracy, cohen_kappa, confusion, per_class, sample_size_curve

LABELS = ["Health", "Education", "Environment", "Immigration",
          "Law and Crime", "Social Welfare", "Other"]"""),
("md", """<a id='data'></a>

# The Data

We use bill titles from the **Comparative Agendas Project** (CAP). Every bill introduced in the U.S. Congress has been coded into a policy topic by trained human coders, which gives us an **answer key**. We keep 1,650 titles (75 per Congress, 1973 to 2016) and collapse the 21 CAP topics into seven labels."""),
("code", """bills = pd.read_csv("../data/processed/bills_sample.csv")
bills["party_name"] = bills["party"].map({100: "Democrat", 200: "Republican"})
print(bills.shape)
bills[["bill_id", "year", "party_name", "label_gold", "title"]].head()"""),
("code", """bills["label_gold"].value_counts(normalize=True).round(3)"""),
("md", """🔔 **Question**: "Other" is 72% of the sample. What accuracy would a classifier get if it answered "Other" for every title? Keep that number in mind."""),
("md", """<a id='prompt'></a>

# From Codebook to Prompt

Three prompts were written for the same task. Think of each one as a **version of a measurement instrument**:

- `prompt_1_names`: the seven category names and nothing else.
- `prompt_2_codebook`: a condensed version of the CAP codebook, with boundary rules.
- `prompt_3_codebook_examples`: the codebook plus fourteen human-coded examples drawn from **outside** the evaluation sample.

Every call asked for a fixed JSON answer: `{"label": ..., "why": ...}`. The label, the model name, the prompt version, the tokens and a timestamp were logged."""),
("code", """print(open("../prompts/prompt_2_codebook.txt").read()[:1500])"""),
("code", """labels = pd.read_csv("../outputs/llm_labels_clean.csv")
labels = labels.drop_duplicates(subset=["model", "prompt", "bill_id"], keep="last")
labels.groupby(["model", "prompt"]).size()"""),
("md", """💡 **Tip**: Never put a document you will score inside the prompt as an example. The examples here come from a separate pool."""),
("md", """<a id='validate'></a>

# Validate

We merge the model labels with the human codes and compute agreement for each instrument."""),
("code", """d = labels.merge(bills, on="bill_id")
d["correct"] = d["label"] == d["label_gold"]

summary = []
for (model, prompt), g in d.groupby(["model", "prompt"]):
    summary.append({"model": model, "prompt": prompt,
                    "accuracy": accuracy(g["label_gold"], g["label"]),
                    "kappa": cohen_kappa(g["label_gold"], g["label"], LABELS)})
pd.DataFrame(summary).round(3)"""),
("md", """## 🥊 Notebook Challenge 1: Always "Other"

Compute the accuracy and the kappa of a classifier that answers `"Other"` for every title. Compare with the table above."""),
("challenge", """always_other = ["Other"] * len(bills)
# YOUR CODE HERE
...""", """always_other = ["Other"] * len(bills)
print("accuracy:", round(accuracy(bills["label_gold"], always_other), 3))
print("kappa:", round(cohen_kappa(bills["label_gold"], always_other, LABELS), 3))"""),
("md", """Accuracy flatters any classifier when one class dominates. **Kappa** measures agreement beyond that chance level, which is why we report both.

## The Confusion Matrix

Rows are the human codes, columns are the model's labels (best cheap instrument: Haiku with prompt 3)."""),
("code", """best = d[(d["model"] == "claude-haiku-5-5") & (d["prompt"] == "prompt_3_codebook_examples")].copy()
cm = confusion(best["label_gold"], best["label"], LABELS)
cm"""),
("md", """Two questions for every category. **Precision**: when the model says Health, how often do the humans agree? **Recall**: when the humans say Health, how often does the model notice? The `per_class()` function computes both from the matrix above."""),
("code", """per_class(best["label_gold"], best["label"], LABELS).round(3)"""),
("md", """🔔 **Question**: Which category has the lowest recall? Where do its missed bills end up?

## 🥊 Notebook Challenge 2: Whose Error Is It?

Print five titles where the model and the humans disagree, with the model's one-sentence reason. For each, decide: is the model wrong, is the human code wrong, or is the codebook ambiguous?"""),
("challenge", """disagreements = best[best["label"] != best["label_gold"]]
# YOUR CODE HERE
...""", """disagreements = best[best["label"] != best["label_gold"]]
for row in disagreements.sample(5, random_state=3).itertuples():
    print(f"HUMAN: {row.label_gold:15s} MODEL: {row.label:15s}")
    print(f"  {row.title[:150]}")
    print(f"  model's reason: {row.why[:150]}\\n")"""),
("md", """## Differential Error

An instrument can be equally accurate for two groups and still bias the comparison between them, if the **direction** of its errors differs. We check the share of health bills by sponsor's party."""),
("code", """best["is_health_gold"] = best["label_gold"] == "Health"
best["is_health_llm"] = best["label"] == "Health"
best[best["party_name"].notna()].groupby("party_name")[["correct", "is_health_gold", "is_health_llm"]].mean().round(3)"""),
("md", """⚠️ **Warning**: The humans see no partisan gap. The model creates one. Accuracy by group would not have warned you; validate the **estimate** by group, not only the labels."""),
("md", """<a id='estimate'></a>

# Estimate

Suppose we can afford to hand-code only 200 titles. We draw them at random. On those 200 we observe both the model's label and the truth, so we can measure the model's average error and subtract it from the model's average over all 1,650 titles. That is **prediction-powered inference** (Angelopoulos et al., 2023), and the design-based estimators of Egami et al. (2023, 2024) generalize it."""),
("md", """First the random hand-coded subset. The seed makes the draw reproducible; in a real project the draw is the one step you must not skip, because the correction only works if the subset is random."""),
("code", """rng = np.random.default_rng(11)
hand_idx = rng.choice(best.index, size=200, replace=False)
best["hand_coded"] = best.index.isin(hand_idx)
best["hand_coded"].sum()"""),
("md", """Now the three estimates, written out. Each label becomes a number: 1 if the title was called Health, 0 if not, so a share is a mean.

- **Naive**: the mean of the model's labels over all N titles, with the usual interval for a mean. It ignores the model's error.
- **Hand-coded subset only**: the mean of the human codes over the n hand-coded titles. Unbiased, but only n titles wide.
- **Corrected (PPI)**: the model's mean over all N titles, minus the model's average error measured on the n hand-coded titles. The standard error adds the two pieces."""),
("code", """def estimate_naive(model_flag_all):
    f = np.asarray(model_flag_all, dtype=float)
    return f.mean(), np.sqrt(f.var(ddof=1) / len(f))

def estimate_human_only(human_flag_subset):
    y = np.asarray(human_flag_subset, dtype=float)
    return y.mean(), np.sqrt(y.var(ddof=1) / len(y))

def estimate_corrected(model_flag_all, model_flag_subset, human_flag_subset):
    f_all = np.asarray(model_flag_all, dtype=float)
    f_sub = np.asarray(model_flag_subset, dtype=float)
    y_sub = np.asarray(human_flag_subset, dtype=float)
    error = f_sub - y_sub                                   # the model's error on the hand-coded titles
    estimate = f_all.mean() - error.mean()                  # the one line
    se = np.sqrt(f_all.var(ddof=1) / len(f_all) + error.var(ddof=1) / len(error))
    return estimate, se"""),
("md", """A small helper prints the three estimates with their 95% intervals, next to the truth we happen to know here because every title was hand-coded."""),
("code", """def three_estimates(df):
    f_all = df["is_health_llm"].astype(float).values
    sub = df[df["hand_coded"]]
    f_sub = sub["is_health_llm"].astype(float).values
    y_sub = sub["is_health_gold"].astype(float).values
    rows = {}
    rows["naive (model, all titles)"] = estimate_naive(f_all)
    rows["hand-coded subset only"] = estimate_human_only(y_sub)
    rows["corrected (PPI)"] = estimate_corrected(f_all, f_sub, y_sub)
    table = pd.DataFrame(rows, index=["estimate", "se"]).T
    table["ci_low"] = table["estimate"] - 1.96 * table["se"]
    table["ci_high"] = table["estimate"] + 1.96 * table["se"]
    table.loc["truth (all titles hand-coded)"] = [df["is_health_gold"].mean(), np.nan, np.nan, np.nan]
    return table

three_estimates(best).round(3)"""),
("md", """💡 **Tip**: the hand-coded-only interval above is the textbook one, which works when the share is not close to 0 or 1 and n is a few hundred. For a small group with a rare label (say 1 health bill in 63), use a Wilson interval instead; the textbook one dips below zero and runs too narrow.

The same three estimates, by sponsor's party. This is where the naive estimate goes wrong in a way that matters."""),
("code", """for party in ["Democrat", "Republican"]:
    print(party)
    print(three_estimates(best[best["party_name"] == party]).round(3), "\\n")"""),
("md", """🔔 **Question**: The naive point estimates for the two parties sit 2.5 points apart, in opposite directions from the truth, and their intervals ignore the model's error. What happens to that gap after the correction? What would happen to the naive intervals if the corpus had 20,000 titles instead of 1,650?

## 🥊 Notebook Challenge 3: Another Outcome

Repeat the three estimates for the share of bills about **Education**. Does the naive estimate over- or under-state it?"""),
("challenge", """best["is_edu_gold"] = best["label_gold"] == "Education"
best["is_edu_llm"] = best["label"] == "Education"
# YOUR CODE HERE
...""", """best["is_edu_gold"] = best["label_gold"] == "Education"
best["is_edu_llm"] = best["label"] == "Education"
f_all = best["is_edu_llm"].astype(float).values
sub = best[best["hand_coded"]]
f_sub, y_sub = sub["is_edu_llm"].astype(float).values, sub["is_edu_gold"].astype(float).values
for name, (est, se) in {"naive": estimate_naive(f_all), "human only": estimate_human_only(y_sub),
                        "corrected": estimate_corrected(f_all, f_sub, y_sub)}.items():
    print(f"{name:12s} {est:.3f}  ({est - 1.96 * se:.3f}, {est + 1.96 * se:.3f})")
print(f"{'truth':12s} {best['is_edu_gold'].mean():.3f}")"""),
("md", """<a id='howmany'></a>

# How Many to Hand-Code?

The corrected interval narrows as the hand-coded sample grows. We resample the subset at several sizes and plot the half-width of the 95% interval."""),
("code", """sizes = [25, 50, 100, 200, 400, 800]
curve = sample_size_curve(best["is_health_llm"].astype(float).values,
                          best["is_health_llm"].astype(float).values,
                          best["is_health_gold"].astype(float).values, sizes, reps=200)
curve.round(4)"""),
("code", """fig, ax = plt.subplots(figsize=(7, 4))
ax.plot(curve["n_hand_coded"], curve["ci_halfwidth"], marker="o")
ax.set_xscale("log")
ax.set_xticks(sizes, [str(s) for s in sizes])
ax.set_xlabel("Titles coded by hand")
ax.set_ylabel("95% interval half-width")
ax.set_title("Doubling n cuts the width by about 30%")
plt.show()"""),
("md", """The two variances in the standard error are all you need to **project the curve to a bigger corpus**: the variance of the model flag over all N documents, and the variance of the model-minus-human difference over the n hand-coded ones. In your own project you would estimate both from a pilot of about 100 hand-coded documents."""),
("code", """var_f = best["is_health_llm"].astype(float).var(ddof=1)
var_r = (best["is_health_llm"].astype(float) - best["is_health_gold"].astype(float)).var(ddof=1)

def halfwidth(n, N=10_000):
    \"\"\"Half-width of the 95% corrected interval with n hand-coded documents out of N.\"\"\"
    return 1.96 * np.sqrt(var_f / N + var_r / n)

for n in [100, 200, 300, 400, 800]:
    print(f"n = {n:4d}: plus or minus {halfwidth(n):.1%}   (hand-coded sample alone: {1.96 * np.sqrt(best['is_health_gold'].mean() * (1 - best['is_health_gold'].mean()) / n):.1%})")"""),
("md", """🔔 **Question**: Your committee wants to know whether Democrats and Republicans differ by 3 points or more in the share of health bills, in a corpus of 20,000 titles. To tell a 3-point gap between two groups from zero, each group's interval must be about plus or minus 2 points or narrower. How many titles per party would you hand-code, and how many hours is that at 100 titles an hour? What would it cost without the model?

🎬 **Demo**: the correction works for a difference between two groups too. The corrected gap is the difference of the two corrected shares, and its standard error combines the two."""),
("code", """gaps = {}
for party in ["Democrat", "Republican"]:
    df = best[best["party_name"] == party]
    sub = df[df["hand_coded"]]
    gaps[party] = estimate_corrected(df["is_health_llm"].astype(float).values,
                                     sub["is_health_llm"].astype(float).values,
                                     sub["is_health_gold"].astype(float).values)
gap = gaps["Democrat"][0] - gaps["Republican"][0]
gap_se = np.sqrt(gaps["Democrat"][1] ** 2 + gaps["Republican"][1] ** 2)
naive_gap = best[best["party_name"] == "Democrat"]["is_health_llm"].mean() - best[best["party_name"] == "Republican"]["is_health_llm"].mean()
print(f"naive gap:     {naive_gap:+.3f}")
print(f"corrected gap: {gap:+.3f}  (95% interval {gap - 1.96 * gap_se:+.3f} to {gap + 1.96 * gap_se:+.3f})")"""),
("md", """<a id='keypoints'></a>

<div class="alert alert-success">

### ❗ Key Points

- An AI label is a measurement with error; the prompt is the instrument, and a changed prompt needs its own validation.
- Report accuracy and kappa, and precision and recall by class and by the groups you compare.
- Equal accuracy across groups does not rule out a spurious difference between them.
- A random hand-coded sample lets you subtract the model's bias and get a valid interval.
- The number of documents to hand-code is a design decision with a curve behind it.
</div>"""),
]

# ===========================================================================
# Notebook 2: Links as Measurements
# ===========================================================================
nb2 = [
("md", """# Validate Before You Estimate: Links as Measurements

* * *

<div class="alert alert-success">

### Learning Objectives

After this notebook, you will be able to:

- Explain why record linkage is a classification problem with the pair as the unit.
- Use rules for the easy pairs and a model (or a person) for the ambiguous band.
- Measure precision and recall of a linkage policy against known truth.
- Show how linkage error carries into a downstream estimate.
</div>

### Icons Used in This Notebook
🔔 **Question**: A quick question to help you understand what's going on.<br>
🥊 **Challenge**: Interactive exercise. We'll work through these in the workshop!<br>
💡 **Tip**: How to do something a bit more efficiently or effectively.<br>
⚠️ **Warning**: Heads-up about tricky stuff or common mistakes.

### Sections
1. [The Data](#data)
2. [Rules First](#rules)
3. [The Model Adjudicates the Band](#model)
4. [Downstream](#downstream)
5. [Key Points](#keypoints)

The rule-based step is in `scripts/04_linkage_rules.py`; the model's decisions on the ambiguous pairs were run beforehand with `scripts/05_adjudicate_pairs.py` and saved in `outputs/`."""),
("code", """import numpy as np
import pandas as pd"""),
("md", """<a id='data'></a>

# The Data

Febrl ships two files of 5,000 person records. File B is a corrupted copy of file A (typos, swapped digits, moved addresses), so every true link is known. We keep four fields two administrative files would realistically share: given name, surname, date of birth, postcode."""),
("code", """scores = pd.read_csv("../outputs/linkage_scores.csv")
print(scores.shape)
scores.head()"""),
("md", """<a id='rules'></a>

# Rules First

Each candidate pair got one point per agreeing field. Four points is a sure match, zero or one a sure non-match, and the band at two is **ambiguous**."""),
("code", """pd.crosstab(scores["score"], scores["truth"])"""),
("md", """## 🥊 Notebook Challenge 1: Precision and Recall of the Rules

Compute precision and recall when the ambiguous band is (a) dropped and (b) kept as matches. There are 5,000 true links in total."""),
("challenge", """N_TRUE = 5000
sure = scores[scores["zone"] == "match"]
band = scores[scores["zone"] == "ambiguous"]
# YOUR CODE HERE
...""", """N_TRUE = 5000
sure = scores[scores["zone"] == "match"]
band = scores[scores["zone"] == "ambiguous"]

def precision_recall(pred):
    tp = int(pred["truth"].sum())
    return tp / len(pred), tp / N_TRUE

for name, pred in [("drop the band", sure), ("keep the band", pd.concat([sure, band]))]:
    p, r = precision_recall(pred)
    print(f"{name:14s} precision {p:.3f}  recall {r:.3f}")"""),
("md", """<a id='model'></a>

# The Model Adjudicates the Band

Each ambiguous pair was shown to the model as two short records, with the instruction to decide whether they refer to the same person."""),
("code", """print(open("../prompts/pairs_prompt.txt").read())"""),
("md", """The cached decisions are one row per pair and model. We attach the truth and the two postcodes, and turn the model's `yes` into a 1."""),
("code", """pairs = pd.read_csv("../outputs/llm_pairs_clean.csv").drop_duplicates(subset=["model", "pair_id"], keep="last")
amb = pd.read_csv("../outputs/linkage_ambiguous_pairs.csv", dtype={"pair_id": str})
pairs = pairs.merge(amb[["pair_id", "truth", "postcode_a", "postcode_b"]], on="pair_id")
pairs["pred"] = (pairs["match"] == "yes").astype(int)
pairs.groupby("model")[["pred", "truth"]].mean().round(3)"""),
("md", """Precision and recall inside the band, and then for the whole policy (rules plus model), with the 5,000 true links as the denominator of recall."""),
("code", """for model, g in pairs.groupby("model"):
    tp = int(((g["pred"] == 1) & (g["truth"] == 1)).sum())
    fp = int(((g["pred"] == 1) & (g["truth"] == 0)).sum())
    print(f"{model:20s} inside the band: precision {tp / (tp + fp):.3f}  recall {tp / g['truth'].sum():.3f}")
    tp_all = int(sure["truth"].sum()) + tp
    fp_all = int((sure["truth"] == 0).sum()) + fp
    print(f"{'':20s} rules + model:   precision {tp_all / (tp_all + fp_all):.3f}  recall {tp_all / N_TRUE:.3f}")"""),
("md", """🔔 **Question**: The stronger model is slightly more precise and finds fewer true pairs. Which error would you rather make, and does it depend on what you will estimate?

<a id='downstream'></a>

# Downstream

The research question: what share of people changed postcode between the two files? Postcode is also a matching field, so linkage errors are correlated with the outcome."""),
("md", """A pair "moved" when the postcode field disagrees. We compute the share who moved under each policy: among all 5,000 true links (including the 13 that blocking never paired, all of them movers), among the sure matches only, among sure matches plus the whole band, and among sure matches plus the pairs the model accepted."""),
("code", """from recordlinkage.datasets import load_febrl4

dfA, dfB, true_links = load_febrl4(return_links=True)
postcode_a = dfA.loc[true_links.get_level_values(0), "postcode"].astype(str).values
postcode_b = dfB.loc[true_links.get_level_values(1), "postcode"].astype(str).values
truth_moved = (postcode_a != postcode_b).mean()

scores["moved"] = (scores["postcode"] == 0).astype(int)
strict = scores.loc[sure.index, "moved"].mean()
lenient = scores.loc[pd.concat([sure, band]).index, "moved"].mean()
g = pairs[(pairs["model"] == "claude-haiku-5-5") & (pairs["pred"] == 1)]
model_moved = pd.concat([scores.loc[sure.index, "moved"],
                         (g["postcode_a"].astype(str) != g["postcode_b"].astype(str)).astype(int)]).mean()
pd.Series({"truth": truth_moved, "rules, strict": strict, "rules, lenient": lenient,
           "rules + model": model_moved}).round(3)"""),
("md", """## 🥊 Notebook Challenge 2: Review a Sample of the Band

In a real project the truth is unknown. Draw 100 pairs from the band at random, pretend a research assistant reviewed them, and estimate the model's precision and recall inside the band from that sample alone. How close are you to the full-band numbers above?

⚠️ **Warning**: a sample of the band tells you about the band. The true links the rules already missed (score 1, below the band) need their own reviewed sample, or your recall will look better than it is."""),
("challenge", """band_model = pairs[pairs["model"] == "claude-haiku-5-5"]
reviewed = band_model.sample(100, random_state=7)
# YOUR CODE HERE
...""", """band_model = pairs[pairs["model"] == "claude-haiku-5-5"]
reviewed = band_model.sample(100, random_state=7)
tp = int(((reviewed["pred"] == 1) & (reviewed["truth"] == 1)).sum())
fp = int(((reviewed["pred"] == 1) & (reviewed["truth"] == 0)).sum())
fn = int(((reviewed["pred"] == 0) & (reviewed["truth"] == 1)).sum())
print(f"from 100 reviewed pairs: precision {tp / (tp + fp):.2f}  recall {tp / (tp + fn):.2f}")"""),
("md", """<a id='keypoints'></a>

<div class="alert alert-success">

### ❗ Key Points

- A link is a label on a pair, with error like any other label.
- Rules handle the easy pairs cheaply; the ambiguous band is where a person or a model has to decide.
- Measure precision and recall on a reviewed random sample of the band, by the groups you compare.
- Missed links and invented links push a downstream estimate in opposite directions; the right policy depends on the number you want to know.
</div>"""),
]

for name, cells in [("1_Labels_as_Measurements", nb1), ("2_Links_as_Measurements", nb2)]:
    nbf.write(build(cells, solution=False), LESSONS / f"{name}.ipynb")
    nbf.write(build(cells, solution=True), SOLUTIONS / f"{name}_Solutions.ipynb")
    print("wrote", name)

# Execute the solution notebooks in place to check they run (outputs kept).
for name in ["1_Labels_as_Measurements", "2_Links_as_Measurements"]:
    cmd = ["python3", "-m", "nbconvert", "--to", "notebook", "--execute", "--inplace",
           "--ExecutePreprocessor.timeout=300", str(SOLUTIONS / f"{name}_Solutions.ipynb")]
    res = subprocess.run(cmd, cwd=str(SOLUTIONS), capture_output=True, text=True)
    print(name, "executed OK" if res.returncode == 0 else "FAILED\n" + res.stderr[-2000:])
