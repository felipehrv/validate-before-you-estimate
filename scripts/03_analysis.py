"""
Analysis behind the lesson: agreement, differential error, three estimators,
sample-size curve, and the linkage policies. Produces every number the slides
quote and every figure they show.

Run:  python3 scripts/03_analysis.py
Writes figures/*.png and outputs/results.json
"""

import json
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT / "scripts"))
from lesson_tools import (accuracy, cohen_kappa, confusion, estimate_corrected,
                          estimate_human_only, estimate_naive, per_class,
                          sample_size_curve, wilson_interval)

FIG = PROJECT / "figures"
FIG.mkdir(exist_ok=True)
LABELS = ["Health", "Education", "Environment", "Immigration",
          "Law and Crime", "Social Welfare", "Other"]
PROMPTS = {"prompt_1_names": "P1: category names",
           "prompt_2_codebook": "P2: codebook",
           "prompt_3_codebook_examples": "P3: codebook + examples"}
CHEAP, STRONG = "claude-haiku-5-5", "claude-sonnet-5-5"
N_HAND = 200          # the hand-coded validation sample in the lesson
SEED = 11
ERAS = [(1973, 1986), (1987, 2000), (2001, 2016)]

# ---------------------------------------------------------------------------
# Palette and chart style (validated categorical slots; thin marks; recessive grid)
# ---------------------------------------------------------------------------
C = {"blue": "#2a78d6", "orange": "#eb6834", "aqua": "#1baf7a", "yellow": "#eda100",
     "magenta": "#e87ba4", "green": "#008300", "violet": "#6250d6", "red": "#e34948",
     "ink": "#0b0b0b", "ink2": "#52514e", "muted": "#898781", "grid": "#e1e0d9",
     "axis": "#c3c2b7", "surface": "#fcfcfb"}
SEQ = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
plt.rcParams.update({
    "font.family": "sans-serif", "font.sans-serif": ["DejaVu Sans", "Arial"],
    "font.size": 12, "axes.edgecolor": C["axis"], "axes.linewidth": 1,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": C["grid"], "grid.linewidth": 1,
    "axes.axisbelow": True, "xtick.color": C["ink2"], "ytick.color": C["ink2"],
    "axes.labelcolor": C["ink2"], "axes.titlecolor": C["ink"],
    "axes.titleweight": "bold", "axes.titlesize": 14, "axes.titlelocation": "left",
    "figure.facecolor": C["surface"], "axes.facecolor": C["surface"],
    "savefig.dpi": 200, "savefig.bbox": "tight", "legend.frameon": False,
})

def save(fig, name):
    fig.savefig(FIG / f"{name}.png")
    plt.close(fig)
    print(f"  wrote figures/{name}.png")

def era_of(year):
    for lo, hi in ERAS:
        if lo <= year <= hi:
            return f"{lo}-{hi}"
    return None

# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
gold = pd.read_csv(PROJECT / "data/processed/bills_sample.csv")
gold["era"] = gold["year"].apply(era_of)
gold["party_name"] = gold["party"].map({100: "Democrat", 200: "Republican"})
labels = pd.read_csv(PROJECT / "outputs/llm_labels_clean.csv")
labels = labels.drop_duplicates(subset=["model", "prompt", "bill_id"], keep="last")
d = labels.merge(gold, on="bill_id")
d["ok"] = d["label"] == d["label_gold"]
R = {}  # results for the slides

# ---------------------------------------------------------------------------
# 1. Sample composition
# ---------------------------------------------------------------------------
comp = gold["label_gold"].value_counts().reindex(LABELS)
R["n_titles"] = int(len(gold))
R["composition"] = comp.to_dict()
fig, ax = plt.subplots(figsize=(8, 4.2))
y = np.arange(len(LABELS))[::-1]
ax.barh(y, comp.values, height=0.55, color=C["blue"])
ax.set_yticks(y, LABELS)
for yi, v in zip(y, comp.values):
    ax.text(v + 8, yi, f"{v}  ({v/len(gold):.0%})", va="center", color=C["ink2"])
ax.set_xlim(0, comp.max() * 1.22)
ax.set_title(f"The sample: {len(gold):,} bill titles, coded by trained humans")
ax.set_xlabel("Titles (75 per Congress, 93rd to 114th, 1973 to 2016)")
ax.grid(axis="y", visible=False)
save(fig, "f01_sample_composition")

# ---------------------------------------------------------------------------
# 2. The prompt showdown: accuracy and kappa by instrument
# ---------------------------------------------------------------------------
rows = []
for (model, prompt), g in d.groupby(["model", "prompt"]):
    rows.append({"model": model, "prompt": prompt,
                 "accuracy": accuracy(g.label_gold, g.label),
                 "kappa": cohen_kappa(g.label_gold, g.label, LABELS), "n": len(g)})
show = pd.DataFrame(rows)
R["showdown"] = show.to_dict("records")
order = [(CHEAP, "prompt_1_names"), (CHEAP, "prompt_2_codebook"),
         (CHEAP, "prompt_3_codebook_examples"), (STRONG, "prompt_3_codebook_examples")]
names = ["Haiku\nP1 names", "Haiku\nP2 codebook", "Haiku\nP3 codebook\n+ examples",
         "Sonnet\nP3 codebook\n+ examples"]
acc = [show[(show.model == m) & (show.prompt == p)].accuracy.iloc[0] for m, p in order]
kap = [show[(show.model == m) & (show.prompt == p)].kappa.iloc[0] for m, p in order]
fig, axes = plt.subplots(1, 2, figsize=(11, 4.4))
for ax, vals, title, fmt in [(axes[0], acc, "Accuracy against human codes", "{:.1%}"),
                             (axes[1], kap, "Cohen's kappa (agreement beyond chance)", "{:.2f}")]:
    cols = [C["blue"]] * 3 + [C["orange"]]
    ax.bar(range(4), vals, width=0.55, color=cols)
    ax.set_xticks(range(4), names)
    for i, v in enumerate(vals):
        ax.text(i, v + 0.012, fmt.format(v), ha="center", color=C["ink"], fontweight="bold")
    ax.set_ylim(0, 1.0 if title.startswith("Acc") else 0.9)
    ax.set_title(title)
    ax.grid(axis="x", visible=False)
fig.suptitle("Same 1,650 titles, four instruments", x=0.01, ha="left",
             fontsize=15, fontweight="bold", color=C["ink"])
fig.tight_layout()
save(fig, "f02_prompt_showdown")

# ---------------------------------------------------------------------------
# 3. Confusion matrix and per-class metrics (best cheap instrument)
# ---------------------------------------------------------------------------
best = d[(d.model == CHEAP) & (d.prompt == "prompt_3_codebook_examples")].copy()
cm = confusion(best.label_gold, best.label, LABELS)
R["confusion_haiku_p3"] = cm.to_dict()
pc = per_class(best.label_gold, best.label, LABELS)
R["per_class_haiku_p3"] = pc.round(3).to_dict("records")
rownorm = cm.div(cm.sum(axis=1), axis=0)
fig, ax = plt.subplots(figsize=(8.6, 6.6))
im = ax.imshow(rownorm.values, cmap=matplotlib.colors.LinearSegmentedColormap.from_list("seq", SEQ),
               vmin=0, vmax=1)
ax.set_xticks(range(7), LABELS, rotation=30, ha="right", fontsize=10.5)
ax.set_yticks(range(7), LABELS, fontsize=11)
ax.set_xlabel("Model label (Haiku, P3)")
ax.set_ylabel("Human label")
for i in range(7):
    for j in range(7):
        v = cm.values[i, j]
        if v:
            ax.text(j, i, f"{v}", ha="center", va="center",
                    color="white" if rownorm.values[i, j] > 0.55 else C["ink"], fontsize=11)
ax.grid(False)
ax.set_title("Where the model and the humans disagree (row = human label)", pad=14)
cb = fig.colorbar(im, ax=ax, fraction=0.04, pad=0.02)
cb.set_label("Share of the human row")
save(fig, "f03_confusion_heatmap")

fig, ax = plt.subplots(figsize=(8, 4.4))
x = np.arange(len(LABELS))
ax.bar(x - 0.18, pc.precision, width=0.34, color=C["blue"], label="Precision (model says X: how often right)")
ax.bar(x + 0.18, pc.recall, width=0.34, color=C["orange"], label="Recall (human says X: how often found)")
ax.set_xticks(x, [f"{l.replace(' ', chr(10), 1) if l.startswith(('Law', 'Social')) else l}\n(n={n})" for l, n in zip(LABELS, pc.n_human)], fontsize=9)
ax.set_ylim(0, 1.32)
ax.set_yticks([0, 0.2, 0.4, 0.6, 0.8, 1.0])
ax.set_title("Accuracy is one number; the categories are seven stories")
ax.legend(loc="upper left", fontsize=10, ncol=2)
ax.grid(axis="x", visible=False)
save(fig, "f04_precision_recall_by_class")

# ---------------------------------------------------------------------------
# 4. Differential error: by party and by era
# ---------------------------------------------------------------------------
best["is_health_gold"] = best.label_gold == "Health"
best["is_health_llm"] = best.label == "Health"
bp = best[best.party_name.notna()]
by_party = bp.groupby("party_name").agg(accuracy=("ok", "mean"),
                                        health_gold=("is_health_gold", "mean"),
                                        health_llm=("is_health_llm", "mean"),
                                        n=("ok", "size"))
by_era = best.groupby("era").agg(accuracy=("ok", "mean"),
                                 health_gold=("is_health_gold", "mean"),
                                 health_llm=("is_health_llm", "mean"), n=("ok", "size"))
R["by_party_haiku_p3"] = by_party.round(4).to_dict("index")
R["by_era_haiku_p3"] = by_era.round(4).to_dict("index")
# the same for the strong model, to show it does not vanish
s = d[(d.model == STRONG)].copy()
s["is_health_gold"] = s.label_gold == "Health"; s["is_health_llm"] = s.label == "Health"
R["by_party_sonnet_p3"] = (s[s.party_name.notna()].groupby("party_name")
                           .agg(accuracy=("ok", "mean"), health_gold=("is_health_gold", "mean"),
                                health_llm=("is_health_llm", "mean")).round(4).to_dict("index"))

fig, axes = plt.subplots(1, 2, figsize=(11, 4.4))
ax = axes[0]
groups = ["Democrat", "Republican"]
xg = np.arange(2)
ax.bar(xg - 0.18, [by_party.loc[g, "health_gold"] for g in groups], width=0.34,
       color=C["ink2"], label="Human codes")
ax.bar(xg + 0.18, [by_party.loc[g, "health_llm"] for g in groups], width=0.34,
       color=C["blue"], label="Model labels (Haiku, P3)")
for i, g in enumerate(groups):
    ax.text(i - 0.18, by_party.loc[g, "health_gold"] + 0.004, f"{by_party.loc[g, 'health_gold']:.1%}", ha="center", fontsize=10)
    ax.text(i + 0.18, by_party.loc[g, "health_llm"] + 0.004, f"{by_party.loc[g, 'health_llm']:.1%}", ha="center", fontsize=10)
ax.set_xticks(xg, [f"{g}\n(n={int(by_party.loc[g, 'n'])})" for g in groups])
ax.set_ylim(0, 0.19)
ax.set_yticks([0, 0.04, 0.08, 0.12, 0.16])
ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_title("By sponsor's party")
ax.legend(fontsize=10, loc="upper right")
ax.grid(axis="x", visible=False)
ax = axes[1]
eras = [f"{lo}-{hi}" for lo, hi in ERAS]
xe = np.arange(3)
ax.bar(xe - 0.18, [by_era.loc[e, "health_gold"] for e in eras], width=0.34, color=C["ink2"], label="Human codes")
ax.bar(xe + 0.18, [by_era.loc[e, "health_llm"] for e in eras], width=0.34, color=C["blue"], label="Model labels (Haiku, P3)")
for i, e in enumerate(eras):
    ax.text(i - 0.18, by_era.loc[e, "health_gold"] + 0.004, f"{by_era.loc[e, 'health_gold']:.1%}", ha="center", fontsize=10)
    ax.text(i + 0.18, by_era.loc[e, "health_llm"] + 0.004, f"{by_era.loc[e, 'health_llm']:.1%}", ha="center", fontsize=10)
ax.set_xticks(xe, eras)
ax.set_ylim(0, 0.19)
ax.set_yticks([0, 0.04, 0.08, 0.12, 0.16])
ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_title("By era")
ax.legend(fontsize=10, loc="upper left")
ax.grid(axis="x", visible=False)
fig.suptitle("Share of bills about health: the model's error differs by group",
             x=0.01, ha="left", fontsize=14, fontweight="bold", color=C["ink"])
fig.tight_layout()
save(fig, "f05_differential_error")

# ---------------------------------------------------------------------------
# 5. Three estimators (naive, human-only, corrected) of the health share,
#    overall, by party and by era, with a hand-coded validation sample of N_HAND
# ---------------------------------------------------------------------------
rng = np.random.default_rng(SEED)
hand_idx = rng.choice(best.index, size=N_HAND, replace=False)
best["hand_coded"] = best.index.isin(hand_idx)
R["n_hand"] = N_HAND

def three(df):
    f_all = df.is_health_llm.values.astype(float)
    sub = df[df.hand_coded]
    f_sub = sub.is_health_llm.values.astype(float)
    y_sub = sub.is_health_gold.values.astype(float)
    out = {"truth": float(df.is_health_gold.mean()), "n_all": int(len(df)), "n_hand": int(len(sub))}
    for name, (est, se) in {"naive": estimate_naive(f_all),
                            "human": estimate_human_only(y_sub),
                            "corrected": estimate_corrected(f_all, f_sub, y_sub)}.items():
        out[name] = {"est": float(est), "se": float(se),
                     "ci_low": float(est - 1.96 * se), "ci_high": float(est + 1.96 * se)}
    # Wilson interval for the hand-coded-only share: the textbook interval fails near zero
    lo, hi = wilson_interval(int(y_sub.sum()), len(y_sub))
    out["human"]["ci_low"], out["human"]["ci_high"] = lo, hi
    out["human"]["interval"] = "wilson"
    return out

est = {"overall": three(best)}
for g in groups:
    est[g] = three(best[best.party_name == g])
for e in eras:
    est[e] = three(best[best.era == e])
R["estimates"] = est

def plot_three(ax, keys, title, xlabels):
    xs = np.arange(len(keys))
    w = 0.22
    specs = [("naive", -w, C["blue"], "Model labels, all titles (naive)"),
             ("human", 0, C["orange"], f"Hand-coded subset only (n={N_HAND} overall)"),
             ("corrected", w, C["aqua"], "Model labels corrected by the subset (PPI)")]
    for key, off, col, lab in specs:
        vals = np.array([est[k][key]["est"] for k in keys])
        lo = np.array([est[k][key]["ci_low"] for k in keys])
        hi = np.array([est[k][key]["ci_high"] for k in keys])
        ax.errorbar(xs + off, vals, yerr=[vals - lo, hi - vals], fmt="o", color=col,
                    ecolor=col, elinewidth=2, capsize=4, markersize=8, label=lab)
    truth = [est[k]["truth"] for k in keys]
    ax.scatter(xs, truth, marker="_", s=900, color=C["ink"], linewidths=2.5, zorder=5,
               label="All titles hand-coded (the truth we usually cannot afford)")
    ax.set_xticks(xs, xlabels)
    ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
    ax.set_title(title)
    ax.grid(axis="x", visible=False)

fig, ax = plt.subplots(figsize=(10, 5))
plot_three(ax, ["overall", "Democrat", "Republican"], "Share of bills about health: three ways to estimate it",
           ["All bills", "Democratic sponsors", "Republican sponsors"])
ax.legend(fontsize=9.5, loc="upper right")
ax.set_ylim(0.02, 0.19)
ax.set_yticks([0.04, 0.08, 0.12, 0.16])
save(fig, "f06_three_estimators_party")

fig, ax = plt.subplots(figsize=(10, 5))
plot_three(ax, eras, "Health bills over time: 63 to 72 hand-coded titles per era is not enough", eras)
ax.legend(fontsize=9.5, loc="lower center", bbox_to_anchor=(0.5, -0.30), ncol=2)
ax.set_ylim(0, 0.27)
ax.set_yticks([0, 0.05, 0.10, 0.15, 0.20, 0.25])
save(fig, "f07_three_estimators_era")

# The opening chart of the lesson: the naive series alone, as a student would plot it
fig, ax = plt.subplots(figsize=(8.5, 4.6))
naive_series = [est[e]["naive"]["est"] for e in eras]
ax.plot(range(3), naive_series, marker="o", markersize=11, lw=3, color=C["blue"])
for i, v in enumerate(naive_series):
    ax.text(i, v + 0.008, f"{v:.1%}", ha="center", fontsize=13, fontweight="bold", color=C["ink"])
ax.set_xticks(range(3), eras)
ax.set_xlim(-0.4, 2.4)
ax.set_ylim(0, 0.17)
ax.set_yticks([0, 0.04, 0.08, 0.12, 0.16])
ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_title("Health kept climbing on Congress's agenda, era after era")
ax.set_xlabel("Source: 1,650 bill titles, 93rd to 114th Congress, coded by an LLM (88% accuracy on a validation sample)",
              fontsize=10, color=C["muted"], loc="left")
ax.grid(axis="x", visible=False)
save(fig, "f00_hook_naive_series")

# ---------------------------------------------------------------------------
# 6. How many titles to hand-code? Interval width against n
# ---------------------------------------------------------------------------
sizes = [25, 50, 100, 200, 400, 800]
curve = sample_size_curve(best.is_health_llm.values.astype(float),
                          best.is_health_llm.values.astype(float),
                          best.is_health_gold.values.astype(float), sizes, reps=300, seed=SEED)
# human-only width for comparison
curve["human_only_halfwidth"] = [1.96 * np.sqrt(best.is_health_gold.mean() * (1 - best.is_health_gold.mean()) / n) for n in sizes]
R["sample_size_curve"] = curve.round(4).to_dict("records")

# Challenge 3 uses a projection: the same instrument on a group of N_GROUP titles.
# SE^2 = var(model flag) / N + var(model flag - human flag) / n, both variances
# measured on the 1,650 titles. In a real project they come from a pilot sample.
N_GROUP = 10_000
var_f = float(best.is_health_llm.astype(float).var(ddof=1))
var_r = float((best.is_health_llm.astype(float) - best.is_health_gold.astype(float)).var(ddof=1))
p_gold = float(best.is_health_gold.mean())
TARGET = 0.02   # plus or minus 2 points per group: enough to see a 3-point gap between two groups
sizes_c3 = [50, 100, 200, 300, 400, 800]
def hw_corrected(n): return 1.96 * np.sqrt(var_f / N_GROUP + var_r / n)
def hw_human(n): return 1.96 * np.sqrt(p_gold * (1 - p_gold) / n)
n_target_corr = var_r / ((TARGET / 1.96) ** 2 - var_f / N_GROUP)
n_target_human = p_gold * (1 - p_gold) / (TARGET / 1.96) ** 2
R["sample_size_curve_projected"] = {
    "n_group": N_GROUP, "var_model_flag": round(var_f, 5), "var_rectifier": round(var_r, 5),
    "target_halfwidth": TARGET, "n_for_target_corrected": int(round(n_target_corr)),
    "n_for_target_human_only": int(round(n_target_human)),
    "rows": [{"n_hand_coded": n, "ci_halfwidth": round(float(hw_corrected(n)), 4),
              "human_only_halfwidth": round(float(hw_human(n)), 4)} for n in sizes_c3]}
grid = np.linspace(30, 1000, 300)
fig, ax = plt.subplots(figsize=(8.5, 4.6))
ax.plot(grid, hw_human(grid), lw=2, color=C["orange"], label="Hand-coded sample only")
ax.plot(grid, hw_corrected(grid), lw=2, color=C["aqua"], label="Model labels corrected by the sample (PPI)")
ax.scatter(sizes_c3, [hw_human(n) for n in sizes_c3], s=50, color=C["orange"], zorder=4)
ax.scatter(sizes_c3, [hw_corrected(n) for n in sizes_c3], s=50, color=C["aqua"], zorder=4)
ax.axhline(TARGET, ls="--", lw=1.5, color=C["ink2"])
ax.text(32, TARGET - 0.0075, "Plus or minus 2 points: enough to see\na 3-point gap between two groups",
        fontsize=9.5, color=C["ink2"], va="center")
for n_t, col, dy in [(n_target_corr, C["aqua"], 0.0105), (n_target_human, C["orange"], 0.0035)]:
    ax.plot([n_t, n_t], [0, TARGET], ls=":", lw=1.5, color=col)
    ax.text(n_t * 1.04, dy, f"n = {n_t:.0f}", ha="left", fontsize=10, color=col, fontweight="bold")
ax.set_xscale("log"); ax.set_xticks(sizes_c3, [str(s) for s in sizes_c3])
ax.xaxis.set_minor_formatter(matplotlib.ticker.NullFormatter())
ax.set_xlim(30, 1000); ax.set_ylim(0, 0.09)
ax.set_yticks([0, 0.02, 0.04, 0.06, 0.08])
ax.yaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_xlabel("Titles coded by hand in the group (this lesson's error rates, projected to a group of 10,000 titles)")
ax.set_ylabel("95% interval, plus or minus")
ax.set_title("The budget question: how many titles do you have to code by hand?")
ax.legend(fontsize=10, loc="upper right")
save(fig, "f08_sample_size_curve")

# ---------------------------------------------------------------------------
# 7. Record linkage: the ambiguous band, three policies, and the downstream rate
# ---------------------------------------------------------------------------
scores = pd.read_csv(PROJECT / "outputs/linkage_scores.csv")
pairs = pd.read_csv(PROJECT / "outputs/llm_pairs_clean.csv").drop_duplicates(subset=["model", "pair_id"], keep="last")
amb = pd.read_csv(PROJECT / "outputs/linkage_ambiguous_pairs.csv", dtype={"pair_id": str})
amb["truth"] = amb["truth"].astype(int)
pairs = pairs.merge(amb[["pair_id", "truth", "postcode_a", "postcode_b"]], on="pair_id")
pairs["pred"] = (pairs["match"] == "yes").astype(int)
n_true = 5000
sure = scores[scores.zone == "match"]
amb_s = scores[scores.zone == "ambiguous"]
def pr(tp, fp, fn):
    return tp / (tp + fp), tp / (tp + fn)
tp_s, fp_s = int(sure.truth.sum()), int((sure.truth == 0).sum())
policies = {}
policies["Rules, strict (drop the band)"] = pr(tp_s, fp_s, n_true - tp_s)
policies["Rules, lenient (keep the band)"] = pr(tp_s + int(amb_s.truth.sum()), fp_s + int((amb_s.truth == 0).sum()), n_true - tp_s - int(amb_s.truth.sum()))
for model in [CHEAP, STRONG]:
    g = pairs[pairs.model == model]
    tp = int(((g.pred == 1) & (g.truth == 1)).sum()); fp = int(((g.pred == 1) & (g.truth == 0)).sum())
    policies[f"Rules + {model.split('-')[1].title()} on the band"] = pr(tp_s + tp, fp_s + fp, n_true - tp_s - tp)
    R[f"band_{model}"] = {"precision_in_band": tp / max(tp + fp, 1),
                          "recall_in_band": tp / int(g.truth.sum()), "n_band": int(len(g))}
R["linkage_policies"] = {k: {"precision": v[0], "recall": v[1]} for k, v in policies.items()}
R["blocking"] = {"candidates": int(len(scores)), "true_in_candidates": int(scores.truth.sum()),
                 "sure": int(len(sure)), "band": int(len(amb_s)), "band_true": int(amb_s.truth.sum())}

fig, ax = plt.subplots(figsize=(8.5, 4.4))
counts = scores.groupby(["score", "truth"]).size().unstack(fill_value=0).reindex(range(5), fill_value=0)
xs = np.arange(5)
ax.bar(xs - 0.2, counts[0].clip(lower=0.5), width=0.38, color=C["ink2"], label="Different people")
ax.bar(xs + 0.2, counts[1].clip(lower=0.5), width=0.38, color=C["blue"], label="Same person (true link)")
for i in xs:
    ax.text(i - 0.2, counts[0][i] * 1.25 + 0.5, f"{counts[0][i]:,}", ha="center", fontsize=9, color=C["ink2"])
    ax.text(i + 0.2, counts[1][i] * 1.25 + 0.5, f"{counts[1][i]:,}", ha="center", fontsize=9, color=C["blue"])
ax.set_yscale("log")
ax.set_ylim(0.5, 6e5)
ax.set_xticks(range(5), ["0", "1", "2\n(ambiguous)", "3", "4"])
ax.set_xlabel("Fields in agreement (given name, surname, date of birth, postcode)")
ax.set_ylabel("Candidate pairs (log scale)")
ax.set_title("Rules sort most pairs; the band at 2 of 4 is where the work is")
ax.legend(fontsize=10, loc="upper right")
ax.grid(axis="x", visible=False)
save(fig, "f09_linkage_score_band")

fig, ax = plt.subplots(figsize=(9, 4.4))
names_p = list(policies.keys())
y = np.arange(len(names_p))[::-1]
ax.barh(y + 0.18, [policies[k][0] for k in names_p], height=0.34, color=C["blue"], label="Precision")
ax.barh(y - 0.18, [policies[k][1] for k in names_p], height=0.34, color=C["orange"], label="Recall")
for yi, k in zip(y, names_p):
    ax.text(policies[k][0] + 0.005, yi + 0.18, f"{policies[k][0]:.1%}", va="center", fontsize=10)
    ax.text(policies[k][1] + 0.005, yi - 0.18, f"{policies[k][1]:.1%}", va="center", fontsize=10)
ax.set_yticks(y, names_p)
ax.set_xlim(0.8, 1.02)
ax.set_xticks([0.80, 0.85, 0.90, 0.95, 1.00])
ax.xaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_title("Four ways to handle the band, measured against the truth")
ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=10)
ax.grid(axis="y", visible=False)
save(fig, "f10_linkage_policies")

# downstream: share who changed postcode, under each policy
sc = scores.copy()
sc["moved"] = (sc.postcode == 0).astype(int)
# The truth over all 5,000 true links, including the 13 that blocking never paired
from recordlinkage.datasets import load_febrl4
dfA, dfB, true_links = load_febrl4(return_links=True)
pa = dfA.loc[true_links.get_level_values(0), "postcode"].astype(str).values
pb = dfB.loc[true_links.get_level_values(1), "postcode"].astype(str).values
moved_all = (pa != pb).astype(int)
truth_moved = float(moved_all.mean())
R["moved_truth_candidates_only"] = float(sc[sc.truth == 1].moved.mean())
in_cand = pd.MultiIndex.from_frame(scores[["rec_id_1", "rec_id_2"]])
lost_mask = ~true_links.isin(in_cand)
R["blocking"]["lost_links"] = int(lost_mask.sum())
R["blocking"]["lost_links_moved_share"] = float(moved_all[lost_mask].mean())
moved = {"Truth (all true links)": truth_moved}
moved["Rules, strict"] = float(sure.assign(moved=(sure.postcode == 0)).moved.mean())
moved["Rules, lenient"] = float(pd.concat([sure, amb_s]).assign(moved=lambda x: x.postcode == 0).moved.mean())
for model in [CHEAP, STRONG]:
    g = pairs[(pairs.model == model) & (pairs.pred == 1)]
    g_moved = (g.postcode_a.astype(str) != g.postcode_b.astype(str)).astype(int)
    all_moved = pd.concat([pd.Series((sure.postcode == 0).astype(int).values), g_moved])
    moved[f"Rules + {model.split('-')[1].title()}"] = float(all_moved.mean())
R["moved_rate"] = moved
fig, ax = plt.subplots(figsize=(8.5, 4.2))
ks = list(moved.keys())
y = np.arange(len(ks))[::-1]
cols = [C["ink"]] + [C["blue"]] * (len(ks) - 1)
ax.barh(y, [moved[k] for k in ks], height=0.5, color=cols)
for yi, k in zip(y, ks):
    ax.text(moved[k] + 0.004, yi, f"{moved[k]:.1%}", va="center", fontsize=10)
ax.set_yticks(y, ks)
ax.set_xlim(0, 0.3)
ax.xaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0, decimals=0))
ax.set_title("A downstream estimate: share of people whose postcode changed")
ax.grid(axis="y", visible=False)
save(fig, "f11_linkage_downstream")

# ---------------------------------------------------------------------------
# 8. Token and cost accounting
# ---------------------------------------------------------------------------
tok = labels.groupby(["model", "prompt"])[["input_tokens", "output_tokens"]].sum()
R["tokens_titles"] = tok.astype(int).to_dict("index") if False else {f"{m}|{p}": {"in": int(r.input_tokens), "out": int(r.output_tokens)} for (m, p), r in tok.iterrows()}
ptok = pairs.groupby("model")[["input_tokens", "output_tokens"]].sum()
R["tokens_pairs"] = {m: {"in": int(r.input_tokens), "out": int(r.output_tokens)} for m, r in ptok.iterrows()}

with open(PROJECT / "outputs/results.json", "w") as fh:
    json.dump(R, fh, indent=2, default=float)
print("results.json written")

# Print the headline numbers
print("\nHEADLINES")
for r in R["showdown"]:
    print(f"  {r['model']:20s} {r['prompt']:28s} acc {r['accuracy']:.3f} kappa {r['kappa']:.3f}")
e = R["estimates"]
for k in ["overall", "Democrat", "Republican"] + eras:
    print(f"  {k:14s} truth {e[k]['truth']:.3f} | naive {e[k]['naive']['est']:.3f}±{1.96*e[k]['naive']['se']:.3f} "
          f"| human {e[k]['human']['est']:.3f}±{1.96*e[k]['human']['se']:.3f} (n={e[k]['n_hand']}) "
          f"| corrected {e[k]['corrected']['est']:.3f}±{1.96*e[k]['corrected']['se']:.3f}")
for k, v in R["linkage_policies"].items():
    print(f"  {k:36s} precision {v['precision']:.3f} recall {v['recall']:.3f}")
for k, v in R["moved_rate"].items():
    print(f"  moved: {k:28s} {v:.3f}")
