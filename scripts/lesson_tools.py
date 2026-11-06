"""
Small toolbox for the lesson: agreement metrics, the three estimators, and
the sample-size curve. Written to be read by graduate students, so every
function does one thing and the formulas are spelled out.
"""

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# Agreement between the model and the human coders
# ---------------------------------------------------------------------------

def confusion(human, model, labels):
    """Rows: human (gold) label. Columns: model label."""
    return pd.crosstab(pd.Categorical(human, categories=labels),
                       pd.Categorical(model, categories=labels),
                       rownames=["human"], colnames=["model"], dropna=False)


def accuracy(human, model):
    return float(np.mean(np.asarray(human) == np.asarray(model)))


def cohen_kappa(human, model, labels):
    """Agreement beyond chance: (p_o - p_e) / (1 - p_e)."""
    cm = confusion(human, model, labels).to_numpy().astype(float)
    n = cm.sum()
    p_o = np.trace(cm) / n
    p_e = (cm.sum(axis=1) @ cm.sum(axis=0)) / n ** 2
    return float((p_o - p_e) / (1 - p_e))


def per_class(human, model, labels):
    """Precision, recall and support for every label."""
    cm = confusion(human, model, labels)
    rows = []
    for lab in labels:
        tp = cm.loc[lab, lab]
        predicted = cm[lab].sum()
        actual = cm.loc[lab].sum()
        rows.append({"label": lab,
                     "precision": tp / predicted if predicted else np.nan,
                     "recall": tp / actual if actual else np.nan,
                     "n_human": int(actual), "n_model": int(predicted)})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Three estimators of a proportion (e.g. the share of bills about health)
# ---------------------------------------------------------------------------

def estimate_naive(model_flag):
    """Use the model's labels on every document as if they were true."""
    y = np.asarray(model_flag, dtype=float)
    return y.mean(), y.std(ddof=1) / np.sqrt(len(y))


def estimate_human_only(human_flag_subset):
    """Use only the hand-coded subset. Unbiased, but small n."""
    y = np.asarray(human_flag_subset, dtype=float)
    return y.mean(), y.std(ddof=1) / np.sqrt(len(y))


def estimate_corrected(model_flag_all, model_flag_subset, human_flag_subset):
    """
    Prediction-powered inference for a mean (Angelopoulos et al., 2023).

    The model's average over all N documents is biased by its errors. On the
    n hand-coded documents we observe those errors directly, so we measure the
    average error (model minus human) and subtract it:

        theta_hat = mean_N(model) - mean_n(model - human)

    Standard error combines both pieces: the model mean over N and the
    error mean over n are independent estimates.
    """
    f_all = np.asarray(model_flag_all, dtype=float)
    f_sub = np.asarray(model_flag_subset, dtype=float)
    y_sub = np.asarray(human_flag_subset, dtype=float)
    rectifier = f_sub - y_sub
    theta = f_all.mean() - rectifier.mean()
    se = np.sqrt(f_all.var(ddof=1) / len(f_all)
                 + rectifier.var(ddof=1) / len(rectifier))
    return theta, se


def wilson_interval(successes, n, z=1.96):
    """
    Interval for a share from a small sample. The textbook (Wald) interval
    fails when the share is near 0 or 1: with 1 success in 63 it dips below
    zero and runs too narrow. Wilson's interval does not.
    """
    if n == 0:
        return (np.nan, np.nan)
    p = successes / n
    denom = 1 + z ** 2 / n
    center = (p + z ** 2 / (2 * n)) / denom
    half = z * np.sqrt(p * (1 - p) / n + z ** 2 / (4 * n ** 2)) / denom
    return (float(center - half), float(center + half))


def sample_size_curve(model_flag_all, model_flag_subset, human_flag_subset,
                      sizes, reps=200, seed=0):
    """
    How wide is the corrected interval for different numbers of hand-coded
    documents? Resample the hand-coded subset at each size.
    """
    rng = np.random.default_rng(seed)
    f_sub = np.asarray(model_flag_subset, dtype=float)
    y_sub = np.asarray(human_flag_subset, dtype=float)
    rows = []
    for n in sizes:
        ses = []
        for _ in range(reps):
            idx = rng.choice(len(f_sub), size=n, replace=True)
            _, se = estimate_corrected(model_flag_all, f_sub[idx], y_sub[idx])
            ses.append(se)
        rows.append({"n_hand_coded": n, "se": float(np.mean(ses)),
                     "ci_halfwidth": 1.96 * float(np.mean(ses))})
    return pd.DataFrame(rows)
