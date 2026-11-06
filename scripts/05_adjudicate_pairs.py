"""
Record linkage, step 2: the LLM adjudicates the ambiguous band.

Each ambiguous pair (score 2 of 4 under the rules in 04_linkage_rules.py)
is shown to the model as two short records. The model answers whether they
refer to the same person. Answers are appended to outputs/llm_pairs.jsonl,
with model and prompt version recorded, and runs can be resumed.

Usage
-----
    python3 scripts/05_adjudicate_pairs.py --models MODEL [MODEL2] [--limit 20]
"""

import argparse
import importlib.util
import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd

PROJECT = Path(__file__).resolve().parents[1]
PAIRS_FILE = PROJECT / "outputs" / "linkage_ambiguous_pairs.csv"
PROMPT_FILE = PROJECT / "prompts" / "pairs_prompt.txt"
OUTPUT_FILE = PROJECT / "outputs" / "llm_pairs.jsonl"

# Reuse the provider plumbing from the title classifier.
spec = importlib.util.spec_from_file_location(
    "clf", PROJECT / "scripts" / "02_classify_titles.py")
clf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(clf)

FIELDS = ["given_name", "surname", "date_of_birth", "postcode"]


def format_record(row, suffix):
    dob = str(row[f"date_of_birth{suffix}"])
    if dob not in ("nan", "") and len(dob) == 8:
        dob = f"{dob[:4]}-{dob[4:6]}-{dob[6:]}"
    return (f"given name: {row[f'given_name{suffix}']}; "
            f"surname: {row[f'surname{suffix}']}; "
            f"date of birth: {dob}; "
            f"postcode: {row[f'postcode{suffix}']}")


def parse_pair_answer(text):
    label, why, raw = "INVALID", "", text
    import re
    match = re.search(r"\{.*\}", text, flags=re.S)
    if match:
        try:
            obj = json.loads(match.group(0))
            ans = str(obj.get("match", "")).strip().lower()
            if ans in ("yes", "no"):
                label = ans
            why = str(obj.get("why", "")).strip()
        except json.JSONDecodeError:
            pass
    if label == "INVALID":
        m2 = re.search(r'"match"\s*:\s*"(yes|no)"', text, flags=re.I)
        if m2:
            label = m2.group(1).lower()
    return label, why, raw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--models", nargs="+", required=True)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()

    key, provider = clf.read_key()
    print(f"Provider: {provider}")
    pairs = pd.read_csv(PAIRS_FILE, dtype=str)
    if args.limit:
        pairs = pairs.head(args.limit)
    template = PROMPT_FILE.read_text()

    done = set()
    if OUTPUT_FILE.exists():
        for line in OUTPUT_FILE.open():
            rec = json.loads(line)
            if rec["match"] == "INVALID":
                continue
            done.add((rec["model"], rec["pair_id"]))
    jobs = [(model, row) for model in args.models
            for _, row in pairs.iterrows()
            if (model, row["pair_id"]) not in done]
    print(f"{len(done)} answers on disk, {len(jobs)} to request.")
    if not jobs:
        return

    def work(job):
        model, row = job
        prompt = (template.replace("{record_a}", format_record(row, "_a"))
                          .replace("{record_b}", format_record(row, "_b")))
        text, usage = clf.call_with_retries(key, provider, model, prompt)
        label, why, raw = parse_pair_answer(text)
        return {"pair_id": row["pair_id"], "model": model,
                "prompt": "pairs_prompt", "match": label, "why": why,
                "raw": raw, "input_tokens": usage.get("input_tokens"),
                "output_tokens": usage.get("output_tokens"),
                "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S")}

    started = time.time()
    n_fail = 0
    with OUTPUT_FILE.open("a") as out, \
            ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(work, job): job for job in jobs}
        for i, fut in enumerate(as_completed(futures), 1):
            try:
                out.write(json.dumps(fut.result()) + "\n")
                out.flush()
            except Exception as err:  # noqa: BLE001
                n_fail += 1
                print(f"FAILED {futures[fut][1]['pair_id']}: {err}")
            if i % 100 == 0 or i == len(jobs):
                print(f"{i}/{len(jobs)} done, {n_fail} failed, "
                      f"{(time.time()-started)/60:.1f} min")
    print("Finished.")


if __name__ == "__main__":
    main()
