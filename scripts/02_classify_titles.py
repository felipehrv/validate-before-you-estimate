"""
Classify congressional bill titles with an LLM, under a fixed codebook.

Design, in brief
----------------
Each prompt is treated as a version of a measurement instrument. The same
titles are sent to every (model, prompt) pair, with the
model name and prompt version recorded next to every label. Outputs are
appended to a JSONL file so a run can be interrupted and resumed without
re-paying for labels already obtained.

Usage
-----
    python3 scripts/02_classify_titles.py --list-models
    python3 scripts/02_classify_titles.py --models MODEL [MODEL2] --limit 20
    python3 scripts/02_classify_titles.py --models MODEL [MODEL2]

The API key is read from secrets/api_key.txt (one line). The provider is
inferred from the key prefix: sk-ant- (Anthropic), sk-or- (OpenRouter),
anything else starting with sk- (OpenAI).
"""

import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests

PROJECT = Path(__file__).resolve().parents[1]
SAMPLE_FILE = PROJECT / "data" / "processed" / "bills_sample.csv"
PROMPT_DIR = PROJECT / "prompts"
OUTPUT_FILE = PROJECT / "outputs" / "llm_labels.jsonl"
KEY_FILE = PROJECT / "secrets" / "api_key.txt"

# The seven labels the codebook allows. Anything else is recorded as INVALID,
# which is itself a failure mode worth counting.
LABELS = ["Health", "Education", "Environment", "Immigration",
          "Law and Crime", "Social Welfare", "Other"]

SYSTEM_MESSAGE = (
    "You are a careful research assistant coding U.S. congressional bill "
    "titles into policy topics. Follow the codebook exactly. Answer only "
    "with the JSON object requested."
)


# ---------------------------------------------------------------------------
# Provider plumbing
# ---------------------------------------------------------------------------

def read_key():
    key = KEY_FILE.read_text().strip().splitlines()[0].strip()
    if key.startswith("sk-ant-"):
        provider = "anthropic"
    elif key.startswith("sk-or-"):
        provider = "openrouter"
    elif key.startswith("sk-"):
        provider = "openai"
    else:
        sys.exit("Could not infer the provider from the key prefix.")
    return key, provider


def list_models(key, provider):
    if provider == "anthropic":
        r = requests.get("https://api.anthropic.com/v1/models",
                         headers={"x-api-key": key,
                                  "anthropic-version": "2023-06-01"},
                         timeout=30)
        r.raise_for_status()
        names = [m["id"] for m in r.json().get("data", [])]
    else:
        base = ("https://openrouter.ai/api/v1" if provider == "openrouter"
                else "https://api.openai.com/v1")
        r = requests.get(f"{base}/models",
                         headers={"Authorization": f"Bearer {key}"},
                         timeout=30)
        r.raise_for_status()
        names = [m["id"] for m in r.json().get("data", [])]
    for name in sorted(names):
        print(name)


def call_model(key, provider, model, prompt):
    """One request. Returns (text, usage dict). Raises on hard failure."""
    if provider == "anthropic":
        headers = {"x-api-key": key,
                   "anthropic-version": "2023-06-01",
                   "content-type": "application/json"}
        payload = {"model": model,
                   "max_tokens": 300,
                   "thinking": {"type": "disabled"},
                   "system": SYSTEM_MESSAGE,
                   "messages": [{"role": "user", "content": prompt}]}
        r = requests.post("https://api.anthropic.com/v1/messages",
                          headers=headers, json=payload, timeout=60)
        # Models differ in how reasoning is switched off; follow the hint.
        if r.status_code == 400 and "between_tools" in r.text:
            payload["thinking"] = {"type": "between_tools"}
            r = requests.post("https://api.anthropic.com/v1/messages",
                              headers=headers, json=payload, timeout=60)
        elif r.status_code == 400 and "thinking" in r.text:
            payload.pop("thinking", None)
            r = requests.post("https://api.anthropic.com/v1/messages",
                              headers=headers, json=payload, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"{r.status_code}: {r.text[:200]}")
        body = r.json()
        text = "".join(part.get("text", "") for part in body["content"])
        if not text.strip():
            text = json.dumps(body["content"])
        usage = body.get("usage", {})
        return text, {"input_tokens": usage.get("input_tokens"),
                      "output_tokens": usage.get("output_tokens")}

    base = ("https://openrouter.ai/api/v1" if provider == "openrouter"
            else "https://api.openai.com/v1")
    payload = {"model": model,
               "temperature": 0,
               "max_tokens": 300,
               "messages": [{"role": "system", "content": SYSTEM_MESSAGE},
                            {"role": "user", "content": prompt}]}
    r = requests.post(f"{base}/chat/completions",
                      headers={"Authorization": f"Bearer {key}",
                               "content-type": "application/json"},
                      json=payload, timeout=60)
    # Some reasoning models reject the temperature parameter; try once without.
    if r.status_code == 400 and "temperature" in r.text:
        payload.pop("temperature")
        r = requests.post(f"{base}/chat/completions",
                          headers={"Authorization": f"Bearer {key}",
                                   "content-type": "application/json"},
                          json=payload, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"{r.status_code}: {r.text[:200]}")
    body = r.json()
    text = body["choices"][0]["message"]["content"]
    usage = body.get("usage", {})
    return text, {"input_tokens": usage.get("prompt_tokens"),
                  "output_tokens": usage.get("completion_tokens")}


def call_with_retries(key, provider, model, prompt, tries=10):
    """Retry on rate limits and transient errors, backing off gently."""
    delay = 3
    for attempt in range(tries):
        try:
            return call_model(key, provider, model, prompt)
        except Exception as err:  # noqa: BLE001 (network errors of any kind)
            message = str(err)
            if attempt == tries - 1 or "credit balance" in message.lower():
                raise
            # A 429 or 529 means slow down; anything else is retried too,
            # since dropped connections are common in long runs.
            time.sleep(delay)
            delay = min(delay * 1.7, 45)


# ---------------------------------------------------------------------------
# Parsing the structured answer
# ---------------------------------------------------------------------------

def parse_answer(text):
    """Pull the JSON object out of the reply; tolerate stray prose."""
    label, why = "", ""
    match = re.search(r"\{.*\}", text, flags=re.S)
    if match:
        try:
            obj = json.loads(match.group(0))
            label = str(obj.get("label", "")).strip()
            why = str(obj.get("why", "")).strip()
        except json.JSONDecodeError:
            pass
    if not label:
        # Truncated or malformed JSON: recover the label field on its own.
        m2 = re.search(r'"label"\s*:\s*"([^"]+)"', text)
        if m2:
            label = m2.group(1).strip()
        m3 = re.search(r'"why"\s*:\s*"([^"]*)', text)
        if m3:
            why = m3.group(1).strip()
    # Accept case and whitespace differences, nothing else.
    canon = {l.lower(): l for l in LABELS}
    label = canon.get(label.lower(), "INVALID")
    return label, why, text


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def load_done():
    done = set()
    if OUTPUT_FILE.exists():
        with OUTPUT_FILE.open() as fh:
            for line in fh:
                rec = json.loads(line)
                if rec["label"] == "INVALID":
                    continue
                done.add((rec["model"], rec["prompt"], rec["bill_id"]))
    return done


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list-models", action="store_true")
    ap.add_argument("--models", nargs="+", default=[])
    ap.add_argument("--prompts", nargs="+",
                    default=["prompt_1_names", "prompt_2_codebook",
                             "prompt_3_codebook_examples"])
    ap.add_argument("--limit", type=int, default=None,
                    help="only the first N titles (for a test run)")
    ap.add_argument("--workers", type=int, default=4)
    args = ap.parse_args()

    key, provider = read_key()
    print(f"Provider: {provider}")

    if args.list_models:
        list_models(key, provider)
        return
    if not args.models:
        sys.exit("Pass --models MODEL [MODEL2 ...] or --list-models.")

    sample = pd.read_csv(SAMPLE_FILE)
    if args.limit:
        sample = sample.head(args.limit)
    templates = {name: (PROMPT_DIR / f"{name}.txt").read_text()
                 for name in args.prompts}

    done = load_done()
    jobs = [(model, pname, row.bill_id, row.title)
            for model in args.models
            for pname in args.prompts
            for row in sample.itertuples(index=False)
            if (model, pname, row.bill_id) not in done]
    print(f"{len(done)} labels already on disk, {len(jobs)} to request.")
    if not jobs:
        return

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    started = time.time()
    n_ok = n_fail = 0

    def work(job):
        model, pname, bill_id, title = job
        prompt = templates[pname].replace("{title}", str(title))
        text, usage = call_with_retries(key, provider, model, prompt)
        label, why, raw = parse_answer(text)
        return {"bill_id": bill_id, "model": model, "prompt": pname,
                "label": label, "why": why, "raw": raw,
                "input_tokens": usage.get("input_tokens"),
                "output_tokens": usage.get("output_tokens"),
                "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S")}

    with OUTPUT_FILE.open("a") as out, \
            ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(work, job): job for job in jobs}
        for i, fut in enumerate(as_completed(futures), 1):
            try:
                rec = fut.result()
                out.write(json.dumps(rec) + "\n")
                out.flush()
                n_ok += 1
            except Exception as err:  # noqa: BLE001
                n_fail += 1
                job = futures[fut]
                print(f"FAILED {job[0]} {job[1]} {job[2]}: {err}")
            if i % 50 == 0 or i == len(jobs):
                elapsed = time.time() - started
                print(f"{i}/{len(jobs)} done, {n_fail} failed, "
                      f"{elapsed/60:.1f} min elapsed")

    print(f"Finished: {n_ok} labels written, {n_fail} failed.")


if __name__ == "__main__":
    main()
