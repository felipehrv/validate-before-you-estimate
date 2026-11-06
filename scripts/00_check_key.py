"""
Check the API key and pick the two models for the lesson.

Run from the project folder:  python3 scripts/00_check_key.py
"""

import importlib.util
import sys
from pathlib import Path

import requests

PROJECT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location(
    "clf", PROJECT / "scripts" / "02_classify_titles.py")
clf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(clf)

key, provider = clf.read_key()
print(f"Provider inferred from the key: {provider}")

# 1. Can we list models?
try:
    if provider == "anthropic":
        r = requests.get("https://api.anthropic.com/v1/models",
                         headers={"x-api-key": key,
                                  "anthropic-version": "2023-06-01"},
                         timeout=30)
    else:
        base = ("https://openrouter.ai/api/v1" if provider == "openrouter"
                else "https://api.openai.com/v1")
        r = requests.get(f"{base}/models",
                         headers={"Authorization": f"Bearer {key}"}, timeout=30)
    print(f"Model list: HTTP {r.status_code}")
    if r.status_code != 200:
        print(r.text[:400])
        sys.exit("The key was not accepted. Check it in the console and "
                 "paste it again into secrets/api_key.txt.")
    models = sorted(m["id"] for m in r.json().get("data", []))
    print("Available models:")
    for m in models:
        print("  ", m)
except requests.RequestException as err:
    sys.exit(f"Network problem: {err}")

# 2. One real call with a cheap model, to confirm billing works.
cheap = next((m for m in models if "haiku" in m), None) or \
        next((m for m in models if "mini" in m), None) or models[0]
print(f"\nTest call with {cheap} ...")
text, usage = clf.call_model(key, provider, cheap,
                             'Reply with the JSON object {"label": "Other", '
                             '"why": "test"} and nothing else.')
print("Reply:", text.strip()[:120])
print("Usage:", usage)
print("\nThe key works. Suggested pair for the lesson:")
strong = next((m for m in models if "sonnet" in m), None) or \
         next((m for m in models if "gpt-4" in m or "gpt-5" in m), None)
print(f"  cheap:  {cheap}")
print(f"  strong: {strong}")
