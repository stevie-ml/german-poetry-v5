"""
02_surprisal_george.py — GPT-2 surprisal metrics for George poems only.

Runs dbmdz/german-gpt2 on the 52 George poems from poems_raw.json.
Uses original (lowercase) text for surprisal — that's the text being analyzed.

Output: ~/german-poetry-v5/output/george_surprisal.csv
"""

import json, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import torch
from transformers import AutoTokenizer, AutoModelForCausalLM

warnings.filterwarnings("ignore")

OUT_DIR = Path.home() / "german-poetry-v5" / "output"

MODEL_NAME = "dbmdz/german-gpt2"
print(f"Loading {MODEL_NAME}...")
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model     = AutoModelForCausalLM.from_pretrained(MODEL_NAME)
model.eval()
device = "mps" if torch.backends.mps.is_available() else "cpu"
model  = model.to(device)
print(f"Model on {device}")

with open(OUT_DIR / "poems_raw.json", encoding="utf-8") as f:
    all_poems = json.load(f)

poems = [p for p in all_poems if p["poet"] == "george"]
print(f"George poems: {len(poems)}")

MAX_TOKENS = 900


def poem_surprisal(text):
    words = text.split()
    if not words:
        return []

    enc       = tokenizer(text, return_tensors="pt", truncation=True,
                          max_length=MAX_TOKENS).to(device)
    input_ids = enc["input_ids"][0]

    with torch.no_grad():
        logits = model(**enc).logits[0]

    log_probs = torch.log_softmax(logits.float(), dim=-1).cpu()
    input_ids = input_ids.cpu()
    T         = len(input_ids)

    token_surp = torch.zeros(T)
    token_ent  = torch.zeros(T)
    for t in range(1, T):
        lp = log_probs[t - 1]
        token_surp[t] = -lp[input_ids[t]] / np.log(2)
        token_ent[t]  = -(lp.exp() * lp).sum() / np.log(2)

    token_surp = token_surp.numpy()
    token_ent  = token_ent.numpy()
    token_strs = [tokenizer.decode([tid]) for tid in input_ids.tolist()]

    results = []
    t_idx   = 0
    for word in words:
        if t_idx >= T:
            results.append((word, 0.0, 0.0, 0.0))
            continue

        w_surp   = 0.0
        w_ent    = None
        consumed = ""

        while t_idx < T and len(consumed.strip()) < len(word):
            piece  = token_strs[t_idx].lstrip("Ġ▁ ")
            w_surp += float(token_surp[t_idx])
            if w_ent is None:
                w_ent = float(token_ent[t_idx])
            consumed += piece
            t_idx    += 1
            if consumed.strip() == word or consumed.strip().startswith(word):
                break

        if w_ent is None:
            w_ent = 0.0
        s2 = w_surp - w_ent
        results.append((word, w_surp, w_ent, s2))

    return results


poem_rows = []

for i, poem in enumerate(poems):
    pid   = poem["poem_id"]
    title = poem["poem_title"]
    print(f"[{i+1:3}/{len(poems)}] {title[:50]}", flush=True)

    word_data = poem_surprisal(poem["text"])
    if not word_data:
        continue

    poem_surps, poem_ents, poem_s2s = [], [], []
    for _, surp, ent, s2 in word_data:
        poem_surps.append(surp)
        poem_ents.append(ent)
        poem_s2s.append(s2)

    if not poem_surps:
        continue

    n    = len(poem_surps)
    half = n // 2
    arr  = np.array(poem_surps)
    s2a  = np.array(poem_s2s)

    tension  = float(np.mean(arr[half:]) - np.mean(arr[:half])) if half > 0 else 0.0
    peak_pos = float(np.argmax(s2a)) / n if n > 0 else 0.5
    ac1_val  = float(np.corrcoef(s2a[:-1], s2a[1:])[0, 1]) if n > 2 else 0.0
    ttr_val  = len(set(poem["text"].split())) / max(n, 1)

    poem_rows.append({
        "poem_id":        pid,
        "poet":           poem["poet"],
        "collection":     poem["collection"],
        "poem_title":     title,
        "n_lines":        poem["n_lines"],
        "n_stanzas":      poem["n_stanzas"],
        "n_words":        n,
        "pub_place":      poem.get("pub_place", ""),
        "pub_date":       poem.get("pub_date", ""),
        "comp_date_from": poem.get("comp_date_from", ""),
        "comp_date_to":   poem.get("comp_date_to", ""),
        "mean_surprisal": round(float(np.mean(arr)), 4),
        "mean_entropy":   round(float(np.mean(poem_ents)), 4),
        "s2_mean":        round(float(np.mean(s2a)), 4),
        "uid_sigma":      round(float(np.std(arr)), 4),
        "tension":        round(tension, 4),
        "peak_pos":       round(peak_pos, 4),
        "ac1":            round(ac1_val, 4),
        "ttr":            round(ttr_val, 4),
    })

out_path = OUT_DIR / "george_surprisal.csv"
pd.DataFrame(poem_rows).to_csv(out_path, index=False)
print(f"\nDone. {len(poem_rows)} poems -> {out_path}")
