"""
02_compute.py  —  Token-level surprisal/entropy for all poems.

Input:   output/poems_v2.json
Output:  output/poems_metrics_v2.csv

For each poem:
  - Use first_12_text (first 12 verse lines, pre-extracted)
  - Tokenize with dbmdz/german-gpt2 native tokenizer
  - One forward pass
  - mean_token_surprisal = mean of -log2 P(token_t | context_t-1)  over t=1..T
  - mean_token_entropy   = mean of H(context_t-1) = -sum p log2 p  over t=1..T
  - s2                   = mean_token_surprisal - mean_token_entropy
  - n_tokens_used        = T (number of tokens in the 12-line text)

No word reconstruction. Native tokenization only.
Everything recomputed from scratch; no values imported from prior runs.
"""

import json, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import torch
from transformers import AutoTokenizer, AutoModelForCausalLM

warnings.filterwarnings("ignore")

OUT_DIR    = Path.home() / "german-poetry-v5" / "output"
MODEL_NAME = "dbmdz/german-gpt2"

print(f"Loading {MODEL_NAME}...")
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model     = AutoModelForCausalLM.from_pretrained(MODEL_NAME)
model.eval()
device = "mps" if torch.backends.mps.is_available() else "cpu"
model  = model.to(device)
print(f"Model on {device}")

with open(OUT_DIR / "poems_v2.json", encoding="utf-8") as f:
    poems = json.load(f)

MAX_TOKENS = 512   # safety cap; 12 lines typically ≪ 512 tokens


def compute_metrics(text: str):
    """
    Returns (mean_surprisal, mean_entropy, s2, n_tokens) or None on failure.
    All values in bits (log base 2).
    Position 0 has no left context — excluded from averages.
    """
    if not text.strip():
        return None

    enc = tokenizer(text, return_tensors="pt", truncation=True,
                    max_length=MAX_TOKENS).to(device)
    input_ids = enc["input_ids"][0]
    T = len(input_ids)
    if T < 2:
        return None

    with torch.no_grad():
        logits = model(**enc).logits[0]   # [T, vocab]

    log_probs = torch.log_softmax(logits.float(), dim=-1).cpu()
    input_ids = input_ids.cpu()

    surps = []
    ents  = []
    for t in range(1, T):
        lp = log_probs[t - 1]             # distribution predicted before token t
        surps.append(-lp[input_ids[t]].item() / np.log(2))
        ents.append(-(lp.exp() * lp).sum().item() / np.log(2))

    ms = float(np.mean(surps))
    me = float(np.mean(ents))
    return ms, me, ms - me, T - 1         # n_tokens_used = positions with context


rows = []
for i, poem in enumerate(poems):
    if i % 100 == 0:
        print(f"  [{i+1:5}/{len(poems)}] {poem['poet'][:20]:20s} {poem['poem_title'][:35]}",
              flush=True)

    text   = poem.get("first_12_text", "")
    result = compute_metrics(text)

    row = {
        "poem_id":                       poem["poem_id"],
        "poet":                          poem["poet"],
        "poem_title":                    poem["poem_title"],
        "collection_title":              poem["collection_title"],
        "collection_pub_year":           poem["collection_pub_year"],
        "collection_pub_year_uncertain": poem["collection_pub_year_uncertain"],
        "collection_pub_year_note":      poem["collection_pub_year_note"],
        "source_file":                   poem["source_file"],
        "tei_id":                        poem["tei_id"],
        "total_verse_lines":             poem["total_verse_lines"],
        "stanza_pattern":                poem["stanza_pattern"],
        "lines_used_for_metrics":        min(poem["total_verse_lines"], 12),
        "model":                         MODEL_NAME,
    }

    if result is not None:
        ms, me, s2, ntok = result
        row.update({
            "mean_token_surprisal": round(ms,  4),
            "mean_token_entropy":   round(me,  4),
            "s2":                   round(s2,  4),
            "n_tokens_used":        ntok,
            "compute_error":        "",
        })
    else:
        row.update({
            "mean_token_surprisal": None,
            "mean_token_entropy":   None,
            "s2":                   None,
            "n_tokens_used":        0,
            "compute_error":        "empty or too-short text",
        })

    rows.append(row)

df = pd.DataFrame(rows)
df.to_csv(OUT_DIR / "poems_metrics_v2.csv", index=False)

n_ok  = df["mean_token_surprisal"].notna().sum()
n_err = df["compute_error"].astype(bool).sum()
print(f"\nDone.  {n_ok} poems computed, {n_err} errors")
print(f"Model: {MODEL_NAME}")
print(f"Saved → output/poems_metrics_v2.csv")
