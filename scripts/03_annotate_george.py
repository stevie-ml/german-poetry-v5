"""
03_annotate_george.py — Claude Haiku annotations for George poems only.

Output: ~/german-poetry-v5/output/george_annotated.csv
"""

import json, time, re
from pathlib import Path
import pandas as pd
import anthropic

OUT_DIR = Path.home() / "german-poetry-v5" / "output"
client  = anthropic.Anthropic()

with open(OUT_DIR / "poems_raw.json", encoding="utf-8") as f:
    all_poems = json.load(f)

poems = [p for p in all_poems if p["poet"] == "george"]
print(f"George poems to annotate: {len(poems)}")

SYSTEM = """You are annotating German poems using categories drawn from the Princeton Encyclopedia of Poetry and Poetics (Ramazani et al., eds., 4th ed., Princeton UP, 2012).

Read the ENTIRE poem before responding — the final stanzas matter as much as the first. Base your judgments on the whole poem, not just the opening lines.

DEFINITIONS (from PEPP entries):

persona [→ "person"]: The speaker "imagined to be spoken in a distinctive voice or narrated from a determinate vantage" (PEPP, "Persona"). Identify the dominant grammatical person.
  Values: "1st" | "2nd" | "3rd" | "mixed"

apostrophe [→ "has_addressee"]: Direct address to a person, abstraction, or object (PEPP, "Apostrophe"). Does the poem address a "you" (du/Sie/ihr/euch)?
  Values: true | false

imagery_density: "The baffling confluence of concrete and abstract, literal and figurative, body and mind" (PEPP, "Imagery"). Rate density of concrete, sensory particulars.
  Values: 1 (highly abstract) … 5 (densely concrete/sensory)

tone: "Reflects the speaker's awareness of his relation to his listener, his sense of how he stands toward those he is addressing" (Richards, via PEPP, "Tone").
  Values: "lyric" | "elegiac" | "ecstatic" | "ironic" | "meditative" | "dramatic"

closure: "The achievement of an effect of finality, resolution, and stability at the end of a poem... a function of the reader's perception of the concluding portion in relation to the entire composition" (PEPP, "Closure").
  Values: 1 (unresolved/open) … 5 (strongly closed/resolved)

emotional_valence: Overall affective register.
  Values: 1 (very negative/dark) … 7 (very positive/joyful)

emotional_intensity: Strength of emotional charge regardless of direction.
  Values: 1 (calm/detached) … 5 (anguished or ecstatic)

dominant_theme: Central thematic concern.
  Values: "nature" | "death" | "love" | "body" | "city" | "time" | "religion" | "self" | "war" | "myth" | "other"

temporal_frame: Primary temporal orientation.
  Values: "past" | "present" | "future" | "timeless"

setting: Dominant spatial context.
  Values: "urban" | "rural" | "interior" | "abstract" | "mixed"

Return ONLY a JSON object with exactly these 10 fields. No explanation, no markdown fences."""

FIELDS = ["person","has_addressee","emotional_valence","emotional_intensity",
          "dominant_theme","imagery_density","tone","temporal_frame","setting","closure"]


def annotate(poem):
    text = poem["text"]
    prompt = (
        f"Author: {poem['poet']}\n"
        f"Collection: {poem['collection']}\n"
        f"Title: {poem['poem_title']}\n\n"
        f"FULL POEM (read all stanzas before annotating):\n\n"
        f"{text}"
    )
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=300,
        system=SYSTEM,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


results = []
errors  = []

for i, poem in enumerate(poems):
    print(f"[{i+1:3}/{len(poems)}] {poem['poem_title'][:50]}", end=" ... ", flush=True)
    try:
        ann = annotate(poem)
        row = {
            "poem_id":    poem["poem_id"],
            "poet":       poem["poet"],
            "collection": poem["collection"],
            "poem_title": poem["poem_title"],
            **{f: ann.get(f) for f in FIELDS},
        }
        print(f"person={ann.get('person','?'):<5} theme={ann.get('dominant_theme','?')}")
    except Exception as e:
        print(f"ERROR: {e}")
        errors.append((i, poem["poem_title"], str(e)))
        row = {
            "poem_id":    poem["poem_id"],
            "poet":       poem["poet"],
            "collection": poem["collection"],
            "poem_title": poem["poem_title"],
            **{f: None for f in FIELDS},
        }
    results.append(row)
    if (i + 1) % 20 == 0:
        time.sleep(1)

out_path = OUT_DIR / "george_annotated.csv"
pd.DataFrame(results).to_csv(out_path, index=False)
print(f"\nDone. {len(results)} poems -> {out_path}")
if errors:
    print(f"Errors: {len(errors)}")
    for idx, title, err in errors[:5]:
        print(f"  [{idx+1}] {title}: {err}")
