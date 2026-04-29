"""
01_extract.py — Extract poems from TEI XML files for v5 corpus.

Input:  ~/german-poetry-v5/corpus/  (poet subdirectories with .xml files)
Output: ~/german-poetry-v5/output/poems_raw.json

Special handling:
  - Stefan George: lxml recovery mode (embedded <?xml?> declarations in teiCorpus),
    plus george_normalized_text field with standard German capitalization applied
    via spaCy de_core_news_lg (NOUN/PROPN + line-initial capitalization)

Exclusions (same as v4):
  - Poems < MIN_LINES or > MAX_LINES
  - Theatrical content (<sp> / <speaker> tags), except West-östlicher Divan
  - EXCLUDE_TITLES set
"""

import json
import re
from pathlib import Path
import xml.etree.ElementTree as ET

try:
    from lxml import etree as lxml_etree
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

import spacy
nlp = spacy.load("de_core_news_lg")

CORPUS_DIR = Path.home() / "german-poetry-v5" / "corpus"
OUT_DIR    = Path.home() / "german-poetry-v5" / "output"
OUT_DIR.mkdir(exist_ok=True)

NS        = "http://www.tei-c.org/ns/1.0"
MIN_LINES = 4
MAX_LINES = 400

EXCLUDE_TITLES = {
    "fußnoten", "anmerkungen", "register", "inhaltsverzeichnis",
    "vorwort", "nachwort", "widmung", "einleitung", "anhang",
    "erstes buch", "zweites buch", "drittes buch", "viertes buch",
    "fünftes buch", "sechstes buch", "siebtes buch", "achtes buch",
    "erster teil", "zweiter teil", "dritter teil",
}

# ── Helpers ──────────────────────────────────────────────────────────────────

def tag(name):
    return f"{{{NS}}}{name}"

def text_of(elem):
    return re.sub(r"\s+", " ", "".join(elem.itertext())).strip()

def collection_from_path(n_attr):
    parts = [p for p in n_attr.split("/") if p.strip()]
    if len(parts) >= 5:
        return parts[4]
    elif len(parts) >= 4:
        return parts[3]
    return ""

def is_dramatic(tei_elem):
    n = tei_elem.get("n", "")
    if "Divan" in n or "divan" in n:
        return False
    if tei_elem.find(f".//{tag('sp')}") is not None:
        return True
    if tei_elem.find(f".//{tag('speaker')}") is not None:
        return True
    return False

def get_top_level_lgs(tei_elem):
    parent_map = {}
    for parent in tei_elem.iter():
        for child in parent:
            parent_map[child] = parent

    stanzas = []
    for elem in tei_elem.iter(tag("lg")):
        parent = parent_map.get(elem)
        if parent is not None and parent.tag != tag("lg"):
            stanzas.append(elem)
    return stanzas

def extract_poem(tei_elem):
    title = ""
    candidates = []
    for head in tei_elem.iter(tag("head")):
        t = text_of(head)
        t = re.sub(r"^\d+[\.\s]*", "", t).strip()
        t = re.sub(r"\s+", " ", t).strip()
        htype = head.get("type", "")
        if t and htype in ("h4", "h3"):
            candidates.append((htype, t))

    for htype in ("h4", "h3", "h2"):
        for ht, t in reversed(candidates):
            if ht == htype and len(t) > 1:
                title = t
                break
        if title:
            break
    if not title:
        for head in tei_elem.iter(tag("head")):
            t = re.sub(r"^\d+[\.\s]*", "", text_of(head)).strip()
            if len(t) > 1:
                title = t
                break

    stanzas = []
    all_lines = []

    for lg in get_top_level_lgs(tei_elem):
        stanza_lines = []
        for l in lg.findall(tag("l")):
            line_text = text_of(l)
            if line_text:
                stanza_lines.append(line_text)
                all_lines.append(line_text)
        if stanza_lines:
            stanzas.append(stanza_lines)

    return title, stanzas, all_lines

def get_metadata(tei_elem):
    pub_place      = ""
    pub_date       = ""
    comp_date_from = ""
    comp_date_to   = ""

    place = tei_elem.find(f".//{tag('pubPlace')}")
    if place is not None and place.text:
        pub_place = place.text.strip()

    for date in tei_elem.findall(f".//{tag('date')}"):
        val = date.get("when") or date.get("when-iso") or ""
        if not val and date.text:
            val = date.text.strip()
        if val:
            pub_date = val
            break

    comp = tei_elem.find(f".//{tag('creation')}/{tag('date')}")
    if comp is not None:
        comp_date_from = comp.get("notBefore") or comp.get("when") or ""
        comp_date_to   = comp.get("notAfter")  or comp.get("when") or ""

    return pub_place, pub_date, comp_date_from, comp_date_to

# ── George normalization ──────────────────────────────────────────────────────

def normalize_george_line(line):
    """Apply standard German capitalization to a single line via spaCy POS tags."""
    if not line.strip():
        return line
    doc = nlp(line)
    parts = []
    for i, tok in enumerate(doc):
        text = tok.text
        if text and (i == 0 or tok.pos_ in ("NOUN", "PROPN")):
            text = text[0].upper() + text[1:]
        parts.append(text + tok.whitespace_)
    return "".join(parts).rstrip(" ")

def normalize_george_text(text):
    """Normalize full poem text (stanzas separated by blank lines)."""
    normalized_stanzas = []
    for stanza in text.split("\n\n"):
        normalized_lines = [normalize_george_line(line) for line in stanza.split("\n")]
        normalized_stanzas.append("\n".join(normalized_lines))
    return "\n\n".join(normalized_stanzas)

# ── XML parsing ───────────────────────────────────────────────────────────────

def parse_xml_standard(xml_path):
    try:
        tree = ET.parse(xml_path)
        return tree.getroot(), False
    except ET.ParseError:
        return None, False

def parse_xml_george(xml_path):
    """George XML has embedded <?xml?> declarations. Strip them and use lxml recovery."""
    with open(xml_path, "rb") as f:
        content = f.read().decode("utf-8", errors="replace")
    content = re.sub(r"<\?xml[^?]*\?>", "", content)
    content = '<?xml version="1.0" encoding="utf-8"?>\n' + content
    parser = lxml_etree.XMLParser(recover=True, encoding="utf-8")
    lxml_root = lxml_etree.fromstring(content.encode("utf-8"), parser)
    raw = lxml_etree.tostring(lxml_root, encoding="unicode")
    raw = re.sub(r"<\?xml[^?]*\?>", "", raw)
    return ET.fromstring(raw)

def parse_xml_lxml_recovery(xml_path):
    """Generic lxml recovery fallback."""
    parser = lxml_etree.XMLParser(recover=True, encoding="utf-8")
    tree = lxml_etree.parse(str(xml_path), parser)
    lxml_root = tree.getroot()
    raw = lxml_etree.tostring(lxml_root, encoding="unicode")
    raw = re.sub(r"<\?xml[^?]*\?>", "", raw)
    return ET.fromstring(raw)

# ── File processing ───────────────────────────────────────────────────────────

def process_file(xml_path, poet_name):
    is_george = (poet_name.lower() == "george")

    if is_george:
        if not HAS_LXML:
            print(f"  lxml required for George XML but not available")
            return []
        try:
            root = parse_xml_george(xml_path)
            recovered = True
        except Exception as e:
            print(f"  PARSE ERROR (george) {xml_path.name}: {e}")
            return []
    else:
        root, recovered = None, False
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
        except ET.ParseError:
            if HAS_LXML:
                try:
                    root = parse_xml_lxml_recovery(xml_path)
                    recovered = True
                except Exception as e:
                    print(f"  PARSE ERROR (lxml) {xml_path.name}: {e}")
                    return []
            else:
                print(f"  PARSE ERROR (unrecoverable) {xml_path.name}")
                return []

    if root is None:
        return []
    if recovered:
        print(f"  (lxml recovery) {xml_path.name}")

    poems = []
    poem_counter = [0]

    def process_tei(tei_elem):
        n_attr = tei_elem.get("n", "")
        if not n_attr:
            return

        if is_dramatic(tei_elem):
            return

        collection = collection_from_path(n_attr)
        title, stanzas, lines = extract_poem(tei_elem)

        if not title or not lines:
            return
        if len(lines) < MIN_LINES:
            return
        if len(lines) > MAX_LINES:
            return
        if title.lower().strip(".") in EXCLUDE_TITLES:
            return

        pub_place, pub_date, comp_from, comp_to = get_metadata(tei_elem)

        text = "\n\n".join("\n".join(s) for s in stanzas)

        poem_counter[0] += 1
        poem_id = f"{re.sub(r'[^a-zA-Z0-9]', '_', poet_name)}_{poem_counter[0]:04d}"

        record = {
            "poem_id":        poem_id,
            "poet":           poet_name,
            "collection":     collection,
            "poem_title":     title,
            "text":           text,
            "stanzas":        stanzas,
            "n_stanzas":      len(stanzas),
            "n_lines":        len(lines),
            "pub_place":      pub_place,
            "pub_date":       pub_date,
            "comp_date_from": comp_from,
            "comp_date_to":   comp_to,
            "source_file":    xml_path.name,
        }

        if is_george:
            record["george_normalized_text"] = normalize_george_text(text)

        poems.append(record)

    for tei in root.iter(tag("TEI")):
        process_tei(tei)

    return poems

# ── Run ───────────────────────────────────────────────────────────────────────

all_poems = []

for poet_dir in sorted(CORPUS_DIR.iterdir()):
    if not poet_dir.is_dir():
        continue
    xml_files = sorted(poet_dir.glob("*.xml"))
    if not xml_files:
        continue

    poet_name  = poet_dir.name
    poet_poems = []
    for xml_file in xml_files:
        if xml_file.name.startswith("~"):
            continue
        extracted = process_file(xml_file, poet_name)
        poet_poems.extend(extracted)

    print(f"  {poet_name:35s} {len(poet_poems):4d} poems  "
          f"({len(xml_files)} file{'s' if len(xml_files)>1 else ''})")
    all_poems.extend(poet_poems)

for i, p in enumerate(all_poems):
    p["global_id"] = i

out_path = OUT_DIR / "poems_raw.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(all_poems, f, ensure_ascii=False, indent=2)

print(f"\nTotal: {len(all_poems)} poems -> {out_path}")

# ── Diagnostics ───────────────────────────────────────────────────────────────
from collections import Counter
by_poet       = Counter(p["poet"] for p in all_poems)
by_collection = Counter(p["collection"] for p in all_poems)

print("\nBy poet:")
for poet, n in sorted(by_poet.items(), key=lambda x: -x[1]):
    print(f"  {poet:35s} {n}")

print("\nBy collection (top 25):")
for coll, n in by_collection.most_common(25):
    print(f"  {coll:50s} {n}")

print("\nSample titles:")
for p in all_poems[:5]:
    print(f"  [{p['poet']}] {p['collection']} / {p['poem_title']}")

george_poems = [p for p in all_poems if p["poet"] == "george"]
if george_poems:
    print(f"\nGeorge sample (original vs normalized):")
    sample = george_poems[0]
    orig_lines = sample["text"].split("\n")[:4]
    norm_lines = sample.get("george_normalized_text", "").split("\n")[:4]
    for o, n in zip(orig_lines, norm_lines):
        print(f"  orig: {o}")
        print(f"  norm: {n}")
        print()
