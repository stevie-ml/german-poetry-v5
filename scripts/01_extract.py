"""
01_extract.py  —  TEI-XML → poems_v2.json + exclusions_v2.csv

Rules:
  - Verse lines = <l> elements only
  - Skip lines inside <argument>, <epigraph>, <note>, <sp>, <stage>
  - Minimum 4 verse lines; maximum 400
  - Stanza pattern from parent <lg> grouping
  - Collection pub year from hand-coded lookup (ORIGINAL collection, not Sammelband)
  - Uncertainty flagged explicitly; never silently guessed
"""

import json, re, sys
from pathlib import Path
from collections import OrderedDict
import csv

try:
    from lxml import etree as LET
    HAS_LXML = True
except ImportError:
    HAS_LXML = False
    import xml.etree.ElementTree as ET

CORPUS_DIR = Path.home() / "german-poetry-v5" / "corpus"
OUT_DIR    = Path.home() / "german-poetry-v5" / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

NS  = "http://www.tei-c.org/ns/1.0"
def T(name): return f"{{{NS}}}{name}"

MIN_LINES = 4
MAX_LINES = 400

EXCLUDE_TITLES = {
    "fußnoten", "anmerkungen", "register", "inhaltsverzeichnis",
    "vorwort", "nachwort", "widmung", "einleitung", "anhang",
    "erstes buch", "zweites buch", "drittes buch", "viertes buch",
    "fünftes buch", "erster teil", "zweiter teil", "dritter teil",
    "prolog", "epilog",
}

# Tags whose descendants are not verse lines
EXCLUDED_ANCESTORS = {T(x) for x in
    ["argument", "epigraph", "note", "sp", "stage", "figure", "table", "bibl"]}

# ── Collection pub-year lookup ───────────────────────────────────────────────
# Key: (poet_fragment_lower, collection_fragment_lower)
# Value: (year: int|None, uncertain: bool, note: str)
# Longest matching key wins.
COLL_PUB_YEARS = {
    # Trakl
    ("trakl","gedichte"):               (1913, False, "Kurt Wolff, Leipzig, 1913"),
    ("trakl","sebastian im traum"):     (1915, False, "Kurt Wolff, Leipzig, 1915"),
    ("trakl","die dichtungen"):         (1917, True,  "Posth. collected ed., Kurt Wolff 1917"),
    ("trakl","dichtungen"):             (1917, True,  "Posth. collected ed., Kurt Wolff 1917"),
    # Heym
    ("heym","der ewige tag"):           (1911, False, "Ernst Rowohlt, Leipzig, 1911"),
    ("heym","umbra vitae"):             (1912, True,  "Posth., Ernst Rowohlt, Leipzig, 1912"),
    ("heym","dichtungen"):              (1922, True,  "Posth. collected ed., Munich, 1922"),
    ("heym","marathon"):                (1914, True,  "Posth., 1914"),
    # Stadler
    ("stadler","aufbruch"):             (1914, False, "A.R. Meyer, Berlin-Wilmersdorf, 1914"),
    ("stadler","präludien"):            (1905, False, "Strassburg, 1905"),
    # van Hoddis
    ("hvh","weltende"):                 (1918, True,  "Posth. ed. 1918; 'Weltende' poem first publ. Der Demokrat 1911"),
    ("van hoddis","weltende"):          (1918, True,  "Posth. ed. 1918"),
    # Rilke
    ("rilke","stundenbuch"):            (1905, False, "Insel, Leipzig, 1905"),
    ("rilke","buch der bilder"):        (1902, False, "Axel Juncker, Berlin, 1902; expanded 1906"),
    ("rilke","neue gedichte"):          (1907, False, "Insel, Leipzig, 1907"),
    ("rilke","neuen gedichte anderer"): (1908, False, "Insel, Leipzig, 1908"),
    ("rilke","duineser elegien"):       (1923, False, "Insel, Leipzig, 1923"),
    ("rilke","sonette an orpheus"):     (1923, False, "Insel, Leipzig, 1923"),
    ("rilke","marien-leben"):           (1913, False, "Insel, Leipzig, 1913"),
    ("rilke","mir zur feier"):          (1899, False, "Berlin, 1899"),
    ("rilke","larenopfer"):             (1896, False, "Prague, 1896"),
    ("rilke","traumgekrönt"):           (1897, False, "Leipzig, 1897"),
    ("rilke","frühen gedichte"):        (1909, True,  "Rev. of 'Mir zur Feier', Insel 1909; orig. 1899"),
    # Goethe
    ("goethe","west-östlicher divan"):  (1819, False, "Cotta, Stuttgart, 1819"),
    ("goethe","divan"):                 (1819, False, "Cotta, Stuttgart, 1819"),
    ("goethe","gedichte"):              (1789, True,  "First in Goethes Schriften Bd.8, 1789; many later eds"),
    # Schiller
    ("schiller","gedichte"):            (1800, False, "Cotta, Tübingen, 1800 (vol.1)"),
    # Klopstock
    ("klopstock","oden"):               (1771, False, "Bode, Hamburg, 1771"),
    ("klopstock","geistliche lieder"):  (1758, False, "Hamburg, 1758 (pt.1); 1769 (pt.2)"),
    ("klopstock","hinterlassen"):       (1815, True,  "Posth., Hamburg, 1815"),
    # Hölderlin
    ("holderlin","gedichte"):           (1826, True,  "Posth. collected ed. Schwab/Uhland, 1826"),
    ("hölderlin","gedichte"):           (1826, True,  "Posth. collected ed. Schwab/Uhland, 1826"),
    # Mörike
    ("mörike","gedichte"):              (1838, False, "Schweizerbart, Stuttgart, 1838 (1st ed.)"),
    ("morike","gedichte"):              (1838, False, "Schweizerbart, Stuttgart, 1838 (1st ed.)"),
    # Storm
    ("storm","gedichte"):               (1852, False, "Schröder, Kiel, 1852 (1st ed.)"),
    # Droste-Hülshoff
    ("droste","gedichte"):              (1838, False, "Cotta, Stuttgart, 1838 (1st ed.)"),
    ("droste","geistliche jahr"):       (1851, True,  "Posth., Aschendorff, Münster, 1851"),
    # Brockes
    ("brockes","irdisches vergnügen"):  (1721, False, "Hamburg, 1721–1748 (9 vols)"),
    # Haller
    ("haller","versuch schweizerischer"): (1732, False, "Bern, 1732"),
    ("haller","gedichte"):              (1732, True,  "Likely 'Versuch Schweizerischer Gedichten', Bern 1732"),
    # Morgenstern
    ("morgenstern","galgenlieder"):     (1905, False, "Bruno Cassirer, Berlin, 1905"),
    ("morgenstern","palmström"):        (1910, False, "Bruno Cassirer, Berlin, 1910"),
    ("morgenstern","gingganz"):         (1919, True,  "Posth., R. Piper, Munich, 1919"),
    ("morgenstern","phantas schloss"):  (1895, False, "Berlin, 1895"),
    ("morgenstern","ich und du"):       (1911, False, "Bruno Cassirer, Berlin, 1911"),
    ("morgenstern","einkehr"):          (1910, False, "R. Piper, Munich, 1910"),
    ("morgenstern","melancholie"):      (1906, False, "Bruno Cassirer, Berlin, 1906"),
    # Wedekind
    ("wedekind","vier jahreszeiten"):   (1905, False, "Georg Müller, Munich, 1905"),
    ("wedekind","lautenlieder"):        (1920, True,  "Posth., Georg Müller, Munich, 1920"),
    # Dörmann
    ("dörmann","neurotica"):            (1891, False, "Bondy, Wien, 1891"),
    ("dörmann","sensationen"):          (1892, False, "Bondy, Wien, 1892"),
    ("dörmann","gedichte"):             (1896, True,  "Wien, 1896"),
    ("dormann","neurotica"):            (1891, False, "Bondy, Wien, 1891"),
    ("dormann","sensationen"):          (1892, False, "Bondy, Wien, 1892"),
    # Stefan George
    ("george","jahr der seele"):        (1897, False, "Georg Bondi, Berlin, 1897"),
    ("george","teppich des lebens"):    (1900, False, "Georg Bondi, Berlin, 1900"),
    ("george","siebente ring"):         (1907, False, "Georg Bondi, Berlin, 1907"),
    ("george","stern des bundes"):      (1914, False, "Georg Bondi, Berlin, 1914"),
    ("george","neue reich"):            (1928, False, "Georg Bondi, Berlin, 1928"),
    ("george","hymnen"):                (1890, False, "Privately printed, 1890"),
    ("george","pilgerfahrten"):         (1891, False, "Privately printed, 1891"),
    ("george","algabal"):               (1892, False, "Privately printed, 1892"),
    ("george","bücher der hirten"):     (1895, False, "Georg Bondi, Berlin, 1895"),
    ("george","gedichte"):              (None, True,  "Collected; exact original collection unclear"),
}

SAMMELBAND_RE = re.compile(
    r"gesammelt|sämtlich|ausgewählt|\bwerke\b|\bschriften\b|"
    r"nachlass|hinterlassen|gesamtausgabe|gesamte\s+dichtung", re.I
)


def norm(s):
    return re.sub(r"[^a-z0-9äöüß ]", " ", s.lower()).strip()


def lookup_pub_year(poet_name, coll_title):
    """Returns (year|None, uncertain_bool, note, method)."""
    poet_n = norm(poet_name)
    coll_n = norm(coll_title)

    best_score, best_entry = 0, None
    for (pk, ck), val in COLL_PUB_YEARS.items():
        if pk in poet_n and ck in coll_n:
            score = len(pk) + len(ck)
            if score > best_score:
                best_score, best_entry = score, val
    if best_entry:
        year, unc, note = best_entry
        return year, unc, note, "lookup"

    m = re.search(r'\b(1[5-9]\d{2}|20[012]\d)\b', coll_title)
    if m:
        y = int(m.group(1))
        return y, True, f"Year {y} extracted from collection title — verify", "title_regex"

    if SAMMELBAND_RE.search(coll_title):
        return None, True, "Sammelband/collected-works title; original pub year unknown", "sammelband"

    return None, True, "Not in lookup; pub year unknown", "no_match"


# ── XML loading ──────────────────────────────────────────────────────────────

def load_xml(path):
    if HAS_LXML:
        try:
            parser = LET.XMLParser(recover=True)
            tree   = LET.parse(str(path), parser)
            root   = tree.getroot()
            xml_str = LET.tostring(root, encoding="unicode")
            xml_str = re.sub(r'<\?xml[^?]*\?>', '', xml_str)
            import xml.etree.ElementTree as ET2
            return ET2.fromstring(xml_str)
        except Exception:
            pass
    try:
        import xml.etree.ElementTree as ET2
        return ET2.parse(str(path)).getroot()
    except Exception:
        return None


# ── Verse-line extraction ────────────────────────────────────────────────────

def get_verse_lines(tei_elem):
    """Returns (lines: list[str], stanza_counts: list[int])."""
    parent_map = {}
    for parent in tei_elem.iter():
        for child in parent:
            parent_map[child] = parent

    def has_excluded_ancestor(elem):
        curr = parent_map.get(elem)
        while curr is not None:
            if curr.tag in EXCLUDED_ANCESTORS:
                return True
            curr = parent_map.get(curr)
        return False

    def immediate_lg_id(elem):
        curr = parent_map.get(elem)
        while curr is not None:
            if curr.tag == T("lg"):
                return id(curr)
            curr = parent_map.get(curr)
        return None

    lg_order  = OrderedDict()
    line_data = []

    for l_elem in tei_elem.iter(T("l")):
        if has_excluded_ancestor(l_elem):
            continue
        text = re.sub(r"\s+", " ", "".join(l_elem.itertext())).strip()
        if not text:
            continue
        lgid = immediate_lg_id(l_elem)
        if lgid is not None and lgid not in lg_order:
            lg_order[lgid] = len(lg_order)
        st_idx = lg_order.get(lgid, -1)
        line_data.append((text, st_idx))

    lines = [t for t, _ in line_data]

    st_counts: dict = OrderedDict()
    for _, si in line_data:
        st_counts[si] = st_counts.get(si, 0) + 1
    stanza_counts = list(st_counts.values())

    return lines, stanza_counts


def stanza_pattern(counts):
    if not counts:
        return ""
    if len(set(counts)) == 1:
        return str(counts[0])
    return "/".join(str(c) for c in counts)


# ── Metadata helpers ─────────────────────────────────────────────────────────

def get_poem_title(tei_elem, n_attr):
    candidates = []
    for head in tei_elem.iter(T("head")):
        txt   = "".join(head.itertext()).strip()
        txt   = re.sub(r"^\d+[\.\s]*", "", txt).strip()
        htype = head.get("type", "")
        if txt and len(txt) > 1:
            candidates.append((htype, txt))
    for htype in ("h4", "h3", "h2"):
        for ht, t in reversed(candidates):
            if ht == htype:
                return t
    parts = [p for p in n_attr.split("/") if p.strip()]
    return parts[-1] if parts else "Unknown"


def get_collection_title(n_attr):
    parts = [p for p in n_attr.split("/") if p.strip()]
    if len(parts) >= 5:
        return parts[4]
    if len(parts) >= 4:
        return parts[3]
    return ""


def get_tei_id(tei_elem):
    for attr in ["{http://www.w3.org/XML/1998/namespace}id", "xml:id"]:
        v = tei_elem.get(attr)
        if v:
            return v
    return tei_elem.get("n", "")[:120]


# ── Main loop ────────────────────────────────────────────────────────────────

poems      = []
exclusions = []
poem_id    = 0

poet_dirs = sorted(CORPUS_DIR.iterdir())
print(f"Processing {len(poet_dirs)} poet directories in {CORPUS_DIR}")

for poet_dir in poet_dirs:
    if not poet_dir.is_dir():
        continue
    poet_folder = poet_dir.name
    xml_files   = sorted(poet_dir.glob("*.xml"))

    for xml_path in xml_files:
        root = load_xml(xml_path)
        if root is None:
            exclusions.append({"source_file": xml_path.name, "poet": poet_folder,
                                "poem_title": "", "reason": "XML parse failure"})
            continue

        teis = root.findall(f".//{T('TEI')}")
        if not teis and root.tag == T("TEI"):
            teis = [root]

        for tei in teis:
            n_attr    = tei.get("n", "")
            tei_id    = get_tei_id(tei)
            title     = get_poem_title(tei, n_attr)
            coll      = get_collection_title(n_attr)
            parts     = [p for p in n_attr.split("/") if p.strip()]
            poet_name = parts[2] if len(parts) >= 3 else poet_folder

            # Dramatic exclusion (except West-östlicher Divan)
            if "Divan" not in n_attr and "divan" not in n_attr:
                if (tei.find(f".//{T('sp')}") is not None or
                        tei.find(f".//{T('speaker')}") is not None):
                    exclusions.append({"source_file": xml_path.name, "poet": poet_name,
                                       "poem_title": title,
                                       "reason": "Dramatic content (<sp>/<speaker>)"})
                    continue

            if title.lower().strip() in EXCLUDE_TITLES:
                exclusions.append({"source_file": xml_path.name, "poet": poet_name,
                                   "poem_title": title, "reason": f"Excluded title"})
                continue

            lines, stanza_counts = get_verse_lines(tei)

            if len(lines) < MIN_LINES:
                exclusions.append({"source_file": xml_path.name, "poet": poet_name,
                                   "poem_title": title,
                                   "reason": f"Too few verse lines ({len(lines)})"})
                continue

            if len(lines) > MAX_LINES:
                exclusions.append({"source_file": xml_path.name, "poet": poet_name,
                                   "poem_title": title,
                                   "reason": f"Too many verse lines ({len(lines)})"})
                continue

            year, uncertain, note, method = lookup_pub_year(poet_name, coll)

            poem_id += 1
            poems.append({
                "poem_id":                       poem_id,
                "poet":                          poet_name,
                "poem_title":                    title,
                "collection_title":              coll,
                "collection_pub_year":           year,
                "collection_pub_year_uncertain": uncertain,
                "collection_pub_year_note":      note,
                "collection_pub_year_method":    method,
                "comp_year":                     None,
                "comp_year_note":                "Not extracted in this pipeline",
                "source_file":                   xml_path.name,
                "tei_id":                        tei_id,
                "total_verse_lines":             len(lines),
                "stanza_pattern":                stanza_pattern(stanza_counts),
                "first_12_lines":                lines[:12],
                "first_12_text":                 "\n".join(lines[:12]),
            })

print(f"\nExtracted {len(poems)} poems, {len(exclusions)} exclusions")

with open(OUT_DIR / "poems_v2.json", "w", encoding="utf-8") as f:
    json.dump(poems, f, ensure_ascii=False, indent=2)

excl_fields = ["source_file", "poet", "poem_title", "reason"]
with open(OUT_DIR / "exclusions_v2.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=excl_fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(exclusions)

print(f"Saved → output/poems_v2.json")
print(f"Saved → output/exclusions_v2.csv")
