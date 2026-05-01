# Data Curatorial Statement
## German Poetry Corpus v5 — Information-Theoretic Metrics
**Stevie Miller | May 2026**

---

### Overview

This corpus contains 2,806 poems by 18 German-language poets, drawn from 39 identified collections spanning 1721 to 1924, extracted from TEI XML source files in the TextGrid Digital Library. For each poem, three information-theoretic metrics were computed from scratch using the causal language model dbmdz/german-gpt2: mean token surprisal (−log₂ P(token | context)), mean token entropy (−Σ P(w) log₂ P(w) over the model's next-token distribution), and S₂ (surprisal minus entropy, a measure of excess unpredictability above what the context already anticipates). Metrics were computed over the first 12 verse lines of each poem, using the model's native tokenization without word reconstruction. 295 poems were excluded from metrics computation (prose frames, epigraphs, editorial matter, dedications); all exclusions are documented in a separate sheet. 862 poem-level records carry uncertain publication year flags, primarily because their source collection is a Sammelband or collected-works edition whose original constituent books were published at different times.

---

### 1. Data Cleaning: Problems Anticipated

The primary cleaning challenge is **metadata provenance**. The TEI XML files record the year of the digital scholarly edition, not the original publication year, which is exactly the wrong date for historical analysis. The current pipeline addresses this with a hand-coded lookup table of (poet, collection) pairs mapped to first-publication years, cross-referenced against standard bibliographic sources. This is imperfect: for poets with complex publication histories (Rilke, Goethe, Hölderlin), original collection boundaries are sometimes unclear, and the same poem may have appeared in multiple contexts. The uncertainty flag system captures known ambiguities but will not catch cases where the curator was unaware of the complexity.

A second cleaning issue is **verse-line extraction**. The TEI `<l>` element is used inconsistently across files: some editors mark epigraph lines, stage directions in verse drama, and prose section headers as `<l>`. The pipeline excludes `<l>` elements descended from `<argument>`, `<epigraph>`, `<note>`, `<sp>`, `<stage>`, `<figure>`, `<table>`, and `<bibl>`. This captures the most common contamination patterns but edge cases remain, especially in the Goethe and Klopstock files where structural encoding is inconsistent. Spot-checking against printed editions is the recommended next step.

A third issue is **tokenization granularity**. The model's subword tokenizer splits German compounds in ways that do not respect morphological or metrical structure. A single poetic word like "Abendröte" may become two or three tokens of unequal linguistic weight. This creates noise in per-poem averages and makes cross-poet comparisons partly dependent on vocabulary choices rather than purely on stylistic complexity. Future work could weight tokens by their approximate word-boundary status.

---

### 2. Sample Essay Paragraph

The aggregate metrics suggest that the German lyric tradition does not trend monotonically toward greater linguistic predictability over the long eighteenth and nineteenth centuries. Mean token surprisal across the 39 collections in this corpus ranges from approximately 3.1 bits (Brockes, *Irdisches Vergnügen in Gott*, 1721) to 7.2 bits (Trakl, *Sebastian im Traum*, 1915), and the Spearman correlation between first publication year and mean surprisal, while positive (r ≈ 0.41, p < 0.05 on the certain-year subset), accounts for less than a fifth of the variance. What the data reveal instead is a pattern of clustering: collections from the Expressionist moment of 1910–1920 form a high-surprisal, high-entropy cluster clearly separable from both the Baroque devotional verse of the early corpus and the Biedermeier lyric of the mid-nineteenth century, while collections like Mörike's *Gedichte* (1838) and Storm's *Gedichte* (1852) occupy a remarkably stable middle range across all three metrics. This distribution complicates any narrative that equates historical modernity with linguistic defamiliarization: surprise is not simply a function of when a poet wrote, but of what tradition they were working against and what audience they assumed.

---

### 3. Research Questions for Future Scholars

**a. Meter and surprisal.** Do formally stricter meters (sonnets, odes, Knittelvers) produce lower token surprisal than freer forms? This would require adding a metrical annotation layer to the poem-level sheet, likely through a rule-based scansion tool or manual coding of a subset.

**b. Intra-poem trajectories.** Does surprisal rise or fall across the 12 lines used for scoring? This requires moving from poem-level averages to the per-line or per-token output, which the current pipeline stores in intermediate form but does not surface in the workbook. The `poems_v2.json` file contains the raw text; re-running `02_compute.py` with line-by-line output would enable this.

**c. Canonicity as a correlate.** If a canonicity score (e.g. anthology inclusion frequency) were linked to each poem, the corpus could test whether high-S₂ poems are overrepresented in the canon relative to their collection. The poem-level sheet's `poem_id` field is designed to serve as a join key for this purpose.

**d. Cross-language comparison.** Are these surprisal ranges typical of lyric poetry in other languages, or distinctive to German? A parallel corpus in French or English processed with a comparable monolingual GPT-2 model would enable direct comparison, controlling for model architecture.

**e. Genre-specificity of entropy.** The S₂ measure (surprisal minus entropy) is meant to isolate stylistic deviance from contextual uncertainty. Is high S₂ characteristic of specific genres (city poems, nature lyrics, laments) regardless of period? This would require genre tags on the poem-level records.

---

### 4. Related Datasets

Several comparable datasets exist or are in development:

- **TextGrid Digital Library** (source): the full TEI XML archive from which this corpus is drawn. This dataset is a filtered, metric-enriched subset. Linking back to TextGrid IDs (preserved in the `tei_id` column) enables retrieval of full poem text, editorial notes, and variant readings.
- **Projekt Gutenberg-DE**: overlapping poet coverage, plain-text format, no metadata standardization. Useful for text verification but not for structured analysis.
- **CLARIN-D / DTA (Deutsches Textarchiv)**: broader historical German text corpus including prose; some lyric coverage with more standardized metadata. The `collection_pub_year` fields in this corpus were cross-referenced against DTA bibliographic records where available.
- **GerManC / CAB corpora**: historical German with lemmatization and POS annotation, primarily pre-1800. Could serve as a complementary resource for the Baroque and early Enlightenment poets (Brockes, Haller, Klopstock) where this corpus is thin.

No existing dataset combines information-theoretic LM-based metrics with this level of metadata care for canonical German lyric poetry. The closest precedent is work by Peng et al. on English poetry surprisal (using GPT-2), but that corpus lacks first-publication-year disambiguation and verse-line-only extraction.

---

### 5. Unique Identifiers

Each poem record carries a `poem_id` field constructed as `{poet_slug}_{tei_id}`, where `tei_id` is the TEI XML `xml:id` attribute from the source file. This identifier is stable across pipeline runs and can be used to join poem-level data to any external annotation layer. Collection-level records are identified by `(poet, collection_title)` pairs; no globally unique collection identifier is currently assigned, which is a known limitation for linking to external bibliographic databases (VIAF, GND). A future enhancement would add GND (Gemeinsame Normdatei) authority identifiers for both poets and collections.

---

*Dataset produced May 2026. Model: dbmdz/german-gpt2. Code and output: github.com/stevie-ml/german-poetry-v5.*
