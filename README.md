# SEO Meta Tags Generator

A Claude Code skill that generates SEO-compliant **Title Tags**, **Meta Descriptions**, and **H1 Tags** for one URL or a batch of URLs. Performs live Google SERP research, enforces strict character/keyword/grammar/CTR/uniqueness/no-symbol rules, and delivers results as a formatted Excel (.xlsx) workbook.

Built for **Blog**, **Blog Listicle**, **Collection**, **Product**, and **Service** page types — with page-type-aware tonality and anti-templatization rules baked in.

---

## What it produces

A single Excel file with two sheets:

| Sheet | Contents |
|---|---|
| **Suggestions** | One row per URL with Title, Title Length, Meta Description, Meta Length, H1, Rationale, and Compliance summary. Red conditional formatting flags out-of-range character counts. |
| **SERP Insights** | Top 5 competitor domains, common title patterns, search intent, and notable modifiers per URL. |

See [`outputs/`](outputs/) for 10 sample workbooks across SaaS service pages, ecom collection pages, and blog posts.

---

## Quick start

### 1. Install the skill

Copy `SKILL.md` into Claude Code's skills directory:

**Project-scoped (recommended for testing):**
```
<your-project>/.claude/skills/seo-meta-tags-generator/SKILL.md
```

**User-scoped (available in every project):**
```
~/.claude/skills/seo-meta-tags-generator/SKILL.md
```

### 2. Provide the four required inputs per URL

| Field | Required | Example |
|---|---|---|
| **Page URL or Topic** | ✅ | `https://brand.com/shop/yoga-mats` **OR** `Blog Page on Voice AI for Indian Languages` |
| **Page Type** | ✅ | `Collection` (or Blog / Blog Listicle / Product / Service) |
| **Primary Keyword** | ✅ | `yoga mats` |
| **Secondary Keywords** | ✅ | `non slip yoga mat, eco yoga mat` (use `-` if none — the skill will derive one and flag the row) |

The skill **stops and asks** if any required field is missing — it will not auto-derive page type or keywords.

**Field 1 accepts two shapes:**
- A **real URL** (starts with `http`, `https`, or `www`) — the skill fetches the page, runs SERP research, and does site-wide uniqueness checks.
- A **topic description** ("Blog Page on …", "Category Page on …", "Product Page for …") — the skill skips the page fetch, drafts from topic + primary keyword, and notes that site-wide uniqueness was not verified.

Mixed batches are fine — each row is detected and handled appropriately.

### 3. Invoke it

Paste URLs inline, or share a CSV/TXT/XLSX file:

> *"Generate SEO meta tags for the URLs in `urls.csv`"*

Claude Code auto-triggers the skill, runs the workflow, and writes the Excel to the `outputs/` folder.

---

## CSV input format

```csv
URL,Page_Type,Primary_Keyword,Secondary_Keywords
https://brand.com/shop/yoga-mats,Collection Page,yoga mats,"non slip yoga mat
eco yoga mat"
https://brand.com/product/cork-mat,Product Page,cork yoga mat,"natural cork mat
premium yoga mat"
```

Secondary keywords can be comma-separated within the cell or newline-separated.

---

## Non-negotiable rules (enforced before every output)

### Title Tag
- 50–60 characters (spaces count)
- Contains the primary keyword
- Primary keyword starts before the 50% character mark
- Title does **not** end with the primary keyword
- **No punctuation breaks** — no `:` `-` `|` `,` `–` `—` `/` `()`
- Single continuous grammatical sentence/phrase
- Ecommerce pages include a CTA (Shop / Buy / Order); SaaS pages use Book / Try / Get / Discover
- **No emoji or decorative symbols** — letters, digits, spaces only
- **Site-wide unique** (verified via `site:` query before finalizing)

### Meta Description
- 150–160 characters
- Exactly **2 grammatical sentences**
- Sentence 1 contains 1 occurrence of the **primary keyword**
- Sentence 2 contains 1 occurrence of a **secondary keyword**
- Periods, commas, and apostrophes (for possessives) are allowed; emoji and decorative symbols are not
- Site-wide unique

### H1 Tag
- Contains the primary keyword
- Distinct from the Title
- One H1 per page
- No emoji or decorative symbols

---

## Tonality rules

Outputs must read like an invitation to a non-technical reader, not a feature dump for a developer.

1. **Lead with user pain or benefit, not a feature.**
2. **Plain language only** — no in-house jargon.
3. **The two meta sentences must connect logically** — S1 introduces the pain/benefit, S2 explains how the product delivers it.
4. **Name the brand naturally** where it helps.
5. **Use clear active verbs** — Stop wasting, skip, spot, control, cut, run, find.
6. **Address the reader, not the search engine.**

### Capitalization conventions

- Brand names (`Atom11`, `Gayatri Store`): proper case
- Proper nouns (`Amazon`): capitalized
- Product names (`Sponsored Products`, `Marketing Cloud`, `Buy Box`): Title Case
- Acronyms (`PPC`, `AMC`, `AMS`, `ROAS`, `SQL`, `UK`): ALL CAPS
- Generic terms (`tool`, `software`, `dayparting`): lowercase
- Titles: Title Case
- Metas: sentence case with proper nouns capitalized
- H1: Title Case

---

## Anti-templatization (page-type aware)

When processing a batch, the skill enforces structural diversity across rows — but allows legitimate convention-driven repetition:

| Page Type | CTA repetition |
|---|---|
| Ecom Collection | "Shop" / "Browse" may repeat — site-wide consistency wins |
| Ecom Product | "Buy" / "Order" / "Get" may repeat — transactional convention |
| Service / SaaS | Should vary (Try / Get / Book / See / Discover) |
| Blog | Openings should differ by article angle |
| Blog Listicle | "Top X" pattern allowed; topic anchor must differ |

The **honest substance test** runs on every batch:

> *"If proper nouns were covered, could I still tell which page each row was for from the structure and angle alone?"*

If the answer is no, the batch is redrafted until the answer is yes.

---

## Token-efficient workflow

Seven optimizations baked into the default workflow keep batch token cost around **70–75% lower** than a naive implementation:

| # | Optimization |
|---|---|
| A | SERP snippets only — no competitor page fetches |
| B | Single combined `site:` query for uniqueness |
| C | Batch-level domain inventory (uniqueness once per domain, not per URL) |
| D | Lean page-fetch — title, meta, H1/H2, first 300 words only |
| E | Skip page fetch entirely when user provides H1 + intro + USPs |
| F | Single-pass drafting (no speculative variants) + surgical, capped redraft |
| G | SERP cache by primary keyword (reuse across URLs in the same keyword cluster) |

Typical cost per URL: **$0.04–0.08 on Sonnet** in batch mode.

---

## Repository structure

```
SEO-Meta-Tags-Generation/
├── SKILL.md                  # The skill rules and workflow
├── build_seo_excel.py        # Reference Python script (uses openpyxl)
├── outputs/                  # Sample Excel files from real runs
└── README.md                 # This file
```

`build_seo_excel.py` is a reference implementation showing how the skill builds its Excel workbooks — header styling, conditional formatting, freeze panes, and the article-aware PK compliance check.

---

## Requirements

- **Claude Code** (any recent version)
- **Python 3.10+** with `openpyxl` installed (`pip install openpyxl`)
- **WebFetch + WebSearch** tools available (default in Claude Code)

---

## Versioning

Current skill version: **v2.6**

See the `version:` field in `SKILL.md`'s frontmatter. Changelog is implicit in the git history of this repository.

---

## License

Internal use within Growisto. Contact the SEO team for external sharing.
