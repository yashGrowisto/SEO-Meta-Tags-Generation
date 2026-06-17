---
name: seo-meta-tags-generator
description: Generate SEO-compliant Title Tags, Meta Descriptions, and H1 Tags for a single URL or a batch of URLs across page types (Blog, Collection, Product, Service, Blog Listicle). Performs live Google SERP research, enforces strict character/keyword/grammar/CTR/uniqueness/no-symbol rules, and delivers results as a formatted Excel (.xlsx) workbook. Triggers when a user shares a URL, a list of URLs, or a CSV/TXT/XLSX file of URLs and asks for title/meta/H1 suggestions, on-page SEO copy, or SERP-optimized tags.
metadata:
  type: skill
  owner: seo-tools@growisto.com
  version: 2.7
---

# SEO Title, Meta Description & H1 Generator

> **Note:** This skill is **independent** from the existing `seo-meta-generator` skill. It enforces its own stricter rule set (no punctuation breaks in titles, 2-sentence meta with explicit keyword placement per sentence, site-wide uniqueness check, no emoji/special symbols) and follows the workflow defined below. Do not delegate to or reuse logic from `seo-meta-generator`.

## Purpose
Produce unique, SERP-optimized **Title Tag**, **Meta Description**, and **H1 Tag** for any URL provided by the user. Each output must be tailored to the page's actual content and page type — never templatized.

## Operating Modes

### Single URL mode
Triggered when the user shares exactly one URL.

### Batch mode
Triggered when the user shares:
- Multiple URLs in the prompt, OR
- A path to a `.csv`, `.txt`, or `.xlsx` file containing URLs (one per row, or a `url` column).

Use `Read` for text/CSV inputs; use `Bash` + Python (`openpyxl` / `pandas`) for `.xlsx` inputs. In batch mode, process every URL; never silently drop one.

## Inputs (Fixed Contract — ask only for these four)

The user is required to provide **exactly these four fields per URL** — nothing more, nothing less. Do not ask the user for brand name, geography, content paste, USPs, page screenshots, Screaming Frog exports, or anything else. Everything else is the skill's job to figure out.

| # | Field | Required | Description |
|---|---|---|---|
| 1 | **Page URL or Topic** | ✅ | Either a real URL (`https://...`) **or** a topic description string (e.g., "Blog Page on Real Time Voice AI", "Category Page on Yoga Mats", "Product Page for Bluetooth Speakers"). See the URL vs Topic Detection section below for how each is handled. |
| 2 | **Page Type** | ✅ | One of: `Blog`, `Blog Listicle`, `Collection`, `Product`, `Service`. |
| 3 | **Primary Keyword** | ✅ | The single main keyword the page targets. |
| 4 | **Secondary Keywords** | ✅ | 1–3 supporting keywords. |

### Accepted input formats

- **Pasted in prompt** — any clear shape works (one block per URL, table, list).
- **CSV / TXT / XLSX file path** — columns: `url`, `page_type`, `primary_keyword`, `secondary_keywords` (comma-separated values inside the cell).

### Validation gate (run before processing)

- If any of the four fields is missing for any URL, **stop and ask the user only for the missing field(s)**. Do not proceed with partial inputs.
- If the user pastes only URLs (no page type, no keywords), prompt them with: *"This skill needs page type, primary keyword, and secondary keywords for each URL. Please provide them — preferably as a CSV with columns: url, page_type, primary_keyword, secondary_keywords."*
- Do **not** auto-derive page type or keywords from the page. The user has explicitly chosen to provide them.
- **Treat `-`, blank, `N/A`, `none`, or `tbd` in the Secondary Keywords cell as "no SK provided".** In that case, do not stop the batch — derive a sensible secondary keyword from the topic and primary keyword, and flag the row with `(derived - SK was '-' in input)` so the user can review.

### URL vs Topic Detection (mandatory — apply per row)

Field 1 accepts two distinct shapes. Detect which one was provided and branch the workflow accordingly.

**It is a real URL when** the value:
- Starts with `http://`, `https://`, or `www.`
- Contains a `.` followed by a top-level domain (`.com`, `.co.uk`, `.io`, `.in`, etc.) AND has no internal spaces in the domain portion
- Looks like a path after the domain (e.g., `/collections/yoga-mats`)

**It is a topic description when** the value:
- Does not start with `http`, `https`, or `www`
- Contains spaces in what would be the domain part
- Reads like English ("Blog Page on …", "Category Page on …", "Product Page for …", "Service Page about …")
- Has no recognizable domain or path structure

#### Behavior — real URL path
1. Lean `WebFetch` on the URL to read title, meta, H1, H2s, first ~300 words, CTA, price, schema.
2. Run live `WebSearch` on the primary keyword for SERP signals.
3. Run `site:<domain>` uniqueness check (or domain inventory check in batch mode).
4. Use page content + SERP + brand voice to draft compliant Title / Meta / H1.
5. Sheet 2 of the Excel includes real top-5 competitor domains, patterns, intent, and modifiers.

#### Behavior — topic description path
1. **Skip the page fetch entirely** — there is no page to fetch.
2. Use the topic description string itself as the source of page context (the words after "Blog Page on" / "Category Page on" / etc. describe what the page is about).
3. Run live `WebSearch` on the primary keyword for SERP signals — this is still possible and recommended unless the user explicitly says "draft from topic only, do not use competitors".
4. **Skip the `site:` uniqueness check** — no domain is known. Note in the row's `Compliance` column that site-wide uniqueness was not verified.
5. **Inferred-brand handling** — without a real URL, brand name cannot be inferred from the domain. Either ask the user for the brand, or draft brand-agnostic copy and flag the row.
6. Sheet 2 still reports SERP insights if WebSearch ran; otherwise it should explicitly state "Not used in this round - drafted purely from topic and primary keyword".

#### Mixed batches
If the input contains a mix of real URLs and topic descriptions, process each row according to its detected type. Do not block the batch.

#### Output transparency
In the final summary returned to the user, explicitly state how many rows were processed as real URLs vs topic descriptions, so the user knows which rows had full SERP + page-content grounding and which were topic-only drafts.

### Everything else is the skill's responsibility

| Skill figures out (no user input needed) | How |
|---|---|
| Page content (H1, copy, USPs, CTAs, price) | Lean `WebFetch` on the URL |
| Brand name | Inferred from domain + on-page branding |
| Geography / language | Inferred from page content + domain TLD |
| Competitor titles & metas | `WebSearch` on primary keyword, snippets only |
| Search intent | Inferred from SERP snippets + page content |
| Power words, modifiers, year cues | Inferred from competitor SERP snippets |
| Site-wide title/meta inventory | `WebSearch` `site:<domain>` query, cached per batch |
| Cannibalization risk | Batch SERP cache + cross-row uniqueness check |

If a page URL is given but content is inaccessible (JS-rendered, blocked), write the row with `Page Type = UNREACHABLE`, leave Title/Meta/H1 blank, put the reason in Rationale, and flag in the final summary. Do not ask the user to paste content — that's outside the fixed contract.

## Required Tools
This skill expects access to: `WebFetch`, `WebSearch`, `Read`, `Write`, `Bash`, `Glob`. If any are unavailable, surface that to the user before starting.

## User-Welcoming Tonality (mandatory — applies to every Title, Meta, H1)

Outputs must read like an invitation to a non-technical reader, not a feature dump for a developer. A seller scanning the SERP should immediately get *what's in it for me*. Apply these rules to every draft before validating SEO compliance.

### Tonality rules
1. **Lead with a user pain or benefit, not a feature.** Open with the outcome the user wants (save spend, find peak hours, skip the SQL grind) — not the technical mechanism behind it.
2. **Plain language only.** Avoid in-house jargon like "bid boosters," "dayparting at scale," "AMS signals," "codeless pipeline." If a seller without SEO/PPC training cannot parse the line in one read, redraft.
3. **The two meta sentences must connect logically.** Sentence 1 introduces the pain or benefit, sentence 2 explains how the product delivers it. They should read like one continuous thought, not two disconnected bullets.
4. **Name the brand naturally where it helps.** A welcoming meta usually contains the brand once (e.g., "Atom11's [PK]…") in a way that feels like a recommendation, not a product code.
5. **Use clear active verbs.** "Stop wasting," "skip," "spot," "control," "cut," "run," "find" — these speak directly to the user. Avoid passive constructions and abstract verbs like "leverages," "delivers," "facilitates," "automates" (in isolation).
6. **Address the seller, not the search engine.** Write the line as if it will be read aloud to a busy Amazon seller scrolling Google. If it sounds robotic when read aloud, redraft.

### Capitalization rules
- **Brand names** ("Atom11"): always proper case.
- **Proper nouns** ("Amazon"): always capitalized.
- **Amazon product / feature names** ("Sponsored Products," "Sponsored Brands," "Sponsored Display," "Marketing Cloud," "Marketing Stream," "Buy Box"): Title Case — they are proper product nouns.
- **Acronyms** ("PPC," "AMC," "AMS," "ROAS," "ACOS," "SQL"): ALL CAPS.
- **Generic terms** ("tool," "software," "dayparting," "audiences," "campaigns"): lowercase unless inside a title that uses Title Case.
- **Titles**: Title Case across content words.
- **Meta descriptions**: sentence case with proper nouns / product names / acronyms capitalized.
- **H1 tags**: Title Case.

### Punctuation allowance (extends earlier symbol rules)
- **Titles**: letters, digits, spaces only — no commas, no apostrophes, no colons, no dashes, no pipes.
- **Meta descriptions**: letters, digits, spaces, periods (`.`), commas (`,`), and apostrophes (`'` for possessives like "Atom11's") allowed. No emoji, no decorative symbols, no other punctuation.
- **H1 tags**: letters, digits, spaces, and apostrophes allowed for possessives. No commas, colons, pipes.

### The welcoming test
> "If a seller saw this title and meta on a Google SERP, would they feel curious and click — or would they bounce because it reads like a tool spec?"

If the answer leans toward "tool spec," the draft fails the welcoming test. Redraft until it passes.

## Anti-Templatization Rules (page-type aware)

LLMs naturally fall into a single sentence skeleton once a structure passes validation for the first row of a batch — then reuse it for every subsequent row. The result is rows that are word-different but structurally identical. That is templatization, and it is forbidden — but with nuance, because some kinds of repetition across a batch are actually correct.

### The honest substance test (apply to every batch before saving)

> "If I covered the proper nouns in each row, could I still tell which page each row was for from the structure and angle alone?"

If **yes** → outputs are page-specific. Pass.
If **no** → outputs are templatized. Redraft until the answer is yes.

### Lexical repetition vs structural repetition

| What's OK (often *correct*) | What's NOT OK |
|---|---|
| Same CTA verb across pages when the page type has a strong convention (Buy/Shop/Order for ecom; Book/Try for SaaS) | Same **sentence skeleton** — i.e., the slot pattern of the title repeating across rows |
| Same brand voice or tone across a domain | Same **meta opening** ("[Brand] [PK] automates...") repeated across rows |
| Same value-prop theme when it is genuinely the brand's positioning | Same **meta closer** ("Book a demo of [SK] trusted by [N] brands") repeated across rows |
| Same modifier (e.g., year "2026") across listicles where the convention applies | Same **H1 skeleton** (e.g., "[PK] for [benefit]") used for more than one row in a batch |

A "skeleton" is the slot pattern — not the words. Example title skeletons:
- `[CTA] [PK] With [USP-noun]`
- `[CTA] [object] In [PK] Without [pain]`
- `[CTA] [PK] At [granularity]`
- `[Adjective phrase] [PK] for [outcome]`

Two rows may share the CTA verb (e.g., both use "Shop") but must NOT share the full skeleton.

### Page-type CTA-verb convention table (applied per batch)

| Page Type | CTA verb across batch | Reason |
|---|---|---|
| Ecom Collection | **Allowed to repeat** (Shop / Browse) | Site-wide consistency improves brand UX |
| Ecom Product | **Allowed to repeat** (Buy / Order / Get) | Strong transactional convention |
| Service / SaaS | **Should vary** (Book / Try / Get / See / Discover / Start) | No fixed convention; variation lifts CTR |
| Blog | **Should vary** | No CTA required; openings should differ by article angle |
| Blog Listicle | Number + topic anchor must differ | "Top X" pattern may repeat; the topic anchor must change |

### Per-batch structural diversity scan (mandatory before saving)

Run this scan after all rows are drafted, before writing the Excel:

1. **Title skeleton check** — group titles by skeleton. No skeleton may be used by more than one row.
2. **Meta S1 opening** — group by the opening 2–3 words (or by structure). If more than 50% of rows in a Service/SaaS batch share an S1 opening, redraft the duplicates. For ecom batches, allow shared openings only if they reflect a documented brand voice pattern (e.g., always opening with the product type).
3. **Meta S2 closing pattern** — same rule as S1 opening, applied to the last 4–6 words of S2.
4. **H1 skeleton check** — same rule as Title.

If any group has more rows than allowed by the convention table, redraft the colliding rows to satisfy a different skeleton — keeping all other non-negotiable rules intact.

### Surgical, capped redraft (extends step 8 of the workflow)

If the diversity scan flags a row, redraft only the failing element (Title, Meta, or H1) — not the whole row. Cap at 2 redraft attempts per element. If still failing after 2 attempts, accept the closest-compliant draft and flag the specific templatization issue in the row's `Compliance` column.

## Token-Efficient Workflow (default)

This workflow is engineered to minimize token cost without compromising the non-negotiable SEO rules. Five cost optimizations are baked in:

- **(A)** Use SERP snippets for competitor analysis — do **not** fetch competitor pages.
- **(B)** Combine uniqueness checks into a single `site:` query.
- **(C)** In batch mode, run uniqueness once per domain (not per URL) and reuse the result.
- **(D)** Lean page-fetch — extract only title, meta, H1/H2s, first ~300 words, price/schema.
- **(E)** User provides keywords and page type upfront → skill never spends tokens deriving them.
- **(F)** Single-pass drafting — draft once, validate; only redraft on actual rule failure. No multi-variant generation.
- **(G)** Domain + keyword clustering in batches — reuse SERP data and inventory across URLs that share a domain or keyword cluster.

### Batch Pre-Processing (run once before the per-URL loop)
1. **Cluster URLs by domain (Optimization G).** Process all URLs from the same domain together so domain-level data can be reused.
2. **Within each domain, sub-cluster by keyword theme.** Quickly scan each URL's slug + (if available) its existing title to group URLs that likely share a primary keyword cluster (e.g., all "running shoes" PDPs vs. all "yoga mats" PDPs). URLs in the same sub-cluster will reuse the same SERP data — do not re-run `WebSearch` for the same primary keyword more than once per batch.
3. **For each unique domain**, run a single `WebSearch` for `site:<domain>` (optionally with a broad keyword theme) to build a lightweight in-memory inventory of existing titles + URLs on that domain. Store it for the whole batch — do not repeat per URL.
4. **Maintain a SERP cache for the batch:** a dictionary keyed by primary keyword that stores the top-10 SERP snippets. Before running `WebSearch` for any URL's primary keyword in step 4 of the per-URL loop, check the cache first; reuse if present.

### Per-URL Loop

1. **Read the user-provided inputs for this URL** — page type, primary keyword, secondary keywords. These are authoritative; do not override them based on page content.
1a. **Detect URL vs Topic** (per the "URL vs Topic Detection" section above). Branch the rest of the workflow based on which shape was provided.
2. **Get page content (lean) — REAL URL ONLY:**
   - Call `WebFetch` once. Extract **only**: `<title>`, `<meta name="description">`, H1, H2 headings, first ~300 words, primary CTA, product/service name, price, and any visible schema fields (Optimization D). Discard the rest.
   - The page content is used only to understand the page's actual offering, USP, and tone — NOT to override the user's page type or keywords.
   - **If the row was detected as a topic description, skip this step entirely.** Use the topic string itself as page context.
3. **Infer supporting context:**
   - **Real URL path:** brand name (from domain or on-page branding), geography / language (from page copy + domain TLD), primary CTA verb for ecommerce/service pages.
   - **Topic path:** brand cannot be inferred — either ask user or draft brand-agnostic copy. Geography / language defaults to global English unless the topic string says otherwise.
4. **Live SERP research using snippets only (Optimizations A + G):**
   - **First check the batch SERP cache** for this primary keyword. If a cached entry exists, reuse it — do not call `WebSearch`.
   - Otherwise, run **one** `WebSearch` on the primary keyword, capture the top 10 organic results' **titles, meta descriptions, and URLs from the snippets**, and store them in the cache under this keyword.
   - **Do NOT `WebFetch` competitor pages.** Snippets already contain their title and meta — that is all we need.
   - From the snippets, infer: search intent, common modifiers (year, "best", "guide", price, location), power words, and any obvious PAA / featured-snippet patterns.
5. **Single-pass drafting (Optimization F):**
   - Generate **one** Title, **one** Meta, **one** H1 — your strongest attempt, informed by the SERP snippets and page content. Do not generate multiple variants speculatively.
   - Move immediately to validation (step 8). Only redraft a specific element if it fails a specific rule.
6. **Uniqueness check (Optimizations B + C):**
   - **Batch mode:** check the draft against the domain inventory built in pre-processing. No new `WebSearch` calls. If a match or near-match exists, redraft to differentiate.
   - **Single-URL mode:** run **one** combined `WebSearch` for `site:<domain> <primary keyword>` — a hit on this query surfaces both keyword cannibalization and likely title duplicates in one call (Optimization B). If a match exists, redraft.
   - If live SERP queries are unavailable, ask the user for the site's existing title/meta inventory (CSV / Screaming Frog export) and check against it.
7. **Symbol scan** — programmatically reject any draft containing emoji, ™, ®, ©, ★, ✓, →, …, or any non-alphanumeric/non-space character (titles and H1s = letters, digits, spaces only; meta descriptions additionally allow `.` and `,` for the two sentences).
8. **Validate** against every non-negotiable criterion below.
   - **Surgical redraft only.** If a specific element fails a specific rule (e.g., title is 62 chars), fix only that element to satisfy only that rule — do not regenerate the whole set. This keeps redraft cost minimal.
   - Cap the redraft loop at **2 attempts per element**. If still failing after 2 surgical fixes, accept the closest-compliant draft and flag the specific rule violation in the row's `Compliance` column.
   - Never output a non-compliant row silently — if compromised, it must be flagged.
9. **Cross-row content uniqueness (batch only)** — ensure no two output rows duplicate each other's Title / Meta / H1 word-for-word. If two URLs on the same domain target the same primary keyword, flag this in the final summary as a cannibalization risk. Append each new draft to the domain inventory so subsequent URLs in the same batch check against it.
10. **Cross-row structural diversity scan (batch only) — MANDATORY before saving the Excel.** Apply the page-type-aware Anti-Templatization Rules section above:
    a. Group titles by skeleton; no two rows may share a skeleton.
    b. Check meta S1 openings; redraft duplicates per the page-type convention table.
    c. Check meta S2 closing patterns; redraft duplicates per the page-type convention table.
    d. Group H1s by skeleton; no two rows may share a skeleton.
    e. Apply the honest substance test ("if proper nouns were covered, could I still tell which page each row was for?") — if no, redraft.
    f. Allow legitimate convention-driven repetition (ecom CTA verbs, brand voice patterns) per the CTA convention table.
11. **Deliver** as a formatted Excel workbook (spec below).

### Why this is cheap
- Zero competitor page fetches (was 3–5 per URL × ~10K tokens each = ~30–50K tokens saved per URL).
- Lean page extraction cuts the target-page fetch payload by ~50%.
- Domain-level uniqueness in batches replaces 2 `WebSearch` calls per URL with 1 call per domain.
- Batch SERP cache eliminates duplicate keyword searches across URLs in the same keyword cluster.
- Single-pass drafting cuts output tokens by ~40–50% versus multi-variant drafting.
- Surgical, capped redraft loop prevents runaway re-generation on tricky pages.
- Structural diversity scan adds one cheap end-of-batch pass that catches the templatization mode-collapse LLMs are prone to, without inflating per-URL cost.
- Net effect: roughly **70–75% lower token cost per URL** in batch runs versus the v2.0 workflow, while preserving every non-negotiable rule and enforcing page-specific outputs.

## Non-Negotiable Criteria

### Title Tag
- Length: **50–60 characters** (count spaces).
- Must contain the **primary keyword**, ideally combined naturally with one secondary keyword.
- Primary keyword must appear **before the 50% character mark** of the title.
- Title **must NOT end with the primary keyword**.
- **No punctuation breaks** — no `:`, `-`, `|`, `,`, `–`, `—`, `/`, `()`.
- Must read as a **single continuous grammatical sentence/phrase** — not two clauses joined without punctuation.
- Grammatically correct.
- For **ecommerce** (Collection / Product) pages, include a relevant **CTA** (e.g., Shop, Buy, Order, Get).
- Must be compelling / click-worthy.
- **No emoji, pictographs, or special symbols** of any kind — including ™, ®, ©, ★, ✓, →, …, currency symbols used decoratively, or any non-ASCII decorative character. Only standard alphanumeric characters and spaces are permitted.
- **Site-wide uniqueness** — the proposed title must not duplicate (or near-duplicate) any existing title on the same domain. Run a uniqueness check (see workflow below) before finalizing to prevent keyword cannibalization.

### Meta Description
- Length: **150–160 characters**.
- Exactly **2 sentences**, grammatically correct.
- **Sentence 1** must contain **1 occurrence of the primary keyword**.
- **Sentence 2** must contain **1 occurrence of a secondary keyword**.
- Must be relevant to the page and compelling enough to drive a click.
- **No emoji or special/decorative symbols** — standard punctuation (`.`, `,`) needed for two grammatical sentences is allowed; decorative characters (★, ✓, →, ™, ®, emoji, etc.) are not.
- **Site-wide uniqueness** — must not duplicate the meta description of any other page on the domain.

### H1 Tag
- Must be relevant to the page content.
- Must contain the **primary keyword**.
- Should differ from the title (avoid duplication) while staying topic-aligned.
- One H1 per page.
- **No emoji or decorative special symbols.**
- **Site-wide uniqueness** — the H1 should not match the H1 of another page on the same domain targeting overlapping keywords.

## Page-Type Guidance (use to differentiate tone, not as templates)

| Page Type        | Intent           | Tone & CTA cues                                    |
|------------------|------------------|----------------------------------------------------|
| Blog             | Informational    | Curiosity, value, learning. No commercial CTA.     |
| Blog Listicle    | Informational/Commercial | Numbered angle, "top", "best", year, comparison. |
| Collection       | Commercial/Transactional | Shop/Browse CTA, range/variety, offers.      |
| Product          | Transactional    | Buy/Order CTA, key spec, USP, price/offer if known. |
| Service          | Commercial       | Hire/Get/Book CTA, outcome-led benefit.            |

## Multi-Page / Batch Requests
Each page's outputs must be unique — derived from that page's actual content. Never reuse the same skeleton/template across pages. The cross-row uniqueness step (workflow step 9) is mandatory for batches.

## Failure Handling
- **URL unreachable** → write the row with `Page Type = UNREACHABLE`, leave Title/Meta/H1 blank, put the reason in Rationale. Flag in the final summary.
- **JS-rendered / empty WebFetch** → in single-URL mode, ask the user to paste H1 + intro paragraph. In batch mode, mark unreachable and request a follow-up paste for those URLs in the summary.
- **WebSearch unavailable or quota hit** → fall back to general SEO best-practice patterns for that page type and explicitly note in `Rationale` that live SERP data was unavailable for that row.
- Never silently drop a URL. Every input URL gets a row in the output file.

## Output: Excel Workbook (.xlsx)

The deliverable is **always a single formatted Excel file** — for single-URL and batch runs alike. Do not dump the table inline in chat; the file is the deliverable.

Build it with Python (`openpyxl`) via `Bash`. If `openpyxl` isn't installed, run `pip install openpyxl` first.

**Filename**: `seo_meta_<YYYYMMDD>_<HHMMSS>.xlsx`.

**Save location (with local override):**
1. Check whether the path `C:\Users\DKHT2D3\Documents\Test Claude Code\.claude\skills\seo-meta-tags-generator\outputs\` (or its POSIX equivalent) exists, OR whether the parent skill folder `C:\Users\DKHT2D3\Documents\Test Claude Code\.claude\skills\seo-meta-tags-generator\` exists.
2. **If the skill folder exists on this system** (i.e., this is the original author's machine), save the workbook to that skill folder under an `outputs/` subdirectory. Create the `outputs/` subdirectory if missing. This keeps generated files grouped with the skill on the author's machine.
3. **If the skill folder does NOT exist at that exact path** (i.e., the skill has been distributed to another user / another machine), save the workbook to the current working directory — the portable default for any other user.

This branch is intentional: it gives the author a tidy local workspace while keeping the skill fully portable for anyone else who installs it.

### Sheet 1: "Suggestions" — one row per URL

| # | URL | Page Type | Primary Keyword | Secondary Keywords | Title Tag | Title Length | Meta Description | Meta Length | H1 Tag | Rationale | Compliance |

- `Title Length` and `Meta Length` are integers (character counts).
- `Rationale`: 2–3 lines on SERP signals used, intent match, CTR hooks.
- `Compliance`: short summary string like `Title ✅ | Meta ✅ | H1 ✅ | Unique ✅` (true enforcement happens silently in step 8 of the workflow — only compliant rows ever reach the file).

### Sheet 2: "SERP Insights" — one row per URL

| # | URL | Primary Keyword | Top 5 Competitor Domains | Common Title Patterns | Search Intent | Notable Modifiers |

### Formatting (apply via openpyxl)

- Header row: bold, fill `#1F4E78`, white font.
- Freeze the top row.
- Column widths: URL = 50; Title / Meta / H1 / Rationale = 60; others 25–35.
- Wrap text on Title, Meta, H1, and Rationale columns.
- Conditional formatting on Sheet 1:
  - `Title Length` cell red fill if `< 50` or `> 60`, green fill otherwise.
  - `Meta Length` cell red fill if `< 150` or `> 160`, green fill otherwise.

### Example openpyxl skeleton

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from datetime import datetime

wb = Workbook()
ws = wb.active
ws.title = "Suggestions"
# write headers, style header row, write data rows,
# set column widths, freeze panes, apply conditional formats,
# then create the "SERP Insights" sheet similarly.

fname = f"seo_meta_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
wb.save(fname)
```

## Final Summary To Return To The User

After saving the file, return a concise summary — do NOT paste the full table:

```
Processed: N URLs
Output: <absolute path to .xlsx>
Compliance: N/N rows passed every non-negotiable rule
Notable: <1–2 lines on cannibalization risk, intent mismatches, or pages that needed deep redrafts>
```

---

## Additions Beyond the Original Brief (recommended — flag to user)

The user's brief is solid; below are items I'd add as a senior SEO based on current SERP behavior. Confirm before applying.

1. **Pixel-width awareness** — Google truncates titles around ~600px and meta descriptions around ~960px (desktop). Character count alone can pass while pixel width fails (e.g., many `W`/`M` chars). Recommend a pixel check on long titles.
2. **Search intent match** — Classify intent before drafting; mismatched intent kills CTR even with perfect on-page rules.
3. **Brand placement convention** — Decide upfront whether brand goes at the start, end, or is omitted. The "no punctuation" rule limits the usual `Page Title | Brand` pattern, so brand has to be woven in grammatically.
4. **Freshness signals** — For listicles and time-sensitive blogs, including the current year (e.g., "in 2026") often lifts CTR — but only if it fits the no-punctuation, single-statement rule.
6. **Power words & emotional triggers** — Words like *proven, essential, complete, ultimate, fast, free* lift CTR; avoid clickbait that violates Google's spam policies.
7. **Sentence case vs Title Case** — Pick one convention per site and apply consistently. Sentence case often feels more natural under the no-punctuation rule.
8. **YMYL / E-E-A-T pages** — Finance, health, legal pages should avoid hype words and lean on authority/trust cues.
9. **Featured snippet & PAA targeting** — For blogs, align H1 (and ideally the first paragraph) to a question pattern Google already surfaces in PAA.
10. **Localization** — If the page targets a specific geo, the geo modifier should be in the title/H1 (e.g., "in India", "UK") — only if it doesn't push the keyword past 50%.
11. **Structured data alignment** — Ensure the title/H1 phrasing matches schema fields (Product name, Article headline) for consistency in rich results.
12. **CTA verbs for ecommerce** — Standardize verbs (Shop, Buy, Order, Get, Discover) per page type for predictability.
14. **A/B testing note** — For high-value pages, recommend testing 2 title variants via tools like Search Console / SEO testing platforms rather than one-shot.

If you'd like, I can fold any of these into the non-negotiable list and update the skill file.
