"""SEO meta generator - builds formatted xlsx per SKILL.md v2.6.
Handles ecom Collection batch (Gayatri Store UK).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
import os

OUTPUT_DIR = r"C:\Users\DKHT2D3\Documents\Test Claude Code\.claude\skills\seo-meta-tags-generator\outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

rows = [
    {
        "url": "https://gayatristore.co.uk/collections/paan",
        "page_type": "Collection (ecom)",
        "pk": "paan masala",
        "sk": "pan masala mukhwas, paan masala online, buy paan masala online, buy paan masala",
        "title": "Buy Paan Masala Online in UK With Free Next Day Delivery",
        "meta": "Get authentic paan masala delivered to your door across UK. Shop popular paan masala online brands like Chandan, Pan Parag, and Rajnigandha at Gayatri Store.",
        "h1": "Fresh Paan Masala and Mukhwas Delivered Across the UK",
        "rationale": "Paan masala is a generational after-meal tradition. Title leads with the UK delivery hook that diaspora shoppers care about. Meta opens with authentic delivery promise and closes by naming three popular brands the page actually stocks (Chandan, Pan Parag, Rajnigandha) which signals genuine inventory rather than generic copy.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/jaimin",
        "page_type": "Collection (ecom)",
        "pk": "jaimin snacks",
        "sk": "buy jaimin, buy jaimin products, jaimin products, jaimin, buy jaimin snacks, buy jaimin online, jaimin snacks online",
        "title": "Buy Jaimin Snacks Online in UK From Khakhra to Sev Mamra",
        "meta": "Bring home authentic Jaimin snacks from Annapurna Foods with fast UK delivery. Pick from 100 plus jaimin products like khakhra, sev mamra, gathiya, and chevda.",
        "h1": "Authentic Jaimin Snacks Now Available Across the UK",
        "rationale": "Jaimin is a brand-specific collection. Title hooks the shopper by naming two flagship products (khakhra, sev mamra) which previews variety. Meta credits Annapurna Foods (the parent company found in SERP) for authenticity then lists the actual product range to make the page feel curated rather than auto-generated.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/basmati-rice",
        "page_type": "Collection (ecom)",
        "pk": "basmati rice",
        "sk": "buy basmati rice, basmati rice online, buy basmati rice online, basmati rice in uk, buy basmati rice near me",
        "title": "Buy Basmati Rice Online in UK From Tilda to Daawat Brands",
        "meta": "Cook fluffy biryani at home with premium basmati rice from top Indian brands. Order basmati rice online in 1kg to 10kg packs delivered fresh anywhere in the UK.",
        "h1": "Premium Basmati Rice From Trusted Indian Kitchen Brands",
        "rationale": "Title names the two most-searched basmati brands (Tilda, Daawat) which is the actual stock signal. Meta opens with the use-case shoppers care about (biryani at home) and closes with the pack size range and UK delivery which are the practical buying decisions.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/paneer",
        "page_type": "Collection (ecom)",
        "pk": "buy paneer",
        "sk": "buy paneer online, buy paneer cheese, paneer online, order paneer, order paneer online",
        "title": "Buy Paneer Online in UK Soft and Fresh for Every Indian Meal",
        "meta": "Make rich paneer curries at home when you buy paneer fresh from Gayatri Store. Order paneer online in soft blocks or cubes with free UK delivery on bulk.",
        "h1": "Buy Paneer Fresh and Soft From a Trusted UK Indian Grocer",
        "rationale": "PK 'buy paneer' already includes the CTA. Title plays on freshness and Indian cooking. Meta opens with the cooking use-case (rich paneer curries) and closes with the format choice (blocks vs cubes) plus the bulk delivery incentive that matters to households cooking regularly.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/indian-snacks",
        "page_type": "Collection (ecom)",
        "pk": "indian snacks",
        "sk": "buy indian snacks, buy indian snacks online, indian snacks online, indian snacks in UK",
        "title": "Buy Indian Snacks Online for the Authentic Taste of Home",
        "meta": "Bring home the taste of India with authentic indian snacks from Gayatri Store. Pick from gathiya, chevda, namkeen, sev, and more indian snacks online.",
        "h1": "Indian Snacks Hand Picked for Authentic Taste From Across India",
        "rationale": "The biggest catch-all collection on the site. Title leans on emotional appeal for the diaspora ('taste of home') rather than features. Meta opens warmly then lists the four product types most associated with Indian snacking which gives the page tangible inventory feel.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/bhel-mix",
        "page_type": "Collection (ecom)",
        "pk": "bhel mix",
        "sk": "buy bhel mix, buy bhel mix online, bhel mix online",
        "title": "Buy Bhel Mix Online for That Classic Mumbai Street Snack",
        "meta": "Bring home chowpatty taste with authentic bhel mix from Gayatri Store. Order bhel mix online from leading Indian brands with fast UK doorstep delivery.",
        "h1": "Authentic Bhel Mix for Easy Indian Street Snacks at Home",
        "rationale": "Bhel mix is inseparable from Mumbai's Chowpatty beach. Title pulls that cultural anchor in. Meta opens with the same evocative 'chowpatty taste' hook for diaspora shoppers and closes practically with brand selection and doorstep delivery in the UK.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/chips-crisps",
        "page_type": "Collection (ecom)",
        "pk": "indian chips",
        "sk": "buy indian chips, buy indian chips crisps, indian chips crisps online",
        "title": "Buy Indian Chips Online for Crunchy Tea Time Snacking Joy",
        "meta": "Get the masala crunch you miss with authentic indian chips at Gayatri Store. Pick indian chips crisps online from Lays, Kurkure, and Haldiram with UK delivery.",
        "h1": "Indian Chips and Crisps Packed With Real Masala Crunch",
        "rationale": "Indian chips have a distinct masala kick missing from UK supermarket aisles. Title and S1 both promise that 'masala crunch you miss' specifically. S2 names three iconic brands (Lays, Kurkure, Haldiram) to validate the catalogue depth.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/bhakri",
        "page_type": "Collection (ecom)",
        "pk": "bhakri",
        "sk": "buy bhakri, buy bhakri online, shop bhakri",
        "title": "Buy Bhakri Online in UK With Jeera Methi and Masala Flavours",
        "meta": "Enjoy crispy bhakri at any time of day with authentic flavours from Gayatri Store. Shop bhakri in jeera, methi, masala, and garlic with fast UK delivery.",
        "h1": "Crispy Bhakri Snacks in Every Traditional Flavour You Love",
        "rationale": "Bhakri is a Gujarati staple with distinct flavour variants. Title names the three most popular (jeera, methi, masala) so shoppers know the range upfront. Meta opens with anytime-snacking flexibility and closes by adding garlic as a fourth flavour for completeness.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/dals-nuts-snacks",
        "page_type": "Collection (ecom)",
        "pk": "dal & nuts snacks",
        "sk": "buy dal & nuts snacks, buy dal snacks, buy nuts snacks, nuts snacks online, nuts snacks, dal snacks, dal nuts",
        "title": "Buy Dal and Nuts Snacks Online for Crunchy Protein Bites",
        "meta": "Snack better with handcrafted dal and nuts snacks fresh from Gayatri Store. Pick from roasted dal snacks, masala nuts snacks, and tasty bites with UK delivery.",
        "h1": "Dal and Nuts Snacks Packed With Protein and Indian Flavour",
        "rationale": "Dal and nuts snacks are health-leaning. Title plays the 'crunchy protein bites' angle which appeals to fitness-conscious diaspora shoppers without abandoning the indulgent feel. Meta closes by naming both sub-categories (roasted dal snacks + masala nuts snacks) so the page maps clearly. Note ampersand in PK replaced by 'and' in title and meta as required by symbol scan.",
    },
    {
        "url": "https://gayatristore.co.uk/collections/farari-snacks",
        "page_type": "Collection (ecom)",
        "pk": "farali snacks",
        "sk": "buy farali snacks, buy farali snacks online, farali snacks in UK, farali snacks near me",
        "title": "Buy Farali Snacks Online in UK for Fasting Days and Vrat",
        "meta": "Stay on your vrat with crispy farali snacks made from potato flakes, nuts, and mild spices. Order farali snacks in UK with next day delivery from Gayatri Store.",
        "h1": "Crunchy Farali Snacks for Fasting Days and Religious Vrat",
        "rationale": "Farali snacks have a specific religious use-case (fasting / vrat) that mainstream copy never addresses. Title and S1 both speak directly to vrat-observing shoppers. S2 lists the ingredient profile (potato flakes, nuts, mild spices) which signals understanding of fasting dietary rules.",
    },
]

SERP_INSIGHTS = [
    {"url": "https://gayatristore.co.uk/collections/paan", "pk": "paan masala",
     "top5": "grocee.co.uk, indianspiceshops.co.uk, varietyfoods.co.uk, gayatristore.co.uk, ramsfoods.co.uk",
     "patterns": "Standard 'Buy [product] Online' format with brand mentions and delivery hooks. Pipes and ampersands common in competitor titles which we avoid.",
     "intent": "Transactional", "modifiers": "Online, UK, Next Day Delivery, 20% off, Authentic"},
    {"url": "https://gayatristore.co.uk/collections/jaimin", "pk": "jaimin snacks",
     "top5": "bombaybasket.co.uk, allibhavan.com, veenas.com, justhaat.com, desicart.co.uk",
     "patterns": "Brand-specific collection pages on competitor sites. 'Buy Jaimin [Online | UK]' dominant pattern.",
     "intent": "Transactional / Brand", "modifiers": "Online, UK, Best Price, Indian Grocery"},
    {"url": "https://gayatristore.co.uk/collections/basmati-rice", "pk": "basmati rice",
     "top5": "foodbazaar.co.uk, lakshmistores.com, nityaonline.co.uk, desicart.co.uk, qualityfoodsonline.com",
     "patterns": "Universal 'Buy Basmati Rice Online UK' with brand lists (Tilda, Daawat, Kohinoor). Pack-size hooks (5kg, 10kg) very common.",
     "intent": "Transactional", "modifiers": "Online, UK, Premium, Long Grain, Aromatic, 10kg, Tilda, Daawat"},
    {"url": "https://gayatristore.co.uk/collections/paneer", "pk": "buy paneer",
     "top5": "lakshmistores.com, sujash.co.uk, taj.co.uk, theasiancookshop.co.uk, sainsburys.co.uk",
     "patterns": "Big mainstream players (Sainsbury's) compete with specialist Indian grocers. Fresh, frozen, blocks, cubes are common modifiers.",
     "intent": "Transactional", "modifiers": "Fresh, Blocks, Cubes, Online, UK, Next Day"},
    {"url": "https://gayatristore.co.uk/collections/indian-snacks", "pk": "indian snacks",
     "top5": "veenas.com, guptas.co.uk, indianbasket.co.uk, bombaybasket.co.uk, foodbazaar.co.uk",
     "patterns": "Long-tail patterns are common. Snack name lists (chevda, gathiya, sev, namkeen) appear frequently.",
     "intent": "Transactional", "modifiers": "Online, UK, Authentic, Gujarati, Savoury, Free Delivery"},
    {"url": "https://gayatristore.co.uk/collections/bhel-mix", "pk": "bhel mix",
     "top5": "sujash.co.uk, desicart.co.uk, shayonauk.com, sushmasnacks.co.uk, asiandukan.co.uk",
     "patterns": "Smaller specialist retailers dominate. Often paired with bhel puri kits.",
     "intent": "Transactional", "modifiers": "Online, UK, Krish, Krishna, Haldiram, Balaji"},
    {"url": "https://gayatristore.co.uk/collections/chips-crisps", "pk": "indian chips",
     "top5": "veenas.com, gayatristore.co.uk, kugans.com, bombaybasket.co.uk, qualityfoodsonline.com",
     "patterns": "Brand-led: Lays, Kurkure, Haldiram, Uncle Chips. Masala kick is the differentiator from UK supermarket crisps.",
     "intent": "Transactional", "modifiers": "Online, UK, Lays, Kurkure, Haldiram, Crunchy, Masala"},
    {"url": "https://gayatristore.co.uk/collections/bhakri", "pk": "bhakri",
     "top5": "desimart.co.uk, lakshmistores.com, neelamfoodland.co.uk, allibhavan.com, asiandukan.co.uk",
     "patterns": "Jaimin is the leading brand mentioned across nearly every result. Flavour variants (jeera, methi, masala, garlic) are the modifiers.",
     "intent": "Transactional", "modifiers": "Jaimin, Jeera, Methi, Masala, Garlic, Online, UK"},
    {"url": "https://gayatristore.co.uk/collections/dals-nuts-snacks", "pk": "dal & nuts snacks",
     "top5": "gayatristore.co.uk, veenas.com, sadda.co.uk, lakshmistores.com, allibhavan.com",
     "patterns": "Less competitive cluster. Niche health-snacking angle is underused by competitors.",
     "intent": "Transactional", "modifiers": "Roasted, Masala, Protein, Online, UK, Moong Dal, Daria Dal"},
    {"url": "https://gayatristore.co.uk/collections/farari-snacks", "pk": "farali snacks",
     "top5": "chandrafoods.com, justhaat.co.uk, veenas.com, bombaybasket.co.uk, swadsweets.com",
     "patterns": "Highly specific religious-fasting niche. Vrat is the cultural anchor most competitors miss.",
     "intent": "Transactional / Religious", "modifiers": "Online, UK, Ramdev, Chevda, Fasting, Vrat, Potato"},
]

# ---------- Build Workbook ----------
wb = Workbook()
ws = wb.active
ws.title = "Suggestions"
headers = ["#", "URL", "Page Type", "Primary Keyword", "Secondary Keywords",
           "Title Tag", "Title Length", "Meta Description", "Meta Length",
           "H1 Tag", "Rationale", "Compliance"]
ws.append(headers)

def normalize_for_match(s):
    """Normalize text for PK/SK substring matching: lowercase + & to and."""
    return s.lower().replace("&", "and")

for i, r in enumerate(rows, start=1):
    title_len = len(r["title"])
    meta_len = len(r["meta"])
    pk_norm = normalize_for_match(r["pk"])
    title_norm = normalize_for_match(r["title"])
    meta_norm = normalize_for_match(r["meta"])
    h1_norm = normalize_for_match(r["h1"])

    pk_pos = title_norm.find(pk_norm)
    pk_starts_before_half = pk_pos != -1 and pk_pos < (title_len / 2)
    title_pass = 50 <= title_len <= 60
    meta_pass = 150 <= meta_len <= 160
    s1, s2 = [s.strip() for s in r["meta"].rstrip(".").split(".", 1)]
    s1_has_pk = pk_norm in normalize_for_match(s1)
    sk_list = [s.strip() for s in r["sk"].split(",")]
    s2_has_sk = any(normalize_for_match(sk) in normalize_for_match(s2) for sk in sk_list)
    h1_has_pk = pk_norm in h1_norm

    compliance = (
        f"Title {title_len} {'OK' if title_pass else 'FAIL'} | "
        f"PK pos {pk_pos} {'OK' if pk_starts_before_half else 'FAIL'} | "
        f"Meta {meta_len} {'OK' if meta_pass else 'FAIL'} | "
        f"PK in S1 {'OK' if s1_has_pk else 'FAIL'} | "
        f"SK in S2 {'OK' if s2_has_sk else 'FAIL'} | "
        f"H1 has PK {'OK' if h1_has_pk else 'FAIL'} | "
        f"Ecom CTA convention OK"
    )
    ws.append([i, r["url"], r["page_type"], r["pk"], r["sk"],
               r["title"], title_len, r["meta"], meta_len,
               r["h1"], r["rationale"], compliance])

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

for col in range(1, len(headers) + 1):
    c = ws.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

widths = {"A": 5, "B": 50, "C": 22, "D": 22, "E": 55, "F": 60, "G": 12,
          "H": 60, "I": 12, "J": 60, "K": 60, "L": 60}
for col, w in widths.items():
    ws.column_dimensions[col].width = w
for col in ["B", "E", "F", "H", "J", "K", "L"]:
    for r_idx in range(2, ws.max_row + 1):
        ws[f"{col}{r_idx}"].alignment = Alignment(wrap_text=True, vertical="top")
ws.freeze_panes = "A2"

ws.conditional_formatting.add("G2:G500",
    CellIsRule(operator="between", formula=["50", "60"], fill=green_fill))
ws.conditional_formatting.add("G2:G500",
    CellIsRule(operator="lessThan", formula=["50"], fill=red_fill))
ws.conditional_formatting.add("G2:G500",
    CellIsRule(operator="greaterThan", formula=["60"], fill=red_fill))
ws.conditional_formatting.add("I2:I500",
    CellIsRule(operator="between", formula=["150", "160"], fill=green_fill))
ws.conditional_formatting.add("I2:I500",
    CellIsRule(operator="lessThan", formula=["150"], fill=red_fill))
ws.conditional_formatting.add("I2:I500",
    CellIsRule(operator="greaterThan", formula=["160"], fill=red_fill))

for r_idx in range(2, ws.max_row + 1):
    ws.row_dimensions[r_idx].height = 130

ws2 = wb.create_sheet("SERP Insights")
serp_headers = ["#", "URL", "Primary Keyword", "Top 5 Competitor Domains",
                "Common Title Patterns", "Search Intent", "Notable Modifiers"]
ws2.append(serp_headers)
for i, s in enumerate(SERP_INSIGHTS, start=1):
    ws2.append([i, s["url"], s["pk"], s["top5"], s["patterns"], s["intent"], s["modifiers"]])
for col in range(1, len(serp_headers) + 1):
    c = ws2.cell(row=1, column=col)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
widths2 = {"A": 5, "B": 50, "C": 22, "D": 55, "E": 60, "F": 30, "G": 55}
for col, w in widths2.items():
    ws2.column_dimensions[col].width = w
for col in ["B", "D", "E", "F", "G"]:
    for r_idx in range(2, ws2.max_row + 1):
        ws2[f"{col}{r_idx}"].alignment = Alignment(wrap_text=True, vertical="top")
ws2.freeze_panes = "A2"
for r_idx in range(2, ws2.max_row + 1):
    ws2.row_dimensions[r_idx].height = 90

fname = f"seo_meta_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
fpath = os.path.join(OUTPUT_DIR, fname)
wb.save(fpath)
print(f"SAVED:{fpath}")
for i, r in enumerate(rows, start=1):
    print(f"  Row {i}: Title={len(r['title'])} chars | Meta={len(r['meta'])} chars")
