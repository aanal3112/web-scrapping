# School Registry Scraper — England, Norway, Slovakia & Turkey

This project builds a contact-enriched school directory for four countries.
For each country we start from that country's **official government school
register**, then crawl every school's own website to find its **best contact
email and phone number** — combining the official registry data with
verified, up-to-date contact details pulled directly from each school's site.

Each country required its own dedicated crawler, built around that country's
specific official register, site structure, and access quirks.

---

## How the scraping works

For every country, the pipeline follows the same overall process:

1. **Start from the official government register** — the authoritative,
   nationwide list of schools for that country (CSV, XLS, JSON, or a live
   government API/site, depending on the country).
2. **Visit each school's own website** and crawl the pages most likely to
   list real contact details — Contact / Contact Us / About / Staff /
   Administration pages — not just the homepage.
3. **Extract every email and phone number found**, along with exactly which
   page it came from.
4. **Pick the single best email and phone per school**, prioritizing a
   genuine school/office contact found on an official contact page over a
   personal address or one found incidentally elsewhere on the site.
5. **Merge everything into one row per school** — official registry fields
   plus the verified contact details.

Crawling runs are checkpointed, so a run can be safely paused and resumed
without re-crawling schools already completed, and each school can be
independently retried if its site was temporarily unreachable.

## Country-specific work

### 🇬🇧 England
- Source: DfE "Get Information about Schools" (GIAS) national register.
- Scope: all *Open* institutions, every type (primary, secondary, further-ed, higher-ed, etc.) — ~27,200 schools.
- A number of school sites sit behind Cloudflare or block automated traffic outright; a real-browser + VPN retry pass was built specifically to recover those.

### 🇳🇴 Norway
- Source: the Norwegian School Register (NSR).
- Scope: active, publicly-listed entities — 8,778 records, covering both schools and other registered institution types that share the same national registry.

### 🇸🇰 Slovakia
- Source: the CVTI SR national school register, which is split across 18 separate official institution-type files that had to be merged into one consistent dataset — ~12,160 records.
- Some school sites use a login-gated contact platform (Edupage); those are detected and handled explicitly rather than silently failing.

### 🇹🇷 Turkey
- Source: the MEB (Ministry of National Education) nationwide school directory — 55,116 records, the full national listing.
- Turkish school sites actively block plain automated requests, so this crawler drives a real Chrome browser through a Turkey-based VPN to get past that protection.
- Contact emails on MEB pages are published in an obfuscated, encoded form rather than as plain text; the crawler decodes this to recover the real email address for each school.

---

## Project structure

```
scripts/   Crawler code — one script per country, plus shared helpers
data/      Raw official registry files each crawler starts from
```

See `ENGLAND_SCRAPING_GUIDE.md` and `TURKEY_SCRAPING_GUIDE.md` for step-by-step
run instructions.
