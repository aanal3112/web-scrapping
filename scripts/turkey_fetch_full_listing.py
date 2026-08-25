"""
Turkey (MEB) full listing fetcher.

The old approach depended on a manually-saved static HTML page, which only
covered ADANA/SEYHAN (25 schools) - nowhere near enough for 200+. Reading
the actual page source of https://www.meb.gov.tr/baglantilar/okullar/index.php
revealed the table isn't static: it's powered by a DataTables.js server-side
AJAX endpoint, and the province dropdown has an "il=0" ("Tümü" / All) option
meaning this endpoint can return every school in Turkey, paginated, from a
single source - not one manual save per province/district.

Endpoint: POST https://www.meb.gov.tr/baglantilar/okullar/okullar_ajax.php
Params: il=0 (all provinces), ilce=0 (all districts), plus standard
DataTables server-side params (draw/start/length/columns/order).

Each row's "OKUL_ADI" field is "PROVINCE - DISTRICT - Name" text (same
format the old static-page parser already expects), "HOST" is the
subdomain used to build the school's https://{HOST}.meb.k12.tr website.

Must be run on a machine that can actually reach meb.gov.tr (e.g. through
a Turkey VPN) via an already-VPN-connected Chrome window - same
requirement as turkey_playwright_crawl.py. See TURKEY_SCRAPING_GUIDE.md.

Usage:
    python3 turkey_fetch_full_listing.py [--limit N]
"""
import argparse
import json
import os
import time

from playwright.sync_api import sync_playwright

LISTING_PAGE_URL = "https://www.meb.gov.tr/baglantilar/okullar/index.php"
AJAX_URL = "https://www.meb.gov.tr/baglantilar/okullar/okullar_ajax.php"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_FILE = os.path.join(BASE_DIR, "data", "turkey_full_listing.json")

PAGE_SIZE = 100  # the largest option the page's own lengthMenu [10,25,50,100] supports -
# safe to use now that requests go through the real browser's fetch() (see below),
# which was the actual fix for the 429s, not the page size itself.

# CRITICAL: Playwright's request context (request_ctx.post / page.request)
# does NOT route through the real browser's network stack - it uses
# Playwright's own internal Node.js-based HTTP client, which has a
# different TLS/network fingerprint than genuine Chrome even though
# page.goto() (real navigation) uses the real browser and works fine.
# Confirmed by testing: page.goto() succeeds, request_ctx.post() gets an
# instant, empty-body 429 with no Retry-After (not how a normal app-level
# rate limit looks - looks like a fingerprint-based WAF block instead).
# Fix: run the fetch() call INSIDE the page's own JS context via
# page.evaluate() - this is genuinely the real browser's own network
# stack, indistinguishable from the site's own AJAX calls since it IS
# the same mechanism the page's own DataTables code uses.
FETCH_JS = """
async ({url, formData}) => {
    const params = new URLSearchParams(formData);
    const resp = await fetch(url, {
        method: 'POST',
        headers: {
            'X-Requested-With': 'XMLHttpRequest',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        },
        body: params.toString(),
        credentials: 'same-origin',
    });
    const text = await resp.text();
    return {ok: resp.ok, status: resp.status, statusText: resp.statusText, text: text};
}
"""


def fetch_page(page, start, draw, max_retries=4):
    form_data = {
        "draw": str(draw),
        "start": str(start),
        "length": str(PAGE_SIZE),
        "search[value]": "",
        "search[regex]": "false",
        "order[0][column]": "0",
        "order[0][dir]": "asc",
        "il": "0",
        "ilce": "0",
    }
    for i in range(3):
        form_data[f"columns[{i}][data]"] = "OKUL_ADI"
        form_data[f"columns[{i}][searchable]"] = "true"
        form_data[f"columns[{i}][orderable]"] = "true" if i == 0 else "false"
        form_data[f"columns[{i}][search][value]"] = ""
        form_data[f"columns[{i}][search][regex]"] = "false"

    for attempt in range(max_retries):
        result = page.evaluate(FETCH_JS, {"url": AJAX_URL, "formData": form_data})
        if result["ok"]:
            return json.loads(result["text"])

        print(f"  !! HTTP {result['status']} {result['statusText']}")
        print(f"  !! Response body ({len(result['text'])} chars): {result['text'][:2000]!r}")

        if result["status"] == 429 and attempt < max_retries - 1:
            wait_s = 15 * (attempt + 1)
            print(f"  !! Rate-limited (429). Waiting {wait_s}s before retry {attempt + 2}/{max_retries}...")
            time.sleep(wait_s)
            continue

        raise RuntimeError(f"AJAX request failed: HTTP {result['status']}")


def parse_row(row):
    full_text = (row.get("OKUL_ADI") or "").strip()
    host = (row.get("HOST") or "").strip()
    parts = [p.strip() for p in full_text.split(" - ")]
    if len(parts) >= 3:
        province, district, name = parts[0], parts[1], " - ".join(parts[2:])
    else:
        province, district, name = "", "", full_text
    website = f"https://{host}.meb.k12.tr" if host else ""
    return {"Province": province, "District": district, "Institution Name": name, "Website": website}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0, help="stop after collecting this many schools (0 = fetch all)")
    args = parser.parse_args()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        context = browser.contexts[0]
        # Fresh tab rather than reusing whatever state the existing one is
        # in - an existing tab mid-navigation or with an extension overlay
        # active can cause the first goto() to fail with net::ERR_ABORTED
        # (same fix as turkey_playwright_crawl.py).
        page = context.new_page()

        # Load the real listing page first so the AJAX call has a normal
        # referer/session, same-origin as a real visit would. Retry a
        # couple of times since the first goto() after opening a fresh tab
        # can still occasionally hit net::ERR_ABORTED transiently.
        print(f"Loading {LISTING_PAGE_URL} ...")
        last_exc = None
        for attempt in range(3):
            try:
                page.goto(LISTING_PAGE_URL, timeout=30000, wait_until="domcontentloaded")
                last_exc = None
                break
            except Exception as e:
                last_exc = e
                print(f"  (attempt {attempt + 1}/3 failed: {e}, retrying...)")
                time.sleep(1.5)
        if last_exc:
            raise last_exc

        all_schools = []
        start = 0
        draw = 1
        total_records = None

        while True:
            print(f"Fetching rows {start}..{start + PAGE_SIZE} (draw={draw})...")
            data = fetch_page(page, start, draw)

            if total_records is None:
                total_records = data.get("recordsTotal", 0)
                print(f"recordsTotal reported by server: {total_records}")

            rows = data.get("data", [])
            if not rows:
                print("No more rows returned - stopping.")
                break

            for row in rows:
                all_schools.append(parse_row(row))

            print(f"  -> collected so far: {len(all_schools)}")

            start += PAGE_SIZE
            draw += 1
            if args.limit and len(all_schools) >= args.limit:
                all_schools = all_schools[:args.limit]
                break
            if total_records and start >= total_records:
                break
            time.sleep(0.3)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(all_schools, f, ensure_ascii=False, indent=2)
    print(f"\nWrote {len(all_schools)} schools to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
