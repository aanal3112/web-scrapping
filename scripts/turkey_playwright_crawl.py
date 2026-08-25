"""
Turkey (MEB) live crawler using a REAL browser engine (Playwright).

Why this exists: meb.k12.tr sites sit behind a JavaScript-based anti-bot
challenge (browser fingerprinting - checks things like audio/video codec
support via canPlayType). A plain HTTP request (requests/Scrapy/curl)
gets served that challenge page forever, since it can't execute the JS
needed to pass it. A real (even headless) browser runs that JS
automatically and gets through to the actual site, exactly like your own
Chrome browser does when you visit these pages by hand.

This must be run on a machine that can actually reach meb.k12.tr (e.g.
through a Turkey VPN) - it cannot run from the AI sandbox.

GAP RESOLVED (2026-07-22): the old data/Okullar ve Diğer Kurumlar.html was
a manually-saved page covering only ADANA/SEYHAN (25 schools). Reading
the live page's own source revealed its table is powered by a DataTables
AJAX endpoint that supports il=0 ("Tümü"/All provinces) - see
turkey_fetch_full_listing.py, which pages through that endpoint directly
to build data/turkey_full_listing.json covering schools nationwide. Run
that script first (same VPN'd-Chrome requirement as this one) to
(re)generate the full listing before running this script for a batch
larger than 25.

CONCURRENCY (2026-07-24): originally tried multiple OS threads each
driving its own tab via Playwright's sync API + ThreadPoolExecutor -
confirmed broken in real testing ("Cannot switch to a different thread"
greenlet errors). Playwright's sync API is not safe to call from
multiple Python threads, even with separate Page objects, since it's
built on greenlets tied to the thread that started it. Fixed by
switching to Playwright's async API instead: multiple pages processed
concurrently via asyncio.gather() on a single thread/event loop, which
is the officially-supported way to run Playwright concurrently in
Python.

One-time setup (only needed once):
    pip3 install playwright
    playwright install chromium

Usage:
    python3 turkey_fetch_full_listing.py     # once, to (re)build the full listing
    python3 turkey_playwright_crawl.py [--limit N]
"""
import argparse
import asyncio
import json
import os
import re
from datetime import date
from urllib.parse import urljoin

from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

import chrome_vpn_recovery
from checkpoint_utils import append_checkpoint, export_xlsx_from_checkpoint, load_checkpoint
from turkey_extract import CHK_RE, parse_listing, resolve_email

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LISTING_FILE = os.path.join(BASE_DIR, "data", "Okullar ve Diğer Kurumlar.html")
FULL_LISTING_FILE = os.path.join(BASE_DIR, "data", "turkey_full_listing.json")


def load_all_schools():
    """Prefers the full nationwide listing (built by
    turkey_fetch_full_listing.py) if it exists; falls back to the old
    25-school ADANA/SEYHAN-only static page otherwise."""
    if os.path.exists(FULL_LISTING_FILE):
        with open(FULL_LISTING_FILE, encoding="utf-8") as f:
            return json.load(f)
    print(f"WARNING: {FULL_LISTING_FILE} not found - falling back to the old 25-school listing. Run turkey_fetch_full_listing.py first for a full-scale batch.")
    return parse_listing(LISTING_FILE)


CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", "turkey_checkpoint.jsonl")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "turkey_output.xlsx")

PAGE_TIMEOUT_MS = 30000
RETRY_PAGE_TIMEOUT_MS = 50000  # longer than the main pass's 30s - the retry
# pass runs at lower concurrency (2 tabs vs 4), so it can afford to wait
# longer per site without hurting total throughput much, and a site that
# timed out once during the main pass under more contention often just
# needs a bit more patience on a quieter second attempt (2026-08-05).
MAX_WAIT_AFTER_LOAD_MS = 3000  # ceiling - only waited this long if the page genuinely needs it
POLL_INTERVAL_MS = 300  # check this often for real content, exit early as soon as it shows up
REQUEST_DELAY = 0.8
CONCURRENCY = 4  # bumped from 3 after a clean 0/30-failure test (2026-08-05) -
# still deliberately modest vs. the 20 used for England/Slovakia/Norway's
# plain-HTTP crawls - each concurrent task here drives a real, heavier browser tab
# in your actual VPN'd Chrome, not a lightweight HTTP request, and different
# schools are independent domains (not the same rate-limit-sensitive endpoint we
# hit building the listing), so concurrency itself is safe - just kept modest to
# not overload your one real browser process with too many simultaneous tabs.

MAX_RECOVERY_ATTEMPTS = 5  # UPDATED (2026-08-21): user's request - when the
# crash-detection logic fires (see worker()'s chrome_dead_flags/
# chrome_dead_streak), automatically attempt full self-recovery (kill Chrome,
# relaunch, reconnect VPN, verify Turkey IP - see chrome_vpn_recovery.py)
# instead of just stopping and asking for a manual restart. Capped at 5
# attempts per run so a genuinely unrecoverable problem (e.g. the VPN
# subscription itself expired, not just a crash) still stops the run for
# real eventually, rather than looping forever.

GONDERMEK_TEXT = "Göndermek için tıklayınız"  # the exact anchor text of the
# email-send link (spec, 2026-08-14) - only a CHK found in THIS specific
# link may be used; unrelated CHK params found elsewhere on a page must be
# ignored, so this is matched by visible link text, not just by scanning
# the raw HTML for anything that looks like an eposta_gonder.php URL.
GONDERMEK_RECHECK_ATTEMPTS = 2  # UPDATED (2026-08-20): found live on 2 real
GONDERMEK_RECHECK_INTERVAL_MS = 500  # records ("CHK not present" that turned
# out to be wrong) - the contact widget carrying this link can render a
# moment AFTER fetch_rendered_html() already captured the page (which only
# waits for domcontentloaded, plus extra polling for the unrelated anti-bot
# challenge screen specifically). If we don't find the link on the first
# read, re-read the same already-loaded page's content up to
# GONDERMEK_RECHECK_ATTEMPTS more times, GONDERMEK_RECHECK_INTERVAL_MS apart,
# before concluding it's genuinely not there. Deliberately NOT a blanket
# wait-longer-on-every-page change - this only costs time on the rare page
# where the link isn't found on the first read (confirmed 2/1000 in the
# sample); every page where it's already there pays nothing extra.
PHONE_RE = re.compile(r"(?<!\d)(?:\+?90[\s.-]*|0[\s.-]*)?\(?\d{3}[\s.-]*\)?[\s.-]*\d{3}[\s.-]*\d{2}[\s.-]*\d{2}(?!\d)")
# UPDATED (2026-08-20, round 1): found live on 2 real records - a single
# optional whitespace between groups was too strict. "0322  256 2363"
# (Alparslan Ilkokulu) has a DOUBLE space between the first two groups, an
# artifact of how BeautifulSoup's get_text(" ") joins text across separate
# HTML tags; "0.322.2392804" (Aksemsettin Ortaokulu) uses "." as the group
# separator (including right after the leading "0") and runs the last two
# groups together as one 7-digit block with no separator at all.
#
# UPDATED (2026-08-20, round 2): found live on 6 more real records once
# the first fix was applied, all genuine page formats:
#  - "(322) 332-7475" / "(553) 639-3089" / "(322) 429-0920" / "(322)
#    502-0147" use "-" as a group separator - added "-" alongside space/dot
#    everywhere.
#  - "0(322 )432 97 70" has a space INSIDE the parens, before the closing
#    ")" - added an [\s.-]* allowance between the area code digits and the
#    closing paren too (was \d{3}\)? with nothing allowed between them).
#  - "903224353534" uses the country code "90" with no leading "+" at all
#    (bare, not "+90") - changed \+90 to \+?90 so the "+" is optional
#    instead of two separate literal alternatives.
# All of these are real page formats, not made up - confirmed live via
# direct fetch, cross-checked against 177+ already-correct phone numbers
# each round to confirm zero regressions.
ADRES_RE = re.compile(r"adresi?", re.IGNORECASE)
EPOSTA_ADRESI_RE = re.compile(r"e[- ]?posta\s*$", re.IGNORECASE)  # "E-Posta
# Adresi" (Email Address) contains the same "adresi" substring as a real
# physical-address label - excluded so it isn't mistaken for one.
ADRES_LABEL_ONLY_RE = re.compile(r"^adresi?\s*:?\s*$", re.IGNORECASE)  # matches
# an element whose ENTIRE text is just the label itself (e.g. a table cell
# containing only "Adres :") - on templates that lay the label and value
# out as separate elements, the very next sibling element is the value,
# cleanly, with no need to guess where it ends via a stop-word list at
# all. Already naturally excludes "E-Posta Adresi" too, since that starts
# with "E-Posta", not "adres" - the ^ anchor rules it out on its own.


NON_ADDRESS_ICON_LABELS = ("yerleşim yeri", "yerlesim yeri", "ulaşım", "ulasim")
# fa-map-marker is reused on some templates for "Yerleşim Yeri" (settlement/
# location description) and "Ulaşım" (how to get there) too, not just the
# real "Adres" - confirmed on Akören Ortaokulu's page, which has 3
# fa-map-marker icons, only one of them the actual address. Any icon match
# whose own text starts with one of these other labels is skipped.


def extract_address(page_text, soup=None):
    """Finds the physical address, labelled "Adres"/"Adresi" (Turkish for
    "address"). Two real bugs fixed here (confirmed live on real MEB pages,
    2026-08-18):
    1. Position misalignment: the previous version lowercased the whole
       page text first (page_text.lower()) to search case-insensitively,
       then used that match's INDEX to slice the ORIGINAL, non-lowered
       text. Turkish "İ" (dotted capital I) does not lowercase to a single
       character in Python's default Unicode handling - "İ".lower() is
       "i" plus a separate combining-dot character, two characters, not
       one - so every "İ" appearing anywhere before the real match point
       silently shifted the index out of alignment with the real text,
       and the slice started from a wrong, later position instead
       (confirmed: this pulled in unrelated trailing footer text on a
       page that had a perfectly normal "Adres: ..." label). Fixed by
       matching case-insensitively directly on the original text
       (re.IGNORECASE) instead of lowering a copy first.
    2. False match on "E-Posta Adresi" (Email Address) - contains the same
       "adresi" substring as a real physical-address label, so a page with
       no real address but a mention of an email address got its trailing
       content (often just page footer boilerplate) mistaken for one.
       Now skipped explicitly.
    Uses the LAST valid (non-email) occurrence, walking backward - a
    school's own labelled contact block (Telefon/E-Posta/Adres) tends to
    sit later in the page than menu/breadcrumb text also containing
    "adres".

    Tried FIRST, before any of the above: some templates lay the label and
    its value out as two separate elements (e.g. a table row with one
    <td>Adres :</td> cell followed by a <td>the actual address</td> cell -
    confirmed on Aladağ Halk Eğitimi Merkezi's page). When that structure
    exists, the value is simply "whatever's in the next sibling element" -
    completely unambiguous, no stop-word list needed at all, and no risk
    of running on into an unrelated next section (confirmed real bug: a
    "VİZYON" (Vision statement) section right after the address in the
    flat page text got swept in by the old character-count approach,
    since "VİZYON" wasn't a recognized stop word).

    If that structure isn't there, falls back to a flat-text search for
    "Adres"/"Adresi" (this also needs the same İ/E-Posta-Adresi fixes as
    above). If no "Adres" label exists anywhere on the page at all, falls
    back further to looking for a map-marker icon (Bootstrap/FontAwesome
    "fa-map-marker") - confirmed some school site templates show the
    address with no text label at all, just this icon next to it (e.g.
    Boztahta Şehit Serdar Yıldırım İlkokulu). The icon fallback is only
    tried when soup is provided and nothing above found anything, since
    it's the least certain signal - the icon is sometimes reused for
    other location-ish info too (see NON_ADDRESS_ICON_LABELS)."""
    if soup is not None:
        for tag in soup.find_all(["td", "div", "span", "li", "p"]):
            if not ADRES_LABEL_ONLY_RE.match(tag.get_text(strip=True)):
                continue
            sib = tag.find_next_sibling()
            while sib is not None:
                value = sib.get_text(" ", strip=True)
                # Skip a sibling that's just punctuation (e.g. a lone ":"
                # on templates that split label/colon/value into three
                # separate elements instead of two) - confirmed live
                # 2026-08-18: taking the very next sibling unconditionally
                # returned ":" as the "address" on several real schools.
                # Keep walking until something with real content turns up.
                if value and not re.fullmatch(r"[:\-–—.\s]*", value):
                    return value[:200]
                sib = sib.find_next_sibling()

    matches = list(ADRES_RE.finditer(page_text))
    for m in reversed(matches):
        preceding = page_text[max(0, m.start() - 15):m.start()]
        if EPOSTA_ADRESI_RE.search(preceding):
            continue
        snippet = page_text[m.end():m.end() + 200]
        snippet = re.sub(r"^\s*:?\s*", "", snippet, count=1)
        # cut at the first of any of these boundary markers
        for stop in ["Telefon", "Faks", "Devamı", "Kurumumuz"]:
            snippet = snippet.split(stop)[0]
        snippet = re.sub(r"\s+", " ", snippet).strip()
        if snippet:
            return snippet[:200]

    if soup is None:
        return ""
    for icon in soup.find_all("i", class_="fa-map-marker"):
        # Template A: icon sits inside a "row"/column layout, the real text
        # is in a sibling column div, not near the icon itself (Boztahta).
        parent_row = icon.find_parent("div", class_=re.compile(r"\brow\b"))
        if parent_row:
            icon_col = icon.find_parent("div", class_=re.compile(r"\bcol-"))
            for col in parent_row.find_all("div", class_=re.compile(r"\bcol-")):
                if col is icon_col:
                    continue
                text = col.get_text(" ", strip=True)
                if text:
                    return text[:200]
        # Template B: icon sits inline in a text flow, "Label - content" or
        # just "content" right after it, up to the next <br> (Akören).
        parts = []
        node = icon.next_sibling
        while node is not None and getattr(node, "name", None) != "br":
            parts.append(node if isinstance(node, str) else node.get_text(" ", strip=True))
            node = node.next_sibling
        text = re.sub(r"\s+", " ", " ".join(parts)).strip()
        label_m = re.match(r"^([^-]{1,30}?)\s*-\s*(.+)$", text)
        if label_m:
            label, rest = label_m.group(1).strip().lower(), label_m.group(2).strip()
            if label in NON_ADDRESS_ICON_LABELS:
                continue
            text = rest
        if text:
            return text[:200]
    return ""


TELEFON_LABEL_RE = re.compile(r"telefon", re.IGNORECASE)
TELEFON_WINDOW_CHARS = 100  # how far past each "Telefon" occurrence to look
# for a phone-shaped match - generous enough for "Telefon \n : \n <number>"
# style markup (lots of whitespace between label/colon/value on some
# templates) without reaching into unrelated content further down the page.
EPOSTA_NEARBY_RE = re.compile(r"e[- ]?posta", re.IGNORECASE)  # UPDATED
# (2026-08-21): a "Telefon" match is only trusted if "E-Posta" also appears
# in the same window - confirmed live (Ahmet Sabanci Anaokulu) that some
# pages have a SECOND, unrelated "Telefon" label inside a staff/service-info
# section (e.g. "Servis Bilgisi / Ismail KARA-Okul Servis Soforu / Telefon:
# 5379796653" - the school BUS DRIVER's personal mobile, not the school's
# own number), which is otherwise indistinguishable from the real widget by
# shape alone (both are just "Telefon" immediately followed by digits). The
# real institutional contact box always has "E-Posta" right next to
# "Telefon" too (it's the same box that holds the email link we already
# decode from) - the staff-info entry never does. Verified live against 84
# real records before adopting this: correctly rejects the 1 confirmed bad
# match, and does not reject any of the other 83 genuinely correct ones.


def extract_phone(page_text, soup=None):
    """Three tiers, most-targeted (safest) first, blind whole-page search
    last (riskiest, only reached if the two targeted methods both find
    nothing):

    1. For every occurrence of "Telefon" in the flattened page text
    (there can be more than one - some templates repeat it in a nav
    menu/article title, or a staff/service-info section, unrelated to the
    real contact widget), search a short window right after it for a
    phone-shaped match that ALSO has "E-Posta" somewhere in that same
    window (see EPOSTA_NEARBY_RE), and use the first occurrence that
    qualifies. Confirmed live 2026-08-20 (Salbas Sehit Yasin Eroglu
    Ilkokulu): the FIRST "Telefon" on the page was a news-article title
    with no number anywhere near it ("Telefon Can Kardesler Orman
    Dostlari Kampta Bir Gun") - skipped correctly since nothing
    phone-shaped follows it; the SECOND was the real "Telefon :
    3225612031 ... E-Posta : ..." widget, found next. Confirmed live
    2026-08-21 (Ahmet Sabanci Anaokulu): without the E-Posta requirement,
    a "Servis Bilgisi" staff entry's "Telefon:5379796653" (the school bus
    driver's personal mobile) would have matched first, since it's
    phone-shaped too - the E-Posta requirement correctly rejects it (no
    E-Posta anywhere near that entry) and moves on to find the real one.

    2. fa-phone icon fallback (user's suggestion, 2026-08-20, confirmed
    live on 3 real records - Adana Rotary, Cerenli, Yesiloba Anaokulu):
    some templates show the phone with NO "Telefon" text label at all,
    just a Bootstrap/FontAwesome "fa-phone" icon next to it - same
    pattern already handled for address via "fa-map-marker". Tried before
    the blind whole-page fallback since it's still a targeted, structural
    search (only looks inside the icon's own container), not a risk of
    grabbing an unrelated number from elsewhere on the page.

    3. Blind PHONE_RE.search(page_text) over the WHOLE page - last resort
    only. UPDATED (2026-08-20): this used to be the ONLY method, and was
    found live to occasionally grab an unrelated hyphenated number from
    elsewhere on the page (a government document reference number,
    "89692170-10.04-E...", matched purely because it happened to fit the
    digit-grouping pattern once "-" became an allowed separator) INSTEAD
    of the real, correctly-labeled phone number further down - since
    .search() stops at the first match in the page, and the false one
    came first. Demoted to last resort, after both targeted methods,
    specifically to avoid that class of false match as often as possible."""
    for label_m in TELEFON_LABEL_RE.finditer(page_text):
        window = page_text[label_m.end():label_m.end() + TELEFON_WINDOW_CHARS]
        m = PHONE_RE.search(window)
        if m and EPOSTA_NEARBY_RE.search(window):
            return re.sub(r"\s+", " ", m.group(0)).strip()

    if soup is not None:
        for icon in soup.find_all("i", class_="fa-phone"):
            parent_row = icon.find_parent("div", class_=re.compile(r"\brow\b"))
            if parent_row:
                icon_col = icon.find_parent("div", class_=re.compile(r"\bcol-"))
                for col in parent_row.find_all("div", class_=re.compile(r"\bcol-")):
                    if col is icon_col:
                        continue
                    m = PHONE_RE.search(col.get_text(" ", strip=True))
                    if m:
                        return re.sub(r"\s+", " ", m.group(0)).strip()

    m = PHONE_RE.search(page_text)
    if m:
        return re.sub(r"\s+", " ", m.group(0)).strip()
    return ""


HARD_PAGE_CEILING_S = 180  # UPDATED (2026-08-21): confirmed live - a single
# page.goto() can hang WAY past its own `timeout_ms` and never raise
# anything at all, if Chrome itself is too resource-strained to even honor
# its own timeout/cancel signal (confirmed: ~25 leftover Chrome processes
# had piled up from an earlier tab crash; the run went completely silent
# for 35+ minutes with zero errors, zero progress, and no crash-detection
# trigger, since the crash-detection mechanism only counts actual
# exceptions - a silent hang produces none to count). This wraps the whole
# call in our OWN hard deadline, independent of Chrome's cooperation - if
# it's not done within 3 minutes (comfortably longer than the normal
# 30-50s timeout, so it never affects a genuinely slow-but-working page),
# we force it to give up and raise, so the failure flows into the same
# retry/crash-detection handling as any other error instead of hanging
# forever with no signal at all.


async def _fetch_rendered_html_once(page, url, timeout_ms=PAGE_TIMEOUT_MS):
    print(f"    -> loading (real browser): {url}")
    await page.goto(url, timeout=timeout_ms, wait_until="domcontentloaded")

    # Poll instead of always sleeping the full ceiling - most pages are
    # ready almost immediately, only the ones still on the anti-bot
    # challenge page need the extra time to resolve and redirect.
    waited_ms = 0
    html = await page.content()
    while "Antibot solution" in html or "Click for continue" in html:
        if waited_ms >= MAX_WAIT_AFTER_LOAD_MS:
            break
        await page.wait_for_timeout(POLL_INTERVAL_MS)
        waited_ms += POLL_INTERVAL_MS
        html = await page.content()

    print(f"    -> rendered page: {len(html)} characters, final URL: {page.url} (waited {waited_ms}ms)")
    return html


async def _dismiss_dialog(dialog):
    try:
        await dialog.dismiss()
    except Exception:
        pass  # already gone by the time we got to it - exactly the race
        # this handler exists to avoid crashing on (see the docstring on
        # _new_page_with_dialog_handler() below).


async def _new_page_with_dialog_handler(context):
    """UPDATED (2026-08-24): confirmed live - a JS dialog (native browser
    alert()/confirm()/etc, seen live on at least one real school page) with
    NO explicit handler registered can crash Playwright's entire
    underlying Node.js driver process outright ("ProtocolError: No dialog
    is showing" - a race in whatever implicit/default dialog handling
    Playwright falls back to without one), taking down every concurrent
    tab at once, not just the one that hit it - confirmed live: all 4 tabs
    failed simultaneously from a dialog on just one of them. Registering
    an explicit handler on every page as soon as it's created dismisses
    any dialog immediately and reliably ourselves, instead of leaning on
    whatever Playwright does by default."""
    page = await context.new_page()
    page.on("dialog", _dismiss_dialog)
    return page


async def fetch_rendered_html(page, url, timeout_ms=PAGE_TIMEOUT_MS):
    try:
        return await asyncio.wait_for(_fetch_rendered_html_once(page, url, timeout_ms), timeout=HARD_PAGE_CEILING_S)
    except asyncio.TimeoutError:
        raise Exception(f"hard ceiling hit - page.goto() still not done after {HARD_PAGE_CEILING_S}s (Chrome likely unresponsive, not just slow)")


CHK_DECODE_RELOAD_RETRIES = 1  # client feedback (2026-08-20): the CHK value
# is regenerated with a fresh random offset on every page load, so a
# "Not decodable" result can be an unlucky one-off for that particular
# load rather than a genuine failure - reload the same 3 pages and try
# again before giving up. Client's own wording was "make a second
# attempt", i.e. one reload (2 attempts total), not an open-ended loop.


async def fetch_school_contact(website, page, timeout_ms=PAGE_TIMEOUT_MS):
    """Thin retry wrapper around _fetch_school_contact_once() - see that
    function's docstring for the actual page-order/CHK-selection logic.
    Only retries the "Not decodable" outcome (a fresh page load gets a
    fresh CHK, so this is the one outcome a reload can plausibly change);
    "CHK not present" and "Page Load Failed" are left to their own
    existing handling (reloading won't add a CHK where the page genuinely
    has none, and page-load failures already get retried at a higher
    level via retry_multi_round())."""
    result = await _fetch_school_contact_once(website, page, timeout_ms)
    attempts = 0
    while result["chk_status"] == "Not decodable" and attempts < CHK_DECODE_RELOAD_RETRIES:
        attempts += 1
        print(f"    -> Not decodable, reloading pages for a fresh CHK (attempt {attempts})")
        result = await _fetch_school_contact_once(website, page, timeout_ms)
    return result


async def _fetch_school_contact_once(website, page, timeout_ms=PAGE_TIMEOUT_MS):
    """Per the new spec (2026-08-14): checks Adi (homepage), then Bilgi
    (About, okulumuz_hakkinda.php), then Iletisim (Contact, iletisim.php),
    in that exact order. On each page, only the link whose VISIBLE ANCHOR
    TEXT is exactly "Göndermek için tıklayınız" may be used - any other
    CHK= parameter present elsewhere on the page must be ignored, even if
    it matches the same eposta_gonder.php URL pattern. Stops at the first
    source whose CHK decodes successfully; Telephone/Address then come
    ONLY from that same winning page - no cascading across pages. If no
    source's link has a CHK, or none of the CHKs found decode, Email,
    Telephone, and Address are all left blank (client's answer, 2026-08-14:
    "Leave Email, Telephone, and Address blank because no source page
    produced a successfully decoded email")."""
    adi_url = website
    bilgi_url = website.rstrip("/") + "/tema/okulumuz_hakkinda.php"
    iletisim_url = website.rstrip("/") + "/tema/iletisim.php"
    pages_to_try = [
        ("Adi", adi_url),
        ("Bilgi", bilgi_url),
        ("Iletisim", iletisim_url),
    ]

    any_page_loaded = False  # distinguishes "genuinely no CHK link on any
    # page" from "network/VPN blip meant we never even got to check" -
    # confirmed live 2026-08-13: a batch of records came back "CHK Not
    # Present" purely because all 3 page loads failed (timeouts,
    # net::ERR_*), not because the school lacks a CHK. Without this flag
    # those get written as a final, non-blank status, which would make
    # retry_failed_records()'s `not r.get("Email")` check treat them as
    # already-resolved and silently skip retrying them forever.
    last_chk = ""
    last_email = ""
    last_status = None  # set only when at least one CHK was found and
    # attempted but failed to decode - distinguishes "Not decodable" from
    # "CHK not present" per the spec's exact status definitions.
    last_source_label = ""
    last_source_url = ""  # which page the undecodable CHK came from - the
    # spec's own field description ("Page providing successful CHK") only
    # covers the Decoded case, but recording this for "Not decodable" too
    # makes manual review far faster (go straight to that page instead of
    # re-checking all 3 sources by hand). Left blank for "CHK not present",
    # since there's no page to point to there at all.

    for label, url in pages_to_try:
        try:
            html = await fetch_rendered_html(page, url, timeout_ms)
        except Exception as e:
            print(f"    -> {label} page failed: {e}")
            if _looks_like_chrome_dead(e):
                # Don't swallow this one - a dead CDP connection will fail
                # the same way on every remaining page too, and letting it
                # propagate up to process_school() is the only way the
                # the chrome_dead_flags/stop_event mechanism in
                # worker() ever sees it. Previously this except caught and
                # neutralized every page-load failure unconditionally, so a
                # real Chrome crash never reached process_school()'s own
                # try/except at all - chrome_dead stayed False forever, and
                # the crash-detection logic (however it was tuned) could
                # never actually fire. Confirmed live 2026-08-13: Chrome
                # died mid-retry-run with zero CHROME_CONNECTION_LOST
                # warning printed.
                raise
            continue
        if not html:
            continue
        any_page_loaded = True

        soup = BeautifulSoup(html, "html.parser")

        target_href = None
        for a in soup.find_all("a", href=True):
            if a.get_text(strip=True) == GONDERMEK_TEXT:
                target_href = a["href"]
                break

        if not target_href:
            # the link may just not have rendered yet - see
            # GONDERMEK_RECHECK_ATTEMPTS above. Only re-checks the page
            # already open in the browser (no new navigation/reload), so
            # this can't accidentally land on a different random CHK draw.
            for _ in range(GONDERMEK_RECHECK_ATTEMPTS):
                await page.wait_for_timeout(GONDERMEK_RECHECK_INTERVAL_MS)
                soup = BeautifulSoup(await page.content(), "html.parser")
                for a in soup.find_all("a", href=True):
                    if a.get_text(strip=True) == GONDERMEK_TEXT:
                        target_href = a["href"]
                        break
                if target_href:
                    break

        if not target_href:
            continue  # link genuinely not on this page - next source

        full_link = urljoin(url, target_href)
        m = CHK_RE.search(full_link)
        if not m:
            continue  # link present but no CHK - next source (spec section 2)

        chk_value = m.group(1)
        email, status = resolve_email(full_link)
        if status == "Decoded":
            page_text = soup.get_text(" ")
            phone = extract_phone(page_text, soup)
            address = extract_address(page_text, soup)
            print(f"    -> decoded on {label}: {email}")
            return {
                "adi_url": adi_url, "bilgi_url": bilgi_url,
                "contact_source": label, "source_url": url,
                "chk": chk_value, "chk_status": status, "email": email,
                "phone": phone, "address": address,
            }
        last_chk, last_email, last_status = chk_value, email, status
        last_source_label, last_source_url = label, url
        # not decodable on this source - continue to next source

    if not any_page_loaded:
        chk_status, email, chk = "Page Load Failed - Retry Needed", "", ""
        source_label, source_url = "", ""
    elif last_status is not None:
        chk_status, email, chk = last_status, last_email, last_chk
        source_label, source_url = last_source_label, last_source_url
    else:
        chk_status, email, chk = "CHK not present", "CHK not present", ""
        source_label, source_url = "", ""

    print(f"    -> no source decoded: {chk_status}")
    return {
        "adi_url": adi_url, "bilgi_url": bilgi_url,
        "contact_source": source_label, "source_url": source_url,
        "chk": chk, "chk_status": chk_status, "email": email,
        "phone": "", "address": "",
    }


FIELDNAMES = [
    "Institution Name", "Province", "District", "Adi URL", "Bilgi URL",
    "Contact/Email Source", "Contact/Email Source URL", "CHK", "CHK Status",
    "Email", "Telephone", "Address", "Extraction Date",
]


# Signatures seen when the real Chrome instance itself has crashed or the
# CDP connection to it has been severed (confirmed 2026-08-04: Chrome
# crashed twice during long unattended runs, generating real crash dumps
# in ~/.config/google-chrome/Crash Reports/completed/). Without detecting
# this, the crawler just keeps trying the remaining thousands of schools
# one by one, every single one failing uselessly for hours before anyone
# notices - see chrome_dead_flags in worker() below.
CHROME_DEAD_PATTERNS = (
    "connection closed", "pipe closed", "target closed",
    "target page, context or browser has been closed",
    "browser has been closed", "websocket",
    "hard ceiling hit",  # UPDATED (2026-08-21): a page.goto() that hangs
    # silently past HARD_PAGE_CEILING_S (see fetch_rendered_html) with no
    # real exception is just as much a sign of an unhealthy Chrome as any
    # of the patterns above - a lone hang could still be one slow site, but
    # repeated hangs across all tabs at once means Chrome itself is
    # the problem, so it should count toward the same stop-the-run check.
)


def _looks_like_chrome_dead(exc):
    return any(p in str(exc).lower() for p in CHROME_DEAD_PATTERNS)


async def process_school(s, page, timeout_ms=PAGE_TIMEOUT_MS):
    print(f"{s['Institution Name']}")
    row = dict(s)
    row["Extraction Date"] = date.today().isoformat()
    chrome_dead = False
    try:
        result = await fetch_school_contact(s["Website"], page, timeout_ms)
        row["Adi URL"] = result["adi_url"]
        row["Bilgi URL"] = result["bilgi_url"]
        row["Contact/Email Source"] = result["contact_source"]
        row["Contact/Email Source URL"] = result["source_url"]
        row["CHK"] = result["chk"]
        row["CHK Status"] = result["chk_status"]
        row["Email"] = result["email"]
        row["Telephone"] = result["phone"]
        row["Address"] = result["address"]
    except Exception as e:
        print(f"    FAILED ({e})")
        row["Adi URL"] = s["Website"]
        row["Bilgi URL"] = s["Website"].rstrip("/") + "/tema/okulumuz_hakkinda.php"
        row["Contact/Email Source"] = ""
        row["Contact/Email Source URL"] = ""
        row["CHK"] = ""
        row["CHK Status"] = ""
        row["Email"] = ""
        row["Telephone"] = ""
        row["Address"] = ""
        chrome_dead = _looks_like_chrome_dead(e)
    return row, chrome_dead


async def retry_failed_records(pages, name_filter=None, stop_event=None, expanded=False):
    """Retry pass over checkpointed failures, spread across `pages` tabs
    (was single-page/sequential - moved to parallel now that the backlog
    of blanks is large enough for that to matter; kept at a lower
    concurrency than the main crawl since it's still a secondary pass).
    Unlike England/Slovakia/Norway, a failed attempt here still gets
    checkpointed (with blank fields) so the school isn't re-attempted
    forever on a permanently broken site - but that also means a
    genuinely transient failure (confirmed: net::ERR_ABORTED from a stale
    browser tab, fixed by opening a fresh tab) would otherwise get stuck
    blank permanently with no way to retry it. This pass re-attempts
    exactly those cases.

    expanded: UPDATED (2026-08-24, user's request) - "full retry", only
    used for the first 2 of the 4 retry rounds (see RETRY_ROUND_EXPANDED).
    When True, retries EVERY record that isn't "Decoded", including
    "Not decodable" - the CHK's offset is randomized on every page load,
    so a "Not decodable" verdict on one draw isn't necessarily permanent;
    a fresh draw on retry might decode fine (same reasoning as the
    existing CHK_DECODE_RELOAD_RETRIES reload-within-one-fetch, just
    applied again at this outer retry-round level). When False (the
    original, narrower behavior, used for rounds 3-4), only retries
    genuine failures: blank Email, or "CHK not present" (which can itself
    be a false negative - see the UPDATED note below)."""
    if stop_event is None:
        stop_event = asyncio.Event()
    all_rows = load_checkpoint(CHECKPOINT_FILE)
    if expanded:
        def needs_retry(r):
            return r.get("CHK Status") != "Decoded"
    else:
        def needs_retry(r):
            return not r.get("Email") or r.get("CHK Status") == "CHK not present"
            # UPDATED (2026-08-20): "CHK not present" used to be excluded
            # here (its Email field holds the literal text "CHK not
            # present", so `not r.get("Email")` is False for it - looked
            # "already resolved" to this filter and was silently never
            # retried). Confirmed live on 2 real records that "CHK not
            # present" can be a false negative (the contact widget hadn't
            # rendered yet when we checked, not a genuine absence - see
            # GONDERMEK_RECHECK_ATTEMPTS) - included here as a safety net
            # for any that still slip past that fix.
    targets = [
        r for r in all_rows.values()
        if (name_filter is None or r["_id"] in name_filter) and needs_retry(r)
    ]
    if not targets:
        print("Nothing to retry.")
        return 0, 0
    schools = [{k: row[k] for k in ("Institution Name", "Province", "District", "Website")} for row in targets]
    print(f"Retrying {len(schools)} failures across {len(pages)} parallel tabs (timeout {RETRY_PAGE_TIMEOUT_MS}ms)...")
    checkpoint_lock = asyncio.Lock()
    chunks = [schools[i::len(pages)] for i in range(len(pages))]
    chrome_dead_flags = [False] * len(pages)
    chrome_dead_streak = [0]
    await asyncio.gather(*[
        worker(chunk, p, checkpoint_lock, stop_event, chrome_dead_flags, chrome_dead_streak, i, RETRY_PAGE_TIMEOUT_MS)
        for i, (chunk, p) in enumerate(zip(chunks, pages))
    ])
    # UPDATED (2026-08-24): "recovered" used to mean "Email is truthy",
    # which silently miscounted "still CHK not present"/"still Not
    # decodable" as recovered too - those statuses store non-empty text in
    # the Email field ("CHK not present" / "CHK present but not
    # decodable"), not blank, even when the retry changed nothing at all.
    # Checking specifically for "CHK Status == Decoded" is the actually
    # correct definition of "this retry helped" - this was always a bit
    # wrong, but became a much bigger inaccuracy once `expanded` retries
    # started including "Not decodable" targets too.
    retried_ids = {s["Website"] for s in schools}
    recovered = sum(1 for r in load_checkpoint(CHECKPOINT_FILE).values() if r["_id"] in retried_ids and r.get("CHK Status") == "Decoded")
    print(f"Recovered {recovered}/{len(schools)} on retry.")
    return recovered, len(schools)


RETRY_ROUND_CONCURRENCY = [4, 3, 2, 2]  # UPDATED (2026-08-24, user's
# request): each retry round now uses a DIFFERENT concurrency instead of
# one fixed value (RETRY_CONCURRENCY = 2) for every round - higher
# concurrency early, when the most records still need retrying (worth
# going faster), tapering down for later rounds where fewer stragglers are
# left (less benefit from more parallelism, and eases off resource/network
# contention on whatever's still failing by then) - then held at 2 for a
# 4th round rather than dropping further, per explicit request. Number of
# rounds is just however many entries are in this list - no separate
# MAX_RETRY_ROUNDS to keep in sync.
MAX_RETRY_ROUNDS = len(RETRY_ROUND_CONCURRENCY)

RETRY_ROUND_EXPANDED = [True, True, False, False]  # UPDATED (2026-08-24,
# user's request): "full retry" - rounds 1-2 retry EVERYTHING that isn't
# "Decoded" (including "Not decodable" - see retry_failed_records()'s
# `expanded` docstring for why that's not necessarily a permanent
# verdict); rounds 3-4 narrow back down to the original, stricter
# criteria (only genuine failures: blank Email or "CHK not present") once
# the earlier, broader passes have already given every "Not decodable"
# record its extra chances. Must be the same length as
# RETRY_ROUND_CONCURRENCY - one entry per round.
assert len(RETRY_ROUND_EXPANDED) == len(RETRY_ROUND_CONCURRENCY)


async def retry_multi_round(context, name_filter=None, stop_event=None):
    """UPDATED (2026-08-24): now creates its own tabs fresh each round
    (via `context`), sized per RETRY_ROUND_CONCURRENCY - previously took a
    fixed, pre-created `pages` list reused unchanged across every round,
    which meant every round ran at the same concurrency no matter what."""
    if stop_event is None:
        stop_event = asyncio.Event()
    for round_num, (concurrency, expanded) in enumerate(zip(RETRY_ROUND_CONCURRENCY, RETRY_ROUND_EXPANDED), start=1):
        label = "full retry - everything except Decoded" if expanded else "narrow retry - genuine failures only"
        print(f"\n--- Retry round {round_num}/{MAX_RETRY_ROUNDS} ({concurrency} concurrent tabs, {label}) ---")
        pages = [await _new_page_with_dialog_handler(context) for _ in range(concurrency)]
        recovered, total = await retry_failed_records(pages, name_filter=name_filter, stop_event=stop_event, expanded=expanded)
        for p in pages:
            try:
                await p.close()
            except Exception:
                # UPDATED (2026-08-24): confirmed live - a genuine Chrome/
                # driver crash (e.g. an unhandled JS dialog taking down the
                # whole Playwright driver process) means the connection is
                # ALREADY dead by the time we get here, so close() itself
                # throws too - this is exactly the situation stop_event
                # just detected, not a new problem. Previously this wasn't
                # caught, so cleanup itself crashed the whole script with
                # an ugly traceback INSTEAD of the clean "stopping early"
                # message the crash-detection was supposed to produce.
                # Safe to ignore here - there's nothing further to clean
                # up on a connection that's already gone.
                pass
        if total == 0 or stop_event.is_set():
            break


# HISTORY: this used to be a single mechanism, CONSECUTIVE_CHROME_DEAD_LIMIT
# (started at 5, raised to 50 on 2026-08-20 after a brief VPN blip alone
# was enough to trip 5 and stop the whole run at 492/1000 even though
# Chrome/VPN were both fine again moments later). UPDATED (2026-08-21):
# split into two independent triggers, both active at once, either one
# stops the run - see worker() below:
#  1. chrome_dead_flags - per-tab, stops the instant every concurrent tab
#     is simultaneously down at once (the acute case - Chrome/VPN just
#     died right now, no ambiguity).
#  2. chrome_dead_streak / CONSECUTIVE_CHROME_DEAD_LIMIT - a shared count
#     of CHROME_DEAD_PATTERNS failures in a row across every tab combined,
#     RESET TO 0 the instant any tab succeeds (same reset behavior as the
#     original mechanism, just at a higher threshold) - stops once 100
#     failures happen back-to-back with zero successes anywhere in
#     between. Explicitly does NOT fire on an alternating pattern like 50
#     failures, then 50 successes, then 50 more failures - each success
#     zeroes the streak, so that pattern never gets anywhere near 100.
CONSECUTIVE_CHROME_DEAD_LIMIT = 100


async def worker(school_chunk, page, checkpoint_lock, stop_event, chrome_dead_flags, chrome_dead_streak, tab_index, timeout_ms=PAGE_TIMEOUT_MS):
    """One coroutine per tab, all running concurrently on the same event
    loop/thread - this is what actually gives real parallelism here,
    unlike the earlier (broken) multi-thread attempt.

    chrome_dead_flags: one boolean per concurrent tab (sized to however
    many tabs are running - CONCURRENCY for the main pass, whatever that
    round's RETRY_ROUND_CONCURRENCY entry is for the retry pass, no
    hardcoded number), each tab only ever
    setting/clearing its OWN slot - stops the run the instant EVERY slot
    is True at once (all tabs genuinely down simultaneously).

    chrome_dead_streak: a single-element list ([count]) shared by every
    tab, counting CHROME_DEAD_PATTERNS failures in a row across all tabs
    combined, reset to 0 by ANY tab's success - stops the run once it
    reaches CONSECUTIVE_CHROME_DEAD_LIMIT, independently of whether the
    per-tab flags ever all line up as True at the same instant."""
    for s in school_chunk:
        if stop_event.is_set():
            break
        row, chrome_dead = await process_school(s, page, timeout_ms)
        chrome_dead_flags[tab_index] = chrome_dead
        chrome_dead_streak[0] = chrome_dead_streak[0] + 1 if chrome_dead else 0
        async with checkpoint_lock:
            append_checkpoint(CHECKPOINT_FILE, s["Website"], row)
        if not stop_event.is_set():
            if all(chrome_dead_flags):
                stop_event.set()
                print(
                    f"\nCHROME_CONNECTION_LOST: all {len(chrome_dead_flags)} tabs are "
                    "currently failing at once - Chrome has likely crashed or "
                    "disconnected. Stopping this run early instead of wasting time on "
                    "doomed attempts. Relaunch Chrome with --remote-debugging-port=9222, "
                    "reconnect the VPN, and re-run the same command to resume from the "
                    "checkpoint.\n"
                )
            elif chrome_dead_streak[0] >= CONSECUTIVE_CHROME_DEAD_LIMIT:
                stop_event.set()
                print(
                    f"\nCHROME_CONNECTION_LOST: {chrome_dead_streak[0]} connection "
                    "failures in a row with no successes in between - Chrome has "
                    "likely crashed or disconnected. Stopping this run early instead "
                    "of wasting time on doomed attempts. Relaunch Chrome with "
                    "--remote-debugging-port=9222, reconnect the VPN, and re-run the "
                    "same command to resume from the checkpoint.\n"
                )
        await asyncio.sleep(REQUEST_DELAY)


async def async_main():
    global CHECKPOINT_FILE, OUTPUT_FILE
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=200, help="TOTAL records this checkpoint should have (default 200), not how many new ones to add - re-running the same command with the same --limit is always safe and just fills in whatever's still missing, never goes past the cap; use 0 for unlimited (all candidates in the listing file/range)")
    parser.add_argument("--retry-only", action="store_true", help="skip crawling new schools - just retry every existing checkpointed failure")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY, help="schools crawled in parallel, each in its own browser tab")
    parser.add_argument("--start-index", type=int, default=0, help="skip straight to this position in the full listing file before filtering for not-yet-done - lets two machines split the list into non-overlapping ranges (e.g. one runs 0-45000, another runs --start-index 45000) without either touching the other's checkpoint")
    parser.add_argument("--checkpoint-suffix", default="", help="write to a separate turkey_checkpoint_<suffix>.jsonl / turkey_output_<suffix>.xlsx instead of the default files - use this together with --start-index so a range-split run doesn't collide with the main checkpoint, merge the files back afterward")
    parser.add_argument("--listing-file", default="", help="path to a custom JSON listing file (e.g. data/turkey_missing_0k_schools.json) to use instead of the default full listing - useful for running only a specific subset of schools; when set, --start-index is ignored since the file is already pre-filtered")
    parser.add_argument("--no-auto-recovery", action="store_true", help="disable chrome_vpn_recovery.py entirely - Chrome must already be running with the VPN connected (same as before that feature existed), and the run just stops with instructions instead of self-healing if Chrome/VPN dies mid-run")
    args = parser.parse_args()

    if args.checkpoint_suffix:
        CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", f"turkey_checkpoint_{args.checkpoint_suffix}.jsonl")
        OUTPUT_FILE = os.path.join(BASE_DIR, "output", f"turkey_output_{args.checkpoint_suffix}.xlsx")
        print(f"Using separate checkpoint: {CHECKPOINT_FILE}")

    async with async_playwright() as p:
        # Attach to a Chrome window that's already running with the VPN
        # connected (via --remote-debugging-port=9222). UPDATED (2026-08-21,
        # user's request - "if i simple run on command it can auto open and
        # launch"): if nothing's running yet, don't just fail - bootstrap
        # the whole thing ourselves using the exact same sequence as
        # mid-run recovery (kill any stray processes, launch Chrome, click
        # the VPN extension's own connect button, verify a genuine Turkey
        # IP - see chrome_vpn_recovery.py). This means a single command now
        # works from a completely cold start, not just VPN reconnects
        # during an already-running crawl.
        if args.no_auto_recovery:
            # Original behavior, untouched - Chrome must already be running
            # with the VPN connected; let connect_over_cdp fail loudly with
            # its own error if not, rather than ever touching
            # chrome_vpn_recovery.py.
            browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        else:
            try:
                browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222", timeout=5000)
                # UPDATED (2026-08-21): confirmed live - Chrome being
                # ALREADY up (e.g. left over from a previous run that got
                # force-killed without also closing Chrome, since Ctrl+C
                # not working forced a manual `pkill` that only targeted
                # the Python script) used to mean this branch just reused
                # it as-is with ZERO verification - could silently start
                # crawling through a disconnected VPN, or one connected to
                # the wrong country. Now runs the exact same check/fix used
                # for a fresh launch, regardless of which path got us here.
                browser = await chrome_vpn_recovery.ensure_turkey_vpn(browser)
            except Exception:
                print("No Chrome running on the debug port yet - bootstrapping from scratch...")
                browser = await chrome_vpn_recovery.recover(p)
        context = browser.contexts[0]

        if args.retry_only:
            await retry_multi_round(context)
        else:
            if args.listing_file:
                listing_path = args.listing_file if os.path.isabs(args.listing_file) else os.path.join(BASE_DIR, args.listing_file)
                with open(listing_path, encoding="utf-8") as f:
                    all_schools = json.load(f)
                print(f"Using custom listing file: {listing_path} ({len(all_schools)} schools)")
            else:
                all_schools = load_all_schools()
                if args.start_index:
                    all_schools = all_schools[args.start_index:]
                    print(f"Starting from index {args.start_index} in the listing file ({len(all_schools)} schools from there to the end).")
            done = load_checkpoint(CHECKPOINT_FILE)
            print(f"{len(all_schools)} schools available in this range, {len(done)} already checkpointed.")

            # UPDATED (2026-08-21): --limit now means "total records for this
            # checkpoint" (a cap), not "add this many new ones" - confirmed
            # live (three times now, across three different situations - a
            # crash mid-run, an internal auto-recovery loop, and simply
            # re-running the same command by hand a second time after a
            # partial run) that computing target as "first --limit NOT-YET-
            # done schools" silently slides the window forward past the
            # intended range once anything is already checkpointed (e.g.
            # --limit 1000 with 500 already done recomputes as "next 1000
            # NOT-done", landing on schools 500-1499 instead of finishing
            # 0-999). Fixed by reversing the order: slice BY POSITION first
            # (all_schools[:limit] - a fixed set that never depends on
            # `done`, so it's identical no matter how many times or in how
            # many separate processes this same command gets run), and only
            # filter for "not yet done" against that fixed slice afterward,
            # inside the loop below. This means re-running the exact same
            # command repeatedly - by hand, after a crash, whenever - is
            # always safe and just fills in whatever's still missing from
            # the same fixed target, never drifting past it.
            target = all_schools[:args.limit] if args.limit else all_schools
            print(f"Checking {len(target)} schools across {args.concurrency} parallel browser tabs "
                  f"({len(done)} already checkpointed within this range)...\n")

            recovery_attempts = 0
            while True:
                done = load_checkpoint(CHECKPOINT_FILE)
                pending = [s for s in target if s["Website"] not in done]
                if not pending:
                    break

                # One dedicated tab per concurrent worker, created upfront -
                # fresh tabs (not reusing whatever state an existing one is
                # in) avoid the net::ERR_ABORTED issue confirmed earlier.
                worker_pages = [await _new_page_with_dialog_handler(context) for _ in range(args.concurrency)]
                chunks = [pending[i::args.concurrency] for i in range(args.concurrency)]

                stop_event = asyncio.Event()
                checkpoint_lock = asyncio.Lock()
                chrome_dead_flags = [False] * args.concurrency
                chrome_dead_streak = [0]
                await asyncio.gather(*[
                    worker(chunk, wp, checkpoint_lock, stop_event, chrome_dead_flags, chrome_dead_streak, i)
                    for i, (chunk, wp) in enumerate(zip(chunks, worker_pages))
                ])

                if not stop_event.is_set():
                    # No name_filter - retry EVERY outstanding blank in the
                    # whole checkpoint, not just this run's own pending set.
                    # Scoping to just `pending` was a real bug: a school that
                    # got a blank checkpoint entry in an earlier run is
                    # already "done" (so excluded from `pending`) but was
                    # then ALSO excluded from retry (not in `pending`'s name
                    # set) - a permanent blind spot for anything that failed
                    # before the current invocation.
                    await retry_multi_round(context, stop_event=stop_event)
                    if not stop_event.is_set():
                        break  # genuinely finished this pass
                    # else: retry_multi_round's OWN dead-Chrome detection
                    # fired (it shares the same stop_event) - deliberately
                    # NOT breaking here, falls through to the exact same
                    # recovery attempt below as the main dispatch's case.
                    # Missing this the first time around was a real gap -
                    # this pass would have just exited without ever trying
                    # to recover.

                # UPDATED (2026-08-21): stop_event firing used to mean "give
                # up, print instructions, exit" - now attempts automatic
                # recovery instead (see chrome_vpn_recovery.py), and if it
                # succeeds, loops back to re-check `pending` (which will have
                # shrunk to whatever's left of `target`) and keeps going,
                # with no manual restart needed at all. --no-auto-recovery
                # opts back out of all of this, straight back to the
                # original stop-and-print-instructions behavior.
                if args.no_auto_recovery:
                    print("\nCHROME_CONNECTION_LOST: Chrome has likely crashed or disconnected. "
                          "--no-auto-recovery is set, so stopping instead of attempting to self-heal. "
                          "Relaunch Chrome with --remote-debugging-port=9222, reconnect the VPN, and "
                          "re-run the same command to resume from the checkpoint.\n")
                    break
                if recovery_attempts >= MAX_RECOVERY_ATTEMPTS:
                    print(f"\nGiving up after {MAX_RECOVERY_ATTEMPTS} automatic recovery attempts - "
                          "stopping for real. Something may be broken in a way that can't self-heal "
                          "(e.g. the VPN subscription itself). Relaunch Chrome, reconnect the VPN "
                          "manually, and re-run the same command to resume from the checkpoint.\n")
                    break
                recovery_attempts += 1
                print(f"\nAttempting automatic recovery ({recovery_attempts}/{MAX_RECOVERY_ATTEMPTS})...")
                try:
                    browser = await chrome_vpn_recovery.recover(p)
                    context = browser.contexts[0]
                    print("Recovery succeeded - Chrome relaunched, VPN reconnected and verified. Resuming...\n")
                except Exception as e:
                    print(f"Recovery attempt failed: {e}")
                    # loop back around - will try recovery again next pass,
                    # up to MAX_RECOVERY_ATTEMPTS total

        # NOT closing context/browser here - this is YOUR real Chrome
        # window, not one we launched. Closing it would shut down your
        # actual browser session.

    total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Turkey", text_columns={"Telephone"})
    print(f"\nTotal checkpointed: {total}. Wrote {OUTPUT_FILE}")

    all_rows = load_checkpoint(CHECKPOINT_FILE).values()
    decoded = [r for r in all_rows if r.get("Email") not in ("", "CHK not present", "CHK present but not decodable")]
    print(f"Schools with a successfully decoded email: {len(decoded)} / {len(all_rows)}")


def main():
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
