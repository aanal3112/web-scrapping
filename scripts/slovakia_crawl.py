"""
Slovakia (CVTI register) production crawler.
Register gives no contact info (withheld by Ministry instruction), so
email AND phone both come entirely from crawling each school's own site.

Register source: 18 of the 20 CVTI register files (see REGISTER_FILES) -
primary schools, kindergartens, gymnasiums, secondary vocational schools,
conservatories, special-needs secondary/primary/kindergarten schools,
schools attached to healthcare facilities, language schools, art schools,
residential-care/prevention centres, leisure centres, vocational training
centres, outdoor/nature schools, school dormitories, school canteens, and
university faculties. Only vsi_z.xls/vsj_z.xls (university residences/
canteens) are excluded - confirmed no website column exists in either.
Records are walked in file-then-row order and deduplicated by eduid
across all files.

Resumable: see checkpoint_utils.py - each school's result is appended to
output/slovakia_checkpoint.jsonl immediately, so re-running after an
interruption skips already-processed School IDs instead of starting over.
"""
import argparse
import html
import re
import threading
import os
import time
import urllib.robotparser
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from urllib.parse import unquote, urljoin, urlparse

import requests
import urllib3
import xlrd
from bs4 import BeautifulSoup

from checkpoint_utils import append_checkpoint, export_xlsx_from_checkpoint, load_checkpoint

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", "slovakia_checkpoint.jsonl")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "slovakia_output.xlsx")

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
HEADERS = {"User-Agent": UA}
MAX_PAGES = 10  # brief's explicit limit - not a speed knob, do not lower this
MAX_MAIN_CONTACT_SEARCH_PAGES = 20  # client's 2026-08-12 answer: this higher
# cap applies ONLY to the targeted search for the school's "main Contact
# page" (used by the tie-breaking rules) - not a blanket increase to the
# normal MAX_PAGES=10 crawl used for general email/phone extraction.
CRAWL_DELAY = 0.5
TIMEOUT = 10
CONCURRENCY = 20
EXPORT_EVERY = 10

PRIORITY_KEYWORDS = [
    "kontakt", "kontakty", "o-skole", "o_skole", "oskole",
    "vedenie", "zamestnanci", "sekretariat", "contact", "staff",
]

# Client's rewritten best-email/best-phone spec (2026-08-03, Priority_best_email.docx) -
# exact page-name list per country ("Priority pages" row of their table). Distinct from
# PRIORITY_KEYWORDS above (which only affects crawl order); this list decides RANKING -
# whether a found email/phone counts as coming from a "priority page" at all.
SK_PRIORITY_PAGE_KEYWORDS = [
    "kontakt", "kontakty", "o skole", "o škole", "vedenie skoly", "vedenie školy",
    "zamestnanci", "sekretariat", "sekretariát", "contact",
]

# The single "main Contact page" (narrower than SK_PRIORITY_PAGE_KEYWORDS
# above, which also matches O škole/Vedenie školy/Zamestnanci/Sekretariát) -
# used only for the client's tie-breaking rule 4 ("prefer the email
# appearing on the school's main Contact/Kontakt page"), 2026-08-11 clarification.
SK_MAIN_CONTACT_PAGE_KEYWORDS = ["kontakt", "kontakty", "contact"]

IGNORE_LOCAL_PARTS = {"no-reply", "noreply", "webmaster", "privacy", "gdpr"}
PLACEHOLDER_LOCAL_PARTS = {"example", "test", "sample", "yourname", "youremail", "name", "yourdomain", "someone", "user"}
PLACEHOLDER_DOMAINS = {"example.com", "example.org", "example.net", "domain.com", "yoursite.com", "yourdomain.com", "email.com"}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
CF_EMAIL_RE = re.compile(r'data-cfemail="([a-f0-9]+)"')

# Slovak phone formats: +421 9XX XXX XXX (mobile) / 0XX XXX XX XX (landline), 9 digits after leading 0 or +421.
# Digit boundaries on both ends stop this from matching a 10-digit fragment out of a longer
# run of digits (e.g. an IBAN like SK4383000000000203377976 was matching false positives here).
PHONE_RE = re.compile(
    r"(?<!\d)(?:\+421[\s/]?|0)(?:[\s/]?\d){8,9}(?!\d)"
)

# Client's updated tie-breaking spec (Priority_best_email(1).docx,
# 2026-08-11) - order matters: when two candidate emails are otherwise
# equally ranked, the one whose prefix comes FIRST in this list wins.
GENERIC_EMAIL_PREFIXES = [
    "skola", "sekretariat", "info", "kontakt", "office", "podatelna",
    "riaditelstvo", "kancelaria", "sekretarka", "sekretar",
]
DIRECTOR_EMAIL_PREFIXES = ["riaditel", "riaditelka", "director", "principal", "zastupca"]
SPECIALIST_EMAIL_PREFIXES = ["psycholog", "jedalen", "stravovanie", "ekonom", "uctaren"]
CONSUMER_DOMAINS = {
    "gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "icloud.com",
    "azet.sk", "centrum.sk", "zoznam.sk", "zmail.sk",
}
PHONE_CONTEXT_OFFICE = ["sekretariat", "kancelaria", "office", "vratnica"]
PHONE_CONTEXT_SWITCHBOARD = ["ustredna", "spojovatelka", "switchboard", "recepcia"]
PHONE_CONTEXT_DIRECTOR = ["riaditel", "riaditelka", "director"]

SITEMAP_LOC_RE = re.compile(r"<loc>\s*(?:<!\[CDATA\[)?\s*(.*?)\s*(?:\]\]>)?\s*</loc>", re.IGNORECASE)
SITEMAP_PATHS = ("/sitemap.xml", "/sitemap_index.xml")
MAX_SUB_SITEMAPS = 5


def robust_get(session, url, timeout=TIMEOUT):
    """session.get() with two narrow retries (see england_crawl.py for the
    full rationale): SSL verification failures retried once unverified
    (read-only public-info crawl, no sensitive data sent); timeouts/
    connection errors retried up to 2 more times with backoff, since these
    are frequently transient contention from concurrent crawling rather
    than the site actually being down."""
    try:
        return session.get(url, timeout=timeout)
    except requests.exceptions.SSLError:
        return session.get(url, timeout=timeout, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
        last_exc = None
        for attempt in range(2):
            time.sleep(1.5 * (attempt + 1))
            try:
                return session.get(url, timeout=timeout)
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                last_exc = e
        raise last_exc


def get_robot_parser(base_url):
    parsed = urlparse(base_url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = urllib.robotparser.RobotFileParser()
    try:
        resp = requests.get(robots_url, headers=HEADERS, timeout=TIMEOUT)
        rp.parse(resp.text.splitlines() if resp.status_code == 200 else [])
    except requests.RequestException:
        rp.parse([])
    return rp


def can_fetch(rp, url):
    try:
        return rp.can_fetch(UA, url)
    except Exception:
        return True


def score_link(href, text):
    hay = f"{href} {text}".lower()
    for i, kw in enumerate(PRIORITY_KEYWORDS):
        if kw in hay:
            return len(PRIORITY_KEYWORDS) - i
    return 0


def normalize(url):
    return url.split("#")[0].rstrip("/")


def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        chars = [chr(int(encoded[i:i + 2], 16) ^ r) for i in range(2, len(encoded), 2)]
        email = "".join(chars)
        return email if EMAIL_RE.fullmatch(email) else None
    except (ValueError, IndexError):
        return None


# Joomla's built-in "Email Cloaking" anti-spam plugin (used by default on
# a huge number of Joomla sites worldwide, not something specific to any
# one school) builds the real mailto address at runtime via
# document.write() from a JS variable assembled out of literal characters
# mixed with numeric HTML entities, e.g.:
#   var addy1121 = 'r&#105;&#97;d&#105;t&#101;lk&#97;' + '&#64;';
#   addy1121 = addy1121 + 'zsj&#97;k&#117;b&#111;v' + '&#46;' + 'sk';
# A plain HTML/text extraction never sees this (confirmed: even a full
# Playwright/JS-executing browser render doesn't reliably surface it,
# since Joomla's cloaking is a static document.write, not something that
# needs live JS execution to read back out) - the fix is to parse the
# script source directly. Verified against a real Slovak school site
# (zsjakubov.sk/kontakty): decodes all 7 staff emails correctly.
JOOMLA_CLOAK_RE = re.compile(r"\b(addy[a-zA-Z0-9]+)\s*=\s*(?:\1\s*\+\s*)?((?:'[^']*'\s*\+?\s*)+);")


def decode_joomla_cloaked_emails(html_text):
    found = {}
    for m in JOOMLA_CLOAK_RE.finditer(html_text):
        var_name, expr = m.group(1), m.group(2)
        pieces = re.findall(r"'([^']*)'", expr)
        found.setdefault(var_name, []).extend(pieces)
    emails = []
    for pieces in found.values():
        decoded = html.unescape("".join(pieces))
        local = decoded.partition("@")[0]
        if len(local) < 5:
            # Rare edge case (confirmed on zsjakubov.sk/Lenka Verčimáková):
            # some pages split a name between plain visible text ("lenka.
            # vercimá") and the JS-cloaked remainder ("kova@..."), so the
            # cloaking script alone only encodes a truncated fragment.
            # Reconstructing that reliably needs DOM-position tracking
            # this regex-based approach doesn't have; a suspiciously short
            # local part is a cheap, low-risk signal to just skip it
            # rather than surface a wrong/incomplete address.
            continue
        if EMAIL_RE.fullmatch(decoded):
            emails.append(decoded)
    return emails


NON_EMAIL_TLDS = {
    "png", "jpg", "jpeg", "gif", "svg", "webp", "ico", "bmp", "tiff",
    "css", "js", "woff", "woff2", "ttf", "eot",
}  # image/asset filenames using @2x-style retina naming (e.g. logo@2x.png)
# coincidentally match the email pattern - "png" etc. look like a valid TLD.


def is_placeholder(email):
    local, _, domain = email.lower().partition("@")
    if domain.rsplit(".", 1)[-1] in NON_EMAIL_TLDS:
        return True
    return local in PLACEHOLDER_LOCAL_PARTS or domain in PLACEHOLDER_DOMAINS


def _fetch_locs(url, session):
    try:
        resp = robust_get(session, url)
    except requests.RequestException:
        return []
    if resp.status_code != 200:
        return []
    return SITEMAP_LOC_RE.findall(resp.text)


def fetch_sitemap_urls(start_url, domain, session, log):
    parsed = urlparse(start_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    for path in SITEMAP_PATHS:
        locs = _fetch_locs(base + path, session)
        if not locs:
            continue
        is_index = all(loc.lower().endswith(".xml") for loc in locs)
        if is_index:
            page_urls = []
            for sub in locs[:MAX_SUB_SITEMAPS]:
                p = urlparse(sub)
                if p.netloc.lower().lstrip("www.") != domain:
                    continue
                page_urls.extend(_fetch_locs(sub, session))
            locs = page_urls
        urls = [loc for loc in locs if urlparse(loc).netloc.lower().lstrip("www.") == domain and urlparse(loc).scheme in ("http", "https")]
        if urls:
            log.append(f"  sitemap found at {path}: {len(urls)} same-domain URLs")
            return urls
    return []


NON_PHONE_CONTEXT_KEYWORDS = [
    "iban", "účet", "ucet", "adresa", "psč", "psc",
    "ičo", "ico", "dič", "dic", "swift", "bic",
]


def is_non_phone_context(context):
    return any(kw in context for kw in NON_PHONE_CONTEXT_KEYWORDS)


def normalize_phone(raw):
    """Always return the national 0XXXXXXXXX form so +421.../0... variants
    of the same number collapse to a single dict key instead of duplicating."""
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("00421"):
        digits = digits[2:]
    if digits.startswith("421") and len(digits) == 12:
        return "0" + digits[3:]
    if digits.startswith("0") and len(digits) == 10:
        return digits
    return None


def _url_variants(url):
    """www/bare-domain and https/http are each independently a common
    real-world misconfiguration where only one variant actually resolves
    (confirmed: hampsteadprim.camden.sch.uk needs www stripped;
    zsdstal.edu.sk has no working HTTPS at all, only plain HTTP - "HTTPS
    Protocol not configured for this domain"). Try all 4 combinations."""
    parsed = urlparse(url)
    netlocs = [parsed.netloc]
    alt_netloc = parsed.netloc[4:] if parsed.netloc.lower().startswith("www.") else "www." + parsed.netloc
    netlocs.append(alt_netloc)
    schemes = [parsed.scheme, "http" if parsed.scheme == "https" else "https"]
    seen = set()
    variants = []
    for netloc in netlocs:
        for scheme in schemes:
            candidate = parsed._replace(scheme=scheme, netloc=netloc).geturl()
            if candidate not in seen:
                seen.add(candidate)
                variants.append(candidate)
    return variants


META_REFRESH_RE = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]+content=["\']\s*\d+\s*;\s*(?:url=)?([^"\']+)["\']',
    re.IGNORECASE,
)


def _follow_meta_refresh(resp, session, max_hops=3):
    """Some legacy PHP sites redirect via <meta http-equiv="refresh">
    instead of a real HTTP 3xx - requests only follows HTTP-level
    redirects, so this stub page (confirmed on zsjakubov.sk: the plain-
    HTTP homepage returns a 106-byte body doing nothing but this refresh
    to the HTTPS version) otherwise looks like a genuine, contentless 200
    response, and every subsequent fetch of that URL keeps hitting the
    same empty stub. Follow it manually, same as a browser would."""
    for _ in range(max_hops):
        if len(resp.text) > 2000:
            break
        m = META_REFRESH_RE.search(resp.text)
        if not m:
            break
        target = urljoin(resp.url, m.group(1).strip())
        try:
            resp = session.get(target, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        except requests.RequestException:
            break
    return resp


def _record_contact(store, key, url, title, context):
    """PHONE NUMBERS ONLY now (see _record_email() below for email, which
    needs to track every occurrence, not just the best one). Records (url,
    title, context) the first time a phone number is seen, but upgrades to
    a later occurrence if it's on a genuinely better (priority) page and
    the first-seen one wasn't. Confirmed real bug (2026-08-11, client
    feedback): the homepage is always crawled first, so a phone number
    appearing on both the homepage and the actual priority page (very
    common - contact details are often repeated in a footer) was
    permanently attributed to the homepage - worse for phone than email,
    since pick_best_phone() only accepts priority-page phones at all per
    the client's rule, so a phone stuck on a non-priority first sighting
    was being dropped entirely even when it also genuinely appears on the
    priority page."""
    if key not in store:
        store[key] = (url, title, context)
        return
    existing_url, existing_title, _ = store[key]
    if not is_priority_page(existing_url, existing_title) and is_priority_page(url, title):
        store[key] = (url, title, context)


def _record_email(found_emails, addr, url, title):
    """Records EVERY (url, title) an email is found on - found_emails[addr]
    is a list of all occurrences, not just the first or the best. Needed
    for classify_email()'s tier check (an email must count as "on a
    Priority page" if ANY occurrence is, not just the first one seen) and
    for the client's tie-breaking rules (2026-08-11 clarification), which
    need to know whether an email appears on the school's MAIN Contact page
    specifically, and on how many distinct Priority pages it appears on -
    neither answerable from a single stored occurrence."""
    found_emails.setdefault(addr.lower(), []).append((url, title))


def _extract_contacts(soup, html, page_text, url, title, found_emails, found_phones):
    """All email/phone extraction logic for one already-fetched page -
    factored out of the main crawl loop so the targeted main-Contact-page
    search (crawl_site(), 2026-08-12) can run the exact same extraction on
    its one extra fetched page, instead of only grabbing that page's title
    and missing any contacts on it entirely (confirmed real bug in the
    Norway crawler's first version of this: a targeted fetch that only
    recorded the title meant an email genuinely found on that page never
    got credited as being there, so the tie-break logic couldn't see it -
    built correctly here from the start instead of repeating that)."""
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.lower().startswith("mailto:"):
            addr = unquote(href[7:].split("?")[0]).strip()
            if EMAIL_RE.fullmatch(addr) and not is_placeholder(addr):
                _record_email(found_emails, addr.lower(), url, title)

    for m in EMAIL_RE.finditer(page_text):
        addr = m.group(0)
        if not is_placeholder(addr):
            _record_email(found_emails, addr.lower(), url, title)

    for m in EMAIL_RE.finditer(html):
        addr = m.group(0)
        if not is_placeholder(addr):
            _record_email(found_emails, addr.lower(), url, title)

    for meta in soup.find_all("meta"):
        if meta.get("property") in ("og:description",) or meta.get("name") in ("description",):
            for m in EMAIL_RE.finditer(meta.get("content", "")):
                addr = m.group(0)
                if not is_placeholder(addr):
                    _record_email(found_emails, addr.lower(), url, title)

    for m in CF_EMAIL_RE.finditer(html):
        addr = decode_cf_email(m.group(1))
        if addr and not is_placeholder(addr):
            _record_email(found_emails, addr.lower(), url, title)

    for addr in decode_joomla_cloaked_emails(html):
        if not is_placeholder(addr):
            _record_email(found_emails, addr.lower(), url, title)

    for m in PHONE_RE.finditer(page_text):
        norm_phone = normalize_phone(m.group(0))
        if norm_phone:
            ctx_start = max(0, m.start() - 40)
            context = page_text[ctx_start:m.start()].lower()
            if is_non_phone_context(context):
                continue
            _record_contact(found_phones, norm_phone, url, title, context)


def crawl_site(start_url, log):
    resolved_url = start_url
    probe_ok = False
    probe_session = requests.Session()
    for candidate in _url_variants(start_url):
        try:
            probe = probe_session.get(candidate, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
            probe = _follow_meta_refresh(probe, probe_session)
        except requests.RequestException:
            continue
        if probe.status_code == 200 and len(probe.text) > 200:
            if candidate != start_url or probe.url != candidate:
                log.append(f"  {start_url} failed - variant {candidate} worked (-> {probe.url})")
            resolved_url = probe.url
            start_url = probe.url
            probe_ok = True
            break

    parsed_start = urlparse(resolved_url)
    domain = parsed_start.netloc.lower().lstrip("www.")
    if resolved_url != start_url:
        log.append(f"  redirected: {start_url} -> {resolved_url} (domain={domain})")
    rp = get_robot_parser(start_url)

    # Some sites (e.g. EduPage, widely used by Slovak schools) require the
    # visitor to log in to see any content at all - confirmed with a full
    # Playwright/JS-executing browser that this isn't a rendering gap, the
    # school genuinely configured their site as members-only. Detected via
    # the platform's own "/login/" redirect path.
    requires_login = "/login" in urlparse(resolved_url).path.lower()

    visited = set()
    to_visit = [(start_url, 100)]
    found_emails = {}   # email -> (source_url, page_title, surrounding_context_text)
    found_phones = {}   # normalized_phone -> (source_url, page_title, context_text)
    visited_pages = []  # every (url, title) successfully fetched, regardless
    # of whether an email was found there - needed to determine which
    # Priority page genuinely EXISTS on this school's site (client's
    # 2026-08-12 answer: "main Contact page" = Kontakt if it exists, else
    # fall back through the brief's Priority-pages list in order).
    homepage_nav_links = []  # (anchor_text_lower, resolved_url) for every
    # same-domain link found on the homepage specifically (client's answer:
    # "search the priority page text in the anchor text of links in the
    # main link" = the homepage).
    pages_fetched = 0

    session = requests.Session()
    session.headers.update(HEADERS)
    queued = {normalize(start_url)}

    for sm_url in fetch_sitemap_urls(start_url, domain, session, log):
        normf = normalize(sm_url)
        if normf in queued:
            continue
        to_visit.append((sm_url, score_link(sm_url, "")))
        queued.add(normf)

    while to_visit and pages_fetched < MAX_PAGES:
        to_visit.sort(key=lambda x: -x[1])
        url, _ = to_visit.pop(0)
        norm = normalize(url)
        if norm in visited:
            continue
        if not can_fetch(rp, url):
            log.append(f"  robots.txt disallows {url}, skipping")
            continue
        try:
            resp = robust_get(session, url)
            resp = _follow_meta_refresh(resp, session)
        except requests.RequestException as e:
            log.append(f"  fetch failed {url}: {e}")
            continue
        visited.add(norm)
        pages_fetched += 1
        if "/login" in urlparse(resp.url).path.lower():
            requires_login = True
        if resp.status_code != 200 or "text/html" not in resp.headers.get("Content-Type", ""):
            time.sleep(CRAWL_DELAY)
            continue

        soup = BeautifulSoup(resp.text, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else ""
        page_text = soup.get_text(" ")
        visited_pages.append((url, title))
        is_homepage_fetch = normalize(url) == normalize(start_url)

        _extract_contacts(soup, resp.text, page_text, url, title, found_emails, found_phones)

        if pages_fetched < MAX_PAGES or is_homepage_fetch:
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if href.lower().startswith(("mailto:", "tel:", "javascript:")):
                    continue
                full = urljoin(url, href)
                p = urlparse(full)
                if p.netloc.lower().lstrip("www.") != domain or p.scheme not in ("http", "https"):
                    continue
                if is_homepage_fetch:
                    homepage_nav_links.append((a.get_text(" ", strip=True).lower(), full))
                normf = normalize(full)
                if normf in visited or normf in queued:
                    continue
                text = a.get_text(" ", strip=True)
                to_visit.append((full, score_link(href, text)))
                queued.add(normf)

        time.sleep(CRAWL_DELAY)

    log.append(f"  pages fetched: {pages_fetched}, emails found: {len(found_emails)}, phones found: {len(found_phones)}")

    # Determine the school's "main Contact page" for the tie-breaking rules
    # (client's 2026-08-12 answer): try Kontakt first, then fall back
    # through SK_PRIORITY_PAGE_KEYWORDS in the brief's own order. First
    # check pages we ALREADY fetched (free); if none match, check the
    # homepage's own nav links for a matching anchor text (proves the page
    # exists even though we didn't visit it), and fetch it now as one extra
    # targeted page - up to MAX_MAIN_CONTACT_SEARCH_PAGES total, not the
    # normal MAX_PAGES=10. "Not among the pages we crawled" is treated as
    # "doesn't exist" (also client-confirmed), so if nothing matches after
    # this, there is no main Contact page and that tie-break dimension just
    # doesn't apply.
    main_contact_page = None
    for kw in SK_PRIORITY_PAGE_KEYWORDS:
        for u, t in visited_pages:
            if kw in f"{u} {t}".lower():
                main_contact_page = (u, t)
                break
        if main_contact_page:
            break
        if pages_fetched >= MAX_MAIN_CONTACT_SEARCH_PAGES:
            continue
        nav_match_url = next((u for txt, u in homepage_nav_links if kw in txt), None)
        if nav_match_url and normalize(nav_match_url) not in visited:
            try:
                resp = robust_get(session, nav_match_url)
                resp = _follow_meta_refresh(resp, session)
                if resp.status_code == 200 and "text/html" in resp.headers.get("Content-Type", ""):
                    visited.add(normalize(nav_match_url))
                    pages_fetched += 1
                    nav_soup = BeautifulSoup(resp.text, "html.parser")
                    nav_title = nav_soup.title.string.strip() if nav_soup.title and nav_soup.title.string else ""
                    nav_page_text = nav_soup.get_text(" ")
                    visited_pages.append((nav_match_url, nav_title))
                    _extract_contacts(nav_soup, resp.text, nav_page_text, nav_match_url, nav_title, found_emails, found_phones)
                    log.append(f"  targeted fetch for main-contact-page search: {nav_match_url}")
                    main_contact_page = (nav_match_url, nav_title)
            except requests.RequestException as e:
                log.append(f"  targeted fetch failed {nav_match_url}: {e}")
        if main_contact_page:
            break

    return found_emails, found_phones, pages_fetched, requires_login, main_contact_page


def is_priority_page(url, title, keywords=SK_PRIORITY_PAGE_KEYWORDS):
    combined = f"{url} {title}".lower()
    return any(kw in combined for kw in keywords)


def classify_email(email, occurrences, main_contact_page=None):
    """
    Per client's rewritten spec (Priority_best_email.docx, 2026-08-03) plus
    the 2026-08-11 tie-breaking clarification. occurrences is the FULL list
    of (url, title) pairs this email was found on across the whole crawl -
    not just one - since both the tier itself and the tie-break rules need
    to look at everywhere the email appeared, not just wherever it happened
    to be seen first.

    4 tiers, highest to lowest:
      1. Generic school/office prefix, found on a Priority page   -> High
      2. Decision-maker prefix, found anywhere (any crawled page) -> Medium-high
      3. Freemail/consumer domain, found on a Priority page       -> Medium
      4. Freemail/consumer domain, found on any other page        -> Low
    Two cases the client's table doesn't cover (asked, told to treat their
    table as final rather than wait on an answer) - handled as judgment
    calls, ranked below the tiers the table does define:
      - Generic prefix found on a NON-priority page: still an
        institutional-looking address, just not from the ideal page -
        ranked below Decision-maker but above the freemail tiers.
      - Anything matching none of the above (e.g. a personal-name address
        that isn't a listed prefix and isn't freemail): last-resort
        fallback only, same idea as the old code's "nothing better exists"
        catch-all.

    Within a tier, GENERIC_EMAIL_PREFIXES/DIRECTOR_EMAIL_PREFIXES are now
    ORDERED lists per the client's 2026-08-11 tie-breaking rules - an email
    matching an earlier-listed prefix outranks one matching a later-listed
    prefix when their tier is otherwise identical. The remaining tie-break
    chain (main Contact page > appears on multiple Priority pages > first
    encountered during crawl) is resolved in pick_best_email(), which has
    visibility across candidates that this per-email function doesn't.

    Returns a dict with everything pick_best_email() needs to rank and
    report this candidate, or None if excluded entirely.
    """
    local, _, domain = email.partition("@")
    local_l = local.lower()
    domain_l = domain.lower()

    if local_l in IGNORE_LOCAL_PARTS:
        return None

    priority_occurrences = [(u, t) for u, t in occurrences if is_priority_page(u, t)]
    on_priority = bool(priority_occurrences)
    main_contact_norm = normalize(main_contact_page[0]) if main_contact_page else None
    on_main_contact = main_contact_norm is not None and any(normalize(u) == main_contact_norm for u, _ in occurrences)
    num_priority_pages = len({normalize(u) for u, _ in priority_occurrences})

    def best_source():
        if main_contact_norm is not None:
            for u, t in occurrences:
                if normalize(u) == main_contact_norm:
                    return u, t
        if priority_occurrences:
            return priority_occurrences[0]
        return occurrences[0]

    def result(tier_score, etype, confidence, prefix_rank=None, main_contact=on_main_contact,
               n_priority=num_priority_pages, source=None):
        src, title = source if source else best_source()
        return {
            "tier_score": tier_score, "etype": etype, "confidence": confidence,
            "prefix_rank": prefix_rank, "on_main_contact": main_contact,
            "num_priority_pages": n_priority, "source_url": src, "page_title": title,
        }

    # Decision makers - anywhere across the crawl, not priority-page-gated.
    for idx, pref in enumerate(DIRECTOR_EMAIL_PREFIXES):
        if local_l == pref or pref in local_l:
            return result(400, "Director / Headteacher", "Medium-high", prefix_rank=idx)

    # Generic school/office contact.
    for idx, pref in enumerate(GENERIC_EMAIL_PREFIXES):
        if local_l == pref or pref in local_l:
            if on_priority:
                return result(500, "Generic School Contact", "High", prefix_rank=idx)
            return result(300, "Generic School Contact", "Medium", prefix_rank=idx,
                          main_contact=False, n_priority=0, source=occurrences[0])

    # Freemail/consumer domains - rank depends on whether it's on a priority page.
    if domain_l in CONSUMER_DOMAINS:
        if on_priority:
            return result(200, "Consumer / Freemail", "Medium", main_contact=False, n_priority=0, source=priority_occurrences[0])
        return result(100, "Consumer / Freemail", "Low", main_contact=False, n_priority=0, source=occurrences[0])

    # Specialist contacts (psycholog@, jedalen@, etc.) - not in the client's
    # new table but kept as a low-priority fallback, same as before.
    for pref in SPECIALIST_EMAIL_PREFIXES:
        if pref in local_l:
            return result(50, "Specialist Contact", "Low", main_contact=False, n_priority=0, source=occurrences[0])

    # Last resort - doesn't match any defined tier (e.g. a personal-name
    # address on the school's own domain).
    return result(10, "Other / Unclassified", "Low", main_contact=False, n_priority=0, source=occurrences[0])


def pick_best_email(found_emails, school_website="", main_contact_page=None):
    """found_emails: {email: [(url, title), ...]} - every occurrence, per
    _record_email(). main_contact_page: (url, title) crawl_site() determined
    for this school, or None. Ranks candidates by: tier score -> prefix
    order within the tier (client's 2026-08-11 ordered lists) -> appears on
    the main Contact page -> appears on more than one Priority page -> first
    encountered during the crawl (dict insertion order) - the client's
    tie-breaking spec."""
    candidates = []
    for order, (email, occurrences) in enumerate(found_emails.items()):
        cls = classify_email(email, occurrences, main_contact_page)
        if cls is None:
            continue
        prefix_rank = cls["prefix_rank"] if cls["prefix_rank"] is not None else 999
        sort_key = (
            -cls["tier_score"], prefix_rank, 0 if cls["on_main_contact"] else 1,
            -cls["num_priority_pages"], order,
        )
        candidates.append((sort_key, email, cls))

    if not candidates:
        for order, (email, occurrences) in enumerate(found_emails.items()):
            src, title = occurrences[0]
            cls = {"etype": "Specialist Contact", "confidence": "Low", "source_url": src, "page_title": title}
            candidates.append(((0, 999, 1, 0, order), email, cls))

    if not candidates:
        return None

    candidates.sort(key=lambda c: c[0])
    _, email, cls = candidates[0]
    return {
        "best_email": email, "best_email_type": cls["etype"], "confidence": cls["confidence"],
        "source_url": cls["source_url"], "page_title": cls["page_title"],
    }


def classify_phone(context):
    for kw in PHONE_CONTEXT_OFFICE:
        if kw in context:
            return (100, "School office")
    for kw in PHONE_CONTEXT_SWITCHBOARD:
        if kw in context:
            return (80, "Main switchboard")
    for kw in PHONE_CONTEXT_DIRECTOR:
        if kw in context:
            return (60, "Director's office")
    return (10, "Unlabelled/general contact number")


def pick_best_phone(found_phones):
    """Client instruction (2026-08-03): only report a phone number if it
    was found on a Priority page - phones from any other page are ignored
    entirely, not just ranked lower."""
    if not found_phones:
        return None
    candidates = []
    for phone, (src, title, context) in found_phones.items():
        if not is_priority_page(src, title):
            continue
        rank, ptype = classify_phone(context)
        candidates.append((rank, phone, ptype, src, title))
    if not candidates:
        return None
    candidates.sort(key=lambda x: -x[0])
    rank, phone, ptype, src, title = candidates[0]
    return {"best_phone": phone, "best_phone_type": ptype, "source_url": src, "page_title": title}


DATA_DIR = os.path.join(BASE_DIR, "data")

# Slovakia's register is split into 12 separate files by institution type.
# vs_z.xls (universities) is deliberately excluded here: it has no
# per-institution website column at all (its only per-school sheet,
# "Zoznam VŠ", lists name/address but no URL), so there's nothing in that
# source for this pipeline to crawl.
#
# The other 11 files aren't identical in layout despite looking similar:
#   - zs_z.xls: eduid before kodsko, "jazyk" column present (14 cols)
#   - GYM/KON/SOS/SPECSS/ms/speczz: kodsko before eduid, "jazyk" present (14 cols)
#   - jazyk/zus: kodsko before eduid, no "jazyk" column (13 cols) - shifts
#     name/address/municipality/website one position left vs the above
#   - speczs/specms: extra disclaimer row before the header (data starts
#     one row later) plus an extra "disability type" column that shifts
#     website one position right. Each also has a second sheet ("Špeciálne
#     triedy...") listing special-needs classes inside otherwise-regular
#     schools already counted in zs_z.xls/ms_z.xls - deliberately not
#     merged in, to avoid double-counting those schools.
#
# Tuple: (filename, sheet name, school type label, first data row,
#         eduid col, name col, address col, municipality col, website col)
REGISTER_FILES = [
    ("zs_z.xls", "Základné školy", "Základná škola", 2, 1, 8, 9, 10, 12),
    ("ms_z.xls", "Materské školy", "Materská škola", 2, 2, 8, 9, 10, 12),
    ("GYM_Z.XLS", "Gymnáziá a SŠŠ", "Gymnázium", 2, 2, 8, 9, 10, 12),
    ("SOS_Z.XLS", "Stredné odborné školy", "Stredná odborná škola", 2, 2, 8, 9, 10, 12),
    ("KON_Z.XLS", "Konzervatóriá", "Konzervatórium", 2, 2, 8, 9, 10, 12),
    ("SPECSS_Z.XLS", "Špeciálne stredné školy", "Špeciálna stredná škola", 2, 2, 8, 9, 10, 12),
    ("speczz_z.xls", "Školy pri zdravotníc. zariad", "Škola pri zdravotníckom zariadení", 2, 2, 8, 9, 10, 12),
    ("jazyk_z.xls", "Jazykové školy", "Jazyková škola", 2, 2, 7, 8, 9, 11),
    ("zus_z.xls", "ZUS", "Základná umelecká škola", 2, 2, 7, 8, 9, 11),
    ("speczs_z.xls", "Špeciálne základné školy", "Špeciálna základná škola", 3, 1, 8, 9, 10, 13),
    ("specms_z.xls", "Špeciálne materské školy", "Špeciálna materská škola", 3, 1, 8, 9, 10, 13),
    # 4 more added 2026-08-03 (client confirmed there are 20 files on CVTI's
    # register, not the 12 we originally found) - these 4 are genuine
    # standalone institutions with their own eduid/name/address/website.
    ("ovz_z.xls", "OVZ", "Zariadenie výchovnej prevencie a náhradnej výchovy", 2, 2, 7, 8, 9, 11),
    ("cvc_z.xls", "CVC", "Centrum voľného času", 4, 2, 7, 8, 9, 11),
    ("spv_z.xls", "SOP", "Stredisko odbornej praxe", 2, 2, 7, 8, 9, 11),
    ("svp_z.xls", "Školy v prírode", "Škola v prírode", 2, 2, 7, 8, 9, 11),
    # 3 more added 2026-08-04 (client: "include all files, where you find a
    # website you can scrape"). dm_z.xls/sj_z.xls were originally skipped as
    # facilities of already-counted schools, but the client wants them
    # merged in regardless - both have real per-row websites (128/198 and
    # 1,462/4,619 respectively). vs_z.xls was originally thought to have no
    # website at all - that was checking the wrong sheet ("Zoznam VŠ", 35
    # university-level rows, no website); its "Vysoké školy" sheet (137
    # rows, one per faculty) does have a real website column, 131/137
    # populated - corrected and included. vsi_z.xls/vsj_z.xls (university
    # residences/canteens) remain excluded - confirmed no website column
    # exists anywhere in either file, genuinely nothing to scrape.
    ("dm_z.xls", "Školské internáty", "Školský internát", 2, 2, 7, 8, 9, 11),
    ("sj_z.xls", "školské jedálne a výdajne", "Zariadenie školského stravovania", 3, 2, 7, 8, 9, 11),
    ("vs_z.xls", "Vysoké školy", "Vysoká škola (fakulta)", 3, 2, 6, 7, 8, 10),
]


def load_register():
    """Yields every row across all 11 compatible register files (see
    REGISTER_FILES - each is a live pupil-count census, so every listed row
    is inherently an active school, no separate open/closed filter exists
    in any of these sources), in file-then-row order, deduplicated by eduid
    globally across all files (a school could in principle be listed in
    more than one file). Website may be blank - per the brief, "produce one
    row per school" applies regardless, so blank-website schools are still
    yielded here and simply skip the crawl step downstream."""
    seen = set()
    for filename, sheet_name, school_type, data_start, eduid_col, name_col, addr_col, muni_col, web_col in REGISTER_FILES:
        wb = xlrd.open_workbook(f"{DATA_DIR}/{filename}")
        sh = wb.sheet_by_name(sheet_name)
        for r in range(data_start, sh.nrows):
            eduid = str(sh.cell_value(r, eduid_col)).strip()
            if not eduid or eduid in seen:
                continue
            website = str(sh.cell_value(r, web_col)).strip()
            if website and not website.lower().startswith(("http://", "https://")):
                website = "https://" + website
            seen.add(eduid)
            yield {
                "School ID": eduid,
                "School Name": str(sh.cell_value(r, name_col)).strip(),
                "School Type": school_type,
                "Address": str(sh.cell_value(r, addr_col)).strip(),
                "Municipality": str(sh.cell_value(r, muni_col)).strip(),
                "Website": website,
                "Source Register File": filename,
            }


FIELDNAMES = [
    "School ID", "School Name", "School Type", "Address", "Municipality", "Website",
    "Best Email", "Best Email Type", "Confidence Score", "Best Phone",
    "All Emails Found", "All Phones Found",
    "Source Register File", "Email Source Page", "Phone Source Page",
]


def process_school(school):
    row = dict(school)

    if not school["Website"]:
        print(f"=== {school['School Name']} (no website on file - skipping crawl) ===\n")
        row["Website"] = "No Website"
        row["Best Email"] = ""
        row["Best Email Type"] = ""
        row["Confidence Score"] = ""
        row["Best Phone"] = ""
        row["All Emails Found"] = ""
        row["All Phones Found"] = ""
        row["Email Source Page"] = ""
        row["Phone Source Page"] = ""
        return row

    log = []
    print(f"=== {school['School Name']} ({school['Website']}) ===")
    found_emails, found_phones, pages_fetched, requires_login, main_contact_page = crawl_site(school["Website"], log)
    for line in log:
        print(line)

    best_email = pick_best_email(found_emails, school_website=school["Website"], main_contact_page=main_contact_page)
    best_phone = pick_best_phone(found_phones)

    row["Best Email"] = best_email["best_email"] if best_email else ""
    if best_email:
        row["Best Email Type"] = best_email["best_email_type"]
    elif requires_login:
        # Reuses the existing "Best Email Type" column rather than adding a
        # new one, per instruction - makes clear to the client this is a
        # site that requires a login to view any content (confirmed with a
        # full JS-executing browser), not a crawler failure.
        row["Best Email Type"] = "Authentication Required"
    else:
        row["Best Email Type"] = ""
    row["Confidence Score"] = best_email["confidence"] if best_email else ""
    row["Best Phone"] = best_phone["best_phone"] if best_phone else ""
    row["All Emails Found"] = "; ".join(sorted(found_emails.keys()))
    row["All Phones Found"] = "; ".join(sorted(found_phones.keys()))
    row["Email Source Page"] = best_email["source_url"] if best_email else ""
    row["Phone Source Page"] = best_phone["source_url"] if best_phone else ""
    print(f"  -> best email: {row['Best Email']} ({row['Best Email Type']}, {row['Confidence Score']}) | best phone: {row['Best Phone']}")
    print()
    return row


checkpoint_lock = threading.Lock()


RETRY_CONCURRENCY = 8  # was sequential - same fix as england/norway_crawl.py


def _retry_one(row):
    # Must include "Source Register File" - confirmed real bug (2026-08-05):
    # without it, any school retried after an "Authentication Required"
    # result (always blank Best Email, so always picked up by the retry
    # filter below) permanently lost its source-file tracking, since
    # process_school() only ever carries forward whatever's already in the
    # school dict it's given, it never re-derives this field itself.
    school = {k: row[k] for k in ("School ID", "School Name", "School Type", "Address", "Municipality", "Website", "Source Register File")}
    return process_school(school)


def retry_failed_records(school_id_filter=None):
    """Retry pass over checkpointed failures - see england_crawl.py's
    version of this function for the full rationale (concurrency-induced
    transient failures vs genuine network blocks)."""
    all_rows = load_checkpoint(CHECKPOINT_FILE)
    targets = [
        r for r in all_rows.values()
        if (school_id_filter is None or r["School ID"] in school_id_filter)
        and r["Website"] != "No Website" and not r["Best Email"]
    ]
    if not targets:
        print("Nothing to retry.")
        return 0, 0
    print(f"Retrying {len(targets)} failures with {RETRY_CONCURRENCY} parallel workers...")
    recovered = 0
    with ThreadPoolExecutor(max_workers=RETRY_CONCURRENCY) as executor:
        futures = {executor.submit(_retry_one, row): row for row in targets}
        for future in as_completed(futures):
            row = futures[future]
            new_row = future.result()
            if new_row["Best Email"]:
                recovered += 1
            append_checkpoint(CHECKPOINT_FILE, row["School ID"], new_row)
    print(f"Recovered {recovered}/{len(targets)} via fast retry.")

    return recovered, len(targets)


def main():
    global CHECKPOINT_FILE, OUTPUT_FILE
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=200, help="new records to add this run (default 200); use 0 for unlimited (all remaining candidates)")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY, help="schools crawled in parallel")
    parser.add_argument("--retry-only", action="store_true", help="skip crawling new schools - just retry every existing checkpointed failure (e.g. run this from a different network to recover geo-blocked sites)")
    parser.add_argument("--checkpoint-suffix", default="", help="write to a separate slovakia_checkpoint_<suffix>.jsonl / slovakia_output_<suffix>.xlsx instead of the default files - use this to re-run every school (including already-checkpointed ones) under updated classification logic without touching the main checkpoint, merge back afterward")
    args = parser.parse_args()

    if args.checkpoint_suffix:
        CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", f"slovakia_checkpoint_{args.checkpoint_suffix}.jsonl")
        OUTPUT_FILE = os.path.join(BASE_DIR, "output", f"slovakia_output_{args.checkpoint_suffix}.xlsx")
        print(f"Using separate checkpoint: {CHECKPOINT_FILE}")

    if args.retry_only:
        retry_failed_records()
        total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Slovakia", text_columns={"School ID"})
        print(f"Total checkpointed: {total}. Wrote {OUTPUT_FILE}")
        return

    done = load_checkpoint(CHECKPOINT_FILE)
    print(f"{len(done)} schools already checkpointed from previous runs.")

    candidates = []
    for school in load_register():
        if school["School ID"] in done:
            continue
        candidates.append(school)
        if args.limit and len(candidates) >= args.limit:
            break
    print(f"Crawling {len(candidates)} schools with {args.concurrency} parallel workers...")

    added = 0
    with ThreadPoolExecutor(max_workers=args.concurrency) as executor:
        futures = {executor.submit(process_school, s): s for s in candidates}
        for future in as_completed(futures):
            school = futures[future]
            try:
                row = future.result()
            except Exception as e:
                print(f"FAILED {school['School Name']}: {e}")
                continue
            with checkpoint_lock:
                append_checkpoint(CHECKPOINT_FILE, school["School ID"], row)
                added += 1
                if added % EXPORT_EVERY == 0:
                    export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Slovakia", text_columns={"School ID"})

    # No filter - retry EVERY outstanding blank in the whole checkpoint,
    # not just this run's own new candidates (same fix as England/Turkey -
    # see their comments for the full explanation of the bug this fixes).
    retry_failed_records()

    total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Slovakia", text_columns={"School ID"})
    print(f"Added {added} new schools this run. Total checkpointed: {total}. Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
