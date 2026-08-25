"""
England (GIAS) production crawler.
For each school website: crawl same-domain pages (max 10), prioritising
Contact/About/Staff/Office/Administration/Enquiries/Team pages, extract
emails, and score the best one per the annex priority ladder.

Register source: GIAS "edubasealldata" full-establishment export (52,458
rows). Filtered to EstablishmentStatus=Open with a non-empty SchoolWebsite
(24,659 candidates) - closed schools and ones with no listed website can't
be crawled at all. Records are walked in the register's own file order, so
"first 200" means the first 200 open-with-website schools as the register
itself lists them, not a random sample.

Resumable: see checkpoint_utils.py - each school's result is appended to
output/england_checkpoint.jsonl immediately, so re-running after an
interruption skips already-processed URNs instead of starting over.
"""
import argparse
import asyncio
import csv
import os
import re
import threading
import time
import urllib.robotparser
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from urllib.parse import unquote, urljoin, urlparse

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from checkpoint_utils import append_checkpoint, export_xlsx_from_checkpoint, load_checkpoint

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGISTER_CSV = os.path.join(BASE_DIR, "data", "edubasealldata20260714.csv")
CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", "england_checkpoint.jsonl")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "england_output.xlsx")

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
HEADERS = {"User-Agent": UA}
MAX_PAGES = 10  # brief's explicit limit - not a speed knob, do not lower this
MAX_MAIN_CONTACT_SEARCH_PAGES = 20  # client's 2026-08-12 answer: this higher
# cap applies ONLY to the targeted search for the school's "main Contact
# page" (used by the tie-breaking rules) - not a blanket increase to the
# normal MAX_PAGES=10 crawl used for general email extraction.
CRAWL_DELAY = 0.2  # per-page politeness delay *within* a single school's own
# site - lowered from 0.5s; this is our own conservative choice, not a brief
# requirement, and was adding up to 4.5s of pure sleep() per school that
# visits all 10 pages, across 200 schools that's real wall-clock time for
# zero benefit (a real webserver has no trouble with a request every 0.2s
# from a single client).
TIMEOUT = 10
CONCURRENCY = 30  # bumped from 20 - main pass hits 200 independent domains,
# so more parallelism is close to free; any extra first-pass failures this
# causes on low-capacity sites get caught by the existing retry pass anyway.

RETRY_PROXY = None  # e.g. "socks5://127.0.0.1:1080" or "http://user:pass@host:port" -
# set via --proxy, only ever applied during the retry pass (main 200-record
# pass doesn't need it). Lets a UK VPS/residential-proxy be plugged in for
# whichever schools turn out to be IP-blocked from wherever this runs, at
# whatever scale that turns out to be once the full ~24,000-school run
# happens (not knowable in advance - it's determined by each site's own
# blocking behaviour, not something the crawler controls).


def _requests_proxies():
    return {"http": RETRY_PROXY, "https": RETRY_PROXY} if RETRY_PROXY else None

PRIORITY_KEYWORDS = [
    "contact-us", "contact_us", "contact",
    "about-us", "about",
    "staff", "office", "administration", "admin",
    "enquiries", "enquiry", "team",
]

# Client's rewritten best-email spec (2026-08-03, Priority_best_email.docx) -
# exact page-name list for England ("Priority pages" row of their table).
# Distinct from PRIORITY_KEYWORDS above (which only affects crawl order);
# this list decides RANKING - whether a found email counts as coming from a
# "priority page" at all. England has no phone-crawling logic (Telephone
# comes straight from GIAS), so no phone-side priority-page filtering needed here.
EN_PRIORITY_PAGE_KEYWORDS = [
    "contact", "contact us", "about", "staff", "office",
    "administration", "enquiries", "team",
]

# The single "main Contact page" (narrower than EN_PRIORITY_PAGE_KEYWORDS
# above, which also matches About/Staff/Team/etc.) - used only for the
# client's tie-breaking rule 4 ("prefer the email appearing on the school's
# main Contact page"), 2026-08-11 clarification.
EN_MAIN_CONTACT_PAGE_KEYWORDS = ["contact", "contact us"]

IGNORE_LOCAL_PARTS = {"no-reply", "noreply", "webmaster", "privacy", "gdpr"}

# Third-party website-builder/CMS platforms whose OWN generic support
# address gets embedded as boilerplate on every school site built on them
# (not customized per school) - confirmed real, live bug (2026-08-11
# pilot): chessbrook.herts.sch.uk and lesliemanser.lincs.sch.uk are both
# built on SchoolJotter, and info@schooljotter.com (the platform vendor's
# own support inbox, not the school's) was being picked as the "school's"
# best email whenever the school's own address wasn't found within the
# same crawl - same root cause as the earlier futuralearning.co.uk
# trust-hub bug (a shared operator's address beating the school's own),
# just via a template-vendor's boilerplate rather than a trust's shared
# inbox. Excluded outright, same treatment as IGNORE_LOCAL_PARTS, since an
# address on one of these domains is never a genuine answer for a specific
# school regardless of what prefix it uses.
PLATFORM_VENDOR_DOMAINS = {"schooljotter.com", "schooljotter3.com"}
PLACEHOLDER_LOCAL_PARTS = {"example", "test", "sample", "yourname", "youremail", "name", "yourdomain", "someone", "user"}
PLACEHOLDER_DOMAINS = {"example.com", "example.org", "example.net", "domain.com", "yoursite.com", "yourdomain.com", "email.com"}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

# Client's updated tie-breaking spec (Priority_best_email(1).docx,
# 2026-08-11) - order matters here: when two candidate emails are otherwise
# equally ranked, the one whose prefix comes FIRST in this list wins. Not
# just a set of recognized words like before.
HIGH_PRIORITY_PREFIXES = [
    "office", "schooloffice", "school.office", "school-office",
    "admin", "schooladmin", "school.admin",
    "enquiries", "contact", "info", "reception",
    "generaloffice", "secretary", "schoolsecretary", "officemanager", "bursar",
]
DECISION_MAKER_PREFIXES = [
    "head", "headteacher", "principal", "executivehead",
    "deputyhead", "deputyheadteacher", "assistanthead", "assistantheadteacher",
]
CONSUMER_DOMAINS = {
    "gmail.com", "hotmail.com", "outlook.com", "btinternet.com",
    "btconnect.com", "talk21.com", "icloud.com",
}


def robust_get(session, url, timeout=TIMEOUT):
    """session.get() with two narrow, well-understood retries, aimed at
    the two genuinely fixable failure classes found in testing:
    - SSL verification failures (mismatched hostname / self-signed cert,
      e.g. cavendish-school.co.uk, riverstonschool.co.uk) are retried
      once unverified. These are still the real school's own domain,
      just with a misconfigured certificate - and this is a read-only
      public-information crawl (no credentials or sensitive data ever
      sent), so disabling verification for this one retry is low-risk.
    - Timeouts/connection errors are retried up to 2 more times with
      backoff, since testing showed these are frequently transient
      hiccups from many schools being crawled concurrently (confirmed by
      re-testing several "failed" sites alone immediately afterwards -
      they loaded fine) rather than the site actually being down.
    Sites blocked by Cloudflare-style managed bot-challenges are NOT
    retried here - that's a different, much harder problem (verified by
    testing that even a real Playwright-rendered browser gets the same
    block), not something a retry can fix.
    """
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
        resp = requests.get(robots_url, headers=HEADERS, timeout=TIMEOUT, proxies=_requests_proxies())
        if resp.status_code == 200:
            rp.parse(resp.text.splitlines())
        else:
            rp.parse([])
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


CF_EMAIL_RE = re.compile(r'data-cfemail="([a-f0-9]+)"')

# Easy Email Blocker (EEB) - WordPress plugin used by many UK school sites.
# The JS pattern is: var ml="<alphabet>",mi="<indices>"; decode by mapping each
# char of mi through (ord(c) - 48) as an index into ml, then URL-decode.
EEB_RE = re.compile(
    r'var ml="([^"]+)",mi="([^"]+)"[\s\S]{0,200}?decodeURIComponent\(o\)',
    re.IGNORECASE,
)


def decode_eeb_email(ml, mi):
    """Decode an Easy Email Blocker obfuscated email."""
    import urllib.parse
    try:
        o = "".join(ml[ord(c) - 48] for c in mi)
        email = urllib.parse.unquote(o)
        return email if EMAIL_RE.fullmatch(email) else None
    except (IndexError, ValueError):
        return None


PLATFORM_EXCLUDE_DOMAINS = {
    # Confirmed live bug: Gospel Oak's splash page links to a Google Docs
    # link and a google.com/url? redirect wrapper. The splash-fallback
    # treated google.com itself as "the real school site" and tried to
    # crawl it, which spiralled into checking robots.txt for hundreds of
    # google.com/intl/<locale>/drive/ URLs - minutes wasted per retry for
    # zero chance of an email. These generic platforms are never the real
    # destination site themselves (sites.google.com deliberately excluded
    # from this list - a school's whole site can genuinely live there, as
    # gospeloakschool.com's does).
    "google.com", "docs.google.com", "drive.google.com", "forms.google.com",
    "facebook.com", "twitter.com", "x.com", "instagram.com", "youtube.com",
    "linkedin.com", "microsoft.com", "office.com", "apple.com",
}


def _unwrap_redirect(url):
    """Google search-result / link-shim URLs (google.com/url?q=<real>&...)
    embed the real destination in the q= param instead of linking to it
    directly - confirmed on Gospel Oak's splash page, which pointed at
    https://www.google.com/url?q=https%3A%2F%2Fsites.google.com%2F... for
    what's genuinely their live site. Unwrap it so the real domain gets
    evaluated, not the google.com shim."""
    p = urlparse(url)
    host = p.netloc.lower()
    if host.startswith("www."):
        host = host[4:]
    if host == "google.com" and p.path == "/url":
        from urllib.parse import parse_qs, unquote
        qs = parse_qs(p.query)
        if "q" in qs:
            return unquote(qs["q"][0])
    return url


NON_EMAIL_TLDS = {
    "png", "jpg", "jpeg", "gif", "svg", "webp", "ico", "bmp", "tiff",
    "css", "js", "woff", "woff2", "ttf", "eot",
}  # image/asset filenames using @2x-style retina naming (e.g. logo@2x.png)
# coincidentally match the email pattern - "png" etc. look like a valid TLD.

SCHOOL_TLD_RE = re.compile(
    r'\.(sch\.uk|academy\.org|academy\.co\.uk|school\.co\.uk|'  # noqa: E501
    r'nurseryschool\.co\.uk|primary\.co\.uk|junior\.co\.uk|'  # noqa: E501
    r'infant\.co\.uk|college\.co\.uk|edu\.co\.uk)$',
    re.IGNORECASE,
)  # module-level (was local to crawl_site) so classify_email can also use it
# to recognise a genuine UK-school-sector domain, regardless of which
# function needs the check.


def is_placeholder(email):
    local, _, domain = email.lower().partition("@")
    if domain.rsplit(".", 1)[-1] in NON_EMAIL_TLDS:
        return True
    return local in PLACEHOLDER_LOCAL_PARTS or domain in PLACEHOLDER_DOMAINS


SITEMAP_LOC_RE = re.compile(r"<loc>\s*(?:<!\[CDATA\[)?\s*(.*?)\s*(?:\]\]>)?\s*</loc>", re.IGNORECASE)
SITEMAP_PATHS = ("/sitemap.xml", "/sitemap_index.xml")
MAX_SUB_SITEMAPS = 5


def _fetch_locs(url, session):
    try:
        resp = robust_get(session, url)
    except requests.RequestException:
        return []
    if resp.status_code != 200:
        return []
    return SITEMAP_LOC_RE.findall(resp.text)


def fetch_sitemap_urls(start_url, domain, session, log):
    """Many school sites render their nav menu via JS, so plain <a href>
    link-following misses most pages. sitemap.xml is the standard,
    crawler-facing supplement for exactly this case. Some sitemaps are
    just an index of sub-sitemaps (e.g. WordPress SEO plugins), so one
    extra level is followed when every entry itself ends in .xml."""
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
                netloc_clean = p.netloc.lower()
                if netloc_clean.startswith("www."):
                    netloc_clean = netloc_clean[4:]
                if netloc_clean != domain:
                    continue
                page_urls.extend(_fetch_locs(sub, session))
            locs = page_urls

        urls = []
        for loc in locs:
            p = urlparse(loc)
            loc_netloc = p.netloc.lower()
            if loc_netloc.startswith("www."):
                loc_netloc = loc_netloc[4:]
            if loc_netloc == domain and p.scheme in ("http", "https"):
                urls.append(loc)

        if urls:
            log.append(f"  sitemap found at {path}: {len(urls)} same-domain URLs")
            return urls

    return []


def _url_variants(url):
    """www/bare-domain and https/http are each independently a common
    real-world misconfiguration where only one variant actually resolves
    (confirmed: hampsteadprim.camden.sch.uk needs www stripped;
    zsdstal.edu.sk - a Slovak school site - has no working HTTPS at all,
    only plain HTTP). Try all 4 combinations."""
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


def _record_email(found_emails, addr, url, title):
    """Records EVERY (url, title) an email is found on, not just the first
    or the best - found_emails[addr] is a list of all occurrences. Needed
    for two things classify_email() depends on: (1) whether an email counts
    as "on a Priority page" at all must check ALL its occurrences, not just
    the first one seen (confirmed real bug, 2026-08-11 client feedback: the
    homepage is always crawled first, so an email appearing on both the
    homepage and the actual Contact page - very common, footers repeat the
    office email - was permanently judged by the homepage's non-priority
    status alone); (2) the client's tie-breaking rules (2026-08-11
    clarification) need to know whether an email appears on the school's
    MAIN Contact page specifically, and on how many distinct Priority pages
    it appears on - neither is answerable from a single stored occurrence."""
    found_emails.setdefault(addr.lower(), []).append((url, title))


def extract_emails_from_html(html, url, found_emails):
    """Shared by both the plain-requests crawl and the Playwright fallback
    crawl (crawl_site_playwright) - same extraction rules regardless of how
    the HTML was fetched. Mutates found_emails in place; returns the parsed
    soup (so callers that also need it for link-enqueuing don't re-parse)
    and the page title."""
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.string.strip() if soup.title and soup.title.string else ""

    # mailto links
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.lower().startswith("mailto:"):
            addr = unquote(href[7:].split("?")[0]).strip()
            if EMAIL_RE.fullmatch(addr) and not is_placeholder(addr):
                _record_email(found_emails, addr, url, title)

    # visible text emails
    for m in EMAIL_RE.finditer(soup.get_text(" ")):
        addr = m.group(0)
        if not is_placeholder(addr):
            _record_email(found_emails, addr, url, title)

    # raw HTML emails (catches some JS-embedded / data attributes)
    for m in EMAIL_RE.finditer(html):
        addr = m.group(0)
        if not is_placeholder(addr):
            _record_email(found_emails, addr, url, title)

    # Some sites cloak the visible email behind JS (shows "*protected
    # email*" and only reveals the real address client-side), but the
    # SEO plugin that generates the og:description / meta description
    # summary pulls from the underlying content before that cloaking
    # applies, and includes the real address in plain text (confirmed
    # on torriano.camden.sch.uk: page body says "*protected email*",
    # but <meta property="og:description"> has the genuine
    # admin@torriano.camden.sch.uk). BeautifulSoup already decodes
    # HTML entities in attribute values, so no extra decoding needed.
    for meta in soup.find_all("meta"):
        if meta.get("property") in ("og:description",) or meta.get("name") in ("description",):
            for m in EMAIL_RE.finditer(meta.get("content", "")):
                addr = m.group(0)
                if not is_placeholder(addr):
                    _record_email(found_emails, addr, url, title)

    # Cloudflare-obfuscated emails (data-cfemail="...")
    for m in CF_EMAIL_RE.finditer(html):
        addr = decode_cf_email(m.group(1))
        if addr and not is_placeholder(addr):
            _record_email(found_emails, addr, url, title)

    # Easy Email Blocker (EEB) obfuscated emails - WordPress plugin
    for m in EEB_RE.finditer(html):
        addr = decode_eeb_email(m.group(1), m.group(2))
        if addr and not is_placeholder(addr):
            _record_email(found_emails, addr, url, title)

    return soup, title


def crawl_site(start_url, log):
    # Follow HTTP redirects to handle splash-page sites (e.g. Wentworth) that
    # redirect to the real school site on a different domain. We probe the start
    # URL once before the main crawl so the domain variable tracks the real site.
    resolved_url = start_url
    probe_ok = False
    for candidate in _url_variants(start_url):
        try:
            probe = requests.get(candidate, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True, proxies=_requests_proxies())
        except requests.RequestException:
            continue
        if probe.status_code == 200:
            if candidate != start_url:
                log.append(f"  {start_url} failed - variant {candidate} worked (-> {probe.url})")
            resolved_url = probe.url
            start_url = candidate
            probe_ok = True
            break

    parsed_start = urlparse(resolved_url)
    domain = parsed_start.netloc.lower()
    if domain.startswith("www."):
        domain = domain[4:]
    if resolved_url != start_url:
        log.append(f"  redirected: {start_url} -> {resolved_url} (domain={domain})")
    rp = get_robot_parser(start_url)

    visited = set()
    to_visit = [(start_url, 100)]  # (url, priority) homepage always first
    found_emails = {}  # email -> (source_url, page_title)
    visited_pages = []  # every (url, title) successfully fetched, regardless
    # of whether an email was found there - needed to determine which
    # Priority page genuinely EXISTS on this school's site (client's
    # 2026-08-12 answer: "main Contact page" = Contact if it exists, else
    # fall back through the brief's Priority-pages list in order), which is
    # a different question from "which page did we find an email on."
    homepage_nav_links = []  # (anchor_text_lower, resolved_url) for every
    # same-domain link found on the homepage specifically (client's answer:
    # "search the priority page text in the anchor text of links in the
    # main link" = the homepage). Lets us detect a page like "Kontakt"
    # EXISTS even if it wasn't among the pages this crawl happened to visit.
    pages_fetched = 0

    session = requests.Session()
    session.headers.update(HEADERS)
    if RETRY_PROXY:
        session.proxies = _requests_proxies()

    queued = {normalize(start_url)}

    for sm_url in fetch_sitemap_urls(start_url, domain, session, log):
        normf = normalize(sm_url)
        if normf in queued:
            continue
        pr = score_link(sm_url, "")
        to_visit.append((sm_url, pr))
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
        except requests.RequestException as e:
            log.append(f"  fetch failed {url}: {e}")
            continue
        visited.add(norm)
        pages_fetched += 1
        if resp.status_code != 200 or "text/html" not in resp.headers.get("Content-Type", ""):
            time.sleep(CRAWL_DELAY)
            continue

        soup, title = extract_emails_from_html(resp.text, url, found_emails)
        visited_pages.append((url, title))
        is_homepage_fetch = normalize(url) == normalize(start_url)

        # enqueue same-domain links (also captures homepage nav-link anchor
        # text separately, when this fetch IS the homepage)
        if pages_fetched < MAX_PAGES or is_homepage_fetch:
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if href.lower().startswith(("mailto:", "tel:", "javascript:")):
                    continue
                full = urljoin(url, href)
                p = urlparse(full)
                netloc_clean = p.netloc.lower()
                if netloc_clean.startswith("www."):
                    netloc_clean = netloc_clean[4:]
                if netloc_clean != domain:
                    continue
                if p.scheme not in ("http", "https"):
                    continue
                if is_homepage_fetch:
                    homepage_nav_links.append((a.get_text(" ", strip=True).lower(), full))
                normf = normalize(full)
                if normf in visited or normf in queued:
                    continue
                text = a.get_text(" ", strip=True)
                pr = score_link(href, text)
                to_visit.append((full, pr))
                queued.add(normf)

        time.sleep(CRAWL_DELAY)

    log.append(f"  pages fetched: {pages_fetched}, emails found: {len(found_emails)}")

    # Determine the school's "main Contact page" for the tie-breaking rules
    # (client's 2026-08-12 answer): try Contact first, then fall back through
    # EN_PRIORITY_PAGE_KEYWORDS in the brief's own order. First check pages
    # we ALREADY fetched (free); if none match, check the homepage's own nav
    # links for a matching anchor text (proves the page exists even though
    # we didn't visit it), and fetch it now as one extra targeted page - up
    # to MAX_MAIN_CONTACT_SEARCH_PAGES total, not the normal MAX_PAGES=10.
    # "Not among the pages we crawled" is treated as "doesn't exist" (also
    # client-confirmed), so if nothing matches after this, there is no main
    # Contact page and that tie-break dimension just doesn't apply.
    main_contact_page = None
    for kw in EN_PRIORITY_PAGE_KEYWORDS:
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
                if resp.status_code == 200 and "text/html" in resp.headers.get("Content-Type", ""):
                    visited.add(normalize(nav_match_url))
                    pages_fetched += 1
                    _, nav_title = extract_emails_from_html(resp.text, nav_match_url, found_emails)
                    visited_pages.append((nav_match_url, nav_title))
                    log.append(f"  targeted fetch for main-contact-page search: {nav_match_url}")
                    main_contact_page = (nav_match_url, nav_title)
            except requests.RequestException as e:
                log.append(f"  targeted fetch failed {nav_match_url}: {e}")
        if main_contact_page:
            break

    # Splash-page fallback: some GIAS website entries are landing pages that
    # link to the real school site on a different domain (e.g. Wentworth Nursery
    # School uses wentworth.hackney.sch.uk as a splash redirecting to
    # wentworthnurseryschool.co.uk). If we finished with 0 emails and only
    # visited ≤3 pages, we look for school-domain cross-links on the first page
    # fetched and do a lightweight crawl of up to 2 of those linked domains.
    # (SCHOOL_TLD_RE itself is module-level now - also used by classify_email.)
    DOORWAY_TEXT_RE = re.compile(
        r'\b(school|nursery|academy|college|centre|center|enter|visit)\b',
        re.IGNORECASE,
    )
    if not found_emails and pages_fetched <= 3:
        # Re-fetch the resolved homepage to get its cross-domain links
        try:
            hp = robust_get(session, resolved_url)
            hp_soup = BeautifulSoup(hp.text, "html.parser")
        except requests.RequestException:
            hp_soup = None

        if hp_soup:
            seen_fallback = set()
            fallback_urls = []
            for a in hp_soup.find_all("a", href=True):
                href = a["href"]
                if not href.startswith("http"):
                    continue
                href = _unwrap_redirect(href)
                p = urlparse(href)
                fhost = p.netloc.lower()
                if fhost.startswith("www."):
                    fhost = fhost[4:]
                if fhost == domain or fhost in seen_fallback or fhost in PLATFORM_EXCLUDE_DOMAINS:
                    continue
                link_text = a.get_text(" ", strip=True)
                if (SCHOOL_TLD_RE.search(fhost)
                        or score_link(href, link_text) > 0
                        or DOORWAY_TEXT_RE.search(link_text)):
                    seen_fallback.add(fhost)
                    fallback_urls.append(href)

            for fb_url in fallback_urls[:2]:
                log.append(f"  splash-page fallback: crawling linked domain {fb_url}")
                fb_emails, fb_pages, fb_main_contact = crawl_site(fb_url, log)
                found_emails.update(fb_emails)
                pages_fetched += fb_pages
                if main_contact_page is None and fb_main_contact is not None:
                    main_contact_page = fb_main_contact

        log.append(f"  after splash fallback: emails found: {len(found_emails)}")

    return found_emails, pages_fetched, main_contact_page


def is_priority_page(url, title, keywords=EN_PRIORITY_PAGE_KEYWORDS):
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
    Two cases the client's table doesn't cover - handled as judgment calls,
    ranked below the tiers the table does define (see slovakia_crawl.py's
    classify_email for the full rationale):
      - Generic prefix found on a NON-priority page: ranked below
        Decision-maker but above the freemail tiers.
      - Anything matching none of the above (e.g. a personal-name address):
        last-resort fallback only.

    A genuine UK-school-sector domain (.sch.uk etc.) still adds a large
    bonus on top of whichever tier applies, regardless of tier - confirmed
    necessary on sheringtonprimary.co.uk: it found both
    "contact@inspirar.co.uk" (the web design agency that built the site -
    inspirar.co.uk is not a school) and "sao@sherington.greenwich.sch.uk"
    (the school's own address); a school's own address on its own domain
    is correct essentially by definition, even with an unlisted prefix like
    "sao@", so this bonus is large enough to always win that comparison.

    Within a tier, HIGH_PRIORITY_PREFIXES/DECISION_MAKER_PREFIXES are now
    ORDERED lists per the client's 2026-08-11 tie-breaking rules - an email
    matching an earlier-listed prefix outranks one matching a later-listed
    prefix when their tier is otherwise identical. The client's remaining
    tie-break chain (main Contact page > appears on multiple Priority pages
    > first encountered during crawl) is resolved in pick_best_email(),
    which has visibility across candidates that this per-email function
    doesn't - main_contact_page is the (url, title) crawl_site() already
    determined for this whole school (client's 2026-08-12 answer: Contact
    page if it exists, else fall back through the brief's Priority-pages
    list in order) - not something re-derived per candidate here.

    Returns a dict with everything pick_best_email() needs to rank and
    report this candidate, or None if excluded entirely.
    """
    local, _, domain = email.partition("@")
    local_l = local.lower()
    domain_l = domain.lower()

    if local_l in IGNORE_LOCAL_PARTS or domain_l in PLATFORM_VENDOR_DOMAINS:
        return None  # excluded unless no other option, handled by caller

    school_domain_bonus = 1000 if SCHOOL_TLD_RE.search(domain_l) else 0
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
    for idx, pref in enumerate(DECISION_MAKER_PREFIXES):
        if local_l == pref or pref in local_l:
            return result(school_domain_bonus + 400, "Decision-maker", "Medium-high", prefix_rank=idx)

    # Generic school/office contact.
    for idx, pref in enumerate(HIGH_PRIORITY_PREFIXES):
        if local_l == pref or pref in local_l:
            if on_priority:
                return result(school_domain_bonus + 500, "Generic School Contact", "High", prefix_rank=idx)
            return result(school_domain_bonus + 300, "Generic School Contact", "Medium", prefix_rank=idx,
                          main_contact=False, n_priority=0, source=occurrences[0])

    # Freemail/consumer domains - rank depends on whether it's on a priority page.
    if domain_l in CONSUMER_DOMAINS:
        if on_priority:
            return result(200, "Consumer / Freemail", "Medium", main_contact=False, n_priority=0, source=priority_occurrences[0])
        return result(100, "Consumer / Freemail", "Low", main_contact=False, n_priority=0, source=occurrences[0])

    # Last resort - doesn't match any defined tier (e.g. a personal-name address).
    return result(school_domain_bonus + 10, "Other / Unclassified", "Low",
                  main_contact=False, n_priority=0, source=occurrences[0])


def pick_best_email(found_emails, main_contact_page=None):
    """found_emails: {email: [(url, title), ...]} - every occurrence, per
    _record_email(). main_contact_page: (url, title) crawl_site() determined
    for this school, or None. Ranks candidates by: tier score (includes the
    school-domain bonus) -> prefix order within the tier (client's
    2026-08-11 ordered lists) -> appears on the main Contact page -> appears
    on more than one Priority page -> first encountered during the crawl
    (dict insertion order, since a plain dict preserves it) - this exact
    chain is the client's tie-breaking spec."""
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
        # fall back to ignored (no-reply etc.) addresses if nothing else exists
        for order, (email, occurrences) in enumerate(found_emails.items()):
            src, title = occurrences[0]
            cls = {"etype": "Excluded type (no alternative found)", "confidence": "Low",
                   "source_url": src, "page_title": title}
            candidates.append(((0, 999, 1, 0, order), email, cls))

    if not candidates:
        return None

    candidates.sort(key=lambda c: c[0])
    _, email, cls = candidates[0]
    return {
        "best_email": email,
        "best_email_type": cls["etype"],
        "confidence": cls["confidence"],
        "source_url": cls["source_url"],
        "page_title": cls["page_title"],
    }


OPEN_STATUSES = {"Open", "Open, but proposed to close"}


def load_register(csv_path):
    """Yields every active school (brief's exclusion list is Closed,
    Proposed schools, Duplicate establishments - not "no website"), in the
    register's own file order, deduplicated by URN (the register's unique
    key). Website may be blank - the brief requires one row per active
    GIAS URN regardless, so blank-website schools are still yielded here
    and simply skip the crawl step downstream."""
    seen = set()
    with open(csv_path, encoding="cp1252", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            urn = (row.get("URN") or "").strip()
            website = (row.get("SchoolWebsite") or "").strip()
            status = (row.get("EstablishmentStatus (name)") or "").strip()
            if not urn or urn in seen or status not in OPEN_STATUSES:
                continue
            if website and not website.lower().startswith(("http://", "https://")):
                website = "https://" + website
            seen.add(urn)
            address = ", ".join(
                p.strip() for p in [
                    row.get("Street", ""), row.get("Locality", ""),
                    row.get("Address3", ""), row.get("Town", ""),
                ] if p and p.strip()
            )
            yield {
                "URN": urn,
                "Name": row.get("EstablishmentName", "").strip(),
                "Status": status,
                "Phase": row.get("PhaseOfEducation (name)", "").strip(),
                "Type": row.get("TypeOfEstablishment (name)", "").strip(),
                "LA": row.get("LA (name)", "").strip(),
                "Address": address,
                "Postcode": row.get("Postcode", "").strip(),
                "Telephone": row.get("TelephoneNum", "").strip(),
                "Website": website,
            }


FIELDNAMES = [
    "URN", "Name", "Status", "Phase", "Type", "LA", "Address", "Postcode",
    "Telephone", "Website", "Best Email", "Best Email Type", "Confidence Score",
    "All Emails Found", "Email Source URL", "Extraction Date",
    "Pages Crawled",
]
TEXT_COLUMNS = {"URN", "Postcode", "Telephone"}
EXPORT_EVERY = 10  # refresh the .xlsx every N new records so progress can be checked mid-run


def process_school(school):
    row = dict(school)
    row["Extraction Date"] = date.today().isoformat()

    if not school["Website"]:
        print(f"=== {school['Name']} (no website on file - skipping crawl) ===\n")
        row["Website"] = "No Website"
        row["Pages Crawled"] = 0
        row["Best Email"] = ""
        row["Best Email Type"] = ""
        row["Confidence Score"] = ""
        row["Email Source URL"] = ""
        row["Extraction Page"] = ""
        row["All Emails Found"] = ""
        return row

    log = []
    print(f"=== {school['Name']} ({school['Website']}) ===")
    found_emails, pages_fetched, main_contact_page = crawl_site(school["Website"], log)
    for line in log:
        print(line)
    best = pick_best_email(found_emails, main_contact_page)

    row["Pages Crawled"] = pages_fetched
    if best:
        row["Best Email"] = best["best_email"]
        row["Best Email Type"] = best["best_email_type"]
        row["Confidence Score"] = best["confidence"]
        row["Email Source URL"] = best["source_url"]
        row["Extraction Page"] = best["page_title"]
    else:
        row["Best Email"] = ""
        row["Best Email Type"] = ""
        row["Confidence Score"] = ""
        row["Email Source URL"] = ""
        row["Extraction Page"] = ""
    row["All Emails Found"] = "; ".join(sorted(found_emails.keys()))
    print(f"  -> best: {row['Best Email']} ({row['Best Email Type']}, {row['Confidence Score']})")
    print()
    return row


checkpoint_lock = threading.Lock()

async def crawl_site_playwright(start_url, log, page, max_pages=2):
    """Last-resort fetch using a real rendered browser instead of plain
    requests.get(). Two problems this can solve that requests never can:
    1. Content injected by client-side JS (no raw HTML for requests to see).
    2. Simple Cloudflare-style "Just a moment..." JS challenges that clear
       themselves after a few seconds in a real browser - genuine CAPTCHA/
       Turnstile challenges (confirmed on holytrinitynw3.co.uk) still won't
       clear here, that needs an interactive solve, not just a browser.
    Uses Playwright's async API so multiple schools' pages/tabs can run
    concurrently via asyncio.gather() on a single thread - the sync API is
    what's not safe across OS threads (confirmed the hard way on the Turkey
    crawler), not concurrency itself. Each call gets its own `page` (tab),
    so different schools never share mutable state.
    """
    # Resolve the real target first (same www/https variants + redirect
    # following the requests-based crawl_site does). Confirmed necessary on
    # Gospel Oak: the GIAS-listed domain is dead, and only resolves to the
    # real live site (a Google Sites page) via this redirect chain - without
    # it, Playwright just times out hitting the same dead domain again.
    for candidate in _url_variants(start_url):
        try:
            probe = requests.get(candidate, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True, proxies=_requests_proxies())
        except requests.RequestException:
            continue
        if probe.status_code == 200:
            if probe.url != start_url:
                log.append(f"  [playwright] resolved {start_url} -> {probe.url}")
            start_url = probe.url
            break

    parsed_start = urlparse(start_url)
    domain = parsed_start.netloc.lower()
    if domain.startswith("www."):
        domain = domain[4:]

    found_emails = {}
    visited = set()
    to_visit = [(start_url, 100)]
    queued = {normalize(start_url)}
    pages_fetched = 0

    while to_visit and pages_fetched < max_pages:
        to_visit.sort(key=lambda x: -x[1])
        url, _ = to_visit.pop(0)
        norm = normalize(url)
        if norm in visited:
            continue
        try:
            await page.goto(url, timeout=12000, wait_until="domcontentloaded")
            visited.add(norm)
            pages_fetched += 1

            # Short settle time for JS-heavy sites (Google Sites, Wix,
            # Squarespace) whose nav/content isn't in the DOM yet right at
            # domcontentloaded - without this, page.content() below sees the
            # same near-empty shell a plain requests.get() would.
            await page.wait_for_timeout(1500)

            title = (await page.title() or "")
            if "just a moment" in title.lower():
                # give a simple JS challenge a real chance to clear
                await page.wait_for_timeout(4000)
                title = (await page.title() or "")

            html = await page.content()
        except Exception as e:
            # Covers goto failures AND races like a client-side redirect
            # firing mid-extraction ("page is navigating and changing the
            # content" - confirmed on cavendish-school.co.uk). Either way,
            # one flaky page must not take down the whole batch - previously
            # this crashed the entire Playwright fallback pass partway
            # through, silently skipping every school still queued after it.
            log.append(f"  [playwright] fetch failed {url}: {e}")
            if norm not in visited:
                continue
            html = None

        if html is None:
            continue

        soup, title = extract_emails_from_html(html, url, found_emails)

        if pages_fetched < max_pages:
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if href.lower().startswith(("mailto:", "tel:", "javascript:")):
                    continue
                full = urljoin(url, href)
                p = urlparse(full)
                netloc_clean = p.netloc.lower()
                if netloc_clean.startswith("www."):
                    netloc_clean = netloc_clean[4:]
                if netloc_clean != domain or p.scheme not in ("http", "https"):
                    continue
                normf = normalize(full)
                if normf in visited or normf in queued:
                    continue
                text = a.get_text(" ", strip=True)
                pr = score_link(href, text)
                to_visit.append((full, pr))
                queued.add(normf)

    log.append(f"  [playwright] pages fetched: {pages_fetched}, emails found: {len(found_emails)}")
    return found_emails


CDP_URL = "http://localhost:9222"  # if a real Chrome is running with
# --remote-debugging-port=9222 (see RUN_WITH_REAL_CHROME.md), we drive that
# instead of a fresh headless browser. Confirmed on Holy Trinity/Miles
# Coverdale/Sherington: a real Chrome profile - genuine binary, real
# extensions (incl. any VPN extension the user has toggled on), no
# automation-launch flag - clears challenges/blocks a fresh headless
# Chromium can't, since Cloudflare-style detection and IP-based blocks key
# off exactly those signals. Falls back to a normal headless launch below if
# no such Chrome is running - this must keep working with zero setup.


async def _get_playwright_browser(p):
    """Returns (browser, owns_browser). owns_browser=False means a real,
    user-owned Chrome was attached to via CDP - the caller must never close
    it, only the tabs/pages opened on it."""
    try:
        browser = await p.chromium.connect_over_cdp(CDP_URL, timeout=5000)
        print(f"  [playwright] attached to real Chrome via CDP ({CDP_URL})")
        return browser, False
    except Exception:
        pass

    launch_kwargs = {"headless": True}
    if RETRY_PROXY:
        launch_kwargs["proxy"] = {"server": RETRY_PROXY}
    browser = await p.chromium.launch(**launch_kwargs)
    return browser, True


PLAYWRIGHT_RETRY_ATTEMPTS = 2  # lowered from 3 - confirmed on Miles
# Coverdale/Sherington that a free VPN's exit node can be intermittently
# overloaded (works, then times out on the identical URL minutes later), so
# one retry is worth it, but each attempt costs up to ~12s+cooldown and now
# that records run concurrently (PLAYWRIGHT_CONCURRENCY tabs at once via
# asyncio.gather), a 3rd attempt per-record adds more total wall time than
# it recovers - two tries already catches the "worked on reload" case.

PLAYWRIGHT_CONCURRENCY = 5  # multiple tabs on the *same* browser, run
# concurrently via asyncio.gather() on one thread/event loop - safe (unlike
# multi-threading the sync API) because everything stays on one thread;
# asyncio just interleaves the I/O-bound waits (goto/wait_for_timeout).
# Kept lower than the main pass's concurrency since these are specifically
# the hardest schools (Cloudflare/IP-blocked) and each real Chrome tab costs
# more resources than a plain HTTP request.


async def _playwright_process_one(row, context, semaphore):
    async with semaphore:
        page = await context.new_page()
        log = []
        print(f"=== [playwright retry] {row['Name']} ({row['Website']}) ===")
        found_emails = {}
        try:
            for attempt in range(PLAYWRIGHT_RETRY_ATTEMPTS):
                found_emails = await crawl_site_playwright(row["Website"], log, page)
                if found_emails:
                    break
                if attempt < PLAYWRIGHT_RETRY_ATTEMPTS - 1:
                    log.append(f"  [playwright] attempt {attempt + 1} found nothing, retrying after cooldown...")
                    await page.wait_for_timeout(2000)
        finally:
            await page.close()
        for line in log:
            print(line)
        best = pick_best_email(found_emails)
        if best:
            row["Best Email"] = best["best_email"]
            row["Best Email Type"] = best["best_email_type"]
            row["Confidence Score"] = best["confidence"]
            row["Email Source URL"] = best["source_url"]
            row["Extraction Page"] = best["page_title"]
            row["All Emails Found"] = "; ".join(sorted(found_emails.keys()))
            print(f"  -> RECOVERED via Playwright: {row['Best Email']}")
        else:
            print("  -> still no email (Playwright fallback exhausted)")
        append_checkpoint(CHECKPOINT_FILE, row["URN"], row)
        return bool(best)


async def _playwright_retry_pass_async(rows):
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        browser, owns_browser = await _get_playwright_browser(p)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        semaphore = asyncio.Semaphore(PLAYWRIGHT_CONCURRENCY)
        results = await asyncio.gather(*(_playwright_process_one(row, context, semaphore) for row in rows))
        if owns_browser:
            await browser.close()
        return sum(results)


def playwright_retry_pass(rows):
    """Final fallback for records still blank after the fast requests-based
    retry pass. Runs multiple schools concurrently (see PLAYWRIGHT_CONCURRENCY)
    via Playwright's async API instead of one at a time."""
    return asyncio.run(_playwright_retry_pass_async(rows))


def retry_failed_records(urn_filter=None):
    """Re-crawls every checkpointed record that has a website but no Best
    Email, going straight to the real-Chrome-via-CDP Playwright pass (see
    CDP_URL/_get_playwright_browser) - no plain-HTTP fast-retry stage first.
    Requires a real Chrome already running with --remote-debugging-port=9222
    and the VPN connected in it.
    urn_filter: optional set of URNs to restrict the retry to (used by the
    main crawl's own end-of-run cleanup pass); omit to retry every
    outstanding failure in the whole checkpoint (for a standalone
    `--retry-only` run).
    """
    all_rows = load_checkpoint(CHECKPOINT_FILE)
    targets = [
        r for r in all_rows.values()
        if (urn_filter is None or r["URN"] in urn_filter)
        and r["Website"] != "No Website" and not r["Best Email"]
    ]
    if not targets:
        print("Nothing to retry.")
        return 0, 0
    print(f"Retrying {len(targets)} failures directly via Chrome+VPN (Playwright)...")
    recovered = playwright_retry_pass(targets)
    print(f"Recovered {recovered}/{len(targets)} via Playwright.")
    return recovered, len(targets)


MAX_RETRY_ROUNDS = 3  # a single retry pass doesn't catch everything - some
# blanks are transient (VPN/network blip on that specific attempt) and
# recover on a second or third try rather than the first. Runs automatically
# every round unless a round finds nothing left to retry, in which case it
# stops early rather than wasting 2 more no-op passes.


def _retry_multi_round():
    for round_num in range(1, MAX_RETRY_ROUNDS + 1):
        print(f"\n--- Retry round {round_num}/{MAX_RETRY_ROUNDS} ---")
        recovered, total = retry_failed_records()
        if total == 0:
            break


def main():
    global CHECKPOINT_FILE, OUTPUT_FILE, RETRY_PROXY
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=200, help="new records to add this run (default 200); use 0 for unlimited (all remaining candidates)")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY, help="schools crawled in parallel")
    parser.add_argument("--retry-only", action="store_true", help="skip crawling new schools - just retry every existing checkpointed failure (e.g. run this from a different network to recover geo-blocked sites)")
    parser.add_argument("--proxy", default=None, help="proxy URL applied only during the retry pass, e.g. socks5://127.0.0.1:1080 or http://user:pass@host:port - for recovering IP-blocked schools without needing a different machine")
    parser.add_argument("--start-index", type=int, default=0, help="skip straight to this position in the register file before filtering for not-yet-done - lets two machines split the register into non-overlapping ranges (e.g. one runs 0-13615, another runs --start-index 13615) without either touching the other's checkpoint")
    parser.add_argument("--checkpoint-suffix", default="", help="write to a separate england_checkpoint_<suffix>.jsonl / england_output_<suffix>.xlsx instead of the default files - use this together with --start-index so a range-split run doesn't collide with the main checkpoint, or on its own to re-run every school under updated classification logic - merge back afterward")
    args = parser.parse_args()

    if args.proxy:
        RETRY_PROXY = args.proxy
        print(f"Retry pass will route through proxy: {args.proxy}")

    if args.checkpoint_suffix:
        CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", f"england_checkpoint_{args.checkpoint_suffix}.jsonl")
        OUTPUT_FILE = os.path.join(BASE_DIR, "output", f"england_output_{args.checkpoint_suffix}.xlsx")
        print(f"Using separate checkpoint: {CHECKPOINT_FILE}")

    if args.retry_only:
        _retry_multi_round()
        total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="England", text_columns=TEXT_COLUMNS)
        print(f"Total checkpointed: {total}. Wrote {OUTPUT_FILE}")
        return

    done = load_checkpoint(CHECKPOINT_FILE)
    print(f"{len(done)} schools already checkpointed from previous runs.")

    all_schools = list(load_register(REGISTER_CSV))
    if args.start_index:
        all_schools = all_schools[args.start_index:]
        print(f"Starting from index {args.start_index} in the register ({len(all_schools)} schools from there to the end).")

    candidates = []
    for school in all_schools:
        if school["URN"] in done:
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
                print(f"FAILED {school['Name']}: {e}")
                continue
            with checkpoint_lock:
                append_checkpoint(CHECKPOINT_FILE, school["URN"], row)
                added += 1
                if added % EXPORT_EVERY == 0:
                    export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="England", text_columns=TEXT_COLUMNS)

    # No filter - retry EVERY outstanding blank in the whole checkpoint,
    # not just this run's own new candidates. Scoping to just `candidates`
    # was a real bug (confirmed on Turkey's equivalent code): a school
    # that got a blank checkpoint entry in an earlier separate run is
    # already "done" (excluded from `candidates`) but was then ALSO
    # excluded from retry - a permanent blind spot for anything that
    # failed before the current invocation.
    _retry_multi_round()

    total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="England", text_columns=TEXT_COLUMNS)
    print(f"Added {added} new schools this run. Total checkpointed: {total}. Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
