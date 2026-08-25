"""
Norway (NSR) production crawler.
NSR's own API returns epost/telefon fields but they're null in practice
(register doesn't publish contacts), so email AND phone both come from
crawling each institution's own site, same as Slovakia.

Register source: NSR's public REST API (data-nsr.udir.no, confirmed live
and reachable - no VPN needed, unlike Turkey/Greece). /enheter returns all
~18,200 registered entities but without a Website field; filtered to
active, web-visible institutions physically in Norway (ErSkole, ErAktiv,
VisesPaaWeb, excluding FylkeNr "25" which is NSR's pseudo-fylke for
Norwegian institutions abroad) - 5,700 candidates. The brief says "every
active institution" / "Norwegian educational institutions" with no
mention of primary-only - an earlier version of this script wrongly
narrowed this to ErGrunnSkole (primary schools only, 2,745) because the
first hand-picked sample schools all happened to be primary; that filter
has been removed to match the brief's actual (broader) scope. Each
candidate's website then comes from a per-institution /enhet/{NSRId}
detail call (only made for institutions that already passed the
list-level filters, not the whole register). Records are walked in the
API's own list order, so "first 200" means the first 200 such candidates.

Resumable: see checkpoint_utils.py - each institution's result is appended
to output/norway_checkpoint.jsonl immediately, so re-running after an
interruption skips already-processed NSRIds instead of starting over.
"""
import argparse
import asyncio
import html
import json
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

from checkpoint_utils import append_checkpoint, export_xlsx_from_checkpoint, load_checkpoint

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

NSR_ENHETER_URL = "https://data-nsr.udir.no/enheter"
NSR_ENHET_URL = "https://data-nsr.udir.no/enhet/{id}"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENHETER_CACHE = os.path.join(BASE_DIR, "data", "norway_nsr_enheter.json")
CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", "norway_checkpoint.jsonl")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "norway_output.xlsx")

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
HEADERS = {"User-Agent": UA}
MAX_PAGES = 10  # brief's explicit limit - not a speed knob, do not lower this
MAX_MAIN_CONTACT_SEARCH_PAGES = 20  # client's 2026-08-12 answer: this higher
# cap applies ONLY to the targeted search for the school's "main Contact
# page" (used by the tie-breaking rules) - not a blanket increase to the
# normal MAX_PAGES=10 crawl used for general email/phone extraction.
CRAWL_DELAY = 0.5
TIMEOUT = 10
PROBE_TIMEOUT = 5  # _url_variants can produce up to 8 candidate URLs (2
# netlocs x 2 schemes x 2 path-variants) tried sequentially before the real
# crawl even starts - against a genuinely unresponsive host (connection
# hangs rather than fails fast), that's up to 8 x TIMEOUT worth of pure
# waiting per institution. This is just a reachability probe, not real
# content extraction, so a shorter timeout here doesn't cost accuracy on
# legitimately slow-but-working pages (those still get the full TIMEOUT
# once real crawling starts via robust_get).
CONCURRENCY = 20
EXPORT_EVERY = 10

PRIORITY_KEYWORDS = [
    "kontakt", "om-oss", "om_oss", "omoss", "ansatte", "administrasjon", "contact", "about",
]

# Client's rewritten best-email/best-phone spec (2026-08-03, Priority_best_email.docx) -
# exact page-name list for Norway ("Priority pages" row of their table). Distinct from
# PRIORITY_KEYWORDS above (which only affects crawl order); this list decides RANKING -
# whether a found email/phone counts as coming from a "priority page" at all.
NO_PRIORITY_PAGE_KEYWORDS = ["contact", "kontakt", "om oss", "ansatte", "administration"]

# The single "main Contact/Kontakt page" (narrower than NO_PRIORITY_PAGE_KEYWORDS
# above, which also matches Om oss/Ansatte/Administration) - used only for the
# client's tie-breaking rule 4 ("prefer the email appearing on the school's main
# Contact/Kontakt page"), 2026-08-11 clarification.
MAIN_CONTACT_PAGE_KEYWORDS = ["contact", "kontakt"]

IGNORE_LOCAL_PARTS = {"no-reply", "noreply", "webmaster", "privacy", "gdpr", "personvern"}
PLACEHOLDER_LOCAL_PARTS = {"example", "test", "sample", "yourname", "youremail", "name", "yourdomain", "someone", "user"}
PLACEHOLDER_DOMAINS = {"example.com", "example.org", "example.net", "domain.com", "yoursite.com", "yourdomain.com", "email.com"}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
CF_EMAIL_RE = re.compile(r'data-cfemail="([a-f0-9]+)"')

CLOSED_SITE_EXCLUDE_DOMAINS = {
    # Common hosting-platform/social links that turn up on a closed-site
    # placeholder page but are never themselves "the real new site" - used
    # by crawl_site's closed-site fallback (see Kvalfjord skole).
    "facebook.com", "twitter.com", "x.com", "instagram.com", "youtube.com",
    "linkedin.com", "google.com", "goo.gl", "moava.no",
    # Confirmed live bug on a 500-record test run: Akademiet Vgs Kongsberg's
    # homepage (a perfectly normal, working site - not actually closed) has
    # a legal-disclaimer footer link to lovdata.no (Norway's official law/
    # regulation database - every private school links to the regulation
    # governing them). The fallback wrongly treated that as "the real new
    # site" and wandered into lovdata.no's own huge site structure, then
    # pro.lovdata.no - the same class of bug as chasing google.com on the
    # England crawler. These government/reference sites are never a
    # school's actual destination site.
    "lovdata.no", "regjeringen.no", "udir.no", "brreg.no", "skatteetaten.no",
}

# Some kommune.no sites (e.g. Bergen) split emails into two JS variables and only
# concatenate them in the browser on click, so the full address never appears as
# plain text anywhere in the HTML - a regex over static text alone finds nothing.
# e.g. var a='Ingunn.Aase'; var b='bergen.kommune.no'; ...String.fromCharCode(64)...
JS_SPLIT_EMAIL_RE = re.compile(
    r"var\s+a\s*=\s*'([^']+)'\s*;\s*var\s+b\s*=\s*'([^']+)'\s*;[^;]*fromCharCode\(64\)"
)

# Norwegian phone numbers are 8 digits, no area code, optionally prefixed +47 / 0047.
# Digit boundaries so this can't match an 8-digit fragment out of a longer run of digits
# (postal codes, IBAN-style account numbers, etc.) the way Slovakia's numbers did.
PHONE_RE = re.compile(r"(?<!\d)(?:\+47[\s]?|0047[\s]?)?\d{2}[\s]?\d{2}[\s]?\d{2}[\s]?\d{2}(?!\d)")

# Client's updated tie-breaking spec (Priority_best_email(1).docx,
# 2026-08-11) - order matters: when two candidate emails are otherwise
# equally ranked, the one whose prefix comes FIRST in this list wins.
# "postmottak*" was given as a wildcard (matches postmottak@ and
# postmottak.schoolname@) - the trailing "*" is stripped here since the
# existing substring match (`pref in local_l`) already covers that case.
OFFICE_EMAIL_PREFIXES = [
    "postmottak", "post", "skole", "kontakt", "info", "kontor",
    "administrasjon", "sekretariat", "sekretar", "resepsjon", "sentralbord", "kontorleder",
]
RECTOR_EMAIL_PREFIXES = [
    "rektor", "skoleleder", "leder", "principal",
    "assisterenderektor", "assisterende.rektor", "administrativleder", "avdelingsleder",
]
SPECIALIST_EMAIL_PREFIXES = ["sfo", "radgiver", "helsesykepleier", "ikt", "inntak"]
CONSUMER_DOMAINS = {
    "gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "icloud.com", "online.no",
}
PHONE_CONTEXT_OFFICE = ["kontor", "sekretariat", "resepsjon", "administrasjon"]
PHONE_CONTEXT_HEAD = ["rektor"]

SITEMAP_LOC_RE = re.compile(r"<loc>\s*(?:<!\[CDATA\[)?\s*(.*?)\s*(?:\]\]>)?\s*</loc>", re.IGNORECASE)
SITEMAP_PATHS = ("/sitemap.xml", "/sitemap_index.xml")
MAX_SUB_SITEMAPS = 5

NON_PHONE_CONTEXT_KEYWORDS = [
    "iban", "kontonummer", "konto", "adresse", "postnr", "postnummer",
    "org.nr", "orgnr", "organisasjonsnummer",
]


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


def strip_www(host):
    """host.lstrip('www.') is a bug used elsewhere in this file - lstrip
    strips any of the characters w/./ from the left, not the literal "www."
    prefix (e.g. a domain starting with a run of w's/dots would get
    over-stripped). This does what was actually intended."""
    return host[4:] if host.startswith("www.") else host


def decode_cf_email(encoded):
    try:
        r = int(encoded[:2], 16)
        chars = [chr(int(encoded[i:i + 2], 16) ^ r) for i in range(2, len(encoded), 2)]
        email = "".join(chars)
        return email if EMAIL_RE.fullmatch(email) else None
    except (ValueError, IndexError):
        return None


# Joomla's built-in "Email Cloaking" anti-spam plugin (used by default on
# a huge number of Joomla sites worldwide) builds the real mailto address
# at runtime via document.write() from a JS variable assembled out of
# literal characters mixed with numeric HTML entities. A plain HTML/text
# extraction never sees this. Verified against a real Slovak school site
# (zsjakubov.sk/kontakty): decodes all 7 staff emails correctly - kept
# here too since Joomla isn't country-specific.
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
            continue  # suspiciously short - likely a name split across plain text + cloaked JS, see slovakia_crawl.py
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


def is_non_phone_context(context):
    return any(kw in context for kw in NON_PHONE_CONTEXT_KEYWORDS)


def normalize_phone(raw):
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("0047"):
        digits = digits[4:]
    elif digits.startswith("47") and len(digits) == 10:
        digits = digits[2:]
    if len(digits) == 8:
        return digits
    return None


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
                if strip_www(p.netloc.lower()) != domain:
                    continue
                page_urls.extend(_fetch_locs(sub, session))
            locs = page_urls
        urls = [loc for loc in locs if strip_www(urlparse(loc).netloc.lower()) == domain and urlparse(loc).scheme in ("http", "https")]
        if urls:
            log.append(f"  sitemap found at {path}: {len(urls)} same-domain URLs")
            return urls
    return []


def in_scope(url, domain, required_prefix):
    """Same domain, and if the register's website field pointed at a
    subpage (not a site root), stay under that subpath too. Without this,
    a school hosted at e.g. bergen.kommune.no/.../alvoen-skole would let
    the crawler wander the entire city council site (thousands of
    unrelated pages) instead of staying scoped to that one school."""
    p = urlparse(url)
    if strip_www(p.netloc.lower()) != domain or p.scheme not in ("http", "https"):
        return False
    if required_prefix and not p.path.startswith(required_prefix):
        return False
    return True


def _url_variants(url):
    """www/bare-domain and https/http are each independently a common
    real-world misconfiguration where only one variant actually resolves
    (confirmed on England/Slovakia school sites). Try all 4 combinations,
    then - if the register URL points at a specific subpage, not a site
    root - the bare root domain of each as a last resort. Kommune.no-style
    deep links to a specific article/page ID frequently go stale as
    municipalities reorganize their sites, while the site itself is still
    live at its root (confirmed on multiple Norwegian examples: e.g.
    nesodden.kommune.no/artikkel.aspx?... 404s but nesodden.kommune.no/
    loads fine). Tried last, after the exact-path variants, so a working
    deep link is always preferred over falling back to the root."""
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
    if parsed.path.strip("/") or parsed.query:
        for netloc in netlocs:
            for scheme in schemes:
                candidate = parsed._replace(scheme=scheme, netloc=netloc, path="/", params="", query="", fragment="").geturl()
                if candidate not in seen:
                    seen.add(candidate)
                    variants.append(candidate)
    return variants


def robust_get(session, url, timeout=TIMEOUT):
    """session.get() with two narrow retries (see slovakia_crawl.py for
    the full rationale): SSL verification failures retried once
    unverified; timeouts/connection errors retried once more with backoff.
    Lowered from 2 extra retries to 1 (worst case per call: TIMEOUT + 1.5s +
    TIMEOUT, was TIMEOUT + 1.5s + TIMEOUT + 3s + TIMEOUT) - a genuinely
    unresponsive host rarely starts responding on a 3rd identical attempt
    seconds later, so that extra retry was mostly just adding time, not
    recovering real institutions."""
    try:
        return session.get(url, timeout=timeout)
    except requests.exceptions.SSLError:
        return session.get(url, timeout=timeout, verify=False)
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
        time.sleep(1.5)
        return session.get(url, timeout=timeout)


META_REFRESH_RE = re.compile(
    r'<meta[^>]+http-equiv=["\']refresh["\'][^>]+content=["\']\s*\d+\s*;\s*(?:url=)?([^"\']+)["\']',
    re.IGNORECASE,
)


def _follow_meta_refresh(resp, session, max_hops=3):
    """Some legacy sites redirect via <meta http-equiv="refresh"> instead
    of a real HTTP 3xx - requests only follows HTTP-level redirects, so
    this stub page otherwise looks like a genuine, contentless 200
    response (confirmed on a Slovak school site). Follow it manually."""
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


def _is_dns_failure(exc):
    """True only for "this domain does not exist" (NXDOMAIN) - distinct
    from a connection timeout/refused, where the domain is real but slow or
    temporarily unresponsive. Confirmed necessary: vaaler-he.kommune.no is
    a genuine, reachable site that's just slow (user confirmed it loads,
    "take too much time") - the earlier fast-fail treated it exactly like
    skanland.kommune.no (which really doesn't exist, NXDOMAIN), wrongly
    denying it the main crawl loop's longer timeout+retry a fair chance."""
    msg = str(exc).lower()
    return any(s in msg for s in (
        "name or service not known", "nodename nor servname",
        "getaddrinfo failed", "name resolution", "nameresolutionerror",
    ))


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
    need to know whether an email appears on the school's MAIN Contact/
    Kontakt page specifically, and on how many distinct Priority pages it
    appears on - neither answerable from a single stored occurrence."""
    found_emails.setdefault(addr.lower(), []).append((url, title))


def _extract_contacts(soup, html, page_text, url, title, found_emails, found_phones):
    """All email/phone extraction logic for one already-fetched page -
    factored out of the main crawl loop so the targeted main-Contact-page
    search (crawl_site(), 2026-08-12) can run the exact same extraction on
    its one extra fetched page, instead of only grabbing that page's title
    and missing any contacts on it entirely (confirmed real bug: a targeted
    fetch that only recorded the title meant an email genuinely found on
    that page never got credited as being there, so the tie-break logic
    couldn't see it)."""
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

    for local_part, domain_part in JS_SPLIT_EMAIL_RE.findall(html):
        addr = f"{local_part}@{domain_part}"
        if EMAIL_RE.fullmatch(addr) and not is_placeholder(addr):
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


def crawl_site(start_url, log, _depth=0):
    probe_session = requests.Session()
    probe_ok = False
    all_dns_failures = True
    for candidate in _url_variants(start_url):
        try:
            probe = probe_session.get(candidate, headers=HEADERS, timeout=PROBE_TIMEOUT, allow_redirects=True)
            probe = _follow_meta_refresh(probe, probe_session)
        except requests.RequestException as e:
            if not _is_dns_failure(e):
                all_dns_failures = False
            continue
        all_dns_failures = False  # got a real response, so the domain exists
        if probe.status_code == 200 and len(probe.text) > 200:
            if candidate != start_url or probe.url != candidate:
                log.append(f"  {start_url} failed - variant {candidate} worked (-> {probe.url})")
            start_url = probe.url
            probe_ok = True
            break

    if not probe_ok and all_dns_failures:
        # Every variant failed with a genuine DNS NXDOMAIN - the domain
        # itself doesn't exist, so there's no point repeating the exact
        # same failing connection attempt again via robots.txt + the main
        # crawl loop's own robust_get (which has its own ~35s worst-case
        # retry-with-backoff). Confirmed as a real time sink on a 500-
        # institution run: several dead sites (skanland.kommune.no,
        # byggskolen.no, batsfjord.vgs.no, knarrlagsund.oppvekstsenter.no)
        # were being hit twice for a domain already known not to exist.
        # Timeouts/connection-refused (domain exists, just slow) fall
        # through to the normal crawl below instead of being fast-failed.
        log.append(f"  {start_url}: all variants NXDOMAIN, skipping (dead domain)")
        return {}, {}, 0, None

    parsed_start = urlparse(start_url)
    domain = strip_www(parsed_start.netloc.lower())
    start_path = parsed_start.path.rstrip("/")
    required_prefix = start_path if start_path else ""
    rp = get_robot_parser(start_url)

    visited = set()
    to_visit = [(start_url, 100)]
    found_emails = {}
    found_phones = {}
    visited_pages = []  # every (url, title) successfully fetched, regardless
    # of whether an email was found there - needed to determine which
    # Priority page genuinely EXISTS on this school's site (client's
    # 2026-08-12 answer: "main Contact page" = Contact if it exists, else
    # fall back through the brief's Priority-pages list in order).
    homepage_nav_links = []  # (anchor_text_lower, resolved_url) for every
    # same-domain link found on the homepage specifically (client's answer:
    # "search the priority page text in the anchor text of links in the
    # main link" = the homepage).
    pages_fetched = 0

    session = requests.Session()
    session.headers.update(HEADERS)
    queued = {normalize(start_url)}

    if required_prefix:
        log.append(f"  scoping crawl to same domain + path prefix '{required_prefix}' (register URL points to a subpage, not a site root)")

    for sm_url in fetch_sitemap_urls(start_url, domain, session, log):
        normf = normalize(sm_url)
        if normf in queued or not in_scope(sm_url, domain, required_prefix):
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
                if not in_scope(full, domain, required_prefix):
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

    # Determine the school's "main Contact/Kontakt page" for the tie-
    # breaking rules (client's 2026-08-12 answer): try Contact first, then
    # fall back through NO_PRIORITY_PAGE_KEYWORDS in the brief's own order.
    # First check pages we ALREADY fetched (free); if none match, check the
    # homepage's own nav links for a matching anchor text (proves the page
    # exists even though we didn't visit it), and fetch it now as one extra
    # targeted page - up to MAX_MAIN_CONTACT_SEARCH_PAGES total, not the
    # normal MAX_PAGES=10. "Not among the pages we crawled" is treated as
    # "doesn't exist" (also client-confirmed), so if nothing matches after
    # this, there is no main Contact page and that tie-break dimension just
    # doesn't apply.
    main_contact_page = None
    for kw in NO_PRIORITY_PAGE_KEYWORDS:
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

    # Closed-site fallback: several Norwegian schools' own sites have been
    # formally shut down and replaced with a static placeholder page (often
    # on the "moava.no" hosting platform) that just says "Vi har nye
    # hjemmesider her:" ("We have a new homepage here:") with a single link
    # to where the school actually lives now - typically the municipality's
    # own consolidated site. Confirmed on Kvalfjord skole: kvalfjord.skole.no
    # is exactly this pattern, linking to alta.kommune.no's own page for the
    # school, which has the real email (postkval@alta.kommune.no). These
    # pages are trivially small (no real site structure), hence the low
    # pages_fetched guard - a genuinely large site with 0 emails found isn't
    # this pattern and shouldn't trigger a fallback crawl of an unrelated
    # outbound link.
    if not found_emails and pages_fetched <= 2 and _depth < 1:
        # _depth guard confirmed necessary: without it, a fallback target
        # that *also* looks like a closed site (0 emails, <=2 pages) would
        # recurse into ANOTHER fallback, and so on - on a 500-institution
        # test run this stalled the whole batch for several minutes (one
        # worker thread stuck chaining through fallback after fallback
        # while the other 19 sat idle, since main() waits for every
        # submitted future before moving on). One level of fallback is
        # exactly what the confirmed real case (Kvalfjord -> Alta kommune)
        # needs; anything deeper is very unlikely to be a genuine
        # "moved to a new site" chain and not worth the runaway-time risk.
        try:
            hp = robust_get(session, start_url)
            hp = _follow_meta_refresh(hp, session)
            hp_soup = BeautifulSoup(hp.text, "html.parser")
        except requests.RequestException:
            hp = None
            hp_soup = None

        seen_fallback = set()
        fallback_urls = []

        # If the redirect landed on a broken URL, queue the clean root of
        # that same domain as a fallback candidate too - confirmed on
        # Sonans/privatgymnas.no: the dead link's redirect lands on the bare
        # "http://sonans.no" (no trailing slash), which the server returns
        # a plain-text 404 for, but "https://sonans.no/" works fine and has
        # real contact emails. crawl_site() on the root does full
        # extraction, not just link-scanning, so this recovers cases where
        # the broken landing page itself has no useful outbound links.
        if hp is not None and hp.status_code != 200:
            landed = urlparse(hp.url)
            root_url = f"{landed.scheme}://{landed.netloc}/"
            root_host = strip_www(landed.netloc.lower())
            if root_url != hp.url and root_host != domain and root_host not in CLOSED_SITE_EXCLUDE_DOMAINS:
                seen_fallback.add(root_host)
                fallback_urls.append(root_url)

        if hp_soup:
            for a in hp_soup.find_all("a", href=True):
                href = a["href"]
                if not href.startswith("http"):
                    continue
                full = urljoin(start_url, href)
                fhost = strip_www(urlparse(full).netloc.lower())
                if fhost == domain or fhost in seen_fallback or fhost in CLOSED_SITE_EXCLUDE_DOMAINS:
                    continue
                seen_fallback.add(fhost)
                fallback_urls.append(full)

        if fallback_urls:
            for fb_url in fallback_urls[:2]:
                log.append(f"  closed-site fallback: crawling linked domain {fb_url}")
                fb_emails, fb_phones, fb_pages, fb_main_contact = crawl_site(fb_url, log, _depth=_depth + 1)
                found_emails.update(fb_emails)
                found_phones.update(fb_phones)
                pages_fetched += fb_pages
                if main_contact_page is None and fb_main_contact is not None:
                    main_contact_page = fb_main_contact

            log.append(f"  after closed-site fallback: emails found: {len(found_emails)}")

    return found_emails, found_phones, pages_fetched, main_contact_page


async def crawl_site_playwright(start_url, log, page, max_pages=2):
    """Last-resort fetch using a real rendered browser instead of plain
    requests.get() - for sites that render their contact info via
    client-side JS (confirmed on piasballett.com: a Vue.js site where the
    footer's mailto link only exists after JS runs, invisible to any plain
    HTTP fetch). Ported from england_crawl.py's version of this function -
    see there for the full concurrency rationale (async API + asyncio.gather
    is safe across "tabs" on one thread; the sync API is what's not safe
    across OS threads).
    Deliberately does NOT reuse the CF/EEB/Joomla obfuscation decoders from
    crawl_site() - those exist specifically to decode obfuscation *without*
    running JS. A real browser already runs the JS, so page.content() shows
    the final, already-decoded DOM - a plain mailto/text/regex scan over
    that is sufficient and simpler.
    """
    def _sync_resolve():
        for candidate in _url_variants(start_url):
            try:
                probe = requests.get(candidate, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
            except requests.RequestException:
                continue
            if probe.status_code == 200:
                return candidate, probe.url
        return None, None

    # Confirmed real bug: this used to call requests.get() directly, a
    # *blocking* synchronous call sitting inside an async function. Since
    # up to 5 institutions' crawl_site_playwright calls share one asyncio
    # event loop via asyncio.gather(), a single blocking call (up to
    # TIMEOUT=10s for a slow/unresponsive site like Forus/Vaaler/Hafstad)
    # froze the *entire* event loop - starving every other concurrently
    # running institution's Playwright operations (page.goto,
    # wait_for_timeout) of any progress for that whole window. This is
    # exactly why piasballett.com worked every time in isolation but failed
    # every time inside the real batch: it only failed when a slow sibling
    # institution's blocking call happened to freeze the loop during Pias's
    # own JS-render window. asyncio.to_thread() runs it on a separate
    # thread instead, so it can never block the event loop.
    candidate, resolved_url = await asyncio.to_thread(_sync_resolve)
    if candidate:
        if resolved_url != start_url:
            log.append(f"  [playwright] resolved {start_url} -> {resolved_url}")
        start_url = resolved_url

    parsed_start = urlparse(start_url)
    domain = strip_www(parsed_start.netloc.lower())

    found_emails = {}
    found_phones = {}
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
            # Let client-side JS render/route (e.g. Vue Router redirecting
            # "/" to "/hjem" on piasballett.com). Bumped from 1.5s to 3s -
            # confirmed on a real 40-institution run that 1.5s missed the
            # email on piasballett.com even though it worked reliably in
            # isolated testing - other concurrent tabs loading heavier/
            # slower sites on the same shared browser process can starve a
            # lighter page's JS execution of CPU time under real-world load
            # in a way a same-site isolated test won't reproduce.
            await page.wait_for_timeout(3000)
            html_content = await page.content()
        except Exception as e:
            log.append(f"  [playwright] fetch failed {url}: {e}")
            if norm not in visited:
                continue
            html_content = None

        if html_content is None:
            continue

        soup = BeautifulSoup(html_content, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else ""
        page_text = soup.get_text(" ")

        for a in soup.find_all("a", href=True):
            href = a["href"]
            if href.lower().startswith("mailto:"):
                addr = unquote(href[7:].split("?")[0]).strip()
                if EMAIL_RE.fullmatch(addr) and not is_placeholder(addr):
                    link_context = a.get_text(" ", strip=True).lower()
                    _record_email(found_emails, addr.lower(), url, title)

        for m in EMAIL_RE.finditer(page_text):
            addr = m.group(0)
            if not is_placeholder(addr):
                ctx_start = max(0, m.start() - 60)
                context = page_text[ctx_start:m.start()].lower()
                _record_email(found_emails, addr.lower(), url, title)

        for m in PHONE_RE.finditer(page_text):
            norm_phone = normalize_phone(m.group(0))
            if norm_phone:
                ctx_start = max(0, m.start() - 40)
                context = page_text[ctx_start:m.start()].lower()
                if is_non_phone_context(context):
                    continue
                _record_contact(found_phones, norm_phone, url, title, context)

        if pages_fetched < max_pages:
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if href.lower().startswith(("mailto:", "tel:", "javascript:")):
                    continue
                full = urljoin(url, href)
                p = urlparse(full)
                netloc_clean = strip_www(p.netloc.lower())
                if netloc_clean != domain or p.scheme not in ("http", "https"):
                    continue
                normf = normalize(full)
                if normf in visited or normf in queued:
                    continue
                text = a.get_text(" ", strip=True)
                to_visit.append((full, score_link(href, text)))
                queued.add(normf)

    log.append(f"  [playwright] pages fetched: {pages_fetched}, emails found: {len(found_emails)}")
    return found_emails, found_phones


CDP_URL = "http://localhost:9222"  # see ENGLAND_SCRAPING_GUIDE.md - if a
# real Chrome is running with --remote-debugging-port=9222, we drive that
# instead of a fresh headless browser. Falls back to headless if no such
# Chrome is running - must keep working with zero setup.


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
    browser = await p.chromium.launch(headless=True)
    return browser, True


PLAYWRIGHT_RETRY_ATTEMPTS = 2
PLAYWRIGHT_CONCURRENCY = 5  # multiple tabs on the same browser, concurrent
# via asyncio.gather() on one thread - see england_crawl.py for why this is
# safe (unlike multi-threading the sync API).


async def _playwright_process_one(row, context, semaphore):
    async with semaphore:
        log = []
        print(f"=== [playwright retry] {row['Institution name']} ({row['Website']}) ===")
        found_emails, found_phones = {}, {}
        page = None
        try:
            page = await context.new_page()
            for attempt in range(PLAYWRIGHT_RETRY_ATTEMPTS):
                found_emails, found_phones = await crawl_site_playwright(row["Website"], log, page)
                if found_emails:
                    break
                if attempt < PLAYWRIGHT_RETRY_ATTEMPTS - 1:
                    log.append(f"  [playwright] attempt {attempt + 1} found nothing, retrying after cooldown...")
                    await page.wait_for_timeout(2000)
        except Exception as e:
            # Covers new_page() itself failing too, not just crawl_site_playwright -
            # confirmed necessary: the shared browser can die mid-batch (a
            # Chromium crash under concurrent tabs, observed on a real run),
            # and new_page() sat outside this try/except, so that one
            # uncaught exception propagated through asyncio.gather() and
            # killed every other concurrent institution's task too.
            log.append(f"  [playwright] unhandled error, treating as no email: {e}")
        finally:
            if page is not None:
                try:
                    await page.close()
                except Exception:
                    pass
        for line in log:
            print(line)
        best_email = pick_best_email(found_emails, institution_name=row["Institution name"])
        best_phone = pick_best_phone(found_phones)
        if best_email:
            row["Best email"] = best_email["best_email"]
            row["Best Email Type"] = best_email["best_email_type"]
            row["Confidence Score"] = best_email["confidence"]
            row["Source URL"] = best_email["source_url"]
            row["All emails found"] = "; ".join(sorted(found_emails.keys()))
            print(f"  -> RECOVERED via Playwright: {row['Best email']}")
        else:
            print("  -> still no email (Playwright fallback exhausted)")
        if best_phone and not row.get("Telephone"):
            row["Telephone"] = best_phone["best_phone"]
        append_checkpoint(CHECKPOINT_FILE, row["_id"], row)
        return bool(best_email)


async def _playwright_retry_pass_async(rows):
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        browser, owns_browser = await _get_playwright_browser(p)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        semaphore = asyncio.Semaphore(PLAYWRIGHT_CONCURRENCY)
        # return_exceptions=True as defense-in-depth: _playwright_process_one
        # already catches everything internally, but if some truly
        # unexpected exception slipped through anyway, one bad task must
        # never cancel every other concurrent task's still-running work
        # (confirmed the hard way: without this, one crash took down the
        # whole batch via asyncio.gather()'s default fail-fast behaviour).
        results = await asyncio.gather(
            *(_playwright_process_one(row, context, semaphore) for row in rows),
            return_exceptions=True,
        )
        if owns_browser:
            try:
                await browser.close()
            except Exception:
                pass
        return sum(1 for r in results if r is True)


def playwright_retry_pass(rows):
    """Final fallback for records still blank after the fast requests-based
    retry pass. Runs multiple institutions concurrently (see
    PLAYWRIGHT_CONCURRENCY) via Playwright's async API."""
    return asyncio.run(_playwright_retry_pass_async(rows))


def is_priority_page(url, title, keywords=NO_PRIORITY_PAGE_KEYWORDS):
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
    ranked below the tiers the table does define (same approach as
    Slovakia/England - see slovakia_crawl.py's classify_email for the full
    rationale):
      - Generic prefix found on a NON-priority page: ranked below
        Decision-maker but above the freemail tiers.
      - Anything matching none of the above (e.g. a personal-name address):
        last-resort fallback only.

    Within a tier, OFFICE_EMAIL_PREFIXES/RECTOR_EMAIL_PREFIXES are now
    ORDERED lists per the client's 2026-08-11 tie-breaking rules - an email
    matching an earlier-listed prefix outranks one matching a later-listed
    prefix when their tier is otherwise identical. The remaining tie-break
    chain (main Contact/Kontakt page > appears on multiple Priority pages >
    first encountered during crawl) is resolved in pick_best_email(), which
    has visibility across candidates that this per-email function doesn't.

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
    for idx, pref in enumerate(RECTOR_EMAIL_PREFIXES):
        if local_l == pref or pref in local_l:
            return result(400, "Rector / Headteacher", "Medium-high", prefix_rank=idx)

    # School office / administration (includes "skole@" - covers the old
    # school-name-as-mailbox case like "auglend.skole@..." via substring match).
    for idx, pref in enumerate(OFFICE_EMAIL_PREFIXES):
        if local_l == pref or pref in local_l:
            if on_priority:
                return result(500, "School Office / Administration", "High", prefix_rank=idx)
            return result(300, "School Office / Administration", "Medium", prefix_rank=idx,
                          main_contact=False, n_priority=0, source=occurrences[0])

    # Freemail/consumer domains - rank depends on whether it's on a priority page.
    if domain_l in CONSUMER_DOMAINS:
        if on_priority:
            return result(200, "Consumer / Freemail", "Medium", main_contact=False, n_priority=0, source=priority_occurrences[0])
        return result(100, "Consumer / Freemail", "Low", main_contact=False, n_priority=0, source=occurrences[0])

    # Specialist contacts - not in the client's new table but kept as a
    # low-priority fallback, same as before.
    for pref in SPECIALIST_EMAIL_PREFIXES:
        if pref in local_l:
            return result(50, "Specialist Contact", "Low", main_contact=False, n_priority=0, source=occurrences[0])

    # Last resort - doesn't match any defined tier.
    return result(10, "Other / Unclassified", "Low", main_contact=False, n_priority=0, source=occurrences[0])


def pick_best_email(found_emails, institution_name="", main_contact_page=None):
    """found_emails: {email: [(url, title), ...]} - every occurrence, per
    _record_email(). main_contact_page: (url, title) crawl_site() determined
    for this school, or None. Ranks candidates by: tier score -> prefix
    order within the tier (client's 2026-08-11 ordered lists) -> appears on
    the main Contact/Kontakt page -> appears on more than one Priority page
    -> first encountered during the crawl (dict insertion order) - the
    client's tie-breaking spec."""
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
    for kw in PHONE_CONTEXT_HEAD:
        if kw in context:
            return (60, "Rektor (head teacher)")
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


def _load_enheter_list(session):
    """The full /enheter list (~18k entities, no Website field) is cached
    to disk since it doesn't change within a scraping run and re-fetching
    a multi-MB list on every invocation would be wasteful."""
    if os.path.exists(ENHETER_CACHE):
        with open(ENHETER_CACHE, encoding="utf-8") as f:
            return json.load(f)
    resp = session.get(NSR_ENHETER_URL, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    with open(ENHETER_CACHE, "w", encoding="utf-8") as f:
        json.dump(data, f)
    return data


def _fetch_institution_detail(session, c):
    nsr_id = c["NSRId"]
    try:
        resp = session.get(NSR_ENHET_URL.format(id=nsr_id), timeout=TIMEOUT)
        detail = resp.json()
    except (requests.RequestException, ValueError):
        return None

    # NACE 85.510 filter REMOVED (2026-08-07, client decision 2026-08-06):
    # The previous filter dropped all institutions with ONLY NACE code 85.510
    # ("Sports and recreation instruction" - yoga studios, football academies,
    # martial arts clubs, snowboard clubs, etc.), silently excluding 2,495
    # institutions (1,215 ErSkole=True + 1,280 ErSkole=False) from the output.
    # The client explicitly said on 2026-08-06 to INCLUDE all active NSR entities
    # regardless of whether they are schools - they are already distinguished via
    # the Institution Category column ("School" vs "Other NSR Entity"), so the
    # client can filter them downstream if needed. No filter applied here.

    url = (detail.get("Url") or "").strip()
    if url and not url.lower().startswith(("http://", "https://")):
        url = "https://" + url
    skole_typer = [s["Navn"] for s in detail.get("SkoleTyper", []) if s.get("Navn") != "Grunnopplæring"]
    return {
        "Institution ID": detail.get("OrgNr") or str(nsr_id),
        "_nsr_id": str(nsr_id),
        "Institution name": (detail.get("Navn") or "").strip(),
        "Type": ", ".join(skole_typer),
        "Institution Category": "School" if c.get("ErSkole") else "Other NSR Entity",
        "Municipality": detail.get("KommuneNavn") or c.get("KommuneNavn", ""),
        "County": (detail.get("Fylke") or {}).get("Navn", ""),
        "Website": url,
    }


REGISTER_DETAIL_CONCURRENCY = 30  # the bulk /enheter list (fetched once,
# cached) doesn't include each institution's website/NACE code - that only
# exists on a per-institution detail endpoint, so building a candidate list
# needs one API call per institution - now fetching ALL ~8,778 of them
# concurrently every run (see load_register), not just a slice, so this
# needs to be reasonably high to keep that a ~1-2 minute cost rather than
# longer. Confirmed via direct testing that NSR's API handles 100+
# concurrent requests without any errors or slowdown.


def load_register(limit=None):
    """Yields active, web-visible grunnskoler physically located in Norway
    (see module docstring for the filter/lookup strategy), deduplicated by
    NSRId (the register's unique key). Website may be blank - per the
    brief, "one row per NSR institution" applies regardless (exclusions are
    only "Closed institutions" and "Duplicate records"), so blank-website
    institutions are still yielded here and simply skip the crawl step
    downstream.
    limit is accepted for backwards compatibility but no longer used to
    slice the candidate list - a fixed "limit + 10% buffer" slice was
    confirmed wrong on a real run: NACE-excluded businesses (yoga/fitness
    studios wrongly categorized as "Privatskole") aren't spread evenly
    through the list, they cluster (a run of "Activate/Active/Aktiv..."
    named companies sitting right after the legitimate schools), so a
    small fixed buffer came up 4x short (122 valid out of a requested 500).
    Fetching every candidate's detail concurrently instead - confirmed via
    direct testing that NSR's API handles 100+ concurrent requests fine
    (~2s), so this is a safe, one-time ~1-2 minute cost that can never
    undershoot, regardless of where exclusions happen to cluster. Order is
    no longer the API's own list order (concurrent fetches complete out of
    order) - this was never a meaningful ordering to begin with, and
    checkpoint-based resumability doesn't depend on it."""
    session = requests.Session()
    session.headers.update(HEADERS)

    data = _load_enheter_list(session)
    # NOT filtering on ErSkole anymore (2026-08-06, client decision) - NSR's
    # registry also contains non-school entities (driving schools, yoga
    # studios, sports academies, religious groups) that happen to sit in
    # the same registry. The client's brief only explicitly says to
    # exclude "Closed institutions" and "Duplicate records" - ErSkole
    # filtering was our own interpretation of "educational institutions",
    # not a literal instruction, so per the client's request these are now
    # included, tagged via "Institution Category" below so they can be
    # filtered out downstream if wanted.
    candidates = [
        d for d in data
        if d.get("ErAktiv") and d.get("VisesPaaWeb") and d.get("FylkeNr") != "25"
    ]

    seen = set()
    to_fetch = []
    for c in candidates:
        nsr_id = c["NSRId"]
        if nsr_id in seen:
            continue
        seen.add(nsr_id)
        to_fetch.append(c)

    with ThreadPoolExecutor(max_workers=REGISTER_DETAIL_CONCURRENCY) as executor:
        futures = [executor.submit(_fetch_institution_detail, session, c) for c in to_fetch]
        for future in as_completed(futures):
            result = future.result()
            if result:
                yield result


FIELDNAMES = [
    "Institution ID", "Institution name", "Type", "Institution Category", "Municipality", "County", "Website",
    "Best email", "Best Email Type", "Confidence Score", "Telephone", "All emails found", "Source URL",
    "Phone Source URL", "NSR Register Link",
]

NSR_PORTAL_URL = "https://nsr.udir.no/enheter/{org_nr}"


def process_institution(inst):
    row = {k: v for k, v in inst.items() if not k.startswith("_")}
    row["NSR Register Link"] = NSR_PORTAL_URL.format(org_nr=inst["Institution ID"])

    if not inst["Website"]:
        print(f"=== {inst['Institution name']} (no website on file - skipping crawl) ===\n")
        row["Website"] = "No Website"
        row["Best email"] = ""
        row["Best Email Type"] = ""
        row["Confidence Score"] = ""
        row["Telephone"] = ""
        row["All emails found"] = ""
        row["Source URL"] = ""
        row["Phone Source URL"] = ""
        return row

    log = []
    print(f"=== {inst['Institution name']} ({inst['Website']}) ===")
    found_emails, found_phones, pages_fetched, main_contact_page = crawl_site(inst["Website"], log)
    for line in log:
        print(line)

    best_email = pick_best_email(found_emails, institution_name=inst["Institution name"], main_contact_page=main_contact_page)
    best_phone = pick_best_phone(found_phones)

    row["Best email"] = best_email["best_email"] if best_email else ""
    row["Best Email Type"] = best_email["best_email_type"] if best_email else ""
    row["Confidence Score"] = best_email["confidence"] if best_email else ""
    row["Telephone"] = best_phone["best_phone"] if best_phone else ""
    row["All emails found"] = "; ".join(sorted(found_emails.keys()))
    row["Source URL"] = best_email["source_url"] if best_email else ""
    row["Phone Source URL"] = best_phone["source_url"] if best_phone else ""
    print(f"  -> best email: {row['Best email']} ({row['Best Email Type']}, {row['Confidence Score']}) | telephone: {row['Telephone']}")
    print()
    return row


checkpoint_lock = threading.Lock()

RETRY_CONCURRENCY = 8  # was sequential ("no concurrency contention") - on a
# 500-institution test run, the retry pass alone (44 failures, one at a
# time) was a real chunk of total runtime. Same fix as england_crawl.py:
# lower than the main pass's 20 since these are already-failed institutions,
# but still parallel instead of one-at-a-time.


def _retry_one(row):
    # Must include "Institution Category" - same class of bug confirmed in
    # Slovakia's equivalent function (2026-08-05): process_institution()
    # only ever carries forward whatever's already in the dict it's given,
    # it never re-derives fields itself, so leaving one out here means any
    # institution that goes through a retry permanently loses that field.
    inst = {k: row[k] for k in ("Institution ID", "Institution name", "Type", "Institution Category", "Municipality", "County", "Website")}
    inst["_nsr_id"] = row["_id"]
    return process_institution(inst)


USE_PLAYWRIGHT_FALLBACK = False  # off by default per explicit instruction -
# the fast requests-based retry alone is the plain, no-browser behaviour
# that ran correctly before the Playwright/JS-rendering fallback was added.
# Flip to True (or pass --playwright-fallback) to opt back into it for
# JS-rendered sites like piasballett.com.


def retry_failed_records(nsr_id_filter=None, use_playwright_fallback=USE_PLAYWRIGHT_FALLBACK):
    """Retry pass over checkpointed failures - see england_crawl.py's
    version of this function for the full rationale (concurrency-induced
    transient failures vs genuine network blocks)."""
    all_rows = load_checkpoint(CHECKPOINT_FILE)
    targets = [
        r for r in all_rows.values()
        if (nsr_id_filter is None or r["_id"] in nsr_id_filter)
        and r["Website"] != "No Website" and not r["Best email"]
    ]
    if not targets:
        print("Nothing to retry.")
        return 0, 0
    print(f"Retrying {len(targets)} failures with {RETRY_CONCURRENCY} parallel workers...")
    recovered = 0
    still_blank = []
    with ThreadPoolExecutor(max_workers=RETRY_CONCURRENCY) as executor:
        futures = {executor.submit(_retry_one, row): row for row in targets}
        for future in as_completed(futures):
            row = futures[future]
            new_row = future.result()
            if new_row["Best email"]:
                recovered += 1
            else:
                # process_institution()'s output never carries "_id" (that's
                # only injected by append_checkpoint at write-time) - without
                # this, _playwright_process_one's own append_checkpoint call
                # KeyErrors on row["_id"], silently losing the result even
                # after successfully finding an email (confirmed the hard
                # way on Pias Ballettstudio: log said "RECOVERED", checkpoint
                # stayed blank).
                new_row["_id"] = row["_id"]
                still_blank.append(new_row)
            append_checkpoint(CHECKPOINT_FILE, row["_id"], new_row)
    print(f"Recovered {recovered}/{len(targets)} via fast retry.")

    if still_blank and use_playwright_fallback:
        print(f"Trying {len(still_blank)} remaining failures via Playwright fallback...")
        pw_recovered = playwright_retry_pass(still_blank)
        recovered += pw_recovered
        print(f"Recovered {pw_recovered}/{len(still_blank)} more via Playwright.")

    return recovered, len(targets)


def main():
    global CHECKPOINT_FILE, OUTPUT_FILE
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=200, help="new records to add this run (default 200); use 0 for unlimited (all remaining candidates)")
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY, help="institutions crawled in parallel")
    parser.add_argument("--retry-only", action="store_true", help="skip crawling new institutions - just retry every existing checkpointed failure (e.g. run this from a different network to recover geo-blocked sites)")
    parser.add_argument("--playwright-fallback", action="store_true", help="also try a real rendered browser (no VPN/CDP unless one is already running) for JS-rendered sites still blank after the fast retry - off by default, launches a visible-in-process-list browser")
    parser.add_argument("--checkpoint-suffix", default="", help="write to a separate norway_checkpoint_<suffix>.jsonl / norway_output_<suffix>.xlsx instead of the default files - use this to re-run every institution (including already-checkpointed ones) under updated classification logic without touching the main checkpoint, merge back afterward")
    args = parser.parse_args()

    if args.checkpoint_suffix:
        CHECKPOINT_FILE = os.path.join(BASE_DIR, ".checkpoints", f"norway_checkpoint_{args.checkpoint_suffix}.jsonl")
        OUTPUT_FILE = os.path.join(BASE_DIR, "output", f"norway_output_{args.checkpoint_suffix}.xlsx")
        print(f"Using separate checkpoint: {CHECKPOINT_FILE}")

    if args.retry_only:
        retry_failed_records(use_playwright_fallback=args.playwright_fallback)
        total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Norway", text_columns={"Institution ID", "Telephone"})
        print(f"Total checkpointed: {total}. Wrote {OUTPUT_FILE}")
        return

    done = load_checkpoint(CHECKPOINT_FILE)
    print(f"{len(done)} institutions already checkpointed from previous runs.")

    candidates = []
    for inst in load_register(limit=args.limit):
        if inst["_nsr_id"] in done:
            continue
        candidates.append(inst)
        if args.limit and len(candidates) >= args.limit:
            break
    print(f"Crawling {len(candidates)} institutions with {args.concurrency} parallel workers...")

    added = 0
    with ThreadPoolExecutor(max_workers=args.concurrency) as executor:
        futures = {executor.submit(process_institution, inst): inst for inst in candidates}
        for future in as_completed(futures):
            inst = futures[future]
            try:
                row = future.result()
            except Exception as e:
                print(f"FAILED {inst['Institution name']}: {e}")
                continue
            with checkpoint_lock:
                append_checkpoint(CHECKPOINT_FILE, inst["_nsr_id"], row)
                added += 1
                if added % EXPORT_EVERY == 0:
                    export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Norway", text_columns={"Institution ID", "Telephone"})

    # No filter - retry EVERY outstanding blank in the whole checkpoint,
    # not just this run's own new candidates (same fix as England/Turkey -
    # see their comments for the full explanation of the bug this fixes).
    retry_failed_records(use_playwright_fallback=args.playwright_fallback)

    total = export_xlsx_from_checkpoint(CHECKPOINT_FILE, FIELDNAMES, OUTPUT_FILE, sheet_name="Norway", text_columns={"Institution ID", "Telephone"})
    print(f"Added {added} new institutions this run. Total checkpointed: {total}. Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
