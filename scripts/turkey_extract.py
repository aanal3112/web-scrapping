"""
Turkey (MEB) validation-sample builder.

UPDATED FINDING (2026-07-14): the CHK cipher described in the annex DOES
work - our first decode attempt had a bug (see below), not the mechanism
itself.

  - Most schools' "Eposta Gonder" link carries no CHK parameter at all
    (just KeepThis/width/height/TB_iframe). Per the client's instruction,
    these are marked "CHK not present" in the Email column.
  - Where CHK IS present (found on the school homepage's own "Iletisim"
    widget, not the separate /tema/iletisim.php page's button), the
    annex's decode algorithm works - once corrected to allow leading
    padding before the local part (the annex's line "do not assume the
    part before @ has six digits" was a hint about this, not a red
    herring). Our first version assumed no leading padding and required
    the WHOLE decoded string to be a valid email, which is why it
    initially returned nothing.
  - Verified example: 125. Yil Ortaokulu's CHK decodes to
    "726071@meb.k12.tr". "726071" independently matches this exact
    school's own internal MEB folder ID, which appears in unrelated
    document links elsewhere on the same site - strong corroboration
    this decode is genuinely correct, not coincidental.

Where CHK is present but doesn't decode to a valid address, Email is
marked "CHK present but not decodable" so that case is distinguishable
from "CHK not present" in the output.

Province/District/Institution Name/Website come from the saved MEB
listing page (data/Okullar ve Diğer Kurumlar.html). Telephone/Address/
the Eposta link were read directly off each school's live site.
"""
import os
import re

from output_utils import write_xlsx

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LISTING_FILE = os.path.join(BASE_DIR, "data", "Okullar ve Diğer Kurumlar.html")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "turkey_validation_sample_demo.xlsx")

CHK_RE = re.compile(r"[?&]CHK=([a-zA-Z0-9]+)")

# Client-suggested fallback list (2026-08-03): our decode only ever tried
# to match "@meb.k12.tr" - some schools' CHK may encode a different
# suffix entirely, which the old single-target search could never find
# even if it's perfectly decodable. Tried in order, longest/most specific
# first - stop at the first suffix that yields a consistent single offset
# across its whole window. Shorter suffixes (".com", ".tr", 3-4 chars)
# carry a real risk of a spurious/coincidental match (fewer characters
# need to happen to share one offset), so results decoded via one of the
# last 4 short suffixes should be treated with more suspicion than a
# "@meb.k12.tr" or "@k12.tr" match.
CHK_SUFFIXES = [
    "@meb.k12.tr",
    "@meb.gov.tr",
    "@gmail.com",
    "@hotmail.com",
    "@outlook.com",
]


def decode_chk_email(chk, expected_suffix="@meb.k12.tr"):
    """
    Per the annex's described algorithm, confirmed and refined by the
    client's own reverse-engineering (2026-08-20 feedback round):
    1. Split CHK into consecutive 3-digit blocks.
    2. Each block = ASCII code of a character + a constant numeric offset.
    3. Find the offset using the known suffix (expected_suffix) - search
       for a run of blocks whose (block - ord(char)) is the same constant
       for every character of the suffix.
    4. The FIRST TWO blocks of the CHK are always a fixed structural
       prefix and are never part of the local part - this holds
       regardless of what those 2 blocks decode to, and regardless of
       suffix.
    5. The local part is exactly blocks[2 : suffix_start], decoded with
       the offset found in step 3. Nothing is searched or guessed - it's
       a fixed block position, not a character-format heuristic.

    UPDATED (2026-08-20): earlier versions of this function tried to
    *guess* where the local part starts, by regex-searching the decoded
    string for a plausible-looking numeric or alphanumeric run (with a
    numeric-first preference for "@meb.k12.tr" specifically, added after
    the guess-based approach kept getting fooled). The client's feedback
    on the 200-record sample revealed the real, simpler rule above and
    gave worked examples proving it: e.g. a CHK we'd decoded as
    "3L752678@meb.gov.tr" is actually "752678@meb.gov.tr" ("3L" was
    leftover junk from the first two structural blocks, which happened to
    decode into letters that looked like real content). Verified against
    all 181 "Decoded" records in the live 200-sample before adopting this:
    matches every previous case exactly (including the ones patched by
    hand, like "aladag01@meb.gov.tr" and "731761@meb.k12.tr"), and also
    catches 2 more of the same bug the client's examples didn't cover
    ("0731786@meb.k12.tr" -> "731786@meb.k12.tr", "0731727@meb.k12.tr" ->
    "731727@meb.k12.tr"). No per-suffix special-casing needed any more -
    one fixed rule for all 5 suffixes.

    UPDATED (2026-08-03): expected_suffix is no longer hardcoded to
    "@meb.k12.tr" - see resolve_email(), which now tries CHK_SUFFIXES in
    order.
    """
    if not chk.isdigit():
        return None  # non-numeric CHK (e.g. the unrelated hex session token
                      # seen on document/photo-gallery links) - not this cipher

    blocks = [chk[i:i + 3] for i in range(0, len(chk), 3)]
    if any(len(b) != 3 for b in blocks):
        return None
    codes = [int(b) for b in blocks]

    n = len(expected_suffix)
    if len(codes) < n:
        return None

    for start in range(0, len(codes) - n + 1):
        window = codes[start:start + n]
        offsets = [w - ord(c) for w, c in zip(window, expected_suffix)]
        if len(set(offsets)) == 1:
            offset = offsets[0]
            if start < 2:
                # not enough room for the 2-block structural prefix before
                # the suffix starts - this CHK doesn't fit the confirmed
                # shape, treat as not decodable rather than guessing.
                continue
            local_codes = codes[2:start]
            local = "".join(
                chr(c - offset) if 0 <= (c - offset) <= 0x10FFFF else "�"
                for c in local_codes
            )
            return local + expected_suffix
    return None


def resolve_email(eposta_link):
    """Returns (email_value, status) given the school's Eposta link. status
    is a clean, filterable label distinct from the Email column's text -
    see the "CHK not decodable" project memory for why "Not Decodable" is
    a genuine, verified outcome for a CHK that fails every suffix in
    CHK_SUFFIXES (not a decoder bug): checked every possible offset/
    position for each candidate suffix and none produce a consistent
    match, and in some cases the raw numbers are mathematically incapable
    of representing any text at all under this cipher.

    Tries CHK_SUFFIXES in order (client-suggested, 2026-08-03) and stops
    at the first suffix that decodes successfully - "@meb.k12.tr" was the
    only target tried previously, so a CHK encoding some other domain
    suffix would have been wrongly marked "not decodable" before this."""
    m = CHK_RE.search(eposta_link)
    if not m:
        return "CHK not present", "CHK not present"
    chk = m.group(1)
    for suffix in CHK_SUFFIXES:
        decoded = decode_chk_email(chk, expected_suffix=suffix)
        if decoded:
            return decoded, "Decoded"
    return "CHK present but not decodable", "Not decodable"


def parse_listing(path):
    from bs4 import BeautifulSoup
    with open(path, encoding="utf-8", errors="replace") as f:
        html = f.read()
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", id="icerik-listesi")
    schools = []
    for r in table.find_all("tr")[1:]:
        tds = r.find_all("td")
        if len(tds) < 3:
            continue
        a = tds[0].find("a")
        if not a:
            continue
        full_text = a.get_text(strip=True)
        website = a["href"]
        parts = [p.strip() for p in full_text.split(" - ")]
        if len(parts) >= 3:
            province, district, name = parts[0], parts[1], " - ".join(parts[2:])
        else:
            province, district, name = "", "", full_text
        schools.append({"Province": province, "District": district, "Institution Name": name, "Website": website})
    return schools


# Confirmed live data. Telephone/Address/Eposta link were read directly off
# each school's real site (manual browsing, since meb.k12.tr is unreachable
# from this environment's network).
SCHOOLS = [
    {
        "Institution Name": "125. Yıl Ortaokulu", "Province": "ADANA", "District": "SEYHAN",
        "Website": "https://125yilortaokuluseyhan.meb.k12.tr/",
        "Telephone": "0322 428 11 10",
        "Address": "Onur Mah., 45175 Sk. Okullar Kampüsü No2, 01100 Seyhan/Adana",
        "Eposta Link": "https://125yilortaokuluseyhan.meb.k12.tr/tema/eposta/eposta_gonder.php?CHK=287694253248252246253247262307299296244305247248244314312263280277278280272268269279263&KeepThis=true&width=50&height=75&",
    },
    {
        "Institution Name": "19 Mayıs Anadolu Lisesi", "Province": "ADANA", "District": "SEYHAN",
        "Website": "https://seyhan19mayislisesi.meb.k12.tr/",
        "Telephone": "(322) 435 6734",
        "Address": "Narlıca Mah. Şehit Jandarma Onbaşı Fahri ÖZŞEN Cad. No2 Seyhan/Adana",
        "Eposta Link": "https://seyhan19mayislisesi.meb.k12.tr/tema/eposta/eposta_gonder.php?KeepThis=true&width=50&height=75&TB_iframe=true",
    },
    {
        "Institution Name": "24 Kasım İlkokulu", "Province": "ADANA", "District": "SEYHAN",
        "Website": "https://seyhan24kasimilkokulu.meb.k12.tr/",
        "Telephone": "0322 428 98 00",
        "Address": "Aydınlar Mahallesi Aydınlar Caddesi 49030 Sokak No1 Seyhan/ADANA",
        "Eposta Link": "https://seyhan24kasimilkokulu.meb.k12.tr/tema/eposta/eposta_gonder.php?KeepThis=true&width=50&height=75&TB_iframe=true",
    },
    {
        "Institution Name": "Abdulkadir Paksoy Anadolu Lisesi", "Province": "ADANA", "District": "SEYHAN",
        "Website": "https://abdulkadirpaksoyanadolulisesi.meb.k12.tr/",
        "Telephone": "0322 454 26 03",
        "Address": "Kurtuluş Mah. Atatürk Cad. Abdulkadir Paksoy Lisesi Sitesi No: 103, Seyhan/Adana",
        "Eposta Link": "https://abdulkadirpaksoyanadolulisesi.meb.k12.tr/tema/eposta/eposta_gonder.php?KeepThis=true&width=50&height=75&TB_iframe=true",
    },
    {
        "Institution Name": "Abdurrahim Karakoç İlkokulu", "Province": "ADANA", "District": "SEYHAN",
        "Website": "https://abdurrahimkarakocilkokulu.meb.k12.tr/",
        "Telephone": "(322) 435 46 20",
        "Address": "Gülbahçesi Mh. 13306 Sk. No.53, Seyhan/Adana",
        "Eposta Link": "https://abdurrahimkarakocilkokulu.meb.k12.tr/tema/eposta/eposta_gonder.php?KeepThis=true&width=50&height=75&TB_iframe=true",
    },
]


def main():
    out_rows = []
    for s in SCHOOLS:
        row = dict(s)
        email, _ = resolve_email(s["Eposta Link"])
        row["Email"] = email
        row["Email Source URL"] = s["Eposta Link"]
        out_rows.append(row)
        print(f"{s['Institution Name']}: Email={email!r}")

    fieldnames = [
        "Institution Name", "Province", "District", "Website",
        "Telephone", "Address", "Email", "Email Source URL",
    ]
    write_xlsx(out_rows, fieldnames, OUTPUT_FILE, sheet_name="Turkey")
    print(f"\nWrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
