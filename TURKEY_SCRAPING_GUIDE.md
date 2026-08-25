# How to run the Turkey scraper (with VPN)

This is the full process for running `turkey_playwright_crawl.py` successfully.
Turkish school websites (`*.meb.k12.tr`) block plain automated requests, so
this script needs to drive your own real Chrome browser while your VPN is
genuinely connected. Follow these steps in order, every time.

## One-time setup (only needed once ever)

```bash
pip3 install playwright
playwright install chromium
```

## Every time you want to run the scraper

### Step 1 — Close Chrome completely

```bash
pkill -f chrome
```

Wait a few seconds after running this.

### Step 2 — Start Chrome in debug mode

In a terminal, run:

```bash
google-chrome --remote-debugging-port=9222
```

A normal-looking Chrome window will open. **Leave this terminal open** —
closing it or pressing Ctrl+C here will close Chrome too. You'll do the
rest of the steps in a *different* terminal window/tab.

You should see a line like:
```
DevTools listening on ws://127.0.0.1:9222/devtools/browser/...
```

### Step 3 — Turn on your VPN

In that Chrome window, click your VPN extension and turn it on. Wait for
it to say "Connected."

### Step 4 — Verify you actually have a Turkey IP

In the same Chrome window, open a new tab and go to:
`https://whatismyipaddress.com`

Confirm it shows a **Turkey** location. If it still shows your real
location (India), the VPN isn't actually connected — fix that first.
This step matters a lot: everything fails if this isn't genuinely on.

### Step 5 — Open a NEW/different terminal window

Don't reuse the terminal from Step 2 (it's busy running Chrome).

### Step 6 — Check the debug port is alive

In the new terminal:

```bash
cd "/home/zilion/Rohan Zillion/school-registry-scraper/school-registry-scraper/scripts"
curl -s http://localhost:9222/json/version
```

This should print a small block of JSON (browser version info). If it
doesn't, go back to Step 1 and start over - Chrome isn't listening
properly yet.

### Step 7 — Build the full nationwide school listing (first time only, or to refresh it)

The old `data/Okullar ve Diğer Kurumlar.html` only covered ADANA/SEYHAN
(25 schools). This step fetches every school in Turkey instead, via the
same AJAX endpoint the live site itself uses:

```bash
python3 turkey_fetch_full_listing.py
```

This writes `data/turkey_full_listing.json`. Watch the output - it prints
`recordsTotal` (how many schools MEB reports) and its running collected
count as it pages through. Report back what it prints, especially if it
errors out, since this is a new script being run against the live site
for the first time.

### Step 8 — Run the actual scraper

```bash
python3 turkey_playwright_crawl.py --limit 200
```

This automatically uses the full listing from Step 7 if present (falls
back to the old 25-school file otherwise). `--limit 200` processes the
first 200 new schools; it's resumable, so re-running the same command
later continues from wherever it left off instead of starting over.

### Step 9 — Read the results

The script prints progress per school as it goes, then writes the final
file to:

```
output/turkey_output.xlsx
```

---

## Troubleshooting

**`curl` to port 9222 gives nothing / times out**
Chrome isn't actually running with the debug flag. Run `pkill -f chrome`,
wait, then redo Step 2. Also check with:
```bash
ps aux | grep chrome | grep remote-debugging
```
You should see at least one line with `--remote-debugging-port=9222` in it.

**`ECONNREFUSED ::1:9222`**
This is an IPv6-vs-IPv4 mismatch (`::1` is IPv6's version of "localhost").
The script already uses `127.0.0.1` (IPv4) specifically to avoid this, so
if you see this error, double check the script wasn't reverted to using
`localhost` instead of `127.0.0.1` in the `connect_over_cdp(...)` line.

**`ECONNREFUSED 127.0.0.1:9222`**
Chrome was running a moment ago but isn't anymore (e.g. it restarted when
you turned the VPN on, and the restart didn't keep the debug flag). Redo
Steps 1-2, and this time make sure Step 6's `curl` check succeeds
*immediately* before running the Python script - don't leave a big gap
between them.

**Certificate errors (`certificate verify failed`, `CERT_AUTHORITY_INVALID`)**
Already handled in the script - Turkish government sites use their own
CA (Kamu SM) that most systems don't trust by default. Both the
`requests`-based and Playwright-based scripts already set
`ignore_https_errors=True` / `verify=False` for this specific, understood
reason. If you see this error anyway, the fix wasn't applied somewhere -
tell Claude.

**Redirected to `planet.news` or a "click to continue / Antibot solution" page**
This means the request was detected as automated (not a real
VPN-connected browser). Almost always means the VPN wasn't actually
connected when the script ran, or Chrome wasn't properly attached via the
debug port (Steps 1-6 above). Redo those steps carefully, especially
Step 4 (verifying the Turkey IP) before running the script.

**Email shows "CHK not present"**
This is a real, valid result for some schools - not every school's site
has the CHK-encoded email link. Nothing to fix.

**Email shows "CHK present but not decodable"**
The link had a CHK value but it didn't decode to a valid email. Save that
school's page and send it to Claude to double check.

## Files involved

- `scripts/turkey_extract.py` - shared decode logic (`decode_chk_email`,
  `resolve_email`) and the old 25-school listing-page parser (fallback only)
- `scripts/turkey_fetch_full_listing.py` - fetches the nationwide school
  listing via MEB's own AJAX endpoint (Step 7 above); run this first
- `scripts/turkey_playwright_crawl.py` - the actual scraper (Step 8 above)
- `scripts/turkey_merge_final.py` - merges all three checkpoint sources into
  one final JSONL and XLSX (see Multi-PC Merge section below)
- `data/turkey_full_listing.json` - the full nationwide listing (55,116 schools)
- `data/turkey_missing_0k_schools.json` - the 8,167 schools from the 0–45k
  range that were never attempted by the other PC; generated automatically
- `new/turkey_output.xlsx` - the Excel file uploaded from the other PC
  (44,081 rows covering a mix of 0–45k and some 45k+ schools)
- `output/turkey_output_range45k.xlsx` - this PC's 45k+ range output (complete)
- `output/turkey_output_missing0k.xlsx` - this PC's missing 8,167 run output
- `final/turkey_output_final.xlsx` - the authoritative merged final deliverable
- `.checkpoints/turkey_checkpoint.jsonl` - the other PC's checkpoint (31k rows)
- `.checkpoints/turkey_checkpoint_range45k.jsonl` - this PC's 45k+ checkpoint
- `.checkpoints/turkey_checkpoint_missing0k.jsonl` - this PC's missing0k checkpoint
- `.checkpoints/turkey_checkpoint_final.jsonl` - the final merged checkpoint

---

## Multi-PC Split-Run: Current Status & Gap Analysis

### What happened

The nationwide crawl was split across two machines:

| Machine | Range Intended | Records in File | Reality |
|:--------|:--------------|:----------------|:--------|
| Other PC | 0 – 45k | 44,081 rows | **Did NOT use `--start-index`** — ran from position 0 through the whole list without a range limit, got interrupted at 44,081 records. Of those, only **36,833 belong to the 0–45k range**; **7,248 rows actually belong to the 45k+ range** (overlap). |
| This PC | 45k – 55,116 | 10,116 rows | Complete. ✅ |

### The 8,167 gap

Because the other PC's run was not range-limited and got cut off early:

- **36,833** schools from the 0–45k range were attempted by the other PC
- **8,167** schools from the 0–45k range were **never attempted by anyone**
- These 8,167 schools are saved in `data/turkey_missing_0k_schools.json`

### Full nationwide picture (as of 2026-08-08)

| Metric | Count |
|:-------|------:|
| Total schools in registry | 55,116 |
| Attempted by other PC (0–45k range portion) | 36,833 |
| Attempted by this PC (45k+ range) | 10,116 |
| **Never attempted — GAP** | **8,167** |
| Total decoded emails (combined, so far) | 15,780 (29.1%) |
| Blank/failed entries needing retry (other PC) | 12,129 |

---

## Running the 8,167 Missing Schools (missing0k crawl)

This fills the gap left by the other PC's unfinished run. Uses the mini
listing `data/turkey_missing_0k_schools.json` so the scraper only processes
those exact 8,167 schools and writes to its own separate checkpoint.

### Step 1 — Kill Chrome and launch Profile 1

```bash
pkill -9 -f chrome || true
sleep 2
rm -f ~/.config/google-chrome/SingletonLock
DISPLAY=:1 google-chrome --profile-directory="Profile 1" --remote-debugging-port=9222 > /tmp/chrome_launch.log 2>&1 &
sleep 3
curl -s http://localhost:9222/json/version   # should print browser version JSON
```

### Step 2 — Connect your Turkey VPN

In the open Chrome window, click your VPN extension and enable Turkey.
Verify at `https://whatismyipaddress.com` that it shows a Turkey location.

### Step 3 — Launch the missing0k crawl

```bash
cd "/home/zilion/Rohan Zillion/school-registry-scraper/school-registry-scraper/scripts"

nohup python3 turkey_playwright_crawl.py \
  --listing-file data/turkey_missing_0k_schools.json \
  --checkpoint-suffix missing0k \
  --limit 0 \
  > /tmp/turkey_missing0k_run.log 2>&1 &

echo "PID: $!"
tail -f /tmp/turkey_missing0k_run.log
```

This writes to:
- `.checkpoints/turkey_checkpoint_missing0k.jsonl` (resumable checkpoint)
- `output/turkey_output_missing0k.xlsx` (interim output)
- `/tmp/turkey_missing0k_run.log` (live log)

### Step 4 — Check progress

```bash
# Check if still running
ps aux | grep turkey

# Check remaining count
python3 -c "
import json
from scripts.checkpoint_utils import load_checkpoint
with open('data/turkey_missing_0k_schools.json') as f:
    schools = json.load(f)
ckpt = load_checkpoint('.checkpoints/turkey_checkpoint_missing0k.jsonl')
done = sum(1 for s in schools if s['Website'] in ckpt)
blank = sum(1 for s in schools if s['Website'] in ckpt and not ckpt[s['Website']].get('Email'))
decoded = sum(1 for s in schools if s['Website'] in ckpt and ckpt[s['Website']].get('Email') not in ('','CHK not present','CHK present but not decodable'))
print(f'Total: {len(schools)} | Done: {done} | Remaining: {len(schools)-done} | Decoded: {decoded} | Blank (needs retry): {blank}')
"
```

### If Chrome disconnects mid-run

The scraper auto-stops after 5 consecutive Chrome failures (CHROME_CONNECTION_LOST).
Simply repeat Steps 1–3 above. The checkpoint means it resumes from where it stopped.

---

## Merging All Sources into the Final Output

Once **all three runs are complete**:
1. ✅ Other PC's Excel (`new/turkey_output.xlsx`)
2. ✅ This PC's 45k+ run (`output/turkey_output_range45k.xlsx`)
3. ✅ This PC's missing0k run (`.checkpoints/turkey_checkpoint_missing0k.jsonl`)

Run the merge script:

```bash
cd "/home/zilion/Rohan Zillion/school-registry-scraper/school-registry-scraper"
python3 scripts/turkey_merge_final.py
```

**Priority / deduplication logic** (by Website URL as unique key):
- `turkey_checkpoint_missing0k.jsonl` — **highest priority** (freshest)
- `turkey_checkpoint_range45k.jsonl` — **highest priority** (freshest)
- `new/turkey_output.xlsx` — **lowest priority** (older, fallback only)

**Outputs:**
- `final/turkey_checkpoint_final.jsonl` — authoritative merged checkpoint
- `final/turkey_output_final.xlsx` — the final client deliverable

The script also prints a full summary report:
```
Total unique institutions:   ~54,000+
Decoded emails:              ~19,000+  (~35%)
CHK not present:             ~25,000
Not decodable:               ~1,600
Blank/still needs retry:     ~12,000
```

---

## Retry Pass for Other PC's 12,129 Blank Entries

The other PC's Excel has 12,129 rows where Email is blank (scraper failed
on those schools, possibly due to Chrome crashes or VPN drops). To retry:

1. Copy `.checkpoints/turkey_checkpoint.jsonl` from the other PC to this machine
2. Run a retry-only pass on that checkpoint on the other PC:

```bash
# On the other PC (with Turkey VPN active on Profile 1):
cd "/path/to/school-registry-scraper/scripts"
python3 turkey_playwright_crawl.py --retry-only
```

Or run it on this PC after copying the checkpoint:

```bash
# Ensure Chrome + VPN active first, then:
python3 turkey_playwright_crawl.py \
  --checkpoint-suffix ""   \
  --retry-only
```

Once the retry is done, re-run `turkey_merge_final.py` to regenerate
the final merged output with the recovered emails included.

---

## 1,291 Blank Retry Pass & Final Dataset Summary (Completed 2026-08-09)

A targeted retry pass was executed on all **1,291 remaining blank records** across the nationwide registry via `data/turkey_retry_1291.json`. 

### Key Highlights
- **1,291 / 1,291 records** retried (100% coverage).
- **481 new emails** successfully decoded and recovered during the retry pass.
- All checkpoints (`turkey_checkpoint.jsonl`, `turkey_checkpoint_range45k.jsonl`, `turkey_checkpoint_missing0k.jsonl`, `turkey_checkpoint_retry1291.jsonl`, and `new/turkey_output.xlsx`) were merged with smart quality-prioritized deduplication via `scripts/turkey_merge_final.py`.

### Final Nationwide Dataset Metrics (`final/turkey_output_final.xlsx`)

| Metric | Count | Percentage |
| :--- | :---: | :---: |
| **Total Unique Institutions** | **55,116** | **100.0%** |
| **Decoded Emails** | **21,058** | **38.2%** |
| **CHK Not Present** | 32,169 | 58.4% |
| **CHK Present but Unreadable** | 1,888 | 3.4% |
| **Blank / Unreachable** | **1** | **<0.01%** |

