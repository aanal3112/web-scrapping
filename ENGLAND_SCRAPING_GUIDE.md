# England Scraping Guide

Covers: running the England crawler end-to-end, and specifically how to set up
the real-Chrome + VPN retry pass that recovers Cloudflare-blocked and
IP-blocked schools (Holy Trinity, Miles Coverdale, Gospel Oak, etc.).

## Files involved

- Script: `scripts/england_crawl.py`
- Register source: `data/edubasealldata20260714.csv` (GIAS full-establishment export)
- Checkpoint (resumable state): `.checkpoints/england_checkpoint.jsonl`
- Output deliverable: `output/england_output.xlsx`

## Step 1 - Close all Chrome windows

Remote debugging can only be enabled at Chrome *launch*, not on an
already-running instance. If Chrome is already open, close it fully first:

```bash
pkill -9 -f "opt/google/chrome/chrome"
```

This will close whatever tabs/windows you currently have open (Chrome offers
to restore them next time you open it normally).

## Step 2 - Relaunch Chrome with remote debugging enabled

Use the Chrome profile that has the VPN extension installed (in this setup,
that's `Profile 1` - check `chrome://version` in your own Chrome to confirm
which profile you use, if different):

```bash
nohup google-chrome --remote-debugging-port=9222 --profile-directory="Profile 1" --no-first-run about:blank > /tmp/chrome_cdp.log 2>&1 &
```

## Step 3 - Turn the VPN extension on

In the Chrome window that just opened, click the VPN extension icon and
connect (pick a UK server if it offers a choice). Quickly confirm it's
actually working by visiting `https://icanhazip.com` in that browser and
checking the IP changed from your normal one.

**Known issue:** this specific free VPN extension has crash-looped before
(its background service worker fails to start - visible as repeated
`DidStartWorkerFail` errors in the `chrome_cdp.log` file). If pages hang or
won't load at all after connecting, that's the symptom. Fix: remove and
reinstall the extension via `chrome://extensions` + the Chrome Web Store,
then repeat from Step 1.

## Step 4 - Verify the CDP connection is up

```bash
curl -s http://localhost:9222/json/version
```

Should return a JSON blob with `"Browser": "Chrome/..."`. If this fails
("Connection refused"), Chrome either isn't running or was already running
before you added the debugging flag (go back to Step 1).

## Step 5 - Run the crawler

**Fresh 200-record run** (clears checkpoint, re-crawls everything - use this
whenever the code itself has changed, so all 200 records reflect the same
version):

```bash
cd scripts
rm -f ../.checkpoints/england_checkpoint.jsonl   # only if you want a clean slate
python3 -u england_crawl.py --limit 200
```

**Retry-only** (re-attempts every checkpointed record that has a website but
no email found yet - use this to try recovering stragglers without
re-crawling the whole 200):

```bash
python3 -u england_crawl.py --retry-only
```

Both commands automatically detect the CDP connection from Step 4 and use it
for the hardest cases (Cloudflare challenges, IP-blocked sites). If Chrome
with debugging isn't running, they silently fall back to a plain headless
browser instead - nothing breaks, you just lose the VPN/real-browser
advantage for that run.

## What actually happens when you run it

1. **Main pass** - crawls all candidate schools with 30 parallel workers
   (plain HTTP requests, no browser). Takes a few minutes for 200 records.
2. **Fast retry pass** - re-attempts every record that came back blank, 8
   workers in parallel, 1 attempt each (plain HTTP again). Catches
   transient blips from the main pass's concurrency.
3. **Playwright/CDP fallback** - whatever's *still* blank after that goes
   through a real rendered browser. If your Chrome from Step 2-4 is up, it
   reuses that (real Chrome binary + your VPN extension + no automation
   flag = can pass challenges/blocks a plain script can't). Runs up to 5
   schools *concurrently* as separate tabs on the same browser. 2 attempts
   per school with a short cooldown between them (some blocks are just an
   overloaded VPN exit node, not a permanent thing).

Total time for 200 records end-to-end: roughly 4-5 minutes with a working
CDP connection.

## Checking progress / results

```bash
# is it still running, and for how long
ps aux | grep "england_crawl.py"

# how many records are done so far
python3 -c "
import json
seen=set()
with open('.checkpoints/england_checkpoint.jsonl') as f:
    for l in f:
        if l.strip(): seen.add(json.loads(l)['_id'])
print(len(seen), '/ 200')
"
```

**Important gotcha:** when you background a command, the PID your shell
reports back is not always the real running process (it can be a wrapper
shell). Always confirm with:

```bash
ps aux | grep "python3 -u.*england_crawl.py" | grep -v grep
```

and use *that* PID for anything you do next (monitoring, killing, etc.).
Also always check for duplicate processes before starting a new run -
running two crawls against the same checkpoint file at once corrupts it:

```bash
pkill -9 -f "python3 -u scripts/england_crawl.py"   # kill any stragglers first
```

## Known genuinely-unresolved cases (not bugs)

As of the last verified run, out of 200:
- 13 schools have no website listed in the government register at all -
  nothing to crawl (mostly Charedi/Jewish independent schools).
- **The Cavendish School** (cavendish-school.co.uk) - server returns a
  malformed HTTP response on every attempt, confirmed via multiple methods
  including a real browser. Genuinely broken on their end.
- **Camden Primary Pupil Referral Unit** (robsonhouse.org.uk/main/) - page
  loads fine, just doesn't have an email published on it anywhere.

Everything else recovers automatically given a working CDP+VPN connection.
