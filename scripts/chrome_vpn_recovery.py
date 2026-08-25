"""
Chrome/VPN self-recovery for the Turkey crawler.

Kept in its own file (2026-08-21, user's request) so turkey_playwright_crawl.py
only needs a small, isolated hook into this, instead of this whole sequence
living inline in the main crawl loop.

What this does, end to end, when Chrome/VPN has died mid-crawl (or hasn't
even started yet - also used to bootstrap a completely cold start from a
single command, see turkey_playwright_crawl.py):
1. Kill any leftover Chrome processes (a crashed tab can leave zombies behind
   that block a clean relaunch - confirmed live 2026-08-21, ~25 had piled up).
2. Clear the SingletonLock file (Chrome refuses to start a second instance
   against the same profile otherwise).
3. Relaunch Chrome fresh, with the debug port pointed at a DEDICATED
   automation profile - see DEDICATED_PROFILE_DIR below for why this isn't
   just the user's real/default profile.
4. Wait for the CDP debug port to actually respond.
5. Reconnect Playwright to the new Chrome instance.
6. Open the VPN extension's own popup page (chrome-extension://.../app/index.html)
   as a normal page and click its connect button directly via the DOM - this
   is NOT simulated mouse/screen input, it's the same kind of automation
   already used everywhere else in this script (page.click() over CDP), which
   works here because the popup is just an ordinary web page under the hood.
   Confirmed live 2026-08-21: reconnected in ~2 seconds, no manual click
   needed at all.
7. Verify the connection is genuinely routed through Turkey (checks the real
   IP via ifconfig.co, not just the extension's own "Connected" label, in
   case the extension shows "Connected" before the tunnel is actually up).

Every step here was proven working individually, live, before this was
written as reusable code - nothing here is untested guesswork (on Linux,
where this session's environment actually runs - the Windows/Mac branches
below follow the same proven sequence with OS-appropriate commands/paths,
but haven't been live-tested on those platforms the way the Linux path has).

CROSS-PLATFORM (2026-08-21, user's request): the user runs this same setup
across 3 PCs - one each on Windows, Mac, and Ubuntu. Every OS-specific piece
(killing Chrome, finding its config folder, launching it) is isolated into
the functions below via platform.system(), rather than scattered Linux-only
shell commands throughout - the VPN-click/verify/orchestration logic beneath
that point is identical on all 3 platforms, since it all goes through
Playwright/CDP once Chrome is up, which behaves the same everywhere.

DEDICATED PROFILE, NOT THE REAL ONE (2026-08-21, confirmed live on a real
second PC): the original approach reused the user's actual default Chrome
profile (via --profile-directory, later also tried --user-data-dir pointed
straight at it, then a symlink to it) so the already-installed VPN extension
would be available. All of these were rejected by Chrome with "DevTools
remote debugging requires a non-default data directory" - confirmed the
check resolves symlinks before comparing (a deliberate Google security
hardening against exactly this kind of workaround, closing a real past
security hole), so no path trick around the real profile's location works.
The only reliable fix: use a genuinely SEPARATE, dedicated profile that was
never Chrome's default in the first place, at DEDICATED_PROFILE_DIR below -
Chrome has no restriction at all against remote-debugging a profile like
that.

NO MANUAL SETUP NEEDED (2026-08-21, user's request - "can we script this
manual step?"): the dedicated profile starts completely empty, so it would
otherwise need the VPN extension installed into it by hand once per
machine, via the Chrome Web Store's "Add to Chrome" flow - which itself
can't be automated (it triggers a native browser confirmation dialog
outside the page DOM, not reachable by Playwright). Worked around entirely:
the extension's manifest.json has a fixed signing "key" field, which means
loading it via Chrome's --load-extension flag (developer/unpacked mode - a
plain command-line flag, no dialog involved) registers it under the EXACT
SAME extension ID as a normal Web Store install would. So instead of
installing anything, this finds the extension's already-installed files in
whatever REAL Chrome profile has it (see _find_real_extension_source()),
copies them ONCE to a fixed location, and launches Chrome with
--load-extension pointed at that copy every time from then on - fully
automatic, confirmed working live end to end (2026-08-21).
"""
import asyncio
import glob
import os
import platform
import subprocess

from playwright.async_api import async_playwright

CDP_PORT = 9222
CDP_URL = f"http://127.0.0.1:{CDP_PORT}"

VPN_EXTENSION_ID = "abnmgemfnjndaolfflfpjcckiamamein"  # "VPN Browser" -
# identified live via chrome://extensions (2026-08-21); confirmed by the
# user (2026-08-21) to be the same extension across all 3 crawling PCs, so
# this ID should be valid everywhere (Chrome Web Store extension IDs are
# derived from the extension's own signing key, not randomly assigned per
# install). If the VPN extension is ever changed to a different one, this ID
# (and possibly the selectors below, if the new extension's popup markup
# differs) will need updating to match.
VPN_POPUP_URL = f"chrome-extension://{VPN_EXTENSION_ID}/app/index.html"
VPN_CONNECT_BUTTON_SELECTOR = ".connection-button"
VPN_STATUS_SELECTOR = ".status-value"
VPN_CURRENT_COUNTRY_FLAG_SELECTOR = ".select-selected img"  # the currently
# selected server's flag icon - src is like ".../flags/svg/tr.svg" (the
# 2-letter code identifies the country, "tr" = Turkey). More reliable than
# reading the country NAME text, since that name sits in a plain <span>
# with no distinguishing class, right next to the unrelated "Change" label
# in the same container - identified live 2026-08-21.
VPN_CHANGE_COUNTRY_BUTTON_SELECTOR = ".change-text"
VPN_TURKEY_COUNTRY_ITEM_SELECTOR = '.country-list__item:has(img[src$="/tr.svg"])'  # in
# the country picker list that "Change" reveals, each entry has the same
# flag-icon pattern - matching on that instead of the "Turkey" text avoids
# any risk of matching a differently-labeled but similarly-named entry.

CDP_POLL_INTERVAL_S = 2
CDP_POLL_MAX_ATTEMPTS = 10  # 2s * 10 = 20s ceiling waiting for Chrome to launch
VPN_CONNECT_POLL_INTERVAL_S = 2
VPN_CONNECT_POLL_MAX_ATTEMPTS = 15  # 2s * 15 = 30s ceiling waiting for VPN to connect

_OS = platform.system()  # "Linux", "Darwin" (Mac), or "Windows"

DEDICATED_PROFILE_DIR = os.path.join(os.path.expanduser("~"), ".turkey_crawler_chrome_profile")
# Fixed, same on every OS (just under the home directory) - a completely
# separate Chrome profile, used ONLY by this automation, never the user's
# real one. See the module docstring above for why this is necessary at all.

EXTENSION_COPY_DIR = os.path.join(os.path.expanduser("~"), ".turkey_crawler_vpn_extension")
# A one-time, persistent copy of the VPN extension's own files, loaded into
# the dedicated profile above via --load-extension on every launch - see
# the module docstring's "NO MANUAL SETUP NEEDED" section for why this
# works and replaces the Web Store install entirely.


def _real_chrome_config_dir():
    """Where the user's REAL, everyday Chrome profile(s) live - used ONLY
    to locate the VPN extension's already-installed files to copy from
    (see _find_real_extension_source()), never launched into directly
    (that's exactly the restriction DEDICATED_PROFILE_DIR exists to avoid)."""
    if _OS == "Windows":
        return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data")
    if _OS == "Darwin":
        return os.path.expanduser("~/Library/Application Support/Google/Chrome")
    return os.path.expanduser("~/.config/google-chrome")  # Linux


def _find_real_extension_source():
    """Searches every profile folder in the user's REAL Chrome install for
    the VPN extension, returns the path to its actual code (the versioned
    subfolder, e.g. ".../Extensions/<id>/1.2.0_0"), or None if it isn't
    installed anywhere on this machine at all."""
    matches = glob.glob(os.path.join(_real_chrome_config_dir(), "*", "Extensions", VPN_EXTENSION_ID, "*"))
    version_dirs = [m for m in matches if os.path.isdir(m)]
    return version_dirs[0] if version_dirs else None


def _ensure_extension_copied():
    """One-time (per machine) copy of the VPN extension's files to
    EXTENSION_COPY_DIR, so --load-extension has something fixed and
    persistent to point at on every launch. Safe to call every time -
    does nothing once the copy already exists. Raises RuntimeError only
    if the extension genuinely isn't installed in ANY real Chrome profile
    on this machine at all (a real one-time setup gap, not something this
    can work around - the extension has to exist SOMEWHERE first)."""
    if os.path.exists(os.path.join(EXTENSION_COPY_DIR, "manifest.json")):
        return
    source = _find_real_extension_source()
    if source is None:
        raise RuntimeError(
            f"The VPN extension (id {VPN_EXTENSION_ID}) isn't installed in any "
            f"Chrome profile on this machine yet. Install it once in your normal "
            f"Chrome (Chrome Web Store, search for the VPN extension, Add to "
            f"Chrome) - after that, this script will find and reuse it "
            f"automatically, no further manual steps needed."
        )
    import shutil
    shutil.copytree(source, EXTENSION_COPY_DIR)
    print(f"    (copied VPN extension from {source} to {EXTENSION_COPY_DIR} - one-time, won't repeat)")


def _chrome_binary():
    """Finds the actual Chrome executable - install location differs by OS
    (and even by install method on Windows). Tries the standard/common
    locations for each OS, in order, using the first one that exists."""
    if _OS == "Windows":
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        return "chrome"  # last resort - hope it's on PATH
    if _OS == "Darwin":
        return "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    return "google-chrome"  # Linux - standard symlink on PATH


def _chrome_kill_pattern():
    """Pattern to match AGAINST RUNNING PROCESSES for pkill - NOT
    necessarily the same as the command used to launch Chrome. Confirmed
    live (2026-08-21, on two separate real Ubuntu PCs) - "google-chrome"
    is just a thin wrapper script on PATH that execs the REAL binary at a
    different path (/opt/google/chrome/chrome); pkill -f matching against
    the wrapper name never matches the actual running process, so a
    genuinely still-running Chrome silently survives the "kill" step, and
    the following "launch" just hands off to that already-running
    instance instead of starting fresh with the debug port (visible in
    Chrome's own log as "Opening in existing browser session") - which is
    exactly why the debug port then never comes up. Mac's _chrome_binary()
    already returns the real full binary path (no wrapper script
    involved there), so it's reused as-is for Mac."""
    if _OS == "Linux":
        return "/opt/google/chrome/chrome"
    return _chrome_binary()  # Mac - already the real path


def _kill_chrome_and_clear_locks():
    if _OS == "Windows":
        subprocess.run(["taskkill", "/F", "/IM", "chrome.exe", "/T"], capture_output=True)
    else:
        # pkill exists on both Linux and Mac
        subprocess.run(["pkill", "-9", "-f", _chrome_kill_pattern()], capture_output=True)

    lock_file = os.path.join(DEDICATED_PROFILE_DIR, "SingletonLock")
    try:
        os.remove(lock_file)
    except OSError:
        pass  # doesn't exist - fine, nothing to clear


def _detect_display():
    """UPDATED (2026-08-21): confirmed live on a real second PC - hardcoding
    ":1" as a fallback was simply wrong there (its actual display is ":0"),
    and even the "already set in the environment" case can't be trusted
    blindly - DISPLAY can be set in an interactive shell without actually
    being exported to child processes, so os.environ inside this Python
    process doesn't always see it even when `echo $DISPLAY` in the terminal
    that launched it shows a value. Checking which X11 socket actually
    EXISTS on disk (/tmp/.X11-unix/X<N>) is environment-independent - it
    reflects what's really running, not what a shell variable claims."""
    if os.environ.get("DISPLAY"):
        return os.environ["DISPLAY"]
    sockets = glob.glob("/tmp/.X11-unix/X*")
    if sockets:
        num = os.path.basename(sockets[0])[1:]  # "X0" -> "0"
        return f":{num}"
    return ":0"  # last-resort guess - ":0" is the more common default
    # display number than ":1" anyway, if nothing else can be determined.


def _dedicated_profile_has_manual_install():
    """True once a machine has done the manual Web Store fallback install
    directly into the dedicated profile (see the ERR_BLOCKED_BY_CLIENT
    handling in _connect_vpn_and_verify()) - in that case --load-extension
    must NOT also be passed, since that would register the same extension
    ID twice (once from the real install, once from the loaded copy) and
    conflict with the properly-installed one."""
    return os.path.exists(os.path.join(DEDICATED_PROFILE_DIR, "Default", "Extensions", VPN_EXTENSION_ID))


def _launch_chrome():
    print(f"    (launching Chrome with dedicated profile: {DEDICATED_PROFILE_DIR!r} on {_OS})")
    binary = _chrome_binary()
    args = [binary, f"--user-data-dir={DEDICATED_PROFILE_DIR}"]
    if _dedicated_profile_has_manual_install():
        print("    (using the manually-installed extension in this profile - skipping --load-extension)")
    else:
        args.append(f"--load-extension={EXTENSION_COPY_DIR}")
    args.append(f"--remote-debugging-port={CDP_PORT}")
    env = os.environ.copy()
    if _OS == "Linux":
        env["DISPLAY"] = _detect_display()  # X11 display - only meaningful
        # on Linux; Mac/Windows don't use an X server so this is skipped
        # there. See _detect_display()'s own docstring for why this isn't
        # just a hardcoded/setdefault guess any more.
        print(f"    (using DISPLAY={env['DISPLAY']!r})")
    # UPDATED (2026-08-21): confirmed live on a real second PC - this used
    # to send Chrome's own stdout/stderr to /dev/null, which meant a launch
    # failure here was completely silent (just "didn't come up in 20s",
    # no reason). Logging to a file instead so a genuine failure is
    # actually diagnosable instead of a dead end.
    log_path = os.path.join(os.path.expanduser("~"), ".chrome_recovery_launch.log")
    with open(log_path, "w") as log_file:
        subprocess.Popen(args, env=env, stdout=log_file, stderr=subprocess.STDOUT, start_new_session=True)
    print(f"    (Chrome's own output, if it fails to start, will be in {log_path})")


async def _wait_for_cdp():
    for _ in range(CDP_POLL_MAX_ATTEMPTS):
        try:
            async with async_playwright() as p:
                browser = await p.chromium.connect_over_cdp(CDP_URL)
                await browser.close()
            return True
        except Exception:
            await asyncio.sleep(CDP_POLL_INTERVAL_S)
    return False


async def ensure_turkey_vpn(browser):
    """Checks (and fixes, if needed) an ALREADY-CONNECTED Browser object's
    VPN state, ending with it genuinely routed through Turkey - or raises
    RuntimeError if it can't get there. Returns the same browser back.

    UPDATED (2026-08-21): pulled out of _connect_vpn_and_verify() so this
    same check runs EVERY time, not just during a fresh Chrome
    launch/recovery. Confirmed live - if Chrome happened to already be
    running when the crawler started (e.g. left over from a previous
    session that got force-killed without also closing Chrome), the
    original code would just reuse it as-is with ZERO verification that
    the VPN was even connected, let alone on Turkey - it silently started
    crawling through whatever connection Chrome happened to have. Now
    called on BOTH paths (see turkey_playwright_crawl.py and
    _connect_vpn_and_verify() below) so an already-running Chrome gets the
    exact same scrutiny as a freshly-launched one."""
    context = browser.contexts[0]
    page = await context.new_page()

    try:
        await page.goto(VPN_POPUP_URL, timeout=15000)
    except Exception as e:
        if "ERR_BLOCKED_BY_CLIENT" in str(e):
            # UPDATED (2026-08-21): confirmed live on a real second PC -
            # this specific machine's Chrome refuses to run the extension
            # at all when loaded via --load-extension (a real Chrome
            # security policy against unpacked/developer-mode extensions,
            # seen on some installations/versions but not others - this
            # dev machine's Chrome allows it fine, so it's not universal).
            # Nothing in this script can override that - it's enforced by
            # Chrome itself, not a bug in the copied files. Falls back to
            # the one-time manual Web Store install on machines like this.
            await page.close()
            await browser.close()
            raise RuntimeError(
                f"This machine's Chrome is blocking the VPN extension when "
                f"loaded automatically (a Chrome security policy against "
                f"unpacked/developer-mode extensions - not every Chrome "
                f"install enforces this, but this one does). One-time "
                f"manual step needed here instead: run this once, install "
                f"the VPN extension from the Chrome Web Store into the "
                f"window that opens, connect it, select Turkey, then close "
                f"Chrome:\n\n  {_chrome_binary()} --user-data-dir=\"{DEDICATED_PROFILE_DIR}\"\n\n"
                f"After that one-time step, this command will work on its "
                f"own from then on (the manual install replaces the need "
                f"for --load-extension on this specific machine)."
            )
        raise
    # UPDATED (2026-08-21): confirmed live - the popup's DOM is genuinely
    # empty right when goto() finishes (query_selector immediately after
    # returns None), then fills in via the extension's own JS a moment
    # later. A flat 1.5s sleep was based on testing against an already-
    # warmed-up Chrome instance and wasn't enough on a genuinely fresh
    # launch (the extension's background service worker/offscreen document
    # also has to spin up for the first time). wait_for_selector() waits
    # for the actual element to exist instead of guessing a fixed delay.
    await page.wait_for_selector(VPN_STATUS_SELECTOR, timeout=15000)

    # UPDATED (2026-08-21, user's request): the extension can be genuinely
    # "Connected" while still pointed at some OTHER country's server (e.g.
    # confirmed live: a fresh profile defaulted to United Kingdom) - being
    # "Connected" at all was previously treated as good enough, but for
    # this crawler specifically it has to be Turkey. Checked and switched
    # BEFORE the connect-status check below, since changing server while
    # already connected commonly disconnects/reconnects anyway - simplest
    # to always settle on the right country first, then ensure connected.
    current_flag = await page.eval_on_selector(VPN_CURRENT_COUNTRY_FLAG_SELECTOR, "el => el.getAttribute('src')")
    if not current_flag or "/tr.svg" not in current_flag:
        await page.click(VPN_CHANGE_COUNTRY_BUTTON_SELECTOR)
        await page.click(VPN_TURKEY_COUNTRY_ITEM_SELECTOR)
        await page.wait_for_timeout(1000)  # let the selection actually register before reading status next

    status = await page.eval_on_selector(VPN_STATUS_SELECTOR, "el => el.textContent.trim()")
    if status.lower() not in ("connected", "connecting..."):
        # UPDATED (2026-08-21): confirmed live - switching country (just
        # above) ALSO auto-triggers a reconnect on its own if the
        # extension was already connected/connecting - status shows
        # "Connecting..." immediately after the switch, with no click
        # needed. The OLD check here ("anything that isn't literally
        # Connected") didn't know that, and clicked the connect button
        # AGAIN on top of that already-in-progress auto-reconnect - since
        # it's a single toggle button, that redundant click flips it back
        # OFF instead of helping, which is exactly why a fresh switch to
        # Turkey was timing out here even though the extension was doing
        # the right thing on its own. Only click when genuinely idle
        # ("Not connected") - "Connecting..." just needs waiting for.
        await page.click(VPN_CONNECT_BUTTON_SELECTOR)
    if status.lower() != "connected":
        connected = False
        for _ in range(VPN_CONNECT_POLL_MAX_ATTEMPTS):
            await page.wait_for_timeout(VPN_CONNECT_POLL_INTERVAL_S * 1000)
            status = await page.eval_on_selector(VPN_STATUS_SELECTOR, "el => el.textContent.trim()")
            if status.lower() == "connected":
                connected = True
                break
        if not connected:
            await page.close()
            await browser.close()
            raise RuntimeError("VPN extension never reached 'Connected' status after clicking connect")

    # the extension's own "Connected" label is necessary but not sufficient -
    # confirm the real, live IP is actually routed through Turkey before
    # trusting it (same check used manually throughout this whole project).
    verify_page = await context.new_page()
    await verify_page.goto("https://ifconfig.co/country", timeout=20000)
    country = (await verify_page.inner_text("body")).strip()
    await verify_page.close()
    await page.close()

    if "türkiye" not in country.lower() and "turkey" not in country.lower():
        await browser.close()
        raise RuntimeError(f"VPN extension says 'Connected' but real IP shows '{country}', not Turkey")

    return browser


async def _connect_vpn_and_verify(playwright):
    """Connects to the already-launched dedicated-profile Chrome, then
    hands off to ensure_turkey_vpn() for the actual check/fix/verify work
    (shared with the "Chrome was already running" path in
    turkey_playwright_crawl.py - see that function's docstring)."""
    browser = await playwright.chromium.connect_over_cdp(CDP_URL)
    return await ensure_turkey_vpn(browser)


async def recover(playwright):
    """Full recovery sequence: kill -> relaunch -> wait for CDP -> connect
    VPN -> verify Turkey IP -> return a fresh, ready-to-use Browser object.
    Raises RuntimeError if any step can't be completed. Works the same way
    regardless of which of the 3 supported OSes it's running on - the OS
    differences are all contained in the helper functions above."""
    if not _dedicated_profile_has_manual_install():
        _ensure_extension_copied()  # one-time per machine, no-op after the first call

    _kill_chrome_and_clear_locks()
    await asyncio.sleep(2)
    _launch_chrome()

    if not await _wait_for_cdp():
        raise RuntimeError(f"Chrome did not come up on {CDP_URL} within {CDP_POLL_MAX_ATTEMPTS * CDP_POLL_INTERVAL_S}s of relaunching")

    return await _connect_vpn_and_verify(playwright)
