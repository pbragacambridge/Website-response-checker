"""
URL Checker - the checking engine.

Opens each URL/hostname in Google Chrome (or Microsoft Edge if Chrome isn't installed) and records
what a visitor sees, e.g. "200 - OK" or "522 - Connection timed out".

The desktop app (app.py) uses this file. It also works on its own from a terminal:
    python checker.py "list.xlsx"      (.xlsx, .csv or .txt - URLs in column A)
Saves "<file name> - results.xlsx" next to the original.
"""
import asyncio
import io
import re
import socket
import sys
import time
import zipfile
from dataclasses import dataclass
from http.client import responses
from pathlib import Path
from urllib.parse import urlparse

import dns.resolver
from openpyxl import Workbook, load_workbook
from playwright.async_api import Error as BrowserError, async_playwright

WAIT = 110      # seconds to wait for a page - slow errors like Cloudflare 522 (~20s) still get captured
PARALLEL = 20   # pages open at the same time
CLOUDFLARE = {520: "Web server is returning an unknown error", 521: "Web server is down",
              522: "Connection timed out", 523: "Origin is unreachable", 524: "A timeout occurred",
              525: "SSL handshake failed", 526: "Invalid SSL certificate"}
CERT_WARNINGS = {"ERR_CERT_COMMON_NAME_INVALID": "wrong certificate", "ERR_CERT_DATE_INVALID": "expired certificate",
                 "ERR_CERT_AUTHORITY_INVALID": "untrusted certificate"}
CHROME_ERRORS = {"ERR_NAME_NOT_RESOLVED": "DNS error - host not found",
                 "ERR_CONNECTION_TIMED_OUT": "Timed out - no response", "ERR_TIMED_OUT": "Timed out - no response",
                 "ERR_CONNECTION_REFUSED": "Connection refused", "ERR_CONNECTION_RESET": "Connection reset",
                 "ERR_CONNECTION_CLOSED": "Connection closed", "ERR_EMPTY_RESPONSE": "No data received",
                 "ERR_TOO_MANY_REDIRECTS": "Redirect loop", "ERR_ADDRESS_UNREACHABLE": "Address unreachable",
                 "ERR_SSL_PROTOCOL_ERROR": "SSL error - secure connection failed",
                 "ERR_SSL_VERSION_OR_CIPHER_MISMATCH": "SSL error - secure connection failed",
                 **{error: f"Certificate warning - {warning}" for error, warning in CERT_WARNINGS.items()}}
CHECK_LIMIT = 600  # seconds - a site that somehow takes longer than this is marked failed, so a run always finishes
RELOAD_ON = ("ERR_CONNECTION_RESET", "ERR_CONNECTION_CLOSED", "ERR_EMPTY_RESPONSE")


class Steps:
    """The checks made for one URL, in order, so the app can show them ticking off live.
    Each step is (label, state, detail, started, finished); state is running, ok, warn or fail.
    on_change gets a copy of all steps every time one starts or finishes."""

    def __init__(self, on_change=None):
        self.items = []
        self.on_change = on_change

    def start(self, label):
        self.items.append([label, "running", "", time.monotonic(), None])
        self._changed()
        return len(self.items) - 1

    def end(self, i, state, detail=""):
        self.items[i][1:] = [state, detail, self.items[i][3], time.monotonic()]
        self._changed()

    def note(self, label, state, detail=""):
        self.end(self.start(label), state, detail)

    def end_running(self, state, detail=""):
        """Finish every step still running; returns how many there were."""
        running = [i for i, item in enumerate(self.items) if item[1] == "running"]
        for i in running:
            self.end(i, state, detail)
        return len(running)

    def _changed(self):
        if self.on_change:
            self.on_change([tuple(item) for item in self.items])


def explain(error):
    return f"{CHROME_ERRORS.get(error, 'Error')} ({error})"


def verdict(result):
    """'ok', 'warn' or 'bad' - how a result should be coloured."""
    code = re.match(r"(\d{3}) - ", result or "")
    if code:
        n = int(code.group(1))
        what = result.split(" (at ")[0]  # not the address, which could contain any word
        if "Blocked by Cloudflare" in what or n >= 400:
            return "bad"
        return "warn" if "challenge page" in what or n >= 300 else "ok"
    if re.search(r"(over HTTP|behind the warning): [23]\d\d - ", result or ""):
        return "warn"
    return "bad"


def https_only(host):
    """True if the site's DNS has an HTTPS record (Cloudflare adds one to proxied sites).
    Chrome then never falls back to plain HTTP - the visitor just gets the HTTPS error."""
    try:
        dns.resolver.resolve(host, "HTTPS")
        return True
    except Exception:
        return False


def describe(code, title, html, final_url, host):
    """Turn the page Chrome ended up on into e.g. '522 - Connection timed out'."""
    text = html[:60000]
    cf_error = re.search(r"error ?code:?\s*(1\d{3})|\bError\s+(1\d{3})\b", text, re.I)
    if "you have been blocked" in text.lower():
        result = f"{code} - Blocked by Cloudflare WAF"
    elif title.lower().startswith("just a moment"):
        result = f"{code} - Cloudflare challenge page"
    elif cf_error and "cloudflare" in text.lower():
        result = f"{code} - Cloudflare error {cf_error.group(1) or cf_error.group(2)}"
        if " | " in title:
            result += f": {title.split(' | ')[0]}"
    else:
        result = f"{code} - {CLOUDFLARE.get(code) or responses.get(code, '')}"
    final = urlparse(final_url)
    path = "" if final.path in ("", "/") else final.path
    if final.hostname != host or path:
        result += f" (at {final.hostname}{path})"
    return result


async def look_up(host, steps):
    """Shows which addresses this computer's DNS gives for the site (for information only)."""
    s = steps.start(f"Look up {host} in DNS")
    try:
        found = await asyncio.wait_for(
            asyncio.get_running_loop().getaddrinfo(host, 443, type=socket.SOCK_STREAM), 5)
        ips = sorted(dict.fromkeys(f[4][0] for f in found), key=lambda ip: ":" in ip)  # IPv4 first
        steps.end(s, "ok", ", ".join(ips[:2]) + (f" (+{len(ips) - 2} more)" if len(ips) > 2 else ""))
    except Exception:
        steps.end(s, "warn", "Not found by this computer's DNS - the browser will still try")


async def close_quietly(tab_or_window):
    """Close a tab or window. The browser sometimes never answers when a page is redirecting at that moment -
    give up after 10s instead of freezing the whole run (the result is already known by then)."""
    try:
        await asyncio.wait_for(tab_or_window.close(), 10)
    except (asyncio.TimeoutError, BrowserError):
        pass


def address(url):
    """'host/path' of a URL, as shown in results, e.g. 'login.example.com/sso'."""
    u = urlparse(url)
    return f"{u.hostname}{'' if u.path in ('', '/') else u.path}"


class Moving(Exception):
    """The page is on its way to another address (a script or meta-refresh redirect, or a Cloudflare check
    reloading the page) and hasn't got there yet."""


async def read_now(page, read):
    """page.title / page.content right now. Raises Moving if the page is on its way somewhere else - reading
    then either fails ("page is navigating") or waits until that address answers, which may be never."""
    try:
        return await asyncio.wait_for(read(), 10)
    except asyncio.TimeoutError:
        raise Moving from None
    except BrowserError as e:
        if "navigat" in str(e):  # "...page is navigating..." / "...because of a navigation"
            raise Moving from None
        raise


async def on_challenge(page):
    """Is the page Cloudflare's "Just a moment..." check? None if it's on its way somewhere else."""
    try:
        return (await read_now(page, page.title)).lower().startswith("just a moment")
    except Moving:
        return None


async def where_it_lands(page, navigations, steps):
    """Title and HTML of the page a visitor ends up on. If the page has moved on to another address, wait for
    that to load first - up to WAIT seconds, like a visitor. Returns (title, html, None), or (None, None, problem)
    if that address failed or never answered, e.g. 'No data received (ERR_EMPTY_RESPONSE) (at login.example.com)'."""
    deadline, s = time.monotonic() + WAIT, None
    while True:
        try:
            title, html = await read_now(page, page.title), await read_now(page, page.content)
            break
        except Moving:
            target = address(navigations[-1].url) if navigations else address(page.url)
            if s is None:
                s = steps.start(f"Follow the page on to {target}")
            if time.monotonic() > deadline:
                problem = f"{explain('ERR_TIMED_OUT')} (at {target})"
                steps.end(s, "fail", problem)
                return None, None, problem
            await asyncio.sleep(0.5)
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=10000)
            except BrowserError:
                pass
    if page.url.startswith("chrome-error://"):  # it moved on to an address that failed - Chrome shows its error page
        failed = next((r for r in reversed(navigations) if r.failure), None)
        code = re.search(r"ERR_[A-Z_]+", (failed and failed.failure) or "")
        problem = f"{explain(code.group(0) if code else 'ERR_FAILED')} (at {address(failed.url if failed else page.url)})"
        if s is None:
            s = steps.start("Follow the page on")
        steps.end(s, "fail", problem)
        return None, None, problem
    if s is not None:
        steps.end(s, "ok")
    return title, html, None


async def visit(window, url, host, steps):
    """Open url; if the connection drops, reload like a visitor would (up to 3 tries)."""
    for attempt in range(1, 4):
        result, error = await open_page(window, url, host, steps)
        if error not in RELOAD_ON or attempt == 3:
            return result, error
        s = steps.start(f"Connection dropped - reloading (try {attempt + 1} of 3)")
        await asyncio.sleep(3)
        steps.end(s, "ok")


async def open_page(window, url, host, steps):
    """Open url in a new tab. Returns (result, None), or (None, the browser's error code)."""
    page = await window.new_page()
    pages = []
    page.on("response", lambda r: pages.append(r)
            if r.request.is_navigation_request() and r.frame == page.main_frame else None)
    navigations = []  # every address the tab sets off for, answered or not
    page.on("request", lambda r: navigations.append(r)
            if r.is_navigation_request() and r.frame == page.main_frame else None)
    s = steps.start(f"Open {url}")
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=WAIT * 1000)
        steps.end(s, "ok", "Server answered " + " → ".join(str(p.status) for p in pages) if pages else "")
        s = steps.start("Let redirects and scripts finish")
        try:
            await page.wait_for_load_state("networkidle", timeout=8000)
            steps.end(s, "ok")
        except BrowserError:
            steps.end(s, "ok", "Still busy after 8s - carrying on")
        if await on_challenge(page):
            # a Cloudflare "checking your browser" page - give it time to clear, as it would for a visitor
            s = steps.start("Wait for Cloudflare's browser check to clear")
            for waited in range(1, 21):
                await page.wait_for_timeout(1000)
                if await on_challenge(page) is False:
                    steps.end(s, "ok", f"Cleared after {waited}s")
                    break
            else:
                steps.end(s, "warn", "Still showing after 20s")
        title, html, problem = await where_it_lands(page, navigations, steps)
        if problem:
            return problem, None
        code = pages[-1].status if pages else 0
        result = describe(code, title, html, page.url, host)
        steps.note("Read what the page shows", {"bad": "fail"}.get(verdict(result), verdict(result)), result)
        return result, None
    except BrowserError as e:
        found = re.search(r"net::(ERR_[A-Z_]+)", str(e))
        error = found.group(1) if found else "ERR_TIMED_OUT" if "Timeout" in str(e) else type(e).__name__
        if error == "ERR_HTTP_RESPONSE_CODE_FAILURE" and pages:  # blank error page - Chrome shows e.g. "HTTP ERROR 404"
            result = describe(pages[-1].status, "", "", pages[-1].url, host)
            steps.end_running("fail", result)
            return result, None
        if not steps.end_running("fail", explain(error)):  # failed between steps - still show it
            steps.note("Read what the page shows", "fail", explain(error))
        return None, error
    finally:
        await close_quietly(page)


class Checker:
    """A browser that checks URLs. await start() once, then check() as often as needed, then stop()."""

    def __init__(self):
        self.browser = None
        self.name = ""            # e.g. "Google Chrome 142"
        self.user_agent = None    # the browser's normal user agent (not the "Headless" one)
        self._playwright = None
        self._restarting = asyncio.Lock()

    async def start(self):
        self._playwright = await async_playwright().start()
        try:
            await self._launch()
        except RuntimeError:
            await self.stop()
            raise

    async def _launch(self):
        problems = []
        for channel, name in (("chrome", "Google Chrome"), ("msedge", "Microsoft Edge"), (None, "Chromium")):
            try:
                self.browser = await self._playwright.chromium.launch(channel=channel)
                break
            except BrowserError as e:
                problems.append(f"{name}: {str(e).strip().splitlines()[0][:300]}")
        else:
            raise RuntimeError("Could not start Google Chrome or Microsoft Edge. Please install one of them.\n\n"
                               + "\n".join(problems[:2]))
        major = self.browser.version.split(".")[0]
        self.name = f"{name} {major}"
        self.user_agent = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                           f"Chrome/{major}.0.0.0 Safari/537.36" + (f" Edg/{major}.0.0.0" if channel == "msedge" else ""))

    async def stop(self):
        for close in (self.browser and self.browser.close, self._playwright and self._playwright.stop):
            if close:
                try:
                    await close()
                except Exception:
                    pass
        self.browser = self._playwright = None

    async def _running_browser(self):
        """The browser - started again if it has crashed or closed (e.g. on a computer short of memory)."""
        if not self._playwright:
            raise RuntimeError("no browser - Chrome or Edge could not be started")
        async with self._restarting:
            if not self.browser.is_connected():
                await self._launch()
            return self.browser

    async def check(self, host, steps=None):
        """Check one URL/hostname the way a visitor would. Returns e.g. '200 - OK'."""
        steps = steps or Steps()
        for attempt in (1, 2):
            browser = await self._running_browser()
            try:
                result = await asyncio.wait_for(self._check(browser, host, steps), CHECK_LIMIT)
                if browser.is_connected() or attempt == 2:
                    return result
            except asyncio.TimeoutError:
                steps.end_running("fail", "No answer")
                return f"Failed - no result after {CHECK_LIMIT // 60} minutes"
            except BrowserError:
                if browser.is_connected() or attempt == 2:
                    raise
            # the browser went away mid-check, so this result can't be trusted - check again in a fresh one
            steps.note("The browser closed unexpectedly - restarting it and checking again", "warn")

    async def _check(self, browser, host, steps):
        url = host if "://" in host else "https://" + host
        name = urlparse(url).hostname
        if name:
            await look_up(name, steps)
        # One window per site, so the browser remembers what it learns on the way (e.g. a site that is HTTPS-only)
        window = await browser.new_context(user_agent=self.user_agent)
        try:
            result, error = await visit(window, url, name, steps)
            if result:
                return result
            if error.startswith("ERR_CERT_"):  # like clicking "Proceed (unsafe)" on Chrome's warning page
                warning = CERT_WARNINGS.get(error, f"certificate problem ({error})")
                steps.note("Click 'Proceed (unsafe)' on the warning page, as a visitor could", "warn")
                unsafe = await browser.new_context(user_agent=self.user_agent, ignore_https_errors=True)
                try:
                    behind, error2 = await visit(unsafe, url, name, steps)
                finally:
                    await close_quietly(unsafe)
                return f"SSL error - {warning}; behind the warning: {behind or explain(error2)}"
            if error != "ERR_NAME_NOT_RESOLVED" and url.startswith("https://"):
                s = steps.start("Check DNS for an HTTPS-only record")
                only = await asyncio.to_thread(https_only, name)
                steps.end(s, "ok", "Found - browsers stay on HTTPS, so the error above is what visitors see"
                          if only else "None - a browser tries plain HTTP next")
                if not only:
                    over_http, _ = await visit(window, "http://" + url[len("https://"):], name, steps)
                    if over_http:
                        return f"{explain(error)}; over HTTP: {over_http}"
            return explain(error)
        finally:
            await close_quietly(window)


@dataclass
class Book:
    """A list of URLs read from a file."""
    path: Path
    workbook: Workbook
    rows: list    # the Excel row numbers that hold a URL
    hosts: list   # the URL in each of those rows


def read_bytes(path):
    """Read a file even while it is open in Excel. Excel locks OneDrive files so that normal reading fails
    ("Permission denied"), but it still lets a program read them if that program also allows the file to be
    renamed or deleted while it reads (FILE_SHARE_DELETE)."""
    if sys.platform != "win32":
        return Path(path).read_bytes()
    import ctypes
    import msvcrt
    import os
    from ctypes import wintypes
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateFileW.restype = wintypes.HANDLE
    kernel32.CreateFileW.argtypes = (wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                                     wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE)
    # GENERIC_READ, share read + write + delete, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL
    handle = kernel32.CreateFileW(str(path), 0x80000000, 0x7, None, 3, 0x80, None)
    if handle in (None, wintypes.HANDLE(-1).value):
        raise ctypes.WinError(ctypes.get_last_error())
    with open(msvcrt.open_osfhandle(handle, os.O_RDONLY), "rb") as f:  # closes the handle too
        return f.read()


def read_list(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(io.BytesIO(read_bytes(path)))
    elif path.suffix.lower() in (".csv", ".txt"):  # one URL per line
        wb = Workbook()
        for line in read_bytes(path).decode("utf-8-sig").splitlines():
            wb.active.append([line.split(",")[0].strip()])
    else:
        raise ValueError("Please use an Excel file (.xlsx) - or a .csv or .txt file with one URL per line.\n"
                         "Old .xls files: open in Excel and use Save As > Excel Workbook (.xlsx).")
    ws = wb.worksheets[0]
    rows = [r for r in range(1, ws.max_row + 1) if "." in str(ws.cell(r, 1).value or "")]
    return Book(path, wb, rows, [str(ws.cell(r, 1).value).strip() for r in rows])


def results_name(book):
    return f"{book.path.stem.strip()} - results.xlsx"


def save_results(book, results, out=None, missing="Not checked"):
    """Write each result into column B and save a copy of the list as `out`. Without `out`, saves
    "<name> - results.xlsx" next to the original. The original file is never changed."""
    ws = book.workbook.worksheets[0]
    if 1 not in book.rows:  # row 1 is a header
        ws.cell(1, 2, "Result")
    for r, host in zip(book.rows, book.hosts):
        ws.cell(r, 2, results.get(host, missing))
    if out:
        book.workbook.save(out)
        strip_metadata(out)
        return Path(out)
    out = book.path.with_name(results_name(book))
    stamped = f"{out.stem} {time.strftime('%H%M%S')}.xlsx"
    # if the results file is open in Excel, save a copy next to it; if the folder is read-only, use Downloads
    for target in (out, out.with_name(stamped), Path.home() / "Downloads" / stamped):
        try:
            book.workbook.save(target)
            break
        except OSError:
            if target.parent == Path.home() / "Downloads":
                raise
    strip_metadata(target)
    return target


BLANK_PROPERTIES = {  # empty document properties: no author, dates, app or company details
    "docProps/core.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                         b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
    "docProps/app.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>',
    "docProps/custom.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                           b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>',
}


def strip_metadata(path):
    """Remove author names, dates and saved folder paths from an .xlsx file."""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = BLANK_PROPERTIES.get(item.filename) or src.read(item.filename)
            if item.filename == "xl/workbook.xml":  # Excel records the folder the file was saved in here
                data = re.sub(rb"<mc:AlternateContent(?:(?!</mc:AlternateContent>).)*?absPath.*?</mc:AlternateContent>",
                              b"", data, flags=re.S)
            dst.writestr(item, data)
    tmp.replace(path)


async def check_all(hosts):
    """Terminal version: check every host, printing each result as it comes in."""
    checker, results = Checker(), {}
    await checker.start()
    print(f"Using {checker.name}\n")
    limit = asyncio.Semaphore(PARALLEL)

    async def one(host):
        async with limit:
            try:
                results[host] = await checker.check(host)
            except Exception as e:  # never let one URL stop the run
                results[host] = f"Failed - {type(e).__name__}"
        print(f"[{len(results)}/{len(hosts)}] {host:<50} {results[host]}", flush=True)

    try:
        await asyncio.gather(*(one(h) for h in hosts))
    finally:
        await checker.stop()
    return results


def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else input("Drag your file here and press Enter: ").strip('" '))
    book = read_list(path)
    hosts = list(dict.fromkeys(book.hosts))
    print(f"Checking {len(hosts)} URLs (slow ones can take up to {WAIT}s)...")
    try:
        results = asyncio.run(check_all(hosts))
    except RuntimeError as e:
        sys.exit(str(e))
    print(f"\nSaved: {save_results(book, results)}")


if __name__ == "__main__":
    main()
