"""
Opens each URL/hostname in column A in Google Chrome and writes what a visitor sees into column B ("Result").

Usage:  python url_checker.py "list.xlsx"      (.xlsx, .csv or .txt)
        or drag the file onto "Run URL Checker.bat"
Saves "<file name> - results.xlsx" next to the original.
"""
import asyncio
import re
import sys
import time
import zipfile
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
                 "ERR_SSL_VERSION_OR_CIPHER_MISMATCH": "SSL error - secure connection failed"}
user_agent = None  # Chrome's normal user agent (set once the browser starts)


def explain(error):
    return f"{CHROME_ERRORS.get(error, 'Error')} ({error})"


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


async def visit(window, url, host):
    """Open url; if the connection drops, try again like a visitor pressing reload (up to 3 tries)."""
    for _ in range(3):
        result, error = await open_page(window, url, host)
        if error not in ("ERR_CONNECTION_RESET", "ERR_CONNECTION_CLOSED", "ERR_EMPTY_RESPONSE"):
            break
        await asyncio.sleep(3)
    return result, error


async def open_page(window, url, host):
    """Open url in a new tab. Returns (result, None), or (None, Chrome's error code)."""
    page = await window.new_page()
    pages = []
    page.on("response", lambda r: pages.append(r)
            if r.request.is_navigation_request() and r.frame == page.main_frame else None)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=WAIT * 1000)
        try:
            await page.wait_for_load_state("networkidle", timeout=8000)  # let redirects and scripts finish
        except BrowserError:
            pass
        for _ in range(20):  # let a Cloudflare "checking your browser" page clear, as it would for a visitor
            if not (await page.title()).lower().startswith("just a moment"):
                break
            await page.wait_for_timeout(1000)
        code = pages[-1].status if pages else 0
        return describe(code, await page.title(), await page.content(), page.url, host), None
    except BrowserError as e:
        found = re.search(r"net::(ERR_[A-Z_]+)", str(e))
        error = found.group(1) if found else "ERR_TIMED_OUT" if "Timeout" in str(e) else type(e).__name__
        if error == "ERR_HTTP_RESPONSE_CODE_FAILURE" and pages:  # blank error page - Chrome shows e.g. "HTTP ERROR 404"
            return describe(pages[-1].status, "", "", pages[-1].url, host), None
        return None, error
    finally:
        await page.close()


async def check(browser, host):
    url = host if "://" in host else "https://" + host
    name = urlparse(url).hostname
    # One window per site, so Chrome remembers what it learns on the way (e.g. a site that is HTTPS-only)
    window = await browser.new_context(user_agent=user_agent)
    try:
        result, error = await visit(window, url, name)
        if result:
            return result
        if error.startswith("ERR_CERT_"):  # like clicking "Proceed (unsafe)" on Chrome's warning page
            unsafe = await browser.new_context(user_agent=user_agent, ignore_https_errors=True)
            try:
                behind, error2 = await visit(unsafe, url, name)
            finally:
                await unsafe.close()
            warning = CERT_WARNINGS.get(error, f"certificate problem ({error})")
            return f"SSL error - {warning}; behind the warning: {behind or explain(error2)}"
        if (error != "ERR_NAME_NOT_RESOLVED" and url.startswith("https://")
                and not await asyncio.to_thread(https_only, name)):  # Chrome falls back to plain HTTP
            over_http, _ = await visit(window, "http://" + url[len("https://"):], name)
            if over_http:
                return f"{explain(error)}; over HTTP: {over_http}"
        return explain(error)
    finally:
        await window.close()


async def launch(playwright):
    try:
        return await playwright.chromium.launch(channel="chrome")  # the installed Google Chrome
    except BrowserError:
        try:
            return await playwright.chromium.launch()
        except BrowserError:
            sys.exit("Google Chrome not found. Install Chrome, or run:  python -m playwright install chromium")


async def check_all(hosts):
    global user_agent
    results, done = {}, 0
    async with async_playwright() as p:
        browser = await launch(p)
        user_agent = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      f"Chrome/{browser.version.split('.')[0]}.0.0.0 Safari/537.36")
        limit = asyncio.Semaphore(PARALLEL)

        async def one(host):
            nonlocal done
            async with limit:
                try:
                    results[host] = await check(browser, host)
                except Exception as e:  # never let one URL stop the run
                    results[host] = f"Failed - {type(e).__name__}"
            done += 1
            print(f"[{done}/{len(hosts)}] {host:<50} {results[host]}", flush=True)

        await asyncio.gather(*(one(h) for h in hosts))
        await browser.close()
    return results


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


def main():
    path = Path(sys.argv[1] if len(sys.argv) > 1 else input("Drag your file here and press Enter: ").strip('" '))
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(path)
    else:  # .csv / .txt - one URL per line
        wb = Workbook()
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            wb.active.append([line.split(",")[0].strip()])
    ws = wb.worksheets[0]

    rows = [r for r in range(1, ws.max_row + 1) if "." in str(ws.cell(r, 1).value or "")]
    hosts = [str(ws.cell(r, 1).value).strip() for r in rows]
    print(f"Checking {len(hosts)} URLs in Chrome (slow ones can take up to {WAIT}s)...\n")

    results = asyncio.run(check_all(list(dict.fromkeys(hosts))))
    if 1 not in rows:  # row 1 is a header
        ws.cell(1, 2, "Result")
    for r, host in zip(rows, hosts):
        ws.cell(r, 2, results[host])

    out = path.with_name(f"{path.stem.strip()} - results.xlsx")
    try:
        wb.save(out)
    except PermissionError:  # results file is open in Excel - save a copy next to it instead
        out = out.with_name(f"{out.stem} {time.strftime('%H%M%S')}.xlsx")
        wb.save(out)
    strip_metadata(out)
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
