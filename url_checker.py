"""
Checks each URL/hostname in column A and writes what it returns into column B ("Result").

Usage:  python url_checker.py "list.xlsx"      (.xlsx, .csv or .txt)
        or drag the file onto "Run URL Checker.bat"
Saves "<file name> - results.xlsx" next to the original.
"""
import re
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.client import responses
from pathlib import Path
from urllib.parse import urlparse

import requests
import urllib3
from openpyxl import Workbook, load_workbook

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

WAIT = 110  # seconds to wait for a page - slow errors like Cloudflare 522 (~20s) still get captured
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"}
CLOUDFLARE = {520: "Web server is returning an unknown error", 521: "Web server is down",
              522: "Connection timed out", 523: "Origin is unreachable", 524: "A timeout occurred",
              525: "SSL handshake failed", 526: "Invalid SSL certificate"}


def get(url, verify=True):
    return requests.get(url, headers=HEADERS, timeout=(20, WAIT), verify=verify)


def describe(r, host):
    """Turn a response into e.g. '522 - Connection timed out', plus where it ended up if redirected."""
    code, page = r.status_code, r.text[:60000]
    cf_error = re.search(r"error ?code:?\s*(1\d{3})", page, re.I)
    if "you have been blocked" in page.lower():
        text = f"{code} - Blocked by Cloudflare WAF"
    elif "<title>Just a moment" in page:
        text = f"{code} - Cloudflare challenge page"
    elif cf_error:
        text = f"{code} - Cloudflare error {cf_error.group(1)}"
    else:
        text = f"{code} - {CLOUDFLARE.get(code) or responses.get(code, '')}"
    final = urlparse(r.url).hostname
    return text if final == host else f"{text} (at {final})"


def failure(e):
    """Describe a request that got no page back."""
    if isinstance(e, requests.exceptions.SSLError):
        return "SSL error - secure connection failed"
    if isinstance(e, requests.exceptions.Timeout):
        return "Timed out - no response"
    if isinstance(e, requests.exceptions.TooManyRedirects):
        return "Redirect loop"
    if isinstance(e, requests.exceptions.ConnectionError):
        return "DNS error - host not found" if re.search(r"getaddrinfo|NameResolution", str(e)) else "Connection failed"
    return f"Failed - {type(e).__name__}"


def certificate_warning(e):
    """The certificate warning a browser would show (which you can click past), or None."""
    t = str(e)
    if re.search(r"mismatch|doesn't match|not valid for", t, re.I):
        return "wrong certificate"
    if "expired" in t:
        return "expired certificate"
    if re.search(r"self.signed|local issuer|unknown ca|unable to get", t, re.I):
        return "untrusted certificate"
    return None


def check(host):
    url = host if "://" in host else "https://" + host
    name = urlparse(url).hostname
    try:
        return describe(get(url), name)
    except requests.exceptions.SSLError as e:
        warning = certificate_warning(e)
        if not warning:
            return failure(e)
        try:  # like clicking "Proceed (unsafe)" in the browser
            behind = describe(get(url, verify=False), name)
        except requests.exceptions.RequestException as e2:
            behind = failure(e2)
        return f"SSL error - {warning}; behind the warning: {behind}"
    except requests.exceptions.RequestException as e:
        return failure(e)


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
    ws = wb.active

    rows = [r for r in range(1, ws.max_row + 1) if "." in str(ws.cell(r, 1).value or "")]
    hosts = [str(ws.cell(r, 1).value).strip() for r in rows]
    print(f"Checking {len(hosts)} URLs (slow ones can take up to {WAIT}s)...\n")

    if 1 not in rows:  # row 1 is a header
        ws.cell(1, 2, "Result")
    with ThreadPoolExecutor(10) as pool:  # prints each URL as soon as it finishes
        jobs = {pool.submit(check, host): (r, host) for r, host in zip(rows, hosts)}
        for n, job in enumerate(as_completed(jobs), 1):
            r, host = jobs[job]
            ws.cell(r, 2, job.result())
            print(f"[{n}/{len(hosts)}] {host:<50} {job.result()}", flush=True)

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
