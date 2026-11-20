# URL Checker

A Windows app that opens each website in a real browser (Google Chrome, or Microsoft Edge if Chrome isn't installed) and records what a visitor sees, e.g. `200 - OK` or `522 - Connection timed out`.

## Download

Go to the **Releases** page of this repository and download **URL-Checker.exe**. You don't need to install anything else.

If your computer won't run the .exe, download **URL-Checker.zip** instead. Unzip it, open the `URL-Checker` folder, and run `URL-Checker.exe` from there.

## Use

- **One site:** type the address in the box at the top and press Enter.
- **A list:** drag an Excel file onto the window, or click the dashed box to choose one. The URLs go in column A, and a header in row 1 is fine. `.csv` and `.txt` files with one URL per line also work. You can also drop the file onto the `URL-Checker.exe` icon.

Each site shows up in the list as it's checked. Click a site to watch its checks tick off as they happen: DNS lookup, opening the page, redirects, Cloudflare checks, certificate warnings and the plain-HTTP fallback.

The app waits as long as a real visitor would:

- **Slow sites:** up to 110 seconds, so slow errors like Cloudflare's 522 (about 20 seconds) are still caught.
- **Cloudflare's "Just a moment…" check:** up to 20 seconds for it to clear on its own. If it doesn't clear (because it needs a person to click), the result is `Cloudflare challenge page`. That's what visitors would see.
- **Certificate warnings:** the app looks behind the warning, as if you'd clicked "Proceed (unsafe)".
- **Failed HTTPS:** the app tries plain HTTP, as a browser would.

When the list is done, the app asks where to save the results. It suggests `<your file name> - results.xlsx` next to your file, and you can change the name or folder. The results file is a copy of your list with the result for each URL in column B. Your original file is never changed. If you press **Stop**, you can still save whatever has been checked so far. If you cancel, use **Save results…** later; the app reminds you before the results would be lost.

> On a work network, every request goes through the company proxy, so results show what someone on that network sees. That can differ from what the public sees.

## For maintainers

| File | What it is |
|---|---|
| `checker.py` | The checking engine. It also runs on its own: `python checker.py "list.xlsx"` |
| `app.py` | The window |
| `build.py` | Builds `dist\URL-Checker.exe` and `dist\URL-Checker.zip` |
| `.github/workflows/build.yml` | Builds the app on every push to `main` and publishes it to Releases |

To build locally:

```
python -m pip install -r requirements.txt pyinstaller
python build.py
```

Don't commit URL lists or results. `.gitignore` blocks Excel, CSV and text files for that reason.
