"""
Builds the app into two downloads:
    URL-Checker.exe   one file - double-click to run
    URL-Checker.zip   the same app as a folder - for computers that block the single file

    python -m pip install -r requirements.txt pyinstaller
    python build.py            (puts them in dist\\)
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else HERE / "dist"
NAME = "URL-Checker"


def pyinstaller(work, *extra):
    subprocess.run([
        sys.executable, "-m", "PyInstaller", str(HERE / "app.py"), "--name", NAME, "--noconfirm", "--clean",
        "--windowed",                                   # no black console window
        "--icon", str(HERE / "icon.ico"), "--add-data", f"{HERE / 'icon.ico'}:.",
        "--collect-all", "tkinterdnd2",                 # drag and drop
        "--collect-data", "sv_ttk",                     # Windows 11 look
        "--collect-submodules", "dns",                  # DNS record types are loaded by name
        "--exclude-module", "PIL",                      # optional for openpyxl; keeps local and GitHub builds the same
        "--workpath", str(work / "build"), "--specpath", str(work), *extra,
    ], check=True)


with tempfile.TemporaryDirectory() as tmp:
    tmp = Path(tmp)
    pyinstaller(tmp / "onefile", "--onefile", "--distpath", str(tmp / "onefile" / "dist"))
    pyinstaller(tmp / "onedir", "--distpath", str(tmp / "onedir" / "dist"))
    OUT.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tmp / "onefile" / "dist" / f"{NAME}.exe", OUT / f"{NAME}.exe")
    shutil.make_archive(str(OUT / NAME), "zip", tmp / "onedir" / "dist")  # unzips to a URL-Checker folder

for f in (OUT / f"{NAME}.exe", OUT / f"{NAME}.zip"):
    print(f"Built {f}  ({f.stat().st_size / 1e6:.0f} MB)")
