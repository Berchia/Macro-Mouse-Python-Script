Build Guide (Clean Install on Windows using venv)

This guide walks a first-time builder (no Python installed yet) through creating a self-contained .exe of MacroTool using a clean virtual environment. End users won’t need Python — only the developer who builds the .exe does.

============================================================
0) Prerequisites (Windows)
============================================================

1) Install Python 3.10+ (64-bit)
   - Download from the Microsoft Store or https://www.python.org/downloads/
   - During install, check “Add Python to PATH”.

2) Create a project folder and put these files in it:
   - macro_tool.py (the full script)
   - mouse.ico (optional; your icon)
   - (optional) One Program To Rule Them All.spec if building via spec

Tip: If you don’t have an icon yet, you can build without it and add one later.

============================================================
1) Open a terminal in your project folder
============================================================

Press Win + R → type cmd → Enter, then:

cd /d "C:\path\to\your\project\folder"

============================================================
2) Make a brand-new virtual environment
============================================================

python -m venv .venv

Activate it:
.venv\Scripts\activate

Your prompt should now start with (.venv).

============================================================
3) Install build tools and dependencies
============================================================

pip install --upgrade pip
pip install pyinstaller pynput

============================================================
4) Build the EXE (Use one of these two methods, Not both)
============================================================

A) Quick build from the .py file:
pyinstaller --clean --onefile --windowed --name "One Program To Rule Them All" --icon "mouse.ico" macro_tool.py

B) Build using a .spec file:
Ensure your spec has a string path for the icon:
icon=r"C:\path\to\your\project\folder\mouse.ico",
Then run:
pyinstaller --clean "One Program To Rule Them All.spec"

============================================================
5) Where is the exe?
============================================================

After a successful build:
dist\One Program To Rule Them All.exe

============================================================
6) Common issues & fixes
============================================================

"pyinstaller is not recognized"
- Activate venv: .venv\Scripts\activate
- Install: pip install pyinstaller

"The system cannot find the file specified."
- Usually a bad icon path. Use an absolute path in quotes.

Icon still doesn’t show:
- Clean build: rmdir /s /q build dist __pycache__
- Use a real .ico file (not a renamed .png). Prefer multi-size.

Hotkeys / Recording don’t work:
- Some systems require elevated rights: Right-click the exe → Run as administrator.

============================================================
7) Optional: one-click build script
============================================================

Create build.bat in your project folder:

@echo off
setlocal
cd /d "%~dp0"
if not exist .venv (
  py -m venv .venv
)
call .venv\Scripts\activate
pip install --upgrade pip >nul
pip install pyinstaller pynput >nul
rmdir /s /q build dist __pycache__ 2>nul
set ICON=mouse.ico
if exist "%ICON%" (
  pyinstaller --clean --onefile --windowed --name "One Program To Rule Them All" --icon "%ICON%" macro_tool.py
) else (
  echo [WARN] mouse.ico not found; building without custom icon.
  pyinstaller --clean --onefile --windowed --name "One Program To Rule Them All" macro_tool.py
)
echo.
echo Build complete. Opening dist...
start "" "%cd%\dist"
endlocal


============================================================
8) Some tips/Information
============================================================

- Rebuild after editing Python code.
- For crisp icons, use a 256×256 ICO that embeds multiple sizes.
- If adding external files, bundle them via --add-data or in the .spec.
- As a note, icon=r"C:\path\to\your\project\folder\mouse.ico", would be like, C:\Users\your pc name\Desktop\New folder
- You would want to have the all of the files in the same folder, that folder is considered your "project folder" and is where the "pip install" commands will download the pre-reqs to include with the .exe file you're generating.
- If you want to change the Icon for whatever reason. To keep it clean and crisp on any resolution, You would want to create a 256×256 .ICO file that has 16/32/48/64/128/256 pixel sizes embeded in it.