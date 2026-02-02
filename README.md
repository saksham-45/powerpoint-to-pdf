# PowerPoint to PDF

A set of Python scripts to convert PowerPoint (PPT/PPTX) files to PDF.

- **Convert.py** – Converts a single PowerPoint (PPT/PPTX) file to PDF
- **ConvertAll.py** – Converts all PowerPoint (PPT/PPTX) files in a folder to PDFs
- **ConvertHere.py** – Converts all PowerPoint (PPT/PPTX) files in the script’s folder to PDFs (Windows COM only)

## Quick start: convert all PPTX in current directory

```bash
python ConvertAll.py .
```

Output PDFs are written to the same folder. To use a different output folder:

```bash
python ConvertAll.py . ./pdfs
```

## Platform requirements

| Platform | Engine | Requirements |
|----------|--------|--------------|
| **Windows** | Microsoft PowerPoint (COM) | PowerPoint installed, `pip install comtypes` |
| **macOS** | LibreOffice | `brew install --cask libreoffice` |
| **Linux** | LibreOffice | Install `libreoffice` (e.g. `apt install libreoffice`) |

On Windows, run from a folder with PPT/PPTX files or pass input (and optional output) folder. On macOS/Linux, LibreOffice is used automatically when available.

## Improvements made to ConvertAll.py

- **Single PowerPoint instance (Windows)** – One COM application is created and reused for all files, then closed with `Quit()` to avoid leaving PowerPoint running.
- **Skip non-files** – Only regular files with `.ppt`/`.pptx` are converted; directories and other entries are ignored.
- **Output folder validation** – Input folder must exist; output folder is created if missing.
- **Optional output folder** – Second argument is optional; default is the same as the input folder.
- **Cross-platform** – On macOS/Linux, uses LibreOffice headless when available so the script runs without Windows or PowerPoint.
- **Clear errors** – Usage message when no arguments given; explicit messages when LibreOffice or comtypes are missing.
- **Cleanup** – `try/finally` ensures PowerPoint is quit even if a conversion fails.

## Requirements (Windows only)

```bash
pip install -r requirements.txt
```

(`comtypes` is only needed on Windows for the PowerPoint COM path.)
