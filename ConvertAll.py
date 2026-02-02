#%% Convert a Folder of PowerPoint PPTs to PDFs

# Purpose: Converts all PowerPoint PPT/PPTX files in a folder to PDF
# Author:  Matthew Renze (improvements: reuse app, cleanup, cross-platform fallback)

# Usage:   python ConvertAll.py input-folder [output-folder]
#   - input-folder  = folder containing the PowerPoint files to convert
#   - output-folder = optional; defaults to same as input-folder

# Example: python ConvertAll.py .
# Example: python ConvertAll.py ./slides ./pdfs

# Note: On Windows with PowerPoint installed, uses COM. On macOS/Linux, uses LibreOffice if available.

import sys
import os
import subprocess
import shutil

def convert_with_libreoffice(input_folder: str, output_folder: str) -> bool:
    """Convert PPT/PPTX to PDF using LibreOffice (macOS/Linux). Returns True if any file was converted."""
    # Common locations for soffice
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    soffice = None
    for c in candidates:
        if c and os.path.isfile(c):
            soffice = c
            break
    if not soffice:
        return False

    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    exts = (".ppt", ".pptx")
    files = [f for f in os.listdir(input_folder)
             if os.path.isfile(os.path.join(input_folder, f)) and f.lower().endswith(exts)]
    if not files:
        return True  # nothing to do

    # LibreOffice --convert-to pdf --outdir <dir> file1 file2 ...
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_folder]
    cmd.extend([os.path.join(input_folder, f) for f in files])
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=300)
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False
    return True


def convert_with_powerpoint_com(input_folder: str, output_folder: str) -> None:
    """Convert PPT/PPTX to PDF using Windows COM + Microsoft PowerPoint."""
    import comtypes.client

    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    entries = os.listdir(input_folder)
    files = [f for f in entries
             if os.path.isfile(os.path.join(input_folder, f))
             and f.lower().endswith((".ppt", ".pptx"))]
    if not files:
        print("No PPT/PPTX files found in", input_folder)
        return

    powerpoint = None
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        for input_file_name in files:
            input_file_path = os.path.join(input_folder, input_file_name)
            output_name = os.path.splitext(input_file_name)[0] + ".pdf"
            output_file_path = os.path.join(output_folder, output_name)
            try:
                slides = powerpoint.Presentations.Open(input_file_path)
                slides.SaveAs(output_file_path, 32)  # 32 = ppSaveAsPDF
                slides.Close()
                print("Converted:", input_file_name, "->", output_name)
            except Exception as e:
                print("Error converting", input_file_name, ":", e)
    finally:
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass


def main():
    if len(sys.argv) < 2:
        print("Usage: python ConvertAll.py input-folder [output-folder]")
        print("Example: python ConvertAll.py .")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2] if len(sys.argv) > 2 else input_folder

    if not os.path.isdir(input_folder):
        print("Error: input folder does not exist:", input_folder)
        sys.exit(1)

    if sys.platform == "win32":
        try:
            convert_with_powerpoint_com(input_folder, output_folder)
            return
        except ImportError:
            print("comtypes not installed. Install with: pip install comtypes")
            sys.exit(1)

    # macOS / Linux: try LibreOffice first
    exts = (".ppt", ".pptx")
    root = os.path.abspath(input_folder)
    files = [f for f in os.listdir(input_folder)
             if os.path.isfile(os.path.join(input_folder, f)) and f.lower().endswith(exts)]
    if not files:
        print("No PPT/PPTX files found in", root)
        return
    if convert_with_libreoffice(input_folder, output_folder):
        print("Converted", len(files), "file(s) with LibreOffice. Output folder:", os.path.abspath(output_folder))
        return
    print("LibreOffice not found. On macOS install with: brew install --cask libreoffice")
    print("On Windows run this script with PowerPoint installed and comtypes: pip install comtypes")
    sys.exit(1)


if __name__ == "__main__":
    main()
