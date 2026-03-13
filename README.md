# CTFMS EZ Certificates

CTFMS EZ Certificates is a small Windows application that generates service certificates from a Word template and saves them as PDF files. It supports normal certificates and PEP certificates with separate output folders.

---

## Download and run (recommended for most users)

1. Go to the **Releases** page of this repository:  
   [`CTFMS EZ Certificates Releases`](https://github.com/Jakkells/ctfms-ez-certificates/releases)
2. Download the file: **`CTFMS-EZ-Certificates.exe`** from the latest release.
3. Make sure **Microsoft Word** is installed on your computer (required for PDF conversion).
4. Double-click `CTFMS-EZ-Certificates.exe`.

The application will open directly. No Python or extra setup is needed.

---

## What the app does

- Lets you enter **names** and **dates** for certificates.
- Automatically calculates the **next year’s date**.
- Fills a Word template (`SERVICE CERTIFICATE.docx`) with the data.
- Saves output as **PDF** (or DOCX if PDF conversion fails).
- Can send PEP certificates to a separate configurable folder.

---

## Running from source (for advanced users)

If you prefer to run the Python source code instead of the EXE:

### Requirements

- **Windows**
- **Python 3.8+** installed and added to `PATH`
- **Microsoft Word** installed (required by `docx2pdf` for PDF conversion)

### 1. Clone or download this repository

```bash
git clone https://github.com/Jakkells/ctfms-ez-certificates.git
cd ctfms-ez-certificates
```

Or download the ZIP from GitHub and extract it.

### 2. Install dependencies and run

There is a helper script that sets up a virtual environment and runs the app:

1. In the project folder, double-click:

   - `run_app.bat`

   or run from a terminal:

   ```bash
   cd "path\to\ctfms-ez-certificates"
   run_app.bat
   ```

2. On first run this will:
   - Create a virtual environment (`.venv`).
   - Install the required packages from `requirements.txt`.
   - Launch the GUI (`certificate_app.py`).

---

## Building the EXE yourself (developer)

If you want to rebuild the standalone EXE:

1. Make sure **Python 3** is installed and on your `PATH`.
2. In a terminal:

   ```bash
   cd "path\to\ctfms-ez-certificates"
   build_exe.bat
   ```

3. When the build completes successfully, the EXE will be created at:

   ```text
   dist\CTFMS-EZ-Certificates.exe
   ```

You can then upload that file to GitHub Releases or share it directly.

---

## Notes

- The application uses the template file: **`SERVICE CERTIFICATE.docx`**.  
  Make sure this file stays in the same directory as the EXE or the Python script.
- PDF conversion uses the `docx2pdf` library, which relies on Microsoft Word.

