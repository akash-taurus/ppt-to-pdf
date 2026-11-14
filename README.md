# PPT to PDF Converter

Two ways to turn PowerPoint files (`.ppt`, `.pptx`, and friends) into PDFs:

- **A drag-and-drop desktop app** (`gui.py`) — drop files or folders on the window and they convert.
- **A command-line tool** (`ppt_to_pdf.py`) — scriptable, handles single files, batches, and folders.

Both share the same conversion backends.

## Requirements

- **Python 3.8 or newer**, with `tkinter` for the GUI (bundled with the python.org
  installers on Windows/macOS; on Debian/Ubuntu install `python3-tk`).
- **One conversion backend:**
  - **LibreOffice** (recommended, cross-platform) — the tool runs `soffice` in headless mode.
  - **Microsoft PowerPoint** on Windows — uses COM automation, which requires the
    `comtypes` package:
    ```bash
    pip install comtypes
    ```

Drag & drop in the GUI is optional: install `tkinterdnd2` to enable it, otherwise the app
falls back to its **Add files…** / **Add folder…** buttons.

```bash
pip install tkinterdnd2
```

## Installation

Nothing to build — it is two standalone scripts. Clone the repo and run them:

```bash
git clone https://github.com/akash-taurus/ppt-to-pdf.git
cd ppt-to-pdf
python ppt_to_pdf.py --help
```

The optional packages above are the only extra installs, and both tools work without
them (minus drag & drop / PowerPoint automation respectively).

## Desktop app

```bash
python gui.py
```

- **Drag & drop** presentations (or whole folders) onto the drop zone. Dropping is the
  fastest path — with **Start automatically when files are added** checked, conversion
  begins immediately.
- Use **Add files…** / **Add folder…** if you would rather browse. `Ctrl+O` adds files,
  `Ctrl+Enter` starts or stops.
- **Save to** controls where the PDFs go: *Next to each file* (default) or a folder you
  pick with **Change…**. **Open folder** reveals the destination.
- **Skip files that already have a PDF** avoids re-converting decks that are already done.
- **Remove** / **Clear** take entries off the list; **Convert all** becomes **Stop** while a
  run is in progress and halts after the current file finishes.
- Each row shows the source name, status, and details (output name, size, duration).
  Double-click a finished row to open its PDF (double-clicking a failed row shows the
  error). Right-click for *Open PDF*, *Show in folder*, *Copy error message* /
  *Show error…* (only after a failure), and *Remove from list*.
- Conversion runs on a background thread, so the window stays responsive.

## Command line

Convert a single file (writes `input.pdf` next to it):

```bash
python ppt_to_pdf.py input.pptx
```

Specify an exact output path:

```bash
python ppt_to_pdf.py input.pptx -o output.pdf
```

Convert several files, or every presentation in a folder:

```bash
python ppt_to_pdf.py a.pptx b.pptx
python ppt_to_pdf.py ./decks
```

Folders are scanned one level deep, not recursively.

Send all the PDFs to one directory:

```bash
python ppt_to_pdf.py a.pptx b.pptx ./decks --outdir ./pdfs
```

Other useful options:

```bash
python ppt_to_pdf.py input.pptx -m libreoffice        # force a backend
python ppt_to_pdf.py input.pptx -m powerpoint
python ppt_to_pdf.py input.pptx --timeout 600         # per-file timeout (seconds)
python ppt_to_pdf.py ./decks -q                       # only print errors
```

Full help:

```text
usage: ppt_to_pdf.py [-h] [-o OUTPUT] [--outdir OUTDIR]
                     [-m {auto,powerpoint,libreoffice}] [--timeout TIMEOUT]
                     [-q]
                     input [input ...]
```

| Option | Meaning |
| --- | --- |
| `input ...` | One or more `.ppt`/`.pptx` files, or folders containing them |
| `-o`, `--output` | Exact `.pdf` path; only valid with a single input file |
| `--outdir` | Directory to write PDFs into (default: next to each input) |
| `-m`, `--method` | `auto` (default), `powerpoint`, or `libreoffice` |
| `--timeout` | Per-file timeout in seconds (default: `300`) |
| `-q`, `--quiet` | Only print errors |

With `--method auto` the tool prefers PowerPoint on Windows and falls back to LibreOffice;
on macOS/Linux it uses LibreOffice. In a batch, a failure on one file is reported and the
rest continue; the exit status is `1` if anything failed.

## Supported formats

`.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`, `.ppsm`, `.pot`, `.potx`, `.potm`, `.odp`

## Using it as a library

`ppt_to_pdf.py` is importable if you want to script conversions:

```python
from ppt_to_pdf import convert, convert_many

convert("deck.pptx")                        # -> deck.pdf next to it
convert("deck.pptx", "out/deck.pdf")        # explicit destination
convert("deck.pptx", method="libreoffice")  # force a backend
results = convert_many(["a.pptx", "b.pptx"], "out")
```

`convert_many` returns a `ConversionResult` per file (`source`, `output`, `ok`, `seconds`,
`error`) and accepts an `on_result` callback for progress.

## Notes

- Conversion preserves layout as rendered by the chosen backend (PowerPoint or
  LibreOffice), so results may differ slightly between them.
- LibreOffice must be installed and either on your `PATH` as `soffice`/`libreoffice`, or in
  a standard install location. You can also point at it with the `LIBREOFFICE_PATH`
  environment variable.
- The LibreOffice backend runs each conversion with a private user profile, so several
  conversions can run at once without clashing with a LibreOffice window you already have
  open.

## License

[MIT](LICENSE) © 2026 Akash Maheshwari
