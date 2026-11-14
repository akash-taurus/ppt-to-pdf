# PPT to PDF Converter

Two ways to turn PowerPoint files (`.ppt`, `.pptx`, and friends) into PDFs:

- **A drag-and-drop desktop app** (`gui.py`) — drop files or folders on the window and they convert.
- **A command-line tool** (`ppt_to_pdf.py`) — scriptable, handles single files, batches, and folders.

Both share the same conversion backends.

## Requirements

You need one of the following installed:

- **LibreOffice** (recommended, cross-platform) — the tool uses `soffice` in headless mode.
- **Microsoft PowerPoint** on Windows — uses COM automation. This path requires the `comtypes` package:
  ```bash
  pip install comtypes
  ```

The GUI needs Python's `tkinter` (bundled with most Python installers). Drag & drop is
optional: install `tkinterdnd2` to enable it, otherwise the app falls back to "Add files"
/ "Add folder" buttons.

```bash
pip install tkinterdnd2
```

## Desktop app

```bash
python gui.py
```

- **Drag & drop** presentations (or whole folders) onto the drop zone. Dropping is the
  fastest path — with "Start automatically" checked, conversion begins immediately.
- Use **Add files…** / **Add folder…** if you would rather browse. `Ctrl+O` adds files,
  `Ctrl+Enter` starts or stops.
- **Save to** controls where the PDFs go: *Next to each file* (default) or a folder you pick.
- **Skip files that already have a PDF** avoids re-converting decks that are already done.
- Each row shows the source name, status, and details (output name, size, duration).
  Double-click a finished row to open its PDF; right-click for *Open PDF*, *Show in folder*,
  *Copy error message*, and *Remove from list*.
- Conversion runs on a background thread, so the window stays responsive. **Stop** halts
  after the current file finishes.

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

Send all the PDFs to one directory:

```bash
python ppt_to_pdf.py a.pptx b.pptx ./decks --outdir ./pdfs
```

Other useful options:

```bash
python ppt_to_pdf.py input.pptx --method libreoffice   # force a backend
python ppt_to_pdf.py input.pptx --method powerpoint
python ppt_to_pdf.py input.pptx --timeout 600          # per-file timeout (seconds)
python ppt_to_pdf.py ./decks -q                        # only print errors
```

With `--method auto` (the default) the tool prefers PowerPoint on Windows and falls back
to LibreOffice; on macOS/Linux it uses LibreOffice. In a batch, a failure on one file is
reported and the rest continue.

## Supported formats

`.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`, `.ppsm`, `.pot`, `.potx`, `.potm`, `.odp`

## Using it as a library

`ppt_to_pdf.py` is importable if you want to script conversions:

```python
from ppt_to_pdf import convert, convert_many

convert("deck.pptx")                        # -> deck.pdf next to it
convert("deck.pptx", "out/deck.pdf")        # explicit destination
results = convert_many(["a.pptx", "b.pptx"], "out")
```

## Notes

- Conversion preserves layout as rendered by the chosen backend (PowerPoint or
  LibreOffice), so results may differ slightly between them.
- LibreOffice must be installed and either on your `PATH` as `soffice`/`libreoffice`, or in
  a standard install location. You can also point at it with the `LIBREOFFICE_PATH`
  environment variable.
- The LibreOffice backend runs each conversion with a private user profile, so several
  conversions can run at once without clashing with a LibreOffice window you already have
  open.
