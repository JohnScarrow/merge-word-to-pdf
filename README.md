# merge-word-to-pdf

A zero-configuration Python script to merge multiple `.docx` files into a single `.docx` and `.pdf`, preserving formatting, images, tables and (where possible) hyperlinks and bookmarks.

## Quick summary

- Scans the `to_merge/` directory for `.docx` files (alphabetical order)
- Merges each document's pages in order, preserving page layout
- Re-inserts images at their original sizes and approximate positions
- Attempts to preserve hyperlinks and bookmarks
- Converts the merged `.docx` to PDF using LibreOffice (headless), with a fallback to `mammoth` + `wkhtmltopdf`

## Outputs

- `Merged_Doc.docx` — merged Word document
- `Merged_Doc.pdf` — PDF generated via converter

## Requirements

- Python 3.8+
- LibreOffice (preferred, for PDF conversion)
- Optional fallback tools: `mammoth` (Python) and `wkhtmltopdf` (system binary)

Install Python dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Install LibreOffice on Linux (Debian/Ubuntu):

```bash
sudo apt-get update
sudo apt-get install -y libreoffice
```

### Optional (smaller) fallback converter

If LibreOffice is not available or fails, the script falls back to converting DOCX -> HTML using `mammoth` and then HTML -> PDF using `wkhtmltopdf`.

Install the Python package `mammoth` (already in `requirements.txt`) and the `wkhtmltopdf` binary on Debian/Ubuntu:

```bash
sudo apt-get install -y wkhtmltopdf
```

On macOS you can install `wkhtmltopdf` via Homebrew:

```bash
brew install wkhtmltopdf
```

`wkhtmltopdf` is a small system package compared to LibreOffice and can be a good lightweight alternative when high Word fidelity is not required.

## Usage

Place the `.docx` files you want to merge into the `to_merge/` directory, then run:

```bash
cd merge-word-to-pdf
python merge_docs.py
```

This creates `Merged_Doc.docx` and attempts to produce `Merged_Doc.pdf`.

## Notes

- The script attempts LibreOffice (`soffice`) first. If `soffice` is not found or fails, it will try `mammoth` + `wkhtmltopdf`.
- For best fidelity (images, bookmarks, complex fields) use LibreOffice; the HTML fallback is best-effort.
- The script will attempt to recover very large or problematic `.docx` files by resaving them with LibreOffice before processing.

## Troubleshooting

- If you receive errors about LibreOffice not being found, install `libreoffice` or ensure `soffice` is on your `PATH`.
- To enable the fallback conversion install `mammoth` (Python) and `wkhtmltopdf` (system binary).

## Development

If you want to modify or extend the script locally:

```bash
git clone <repo>
cd merge-word-to-pdf
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# then edit merge_docs.py
```

## License

See `LICENSE` in the repository root.
source .venv/bin/activate
