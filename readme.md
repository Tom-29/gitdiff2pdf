# GitDiff2PDF

Generate clean, PR-style **PDF** and **Word** documents from unified `git diff` output.

## Features

| Feature | Details |
|---|---|
| **Unified view** (default) | One column — deletions in red, additions in green, context grey |
| **Side-by-side view** | Old (left) vs. New (right) with per-side line numbers |
| **Word output** | Optional `.docx` alongside the PDF (requires `python-docx`) |
| **File badges & hunk headers** | Grouped by file then hunk; `@@ … @@` headers can be shown or hidden (`--hunk-header`) |
| **Line numbers** | Subtle gutter numbers for every line |
| **Pagination** | Keep-together per file/hunk, widow/orphan protection |
| **Themes** | Light (default) and dark |
| **Fonts** | System fonts (Consolas / Segoe UI on Windows, DejaVu on Linux) with Courier fallback |
| **Encoding** | Auto-detects UTF-8, UTF-8-BOM, UTF-16 LE/BE, Latin-1 |

---

## Requirements

- **Python 3.9+**
- **PyMuPDF** (`pymupdf`)
- **python-docx** *(optional, for Word output)*

```bash
pip install pymupdf python-docx
```

---

## Quick Start

### 1. Create a diff

```powershell
# PowerShell 5.1 (ensure UTF-8)
git diff <from> <to> -- . ':(exclude)path/to/big/folder' |
  Out-File -Encoding utf8 compare.diff

# PowerShell 7+ / Bash / WSL / Linux / macOS
git diff <from> <to> -- . ':(exclude)path/to/big/folder' > compare.diff
```

### 2. Generate the output

```bash
# Unified PDF (Bitbucket-style)
python gitdiff2pdf.py compare.diff -o output.pdf --title "Changed Code" --landscape

# Only changes (hide context)
python gitdiff2pdf.py compare.diff -o output.pdf --hide-context

# Hide @@ hunk headers (clean separator only)
python gitdiff2pdf.py compare.diff -o output.pdf --hunk-header hide

# Side-by-side
python gitdiff2pdf.py compare.diff -o output.pdf --view side-by-side

# PDF + Word
python gitdiff2pdf.py compare.diff -o output.pdf --word

# PDF + Word with custom path
python gitdiff2pdf.py compare.diff -o output.pdf --word-output review.docx

# Dark theme
python gitdiff2pdf.py compare.diff -o output.pdf --theme dark --word

# Pipe directly from git (no temp file)
git diff <from> <to> | python gitdiff2pdf.py - -o output.pdf
```

---

## CLI Reference

| Flag | Default | Description |
|---|---|---|
| `inputs` | — | One or more `.diff` files, or `-` for STDIN |
| `-o, --output` | *(required)* | Output PDF path |
| `--word` | `false` | Also generate a `.docx` (same base name as `--output`) |
| `--word-output FILE` | — | Custom `.docx` output path |
| `--title TEXT` | `Changed Code` | Document title in the header |
| `--view MODE` | `unified` | `unified` or `side-by-side` |
| `--hide-context` | `false` | Show only `+`/`-` lines in unified view |
| `--hunk-header MODE` | `show` | `show` = display `@@ … @@` headers, `hide` = empty blue separator |
| `--landscape` | `false` | A4 landscape orientation |
| `--font-size PT` | `9.5` | Monospace font size |
| `--tabsize N` | `4` | Tab-to-spaces expansion width |
| `--theme THEME` | `light` | `light` or `dark` |
| `--debug` | `false` | Print parser debug info to stderr |
| `--mono-font-file` | — | Custom TTF/OTF for monospace |
| `--mono-bold-font-file` | — | Custom TTF/OTF for monospace bold |
| `--ui-font-file` | — | Custom TTF/OTF for UI text |
| `--ui-bold-font-file` | — | Custom TTF/OTF for UI bold text |

---

## Layout & Pagination

- **Unified view** — file badge, hunk header, then code lines in a single column
  - **Keep-together**: if a whole file fits on the next page, it won't be split at the bottom of the current one
  - **Widow/orphan protection**: hunk headers are never stranded alone at a page bottom
- **Side-by-side view** — old (left) and new (right) with continuous flowing layout

---

## Encoding Tips (Windows)

PowerShell 5.1 uses UTF-16 LE by default for `>` redirects.
Use `Out-File -Encoding utf8` or pipe directly to the script:

```powershell
git diff <from> <to> | Out-File -Encoding utf8 compare.diff
```

The script auto-detects encoding, but UTF-8 is recommended.

**Not supported:** `--word-diff`, `--name-only`, `--name-status`, binary-only changes.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| *"No parsable diffs found."* | Ensure input is a unified diff with `@@` hunks. Try `--debug`. |
| *Font errors / spaces in font name* | Falls back to Courier automatically. Use `--mono-font-file` for a specific TTF. |
| *Dots/ellipsis at the top* | BOMs, zero-widths, NBSP, and bullet artefacts are stripped automatically. |
| *Broken alignment / wrapping* | Use `--landscape` for long lines, or reduce `--font-size`. |

