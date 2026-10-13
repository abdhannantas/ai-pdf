# JSON Blueprint Document Studio

Turn one JSON "blueprint" into a polished, downloadable document — **PDF** (ReportLab), **Word** (python-docx) or **PowerPoint** (python-pptx). Paste a blueprint in the web UI (or POST it to `/generate`) and get a themed file back with headings, tables, charts, callouts, quotes, code blocks, images, cover pages, tables of contents, headers/footers and page numbers.

The intended workflow: click **Copy AI Instructions** in the UI, paste that prompt into any LLM, describe the document you want, then paste the generated JSON back into the textarea and compile.

---

## Quick start (local)

```bash
pip install -r requirements.txt
python app.py
# open http://127.0.0.1:5000
```

## API

```bash
curl -X POST http://127.0.0.1:5000/generate \
  -H "Content-Type: application/json" \
  -d @examples/report.json \
  --output quarterly-report.pdf
```

`POST /generate` accepts the blueprint either as form field `json_blueprint` (what the web UI sends) or as a raw JSON request body. Errors come back as **HTTP 400** with a plain-text message that names the exact problem (e.g. `elements[3] (table): rows[1] must be a list.`) — feed that message back to the LLM to self-correct.

## Deploy to Vercel

The included `vercel.json` wires `app.py` as a serverless function (legacy `builds`/`routes` format — still fully supported). Push the repo to GitHub and import it at [vercel.com/new](https://vercel.com/new), or run `vercel` from this folder. Dependencies install automatically from `requirements.txt`.

---

## Blueprint reference

### Document-level settings

| Key | Type | Default | Description |
|---|---|---|---|
| `format` | `"pdf"` \| `"docx"` \| `"pptx"` | `"pdf"` | Output format. |
| `filename` | string | `"output"` | Download name; unsafe characters are sanitized, extension fixed per format. |
| `title` | string | `"Document Title"` | Main title (also file metadata). |
| `subtitle` | string | `""` | Line under the title. |
| `author` | string | `""` | Author — document metadata and cover page. |
| `theme_color` | hex string | `#2563EB` | Drives heading, table header, chart and accent colors (also accepts `accent_color`). |
| `orientation` | `"portrait"` \| `"landscape"` | `"portrait"` | Page/slide orientation. |
| `page_size` | `letter` \| `a4` \| `legal` \| `a3` \| `a5` | `letter` | PDF/DOCX page size (PPTX uses 16:9 or portrait slides). |
| `margin` | number | `54` | Page margin in points, clamped 24–144 (PDF/DOCX). |
| `header` | string | `""` | Running header text (PDF top-right, DOCX header). |
| `footer` | string | `""` | Running footer text (all formats). |
| `page_numbers` | boolean | `false` | PDF/DOCX: "Page X of Y". PPTX: slide numbers (cover excluded). |
| `title_page` | boolean | `false` | PDF/DOCX: dedicated centered cover page. PPTX always has a cover slide. |
| `elements` | array | `[]` | Ordered content elements (below). |

### Element types

| `type` | Required fields | Optional fields | Notes |
|---|---|---|---|
| `heading` | `text` | `level` (1–3) | PDF/DOCX: styled section heading + TOC entry. PPTX: level 1 opens a new slide; 2–3 are bold sub-lines. |
| `paragraph` | `text` | — | Body text with inline formatting. |
| `bullet_list` | `items` | per-item `level` (0–2) | Items are strings or `{"text", "level"}` for nesting. |
| `numbered_list` | `items` | per-item `level` | Numbered `1. 2. 3.`; numbering restarts per element. |
| `callout` | `text` | `variant`: `info`/`success`/`warning`/`danger` | Colored highlight box (aliases like `error`, `tip`, `caution` map to these). |
| `quote` | `text` | `attribution` | Indented italic quote with theme-colored bar. |
| `code` | `text` | `language` | Monospace block on a dark background; line breaks and indentation preserved. |
| `table` | `headers`, `rows` | `col_widths` (percents), `title` (PPTX slide title) | Rows are padded/truncated to header count; themed header row + zebra striping. |
| `chart` | `categories`, `series` | `chart_type` (`bar`/`column`/`line`/`pie`), `title` | `series`: `[{"name", "values"}]` or a bare `values` array. Native vector charts in PDF/PPTX; DOCX renders a styled data table (python-docx has no chart API). |
| `stats` | `items: [{label, value}]` | `title` (PPTX) | Big-number metric row, up to 8 per row. |
| `image` | `url` **or** `base64` | `caption`, `width` (px) | Remote http(s) URL or base64 / `data:` URI. Converted to PNG via Pillow, aspect ratio kept, auto-scaled to fit. |
| `divider` | — | — | Horizontal separator line. |
| `spacer` | — | `height` (pt, 0–300) | Vertical gap. |
| `page_break` | — | — | PDF/DOCX: new page. PPTX: next content starts a fresh slide. |
| `toc` | — | — | PDF: clickable TOC + PDF outline bookmarks. DOCX: real TOC field (open in Word → right-click → *Update Field*). PPTX: agenda slide of level-1 headings. |

### Inline formatting

Every `text` field supports `**bold**`, `*italic*` and `` `monospace` ``. Raw HTML characters (`& < >`) are escaped automatically — `"R&D < Growth"` renders literally and never breaks generation.

### Per-format behavior to know

- **PPTX deck model** — the blueprint title becomes the cover slide; each level-1 heading starts a slide; `table`, `chart`, `image` and `stats` each get their own slide; `footer`/`page_numbers` are stamped on every slide except the cover.
- **DOCX TOC** — Word computes the TOC when the field is refreshed (Ctrl+A → F9, or right-click → Update Field).
- **DOCX charts** — rendered as a themed data table with a note, since python-docx cannot embed charts.

### Examples

- [`examples/report.json`](examples/report.json) — PDF quarterly report: cover, TOC, stats, table, column chart, quote, code block.
- [`examples/deck.json`](examples/deck.json) — PPTX product deck: agenda, stats, line chart, pricing table, callout.
- [`examples/memo.json`](examples/memo.json) — DOCX policy memo on A4: TOC, stats, table, warning callout.

The **Load example** button in the web UI fills the textarea with the PDF sample.

---

## Project structure

```
app.py                  Flask app + the three format builders (documented in code)
templates/index.html    Web UI: paste blueprint, copy the AI prompt, load examples
examples/*.json         Ready-to-run sample blueprints
requirements.txt        flask, reportlab, python-docx, python-pptx, Pillow
vercel.json             Serverless deployment config
tests/test_generation.py  End-to-end generation tests (run: python tests/test_generation.py)
```

## Robustness

- Malformed blueprints return **400 with a specific message** (element index + reason) instead of a stack trace — designed for LLM retry loops.
- Text is XML-escaped before PDF rendering; table rows are auto-padded; numbers are coerced; colors/levels/widths are validated and clamped; filenames are sanitized.
- Requests are capped at 16 MB; image downloads time out after 15 s and are converted server-side to PNG.
