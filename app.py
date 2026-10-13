"""
JSON Blueprint Document Studio
==============================

A Flask service that compiles a single JSON "blueprint" into a polished,
downloadable document in one of three formats:

    * PDF   -- ReportLab (styled headings, tables, callouts, charts, TOC, cover)
    * DOCX  -- python-docx (real Word headings, TOC field, page-number fields)
    * PPTX  -- python-pptx (16:9 slide deck, native charts and tables)

Routes
------
GET  /         -- the web UI (paste a blueprint, download the file)
POST /generate -- accepts `json_blueprint` as a form field (web UI) OR a JSON
                  request body (API/curl) and returns the generated file.

Blueprint schema (summary -- full reference in README.md)
---------------------------------------------------------
Document-level settings:
    format         "pdf" | "docx" | "pptx"          (default "pdf")
    filename       output file name                 (default "output")
    title          main document title
    subtitle       secondary line under the title
    author         author name (metadata + cover page)
    theme_color    hex color, e.g. "#2563EB"
    orientation    "portrait" | "landscape"         (default "portrait")
    page_size      "letter" | "a4" | "legal" | "a3" | "a5"
    margin         page margin in points            (default 54)
    header         running header text (PDF/DOCX; ignored on PPTX covers)
    footer         running footer text (all formats)
    page_numbers   true -> "Page X of Y" (slide numbers on PPTX)
    title_page     true -> dedicated cover page / title slide
    elements       the ordered list of content elements below

Element types (any order, any count):
    heading         {level 1-3, text}
    paragraph       {text}
    bullet_list     {items: [str | {text, level 0-2}]}
    numbered_list   {items: [str | {text, level 0-2}]}
    callout         {text, variant: info|success|warning|danger}
    quote           {text, attribution?}
    code            {text, language?}
    table           {headers: [str], rows: [[str]], col_widths?: [percent]}
    chart           {chart_type: bar|column|line|pie, title?, categories,
                     series: [{name, values}] (or bare `values`)}
    stats           {items: [{label, value}]}
    image           {url | base64, width?: px, caption?}
    divider         {}
    spacer          {height?: points}
    page_break      {}
    toc             {}  (PDF: linked TOC + bookmarks -- DOCX: TOC field --
                        PPTX: agenda slide)

Inline formatting: every "text" field supports **bold**, *italic* and
`monospace` markers (markdown style). Raw HTML characters are escaped
automatically and are rendered literally.
"""

import base64
import hashlib
import io
import json
import re
import textwrap
import urllib.request
from datetime import date
from xml.sax.saxutils import escape as xml_escape

from flask import Flask, render_template, request, send_file

# ---------------------------------------------------------------- PDF imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, A5, landscape, letter, legal, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.charts.linecharts import LineChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

# --------------------------------------------------------------- DOCX imports
import docx.enum.section
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor

# --------------------------------------------------------------- PPTX imports
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Emu as PEmu
from pptx.util import Inches as PInches
from pptx.util import Pt as PPt

from PIL import Image as PILImage

PILImage.LOAD_TRUNCATED_IMAGES = True  # tolerate slightly truncated image uploads

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB request cap

# =============================================================================
# Constants
# =============================================================================

#: Page sizes in inches: name -> (width, height)
PAGE_IN = {
    "letter": (8.5, 11.0),
    "a4": (8.27, 11.69),
    "legal": (8.5, 14.0),
    "a3": (11.69, 16.54),
    "a5": (5.83, 8.27),
}

PDF_PAGESIZES = {
    "letter": letter,
    "a4": A4,
    "legal": legal,
    "a3": A3,
    "a5": A5,
}

MIMETYPES = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}

#: Callout box variants: background / border / text colors.
CALLOUTS = {
    "info":    {"bg": "#EBF8FF", "border": "#90CDF4", "text": "#2C5282"},
    "success": {"bg": "#F0FFF4", "border": "#68D391", "text": "#22543D"},
    "warning": {"bg": "#FFFBEB", "border": "#F6E05E", "text": "#975A16"},
    "danger":  {"bg": "#FFF5F5", "border": "#FC8181", "text": "#742A2A"},
}
CALLOUT_ALIASES = {
    "error": "danger", "critical": "danger",
    "caution": "warning", "note": "info", "tip": "success", "positive": "success",
}

VALID_ELEMENTS = {
    "heading", "paragraph", "bullet_list", "numbered_list", "callout", "quote",
    "code", "divider", "spacer", "table", "chart", "stats", "image",
    "page_break", "toc",
}

INK_DARK = "#1A202C"
INK_BODY = "#4A5568"
INK_MUTED = "#718096"
LINE_GRAY = "#CBD5E1"
ZEBRA = "#F7FAFC"
CODE_BG = "#1A202C"
CODE_FG = "#E2E8F0"
NBSP = "\u00A0"


class BlueprintError(ValueError):
    """Raised for any invalid blueprint input; surfaced to the client as HTTP 400."""


# =============================================================================
# Small parsing / coercion helpers
# =============================================================================

def to_color(hexstr):
    """'#RRGGBB' string -> reportlab color (falls back to default blue)."""
    if not isinstance(hexstr, str) or not re.fullmatch(r"#?[0-9a-fA-F]{6}", hexstr.strip()):
        hexstr = "#2563EB"
    return colors.HexColor(hexstr if hexstr.startswith("#") else f"#{hexstr}")


def hex6(hexstr):
    """Normalize any color input to 'RRGGBB' (no hash); default blue if invalid."""
    if not isinstance(hexstr, str) or not re.fullmatch(r"#?[0-9a-fA-F]{6}", hexstr.strip()):
        hexstr = "#2563EB"
    return hexstr.lstrip("#").upper()


def get_str(src, key, default="", where="config", required=False):
    v = src.get(key, default)
    if v is None:
        if required:
            raise BlueprintError(f"{where}: '{key}' is required and cannot be null.")
        return default
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(v)
    if not isinstance(v, str):
        raise BlueprintError(f"{where}: '{key}' must be a string.")
    if required and not v.strip():
        raise BlueprintError(f"{where}: '{key}' must not be empty.")
    return v


def get_bool(src, key, default=False):
    v = src.get(key, default)
    return v if isinstance(v, bool) else default


def get_int(src, key, default, lo, hi, where="config"):
    v = src.get(key, default)
    try:
        return max(lo, min(hi, int(v)))
    except (TypeError, ValueError):
        return default


def get_float(src, key, default, lo, hi, where="config"):
    v = src.get(key, default)
    try:
        return max(lo, min(hi, float(v)))
    except (TypeError, ValueError):
        return default


def cell_str(v, where):
    """Coerce a table cell / list item to a plain string."""
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return str(v)
    raise BlueprintError(f"{where}: cell values must be strings or numbers (got {type(v).__name__}).")


# =============================================================================
# Inline **markdown-ish** formatting (shared by all three builders)
# =============================================================================

INLINE_RE = re.compile(r"(\*\*.+?\*\*|\*[^*\n]+?\*|`[^`\n]+?`)")


def parse_inline(text):
    """Split text into [(chunk, bold, italic, mono), ...] honoring **b** *i* `c`."""
    out = []
    for tok in INLINE_RE.split(text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            out.append((tok[2:-2], True, False, False))
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            out.append((tok[1:-1], False, False, True))
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            out.append((tok[1:-1], False, True, False))
        else:
            out.append((tok, False, False, False))
    return out or [("", False, False, False)]


def plain_text(text):
    """Markdown-stripped, unescaped text (for TOC entries / outlines)."""
    return re.sub(r"\*\*|\*|`", "", text or "")


def pdf_rich(text):
    """Escape XML, then re-inject <b>/<i>/<font> markup from markdown markers."""
    parts = []
    for chunk, bold, italic, mono in parse_inline(text):
        chunk = xml_escape(chunk)
        if mono:
            chunk = f'<font face="Courier">{chunk}</font>'
        if bold:
            chunk = f"<b>{chunk}</b>"
        if italic:
            chunk = f"<i>{chunk}</i>"
        parts.append(chunk)
    return "".join(parts)


def docx_add_runs(paragraph, text, size=None, color=None, italic=False, mono_font="Consolas"):
    """Append runs to a python-docx paragraph honoring inline markdown."""
    for chunk, bold, ital, mono in parse_inline(text):
        run = paragraph.add_run(chunk)
        run.bold = run.bold or bold
        run.italic = ital or italic
        if mono:
            run.font.name = mono_font
        if size is not None:
            run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = RGBColor.from_string(color)


def pptx_add_runs(paragraph, text, size=None, color=None, italic=False, bold=False):
    """Append runs to a python-pptx paragraph honoring inline markdown."""
    for chunk, b, ital, mono in parse_inline(text):
        run = paragraph.add_run()
        run.text = chunk
        run.font.bold = bold or b
        run.font.italic = ital or italic
        if mono:
            run.font.name = "Consolas"
        if size is not None:
            run.font.size = PPt(size)
        if color is not None:
            run.font.color.rgb = PRGBColor.from_string(color)


def pptx_no_bullet(paragraph):
    """Suppress the inherited bullet char on a placeholder paragraph."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.append(parse_xml(
        '<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))


def nbsp_indent(line):
    """Replace leading spaces with non-breaking spaces (Word/PP preserve them)."""
    return re.sub(r"^ +", lambda m: NBSP * len(m.group()), line)


# =============================================================================
# Element normalizers (validate + coerce; raise BlueprintError with context)
# =============================================================================

def normalize_items(el, where):
    raw = el.get("items", [])
    if not isinstance(raw, list):
        raise BlueprintError(f"{where}: 'items' must be a list.")
    items = []
    for i, it in enumerate(raw):
        if isinstance(it, str):
            items.append((it, 0))
        elif isinstance(it, dict):
            txt = get_str(it, "text", "", f"{where}.items[{i}]", required=True)
            lvl = get_int(it, "level", 0, 0, 2, f"{where}.items[{i}]")
            items.append((txt, lvl))
        else:
            raise BlueprintError(f"{where}.items[{i}]: must be a string or an object with 'text'.")
    return items


def normalize_table(el, where):
    headers = el.get("headers", [])
    rows = el.get("rows", [])
    if not isinstance(headers, list) or not all(isinstance(h, (str, int, float)) for h in headers):
        raise BlueprintError(f"{where}: 'headers' must be a list of strings.")
    headers = [cell_str(h, where) for h in headers]
    if not isinstance(rows, list):
        raise BlueprintError(f"{where}: 'rows' must be a list of lists.")
    norm_rows = []
    for r_i, row in enumerate(rows):
        if not isinstance(row, list):
            raise BlueprintError(f"{where}.rows[{r_i}]: must be a list.")
        cells = [cell_str(c, f"{where}.rows[{r_i}]") for c in row]
        # Pad / truncate so every row matches the header count.
        if len(cells) < len(headers):
            cells += [""] * (len(headers) - len(cells))
        norm_rows.append(cells[: len(headers)])
    if not headers and norm_rows:
        width = max(len(r) for r in norm_rows)
        headers = [f"Column {i + 1}" for i in range(width)]
        norm_rows = [r + [""] * (width - len(r)) for r in norm_rows]
    if not headers:
        raise BlueprintError(f"{where}: table needs 'headers' or 'rows'.")

    widths = el.get("col_widths")
    fracs = None
    if isinstance(widths, list) and len(widths) == len(headers):
        try:
            vals = [abs(float(w)) for w in widths]
            total = sum(vals) or 1.0
            fracs = [v / total for v in vals]
        except (TypeError, ValueError):
            fracs = None
    return headers, norm_rows, fracs


def normalize_chart(el, where):
    ctype = get_str(el, "chart_type", "column", where).lower()
    if ctype not in ("bar", "column", "line", "pie"):
        raise BlueprintError(f"{where}: 'chart_type' must be one of bar, column, line, pie.")

    categories = el.get("categories", [])
    if not isinstance(categories, list) or not categories:
        raise BlueprintError(f"{where}: 'categories' must be a non-empty list.")
    categories = [cell_str(c, f"{where}.categories") for c in categories]

    series = []
    raw_series = el.get("series")
    if raw_series is None and isinstance(el.get("values"), list):
        raw_series = [{"name": "Series 1", "values": el["values"]}]
    if not isinstance(raw_series, list) or not raw_series:
        raise BlueprintError(f"{where}: 'series' must be a list of {{name, values}}.")
    for s_i, s in enumerate(raw_series):
        s_where = f"{where}.series[{s_i}]"
        name = get_str(s, "name", f"Series {s_i + 1}", s_where)
        vals = s.get("values", [])
        if not isinstance(vals, list) or not vals:
            raise BlueprintError(f"{s_where}: 'values' must be a non-empty list of numbers.")
        try:
            nums = [float(v) for v in vals]
        except (TypeError, ValueError):
            raise BlueprintError(f"{s_where}: 'values' must contain only numbers.")
        # Pad / truncate to the category count so every renderer stays happy.
        if len(nums) < len(categories):
            nums += [0.0] * (len(categories) - len(nums))
        series.append((name, nums[: len(categories)]))
    return ctype, get_str(el, "title", "", where), categories, series


def normalize_stats(el, where):
    raw = el.get("items", [])
    if not isinstance(raw, list) or not raw:
        raise BlueprintError(f"{where}: 'items' must be a non-empty list of {{label, value}}.")
    items = []
    for i, it in enumerate(raw):
        if not isinstance(it, dict):
            raise BlueprintError(f"{where}.items[{i}]: must be an object with 'label' and 'value'.")
        label = get_str(it, "label", "", f"{where}.items[{i}]", required=True)
        value = get_str(it, "value", "", f"{where}.items[{i}]", required=True)
        items.append((label, value))
    return items[:8]


def load_image(el, where):
    """Fetch an image from `url` or `base64` and return (PNG BytesIO, (w, h))."""
    url = el.get("url") or el.get("src")
    b64 = el.get("base64") or el.get("data")
    if not url and not b64:
        raise BlueprintError(f"{where}: image needs a 'url' or 'base64' source.")
    try:
        if url:
            if not isinstance(url, str) or not url.startswith(("http://", "https://")):
                raise BlueprintError(f"{where}: 'url' must be an http(s) URL.")
            req = urllib.request.Request(url, headers={"User-Agent": "JSON-Document-Studio/2.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read(25 * 1024 * 1024)
        else:
            if isinstance(b64, str) and b64.startswith("data:"):
                b64 = b64.split(",", 1)[-1]
            data = base64.b64decode(b64, validate=False)
        img = PILImage.open(io.BytesIO(data))
        img.load()
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA" if img.mode in ("P", "LA", "PA") else "RGB")
        out = io.BytesIO()
        img.save(out, "PNG")
        out.seek(0)
        return out, img.size
    except BlueprintError:
        raise
    except Exception as exc:
        raise BlueprintError(f"{where}: could not load image ({exc}).")


# =============================================================================
# Shared config validation
# =============================================================================

def validate_config(config):
    """Validate document-level settings; returns a normalized dict."""
    if not isinstance(config, dict):
        raise BlueprintError("The blueprint must be a JSON object.")

    fmt = get_str(config, "format", "pdf").lower()
    if fmt not in ("pdf", "docx", "pptx"):
        raise BlueprintError(f"'format' must be pdf, docx or pptx (got '{fmt}').")

    filename = re.sub(r'[\\/:*?"<>|]+', "_", get_str(config, "filename", "output").strip())
    filename = filename.strip(". ") or "output"
    if not filename.lower().endswith(f".{fmt}"):
        filename = f"{filename}.{fmt}"

    orientation = get_str(config, "orientation", "portrait").lower()
    if orientation not in ("portrait", "landscape"):
        orientation = "portrait"

    page_size = get_str(config, "page_size", "letter").lower()
    if page_size not in PAGE_IN:
        page_size = "letter"

    elements = config.get("elements", [])
    if not isinstance(elements, list):
        raise BlueprintError("'elements' must be a list.")
    for i, el in enumerate(elements):
        if not isinstance(el, dict):
            raise BlueprintError(f"elements[{i}]: must be an object.")
        etype = el.get("type")
        if etype not in VALID_ELEMENTS:
            known = ", ".join(sorted(VALID_ELEMENTS))
            raise BlueprintError(f"elements[{i}]: unknown type '{etype}'. Valid types: {known}.")

    return {
        "format": fmt,
        "filename": filename,
        "title": get_str(config, "title", "Document Title"),
        "subtitle": get_str(config, "subtitle", ""),
        "author": get_str(config, "author", ""),
        "theme": hex6(config.get("theme_color") or config.get("accent_color")),
        "orientation": orientation,
        "page_size": page_size,
        "margin": get_int(config, "margin", 54, 24, 144),
        "header": get_str(config, "header", ""),
        "footer": get_str(config, "footer", ""),
        "page_numbers": get_bool(config, "page_numbers", False),
        "title_page": get_bool(config, "title_page", False),
        "elements": elements,
    }


# =============================================================================
# 1. PDF BUILDER (ReportLab)
# =============================================================================

class DecoratedCanvas(pdfcanvas.Canvas):
    """Two-pass canvas that stamps header / footer / 'Page X of Y' on every page."""

    def __init__(self, *args, decorations=None, **kwargs):
        super().__init__(*args, **kwargs)
        self._decorations = decorations or {}
        self._saved_pages = []

    def showPage(self):
        self._saved_pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_pages)
        for state in self._saved_pages:
            self.__dict__.update(state)
            self._draw_decorations(total)
            pdfcanvas.Canvas.showPage(self)
        pdfcanvas.Canvas.save(self)

    def _draw_decorations(self, total):
        deco, m = self._decorations, self._decorations.get("margin", 54)
        if not (deco.get("header") or deco.get("footer") or deco.get("page_numbers")):
            return
        w, h = self._pagesize
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#94A3B8"))
        if deco.get("header"):
            self.drawRightString(w - m, h - m + 14, deco["header"])
            self.setStrokeColor(colors.HexColor("#E2E8F0"))
            self.setLineWidth(0.5)
            self.line(m, h - m + 8, w - m, h - m + 8)
        if deco.get("footer"):
            self.drawString(m, m - 24, deco["footer"])
        if deco.get("page_numbers"):
            self.drawRightString(w - m, m - 24, f"Page {self._pageNumber} of {total}")
        self.restoreState()


class StudioDocTemplate(SimpleDocTemplate):
    """DocTemplate that feeds headings to the TOC and adds PDF outline bookmarks."""

    def beforeDocument(self):
        # Reset per build pass so TOC keys are identical on every multiBuild
        # pass (otherwise entries never stabilize and multiBuild exhausts passes).
        self._toc_seq = 0
        self._outline_last = 0
        self._toc_seen = {}

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and hasattr(flowable, "_toc_entry"):
            level, text = flowable._toc_entry  # level is already 0-based (0..2)
            tag = hashlib.md5(f"{level}|{text}".encode("utf-8")).hexdigest()[:10]
            seen = self._toc_seen.get(tag, 0)
            self._toc_seen[tag] = seen + 1
            key = f"toc-{tag}-{seen}"
            self.canv.bookmarkPage(key)
            try:
                # Outline levels may only step down one at a time.
                last = getattr(self, "_outline_last", 0)
                outline_lvl = min(level, last + 1)
                self._outline_last = outline_lvl
                self.canv.addOutlineEntry(text, key, level=outline_lvl, closed=False)
            except Exception:
                pass
            self.notify("TOCEntry", (level, text, self.page, key))


def build_pdf(cfg):
    """Compile the blueprint into a styled PDF; returns a BytesIO ready to send."""
    buffer = io.BytesIO()
    m = cfg["margin"]
    page = landscape(PDF_PAGESIZES[cfg["page_size"]]) if cfg["orientation"] == "landscape" \
        else portrait(PDF_PAGESIZES[cfg["page_size"]])
    doc = StudioDocTemplate(
        buffer, pagesize=page,
        leftMargin=m, rightMargin=m, topMargin=m + 16, bottomMargin=m,
        title=cfg["title"], author=cfg["author"] or "JSON Document Studio",
    )
    theme = to_color(cfg["theme"])
    palette = [cfg["theme"], "0EA5E9", "F59E0B", "10B981", "8B5CF6", "EF4444", "94A3B8"]
    today = date.today().strftime("%B %d, %Y")

    styles = getSampleStyleSheet()
    S = dict(
        title=ParagraphStyle("T", parent=styles["Heading1"], fontName="Helvetica-Bold",
                             fontSize=20, leading=24, textColor=colors.HexColor(INK_DARK), spaceAfter=4),
        subtitle=ParagraphStyle("St", parent=styles["Normal"], fontName="Helvetica",
                                fontSize=11, leading=15, textColor=colors.HexColor(INK_MUTED), spaceAfter=16),
        h1=ParagraphStyle("H1", parent=styles["Heading2"], fontName="Helvetica-Bold",
                          fontSize=15, leading=19, textColor=theme, spaceBefore=14, spaceAfter=6),
        h2=ParagraphStyle("H2", parent=styles["Heading3"], fontName="Helvetica-Bold",
                          fontSize=13, leading=17, textColor=colors.HexColor("#2D3748"), spaceBefore=10, spaceAfter=4),
        h3=ParagraphStyle("H3", parent=styles["Heading3"], fontName="Helvetica-Bold",
                          fontSize=11, leading=15, textColor=colors.HexColor(INK_MUTED), spaceBefore=8, spaceAfter=3),
        body=ParagraphStyle("B", parent=styles["Normal"], fontName="Helvetica",
                            fontSize=10, leading=15, textColor=colors.HexColor(INK_BODY), spaceAfter=8),
        th=ParagraphStyle("Th", parent=styles["Normal"], fontName="Helvetica",
                          fontSize=10, leading=13, textColor=colors.white),
        bullet=ParagraphStyle("Bu", fontName="Helvetica", fontSize=10, leading=14,
                              textColor=colors.HexColor(INK_BODY), spaceAfter=4, leftIndent=15, bulletIndent=5),
        caption=ParagraphStyle("Cap", fontName="Helvetica-Oblique", fontSize=8.5, leading=11,
                               textColor=colors.HexColor(INK_MUTED), spaceBefore=4),
        quote=ParagraphStyle("Q", fontName="Helvetica-Oblique", fontSize=11, leading=16,
                             textColor=colors.HexColor("#2D3748")),
        quote_attr=ParagraphStyle("Qa", fontName="Helvetica", fontSize=9, leading=12,
                                  textColor=colors.HexColor(INK_MUTED), alignment=2, spaceBefore=3),
        code=ParagraphStyle("Cd", fontName="Courier", fontSize=8, leading=10.5,
                            textColor=colors.HexColor(CODE_FG)),
        callout=ParagraphStyle("Co", fontName="Helvetica-Oblique", fontSize=10, leading=14),
        chart_title=ParagraphStyle("Ct", fontName="Helvetica-Bold", fontSize=11, leading=14,
                                   textColor=colors.HexColor("#2D3748"), spaceBefore=8, spaceAfter=6),
        stat_v=ParagraphStyle("Sv", fontName="Helvetica-Bold", fontSize=15, leading=18,
                              textColor=theme, alignment=1),
        stat_l=ParagraphStyle("Sl", fontName="Helvetica", fontSize=8.5, leading=11,
                              textColor=colors.HexColor(INK_MUTED), alignment=1),
    )

    def heading_flowable(el, where):
        text = get_str(el, "text", "", where, required=True)
        level = get_int(el, "level", 1, 1, 3, where)
        para = Paragraph(pdf_rich(text), S["h1"] if level == 1 else S["h2"] if level == 2 else S["h3"])
        para._toc_entry = (level - 1, plain_text(text))  # picked up by afterFlowable
        return para

    def callout_flowable(el, where):
        variant = CALLOUTS.get(CALLOUT_ALIASES.get(get_str(el, "variant", "info").lower(),
                                                   get_str(el, "variant", "info").lower()),
                               CALLOUTS["info"])
        text = get_str(el, "text", "", where, required=True)
        style = ParagraphStyle("CoV", parent=S["callout"], textColor=colors.HexColor(variant["text"]))
        box = Table([[Paragraph(pdf_rich(text), style)]], colWidths=[doc.width])
        box.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(variant["bg"])),
            ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor(variant["border"])),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        return [Spacer(1, 4), box, Spacer(1, 8)]

    def table_flowable(el, where):
        headers, rows, fracs = normalize_table(el, where)
        col_ws = [doc.width * f for f in fracs] if fracs else [doc.width / len(headers)] * len(headers)
        data = [[Paragraph(f"<b>{pdf_rich(h)}</b>", S["th"]) for h in headers]]
        for row in rows:
            data.append([Paragraph(pdf_rich(c), S["body"]) for c in row])
        t = Table(data, colWidths=col_ws, repeatRows=1)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), theme),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor(ZEBRA)),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for r in range(2, len(data), 2):  # subtle zebra alternation
            style.append(("BACKGROUND", (0, r), (-1, r), colors.white))
        t.setStyle(TableStyle(style))
        return [Spacer(1, 6), t, Spacer(1, 10)]

    def chart_flowable(el, where):
        ctype, title, categories, series = normalize_chart(el, where)
        out = []
        if title:
            out.append(Paragraph(pdf_rich(title), S["chart_title"]))
        width, height = doc.width, 230
        drawing = Drawing(width, height)
        if ctype == "pie":
            ch = Pie()
            ch.data = series[0][1]  # pies visualize a single series
            ch.labels = categories
            ch.sideLabels = 1
            for i in range(len(categories)):
                ch.slices[i].fillColor = to_color(palette[i % len(palette)])
            ch.x, ch.y = 15, 15
            ch.width, ch.height = width - 30, height - 30
        else:
            ch = {
                "column": VerticalBarChart,
                "bar": HorizontalBarChart,
                "line": LineChart,
            }[ctype]()
            ch.data = [vals for _, vals in series]
            ch.categoryAxis.categoryNames = categories
            for axis in (ch.categoryAxis, ch.valueAxis):
                axis.labels.fontName = "Helvetica"
                axis.labels.fontSize = 8
                axis.labels.fillColor = colors.HexColor(INK_MUTED)
            if ctype == "line":
                for i, _ in enumerate(series):
                    ch.lines[i].strokeColor = to_color(palette[i % len(palette)])
                    ch.lines[i].strokeWidth = 1.5
            else:
                for i, _ in enumerate(series):
                    ch.bars[i].fillColor = to_color(palette[i % len(palette)])
            ch.width, ch.height = width - 40, height - 25
            ch.x, ch.y = 25, 12
        drawing.add(ch)
        out += [Spacer(1, 6), drawing, Spacer(1, 8)]
        return out

    def image_flowable(el, where):
        img_io, (px_w, px_h) = load_image(el, where)
        max_w = doc.width
        want = get_float(el, "width", 0, 0, 2000, where) * 0.75  # px -> pt
        w = min(want or max_w, max_w)
        h = w * px_h / max(px_w, 1)
        rl_img = RLImage(img_io, width=w, height=h)
        caption = get_str(el, "caption", "")
        out = [Spacer(1, 6), rl_img]
        if caption:
            para = Paragraph(pdf_rich(caption), S["caption"])
            out.append(para)
        out.append(Spacer(1, 8))
        return out

    # ---------------- assemble the story ----------------
    story = []
    if cfg["title_page"]:
        cover = [
            Spacer(1, 2.1 * 72),
            HRFlowable(width="30%", thickness=2, color=theme, spaceAfter=26, hAlign="CENTER"),
            Paragraph(pdf_rich(cfg["title"]), ParagraphStyle(
                "CovT", parent=S["title"], fontSize=26, leading=32, alignment=1)),
        ]
        if cfg["subtitle"]:
            cover.append(Paragraph(pdf_rich(cfg["subtitle"]), ParagraphStyle(
                "CovS", parent=S["subtitle"], fontSize=13, leading=18, alignment=1, spaceAfter=10)))
        cover.append(Spacer(1, 14))
        if cfg["author"]:
            cover.append(Paragraph(pdf_rich(cfg["author"]), ParagraphStyle(
                "CovA", parent=S["body"], alignment=1)))
        cover.append(Paragraph(today, ParagraphStyle("CovD", parent=S["body"], alignment=1)))
        cover.append(PageBreak())
        story += cover
    else:
        story.append(Paragraph(pdf_rich(cfg["title"]), S["title"]))
        if cfg["subtitle"]:
            story.append(Paragraph(pdf_rich(cfg["subtitle"]), S["subtitle"]))
        story.append(Spacer(1, 6))

    for i, el in enumerate(cfg["elements"]):
        where = f"elements[{i}] ({el.get('type')})"
        try:
            etype = el.get("type")
            if etype == "heading":
                story.append(heading_flowable(el, where))
            elif etype == "paragraph":
                story.append(Paragraph(pdf_rich(get_str(el, "text", "", where, required=True)), S["body"]))
            elif etype == "bullet_list":
                for text, lvl in normalize_items(el, where):
                    st = ParagraphStyle(f"Bu{lvl}", parent=S["bullet"], leftIndent=15 + 15 * lvl)
                    story.append(Paragraph(f"• {pdf_rich(text)}", st))
                story.append(Spacer(1, 4))
            elif etype == "numbered_list":
                for n, (text, lvl) in enumerate(normalize_items(el, where), 1):
                    st = ParagraphStyle(f"N{lvl}", parent=S["bullet"], leftIndent=22 + 15 * lvl)
                    story.append(Paragraph(f"{n}. {pdf_rich(text)}", st))
                story.append(Spacer(1, 4))
            elif etype == "callout":
                story += callout_flowable(el, where)
            elif etype == "quote":
                text = get_str(el, "text", "", where, required=True)
                attr = get_str(el, "attribution", "")
                bar = Table([["", Paragraph(pdf_rich(text), S["quote"])]],
                            colWidths=[3, doc.width - 3])
                bar.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (0, 0), theme),
                    ("LEFTPADDING", (1, 0), (1, 0), 12),
                    ("RIGHTPADDING", (1, 0), (1, 0), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]))
                story += [Spacer(1, 6), bar]
                if attr:
                    story.append(Paragraph(f"— {pdf_rich(attr)}", S["quote_attr"]))
                story.append(Spacer(1, 10))
            elif etype == "code":
                text = get_str(el, "text", "", where, required=True)
                lines = []
                for line in text.splitlines():
                    lines += textwrap.wrap(line, width=96, subsequent_indent="  ",
                                           replace_whitespace=False, drop_whitespace=False) or [""]
                box = Table([[Preformatted("\n".join(lines), S["code"])]], colWidths=[doc.width])
                box.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(CODE_BG)),
                    ("LEFTPADDING", (0, 0), (-1, -1), 10),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]))
                story += [Spacer(1, 4), box, Spacer(1, 8)]
            elif etype == "divider":
                story += [Spacer(1, 6), HRFlowable(width="100%", thickness=1,
                                                  color=colors.HexColor(LINE_GRAY), spaceAfter=10)]
            elif etype == "spacer":
                story.append(Spacer(1, get_float(el, "height", 12, 0, 300, where)))
            elif etype == "table":
                story += table_flowable(el, where)
            elif etype == "chart":
                story += chart_flowable(el, where)
            elif etype == "stats":
                items = normalize_stats(el, where)
                col_w = doc.width / len(items)
                data = [[Paragraph(pdf_rich(v), S["stat_v"]) for v, _ in items],
                        [Paragraph(pdf_rich(l), S["stat_l"]) for _, l in items]]
                t = Table(data, colWidths=[col_w] * len(items))
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                    ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
                    ("TOPPADDING", (0, 0), (-1, 0), 12),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
                    ("TOPPADDING", (0, 1), (-1, 1), 0),
                    ("BOTTOMPADDING", (0, 1), (-1, 1), 12),
                    ("LINEBEFORE", (1, 0), (-1, -1), 0.5, colors.HexColor("#EDF2F7")),
                ]))
                story += [Spacer(1, 6), t, Spacer(1, 10)]
            elif etype == "image":
                story += image_flowable(el, where)
            elif etype == "page_break":
                story.append(PageBreak())
            elif etype == "toc":
                toc = TableOfContents()
                toc.levelStyles = [
                    ParagraphStyle("TOC1", fontName="Helvetica-Bold", fontSize=11, leading=16,
                                   leftIndent=16, firstLineIndent=-16, spaceBefore=6,
                                   textColor=colors.HexColor(INK_DARK)),
                    ParagraphStyle("TOC2", fontName="Helvetica", fontSize=10, leading=14,
                                   leftIndent=32, firstLineIndent=-16,
                                   textColor=colors.HexColor(INK_BODY)),
                    ParagraphStyle("TOC3", fontName="Helvetica", fontSize=9, leading=13,
                                   leftIndent=48, firstLineIndent=-16,
                                   textColor=colors.HexColor(INK_MUTED)),
                ]
                toc.dotsMinLevel = 0
                story += [Paragraph("Contents", S["h1"]), toc, PageBreak()]
        except BlueprintError:
            raise
        except Exception as exc:
            raise BlueprintError(f"{where}: {exc}")

    deco = {"header": cfg["header"], "footer": cfg["footer"],
            "page_numbers": cfg["page_numbers"], "margin": m}
    doc.multiBuild(
        story,
        canvasmaker=lambda *a, **kw: DecoratedCanvas(*a, decorations=deco, **kw),
    )
    buffer.seek(0)
    return buffer


# =============================================================================
# 2. WORD BUILDER (python-docx)
# =============================================================================

def _shade(element_pr, fill_hex):
    """Append a <w:shd> shading element to a pPr/tcPr."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill_hex)
    element_pr.append(shd)


def _paragraph_border(paragraph, edge, color_hex, sz=12):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "4")
    el.set(qn("w:color"), color_hex)
    pBdr.append(el)


def _field(paragraph, instr):
    """Append a Word field (e.g. PAGE / NUMPAGES) to a paragraph."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText"); instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.append(begin); run._r.append(instr_el); run._r.append(end)
    return run


def build_docx(cfg):
    """Compile the blueprint into a Word document; returns a BytesIO ready to send."""
    buffer = io.BytesIO()
    doc = Document()
    theme = hex6(cfg["theme"])

    sec = doc.sections[0]
    w_in, h_in = PAGE_IN[cfg["page_size"]]
    if cfg["orientation"] == "landscape":
        w_in, h_in = h_in, w_in
        sec.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(w_in), Inches(h_in)
    m_in = Inches(cfg["margin"] / 72.0)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = m_in
    content_emu = sec.page_width - sec.left_margin - sec.right_margin

    props = doc.core_properties
    props.title, props.author = cfg["title"], cfg["author"] or "JSON Document Studio"

    # ---------------- header / footer ----------------
    if cfg["header"]:
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        docx_add_runs(hp, cfg["header"], size=8, color="94A3B8")
    if cfg["footer"] or cfg["page_numbers"]:
        fp = sec.footer.paragraphs[0]
        fp.paragraph_format.tab_stops.add_tab_stop(content_emu, WD_TAB_ALIGNMENT.RIGHT)
        if cfg["footer"]:
            docx_add_runs(fp, cfg["footer"], size=8, color="94A3B8")
        if cfg["page_numbers"]:
            fp.add_run("\t")
            r = _field(fp, "PAGE"); r.font.size = Pt(8); r.font.color.rgb = RGBColor.from_string("94A3B8")
            fp.add_run(" of ").font.size = Pt(8)
            r2 = _field(fp, "NUMPAGES"); r2.font.size = Pt(8); r2.font.color.rgb = RGBColor.from_string("94A3B8")

    # ---------------- title block / cover ----------------
    today = date.today().strftime("%B %d, %Y")
    if cfg["title_page"]:
        for text, size, color, bold, before in [
            (cfg["title"], 28, INK_DARK.replace("#", ""), True, 120),
            (cfg["subtitle"], 13, INK_MUTED.replace("#", ""), False, 6),
            (cfg["author"], 11, INK_BODY.replace("#", ""), False, 18),
            (today, 10, INK_MUTED.replace("#", ""), False, 2),
        ]:
            if not text:
                continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(before)
            docx_add_runs(p, text)
            for r in p.runs:
                r.font.size = Pt(size)
                r.font.color.rgb = RGBColor.from_string(color)
                r.font.bold = bold
        doc.add_page_break()
    else:
        tp = doc.add_heading("", level=0)
        docx_add_runs(tp, cfg["title"])
        for r in tp.runs:
            r.font.color.rgb = RGBColor.from_string(INK_DARK.replace("#", ""))
        if cfg["subtitle"]:
            sp = doc.add_paragraph()
            docx_add_runs(sp, cfg["subtitle"], size=10.5, color=INK_MUTED.replace("#", ""), italic=True)

    def add_heading(el, where):
        text = get_str(el, "text", "", where, required=True)
        level = get_int(el, "level", 1, 1, 3, where)
        p = doc.add_heading("", level=level + 1)
        docx_add_runs(p, text)
        for r in p.runs:
            r.font.color.rgb = RGBColor.from_string(theme if level == 1 else "2D3748")

    def add_table(el, where, header_fill=None, zebra=True):
        headers, rows, fracs = normalize_table(el, where)
        n = len(headers)
        table = doc.add_table(rows=1 + len(rows), cols=n)
        table.style = "Table Grid"
        table.autofit = False
        col_emus = [int(content_emu * f) for f in fracs] if fracs else [int(content_emu / n)] * n
        fill = header_fill or theme
        for c, h in enumerate(headers):
            cell = table.rows[0].cells[c]
            cell.width = Emu(col_emus[c])
            docx_add_runs(cell.paragraphs[0], h)
            for r in cell.paragraphs[0].runs:
                r.font.bold = True
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor.from_string("FFFFFF")
            _shade(cell._tc.get_or_add_tcPr(), fill)
        for r_i, row in enumerate(rows):
            for c_i, val in enumerate(row):
                cell = table.rows[r_i + 1].cells[c_i]
                cell.width = Emu(col_emus[c_i])
                docx_add_runs(cell.paragraphs[0], val, size=10, color=INK_BODY.replace("#", ""))
                if zebra and r_i % 2 == 1:
                    _shade(cell._tc.get_or_add_tcPr(), ZEBRA.replace("#", ""))
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    for i, el in enumerate(cfg["elements"]):
        where = f"elements[{i}] ({el.get('type')})"
        try:
            etype = el.get("type")
            if etype == "heading":
                add_heading(el, where)
            elif etype == "paragraph":
                p = doc.add_paragraph()
                docx_add_runs(p, get_str(el, "text", "", where, required=True),
                              color=INK_BODY.replace("#", ""))
            elif etype == "bullet_list":
                for text, lvl in normalize_items(el, where):
                    p = doc.add_paragraph(style=f"List Bullet {lvl + 1}" if lvl else "List Bullet")
                    docx_add_runs(p, text, color=INK_BODY.replace("#", ""))
            elif etype == "numbered_list":
                for n, (text, lvl) in enumerate(normalize_items(el, where), 1):
                    p = doc.add_paragraph()
                    pf = p.paragraph_format
                    pf.left_indent = Inches(0.3 + 0.25 * lvl)
                    pf.first_line_indent = Inches(-0.22)
                    docx_add_runs(p, f"{n}.  {text}", color=INK_BODY.replace("#", ""))
            elif etype == "callout":
                variant = get_str(el, "variant", "info").lower()
                variant = CALLOUTS.get(CALLOUT_ALIASES.get(variant, variant), CALLOUTS["info"])
                text = get_str(el, "text", "", where, required=True)
                box = doc.add_table(rows=1, cols=1)
                cell = box.rows[0].cells[0]
                cell.width = Emu(int(content_emu))
                _shade(cell._tc.get_or_add_tcPr(), variant["bg"].replace("#", ""))
                docx_add_runs(cell.paragraphs[0], text, size=10,
                              color=variant["text"].replace("#", ""), italic=True)
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
            elif etype == "quote":
                text = get_str(el, "text", "", where, required=True)
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                docx_add_runs(p, f"\u201C{text}\u201D", size=11.5, italic=True, color="2D3748")
                _paragraph_border(p, "left", theme, sz=18)
                attr = get_str(el, "attribution", "")
                if attr:
                    ap = doc.add_paragraph()
                    ap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    ap.paragraph_format.right_indent = Inches(0.3)
                    docx_add_runs(ap, f"— {attr}", size=9, color=INK_MUTED.replace("#", ""))
            elif etype == "code":
                text = get_str(el, "text", "", where, required=True)
                box = doc.add_table(rows=1, cols=1)
                cell = box.rows[0].cells[0]
                cell.width = Emu(int(content_emu))
                _shade(cell._tc.get_or_add_tcPr(), CODE_BG.replace("#", ""))
                first = True
                for line in text.splitlines():
                    p = cell.paragraphs[0] if first else cell.add_paragraph()
                    first = False
                    p.paragraph_format.space_after = Pt(0)
                    run = p.add_run(nbsp_indent(line) or " ")
                    run.font.name = "Consolas"
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor.from_string(CODE_FG.replace("#", ""))
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
            elif etype == "divider":
                p = doc.add_paragraph()
                _paragraph_border(p, "bottom", LINE_GRAY.replace("#", ""), sz=8)
            elif etype == "spacer":
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(get_float(el, "height", 12, 0, 300, where))
            elif etype == "table":
                add_table(el, where)
            elif etype == "chart":
                # Word has no chart API in python-docx: render the data as a styled table.
                ctype, title, categories, series = normalize_chart(el, where)
                if title:
                    add_heading({"text": title, "level": 2}, where)
                add_table({"headers": ["Category"] + [n for n, _ in series],
                           "rows": [[cat] + [f"{vals[j]:g}" for _, vals in series]
                                    for j, cat in enumerate(categories)]}, where)
                note = doc.add_paragraph()
                docx_add_runs(note, f"Chart ({ctype}) rendered as a data table.",
                              size=8.5, italic=True, color=INK_MUTED.replace("#", ""))
            elif etype == "stats":
                items = normalize_stats(el, where)
                table = doc.add_table(rows=1, cols=len(items))
                for c, (label, value) in enumerate(items):
                    cell = table.rows[0].cells[c]
                    cell.width = Emu(int(content_emu / len(items)))
                    vp = cell.paragraphs[0]
                    vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    docx_add_runs(vp, value, size=16, color=theme)
                    for r in vp.runs:
                        r.font.bold = True
                    lp = cell.add_paragraph()
                    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    docx_add_runs(lp, label, size=8.5, color=INK_MUTED.replace("#", ""))
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
            elif etype == "image":
                img_io, (px_w, px_h) = load_image(el, where)
                want_in = get_float(el, "width", 0, 0, 2000, where) / 96.0
                max_in = content_emu / 914400.0
                w_in = min(want_in or max_in, max_in)
                doc.add_picture(img_io, width=Inches(w_in))
                caption = get_str(el, "caption", "")
                if caption:
                    cp = doc.add_paragraph()
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    docx_add_runs(cp, caption, size=8.5, italic=True, color=INK_MUTED.replace("#", ""))
            elif etype == "page_break":
                doc.add_page_break()
            elif etype == "toc":
                doc.add_heading("Contents", level=1)
                p = doc.add_paragraph()
                fld = OxmlElement("w:fldSimple")
                fld.set(qn("w:instr"), r'TOC \o "1-4" \h \z \u')
                r_el = OxmlElement("w:r")
                t_el = OxmlElement("w:t")
                t_el.text = 'Right-click here and choose "Update Field" to build the Table of Contents.'
                r_el.append(t_el)
                fld.append(r_el)
                p._p.append(fld)
                doc.add_page_break()
        except BlueprintError:
            raise
        except Exception as exc:
            raise BlueprintError(f"{where}: {exc}")

    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =============================================================================
# 3. POWERPOINT BUILDER (python-pptx)
# =============================================================================

def build_pptx(cfg):
    """Compile the blueprint into a slide deck; returns a BytesIO ready to send.

    Deck model: the blueprint title becomes a cover slide; every level-1
    heading opens a new content slide (level 2/3 headings become bold body
    lines); table / chart / image / stats elements each get a dedicated slide.
    """
    buffer = io.BytesIO()
    prs = Presentation()
    theme = hex6(cfg["theme"])
    muted = INK_MUTED.replace("#", "")
    body_col = INK_BODY.replace("#", "")

    if cfg["orientation"] == "portrait":
        prs.slide_width, prs.slide_height = PInches(7.5), PInches(13.333)
    else:
        prs.slide_width, prs.slide_height = PInches(13.333), PInches(7.5)
    W, H = prs.slide_width, prs.slide_height

    props = prs.core_properties
    props.title, props.author = cfg["title"], cfg["author"] or "JSON Document Studio"

    def new_slide(title_text="", themed=True):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        if title_text:
            slide.shapes.title.text = plain_text(title_text)
            for p in slide.shapes.title.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = PRGBColor.from_string(theme if themed else "1A202C")
        return slide

    # ---------------- cover slide ----------------
    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = cfg["title"]
    for p in cover.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            r.font.color.rgb = PRGBColor.from_string(theme)
    sub_tf = cover.placeholders[1].text_frame
    sub_tf.text = cfg["subtitle"] or ""
    meta = " — ".join(x for x in [cfg["author"], date.today().strftime("%B %d, %Y")] if x)
    if meta:
        p = sub_tf.add_paragraph()
        pptx_add_runs(p, meta, size=12, color=muted)

    current = None  # the slide body elements are currently flowing into

    def body_slide():
        """Guarantee a content slide exists (never write onto the cover)."""
        nonlocal current
        if current is None:
            current = new_slide()
        return current

    def body_tf(fresh=False):
        slide = body_slide()
        tf = slide.placeholders[1].text_frame
        tf.word_wrap = True
        if fresh or not tf.text.strip():
            return tf.paragraphs[0], tf
        return tf.add_paragraph(), tf

    def visual_slide(title_text):
        """Dedicated slide for tables / charts / images / stats."""
        nonlocal current
        current = new_slide(title_text)
        return current

    # ---------------- first pass: headings for the agenda (toc) ----------------
    agenda = [plain_text(get_str(el, "text", ""))
              for el in cfg["elements"]
              if el.get("type") == "heading" and get_int(el, "level", 1, 1, 3, "") == 1]

    for i, el in enumerate(cfg["elements"]):
        where = f"elements[{i}] ({el.get('type')})"
        try:
            etype = el.get("type")
            if etype == "heading":
                text = get_str(el, "text", "", where, required=True)
                level = get_int(el, "level", 1, 1, 3, where)
                if level == 1:
                    current = new_slide(text)
                else:
                    p, _ = body_tf(fresh=False)
                    pptx_no_bullet(p)
                    pptx_add_runs(p, text, size=15 if level == 2 else 13, bold=True,
                                  color=theme if level == 2 else "2D3748")
                    p.space_before = PPt(10)
            elif etype == "paragraph":
                p, _ = body_tf()
                pptx_add_runs(p, get_str(el, "text", "", where, required=True), size=14, color=body_col)
                p.space_after = PPt(8)
            elif etype == "bullet_list":
                for text, lvl in normalize_items(el, where):
                    p, _ = body_tf(fresh=False)
                    pptx_add_runs(p, text, size=14, color=body_col)
                    p.level = lvl
                    p.space_after = PPt(4)
            elif etype == "numbered_list":
                for n, (text, lvl) in enumerate(normalize_items(el, where), 1):
                    p, _ = body_tf(fresh=False)
                    pptx_no_bullet(p)
                    pptx_add_runs(p, f"{n}.  {text}", size=14, color=body_col)
                    p.space_after = PPt(4)
            elif etype == "callout":
                variant = get_str(el, "variant", "info").lower()
                variant = CALLOUTS.get(CALLOUT_ALIASES.get(variant, variant), CALLOUTS["info"])
                p, _ = body_tf()
                pptx_no_bullet(p)
                pptx_add_runs(p, f"{variant_key_label(el)}: ", size=13, bold=True,
                              color=variant["text"].replace("#", ""))
                pptx_add_runs(p, get_str(el, "text", "", where, required=True), size=13,
                              italic=True, color=variant["text"].replace("#", ""))
                p.space_before, p.space_after = PPt(8), PPt(8)
            elif etype == "quote":
                p, _ = body_tf()
                pptx_no_bullet(p)
                pptx_add_runs(p, f"\u201C{get_str(el, 'text', '', where, required=True)}\u201D",
                              size=16, italic=True, color="2D3748")
                attr = get_str(el, "attribution", "")
                if attr:
                    ap, _ = body_tf(fresh=False)
                    pptx_no_bullet(ap)
                    pptx_add_runs(ap, f"— {attr}", size=11, color=muted)
            elif etype == "code":
                text = get_str(el, "text", "", where, required=True)
                for line in text.splitlines()[:14]:
                    p, _ = body_tf(fresh=False)
                    pptx_no_bullet(p)
                    pptx_add_runs(p, nbsp_indent(line) or " ", size=11, color="2D3748")
                    p.space_after = PPt(0)
            elif etype == "divider":
                slide = body_slide()
                y = PInches(1.55)
                line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, PInches(0.6), y,
                                                  W - PInches(0.6), y)
                line.line.color.rgb = PRGBColor.from_string(LINE_GRAY.replace("#", ""))
                line.line.width = PPt(1.5)
            elif etype == "spacer":
                p, _ = body_tf(fresh=False)
                p.space_after = PPt(get_float(el, "height", 12, 0, 300, where) * 0.75)
            elif etype == "table":
                headers, rows, fracs = normalize_table(el, where)
                slide = visual_slide(get_str(el, "title", "") or "Table")
                left, top = PInches(0.7), PInches(1.9)
                tbl_w = W - PInches(1.4)
                shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top,
                                               tbl_w, PInches(0.45 * (len(rows) + 1)))
                table = shape.table
                if fracs:
                    for c, f in enumerate(fracs):
                        table.columns[c].width = PEmu(int(tbl_w * f))
                for c, h in enumerate(headers):
                    cell = table.cell(0, c)
                    cell.text = plain_text(h)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = PRGBColor.from_string(theme)
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.bold = True
                            r.font.size = PPt(13)
                            r.font.color.rgb = PRGBColor.from_string("FFFFFF")
                for r_i, row in enumerate(rows):
                    for c_i, val in enumerate(row):
                        cell = table.cell(r_i + 1, c_i)
                        cell.text = plain_text(val)
                        if r_i % 2 == 1:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = PRGBColor.from_string(ZEBRA.replace("#", ""))
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                r.font.size = PPt(12)
                                r.font.color.rgb = PRGBColor.from_string(body_col)
            elif etype == "chart":
                ctype, title, categories, series = normalize_chart(el, where)
                slide = visual_slide(title or "Chart")
                chart_type = {
                    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                    "line": XL_CHART_TYPE.LINE_MARKERS,
                    "pie": XL_CHART_TYPE.PIE,
                }[ctype]
                cd = CategoryChartData()
                cd.categories = categories
                for name, vals in series:
                    cd.add_series(name, vals)
                gframe = slide.shapes.add_chart(chart_type, PInches(0.9), PInches(1.9),
                                                W - PInches(1.8), H - PInches(3.1), cd)
                chart = gframe.chart
                chart.has_title = False
                if len(series) > 1 or ctype == "pie":
                    chart.has_legend = True
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                    chart.legend.include_in_layout = False
                else:
                    chart.has_legend = False
            elif etype == "stats":
                items = normalize_stats(el, where)
                slide = visual_slide(get_str(el, "title", "") or "Metrics")
                margin = PInches(0.9)
                usable = W - 2 * margin
                col_w = PEmu(int(usable / len(items)))
                for c, (label, value) in enumerate(items):
                    box = slide.shapes.add_textbox(
                        PEmu(int(margin + c * col_w)), PInches(2.4), col_w, PInches(1.9))
                    tf = box.text_frame
                    tf.word_wrap = True
                    vp = tf.paragraphs[0]
                    vp.alignment = PP_ALIGN.CENTER
                    pptx_add_runs(vp, value, size=34, bold=True, color=theme)
                    lp = tf.add_paragraph()
                    lp.alignment = PP_ALIGN.CENTER
                    pptx_add_runs(lp, label, size=13, color=muted)
            elif etype == "image":
                img_io, (px_w, px_h) = load_image(el, where)
                slide = visual_slide(get_str(el, "caption", "") or "Image")
                want_in = get_float(el, "width", 0, 0, 2000, where) / 96.0
                max_w_in, max_h_in = (W - PInches(1.4)) / 914400, (H - PInches(2.6)) / 914400
                w_in = min(want_in or max_w_in, max_w_in)
                h_in = w_in * px_h / max(px_w, 1)
                if h_in > max_h_in:
                    h_in, w_in = max_h_in, max_h_in * px_w / max(px_h, 1)
                pic = slide.shapes.add_picture(
                    img_io, PEmu(int((W - w_in * 914400) / 2)), PInches(2.0),
                    width=PInches(w_in))
                caption = get_str(el, "caption", "")
                if caption:
                    cp = slide.shapes.add_textbox(PInches(0.9), pic.top + pic.height + PInches(0.15),
                                                  W - PInches(1.8), PInches(0.5))
                    p = cp.text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER
                    pptx_add_runs(p, caption, size=12, italic=True, color=muted)
            elif etype == "page_break":
                current = None  # next content element starts a fresh slide
            elif etype == "toc":
                slide = new_slide("Agenda")
                tf = slide.placeholders[1].text_frame
                tf.word_wrap = True
                first = True
                for entry in agenda:
                    p = tf.paragraphs[0] if first else tf.add_paragraph()
                    first = False
                    pptx_add_runs(p, entry, size=16, color=body_col)
                    p.space_after = PPt(6)
        except BlueprintError:
            raise
        except Exception as exc:
            raise BlueprintError(f"{where}: {exc}")

    # ---------------- footers / slide numbers (skip the cover) ----------------
    for idx, slide in enumerate(prs.slides):
        if idx == 0:
            continue
        if cfg["footer"]:
            box = slide.shapes.add_textbox(PInches(0.5), H - PInches(0.42), PInches(8), PInches(0.35))
            pptx_add_runs(box.text_frame.paragraphs[0], cfg["footer"], size=9, color=muted)
        if cfg["page_numbers"]:
            box = slide.shapes.add_textbox(W - PInches(1.0), H - PInches(0.42),
                                           PInches(0.5), PInches(0.35))
            pptx_add_runs(box.text_frame.paragraphs[0], str(idx), size=9, color=muted)

    prs.save(buffer)
    buffer.seek(0)
    return buffer


def variant_key_label(el):
    """Human label for a callout variant ('warning' -> 'Warning')."""
    v = get_str(el, "variant", "info").lower()
    v = CALLOUT_ALIASES.get(v, v)
    return v.capitalize()


# =============================================================================
# Routes
# =============================================================================

BUILDERS = {"pdf": build_pdf, "docx": build_docx, "pptx": build_pptx}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    # Accept the blueprint as a form field (web UI) or as a JSON body (API).
    json_raw = request.form.get("json_blueprint")
    if json_raw is not None:
        try:
            config = json.loads(json_raw)
        except json.JSONDecodeError as exc:
            return f"Invalid JSON Blueprint: {exc}", 400
    else:
        config = request.get_json(force=True, silent=True)
        if config is None:
            return "Provide the blueprint as form field 'json_blueprint' or a JSON body.", 400

    try:
        cfg = validate_config(config)
        buffer = BUILDERS[cfg["format"]](cfg)
    except BlueprintError as exc:
        return f"Blueprint error: {exc}", 400
    except Exception as exc:  # keep the feedback actionable for AI-driven retries
        return f"Generation failed: {type(exc).__name__}: {exc}", 400

    return send_file(buffer, as_attachment=True, download_name=cfg["filename"],
                     mimetype=MIMETYPES[cfg["format"]], max_age=0)


if __name__ == "__main__":
    app.run(debug=True)
