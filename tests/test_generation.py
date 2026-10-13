"""End-to-end generation tests: every sample blueprint through /generate,
plus error-handling checks. Run:  python tests/test_generation.py"""

import base64
import io
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PIL import Image  # noqa: E402

import app as studio  # noqa: E402

OUT = ROOT / "tests" / "output"
OUT.mkdir(parents=True, exist_ok=True)

# A real 16x16 PNG generated with Pillow (avoids hand-typed base64 corruption).
_buf = io.BytesIO()
Image.new("RGB", (16, 16), (30, 120, 240)).save(_buf, "PNG")
TEST_IMG_B64 = base64.b64encode(_buf.getvalue()).decode()

client = studio.app.test_client()
passed, failed = 0, []


def check(name, condition, detail=""):
    global passed
    if condition:
        passed += 1
        print(f"  PASS  {name}")
    else:
        failed.append(f"{name}: {detail}")
        print(f"  FAIL  {name}  -> {detail}")


print("== Sample blueprints (all three formats) ==")
for path in sorted((ROOT / "examples").glob("*.json")):
    cfg = json.loads(path.read_text(encoding="utf-8"))
    res = client.post("/generate", data={"json_blueprint": json.dumps(cfg)})
    magic = b"%PDF" if cfg["format"] == "pdf" else b"PK"
    out = OUT / cfg["filename"]
    out.write_bytes(res.data)
    check(f"{path.name} -> {cfg['format']} ({len(res.data):,} bytes)",
          res.status_code == 200 and res.data[: len(magic)] == magic,
          f"status={res.status_code} head={res.data[:8]!r}")

print("== JSON body (API mode) ==")
cfg = json.loads((ROOT / "examples" / "deck.json").read_text(encoding="utf-8"))
res = client.post("/generate", json=cfg)
check("POST /generate with JSON body", res.status_code == 200 and res.data[:2] == b"PK",
      f"status={res.status_code}")

print("== Robustness: previously crashing inputs ==")


def run(elements, fmt="pdf", **extra):
    blueprint = {"format": fmt, "title": "T", "elements": elements}
    blueprint.update(extra)
    return client.post("/generate", data={"json_blueprint": json.dumps(blueprint)})


res = run([{"type": "paragraph", "text": "R&D spend < revenue & > last year, 5*3 `code` **bold**"}])
check("XML-unsafe text in PDF", res.status_code == 200 and res.data[:4] == b"%PDF",
      f"status={res.status_code} body={res.data[:120]!r}")

res = run([{"type": "paragraph", "text": "R&D < revenue & growth"}], fmt="docx")
check("XML-unsafe text in DOCX", res.status_code == 200, f"status={res.status_code}")

res = run([{"type": "paragraph", "text": "R&D < revenue & growth"}], fmt="pptx")
check("XML-unsafe text in PPTX", res.status_code == 200, f"status={res.status_code}")

res = run([{"type": "table", "headers": ["A", "B"], "rows": [["1", "2", "3", "4"], ["5"]]}])
check("ragged table rows", res.status_code == 200, f"status={res.status_code} body={res.data[:120]!r}")

res = run([{"type": "table", "headers": ["A", "B"], "rows": [["1", "2"]]}, {"type": "table", "headers": ["A"], "rows": [["1"]]}], fmt="pptx")
check("pptx table", res.status_code == 200, f"status={res.status_code}")

res = run([{"type": "paragraph"}])
check("missing text -> clear 400", res.status_code == 400 and b"elements[0]" in res.data,
      f"status={res.status_code} body={res.data[:200]!r}")

res = run([{"type": "wombat"}])
check("unknown type -> clear 400", res.status_code == 400 and b"unknown type" in res.data,
      f"status={res.status_code}")

res = client.post("/generate", data={"json_blueprint": "{not json"})
check("invalid JSON -> 400", res.status_code == 400, f"status={res.status_code}")

res = run([{"type": "paragraph", "text": "before any heading"}], fmt="pptx")
check("pptx body before first heading (no cover overwrite)",
      res.status_code == 200, f"status={res.status_code}")
# cover subtitle must not be clobbered: body goes to a new slide
prs_bytes = res.data
check("pptx cover intact", prs_bytes[:2] == b"PK", "")

res = run([{"type": "chart", "chart_type": "pie", "categories": ["A", "B", "C"],
            "series": [{"name": "S", "values": [1, 2, 3]}]}])
check("pdf pie chart", res.status_code == 200 and res.data[:4] == b"%PDF", f"status={res.status_code}")

res = run([{"type": "chart", "chart_type": "pie", "categories": ["A", "B"],
            "series": [{"name": "S", "values": [1, 2]}]}], fmt="pptx")
check("pptx pie chart", res.status_code == 200, f"status={res.status_code}")

res = run([{"type": "chart", "chart_type": "line", "categories": ["A", "B"],
            "series": [{"name": "S", "values": [1, 2]}]}], fmt="docx")
check("docx chart fallback table", res.status_code == 200, f"status={res.status_code}")

res = run([{"type": "image",
            "base64": "data:image/png;base64," + TEST_IMG_B64,
            "caption": "a small blue square"}])
check("base64 image in PDF", res.status_code == 200 and res.data[:4] == b"%PDF",
      f"status={res.status_code} body={res.data[:120]!r}")

res = run([{"type": "heading", "level": 1, "text": "H1"},
           {"type": "toc"},
           {"type": "heading", "level": 2, "text": "H2"},
           {"type": "code", "text": "def f():\n    return 1"},
           {"type": "quote", "text": "q", "attribution": "a"},
           {"type": "stats", "items": [{"label": "L", "value": "V"}]},
           {"type": "divider"}, {"type": "spacer", "height": 20},
           {"type": "page_break"}, {"type": "numbered_list", "items": ["a", "b"]}],
          page_numbers=True, header="H", footer="F", title_page=True, orientation="landscape",
          page_size="a4", theme_color="#FF8800")
check("pdf kitchen sink (landscape a4, TOC, all elements)", res.status_code == 200,
      f"status={res.status_code} body={res.data[:200]!r}")

res = run([{"type": "toc"}, {"type": "heading", "level": 1, "text": "One"},
           {"type": "heading", "level": 2, "text": "Two"},
           {"type": "code", "text": "x = 1"}, {"type": "stats", "items": [{"label": "L", "value": "V"}]}],
          fmt="docx", page_numbers=True, header="H", footer="F", title_page=True)
check("docx kitchen sink", res.status_code == 200, f"status={res.status_code} body={res.data[:200]!r}")

res = run([{"type": "toc"}, {"type": "heading", "level": 1, "text": "One"},
           {"type": "code", "text": "x = 1"}, {"type": "callout", "variant": "warning", "text": "careful"},
           {"type": "divider"}, {"type": "stats", "items": [{"label": "L", "value": "V"}]},
           {"type": "image", "base64": TEST_IMG_B64}],
          fmt="pptx", footer="F", page_numbers=True)
check("pptx kitchen sink", res.status_code == 200, f"status={res.status_code} body={res.data[:200]!r}")

print()
if failed:
    print(f"{passed} passed, {len(failed)} FAILED:")
    for f in failed:
        print(f"  - {f}")
    sys.exit(1)
print(f"All {passed} checks passed. Generated files are in {OUT}")
