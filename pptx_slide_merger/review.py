"""
Reviewer-mode slide review surface.

The pptx-slide-merger has no live viewer of its own — decks are reviewed as a
rendered PDF preview (see docs/visual-edits-workflow.md). This module builds a
*reviewable* surface: it renders each slide of a PPTX to a PNG, extracts per-slide
shape geometry from the OOXML so a click can be resolved to the nearest shape, and
emits a self-contained static HTML viewer with a Reviewer Mode overlay.

The viewer mirrors the STACK deck's finalized `js/reviewer.js` in spirit:

  - a collapsible / opt-in "Review" toggle (button or `R` key) — collapsed by
    default so the viewer is a clean slide browser until you opt in
  - click anywhere on a slide image to drop a comment pinned at that location
  - the click is slide-aware (slide index + title), location-aware (x/y % of the
    slide), and element/region-aware (nearest shape: name + text + bbox)
  - the comment bubble auto-saves 2s after the last keystroke, and on click-outside
    (the outside click is consumed so it never drops a stray second bubble)
  - a capture-phase keyboard shield so typing in a comment (letters like r/x/p)
    never triggers viewer navigation/shortcuts
  - comments persist to localStorage, best-effort PUT to `review-comments.json`
    (when served by the review save-server — no manual Export needed), and Export
    + Copy-for-Claude as fallbacks
  - if `review-comments.json` sits next to the viewer it is auto-loaded on startup

The exported JSON is machine-readable: every comment carries the slide index, the
1-based slide number, the slide title, the click location (xPct/yPct), and the
resolved target shape (name, text, bounding box in slide %), so a follow-up CLI /
agent pass can apply the requested change to the right shape on the right slide.
"""

import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from .merger import NS_P, NS_REL, REL_SLIDE, list_slides

# DrawingML main namespace (shape geometry lives here)
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

EMU_PER_INCH = 914400


# --------------------------------------------------------------------------- #
# Shape geometry extraction
# --------------------------------------------------------------------------- #
def _slide_size_emu(pptx_path: Path):
    """Return (cx, cy) slide size in EMU from presentation.xml."""
    with zipfile.ZipFile(pptx_path, "r") as z:
        root = etree.fromstring(z.read("ppt/presentation.xml"))
    sz = root.find(f"{{{NS_P}}}sldSz")
    if sz is None:
        # PowerPoint 16:9 default
        return 12192000, 6858000
    return int(sz.get("cx", 12192000)), int(sz.get("cy", 6858000))


def _ordered_slide_files(pptx_path: Path):
    """Return slide part names (e.g. 'ppt/slides/slide3.xml') in presentation order."""
    with zipfile.ZipFile(pptx_path, "r") as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))

    rid_to_target = {}
    for rel in rels.findall(f"{{{NS_REL}}}Relationship"):
        if rel.get("Type") == REL_SLIDE:
            rid_to_target[rel.get("Id")] = rel.get("Target")

    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ordered = []
    sld_id_lst = pres.find(f"{{{NS_P}}}sldIdLst")
    if sld_id_lst is None:
        return ordered
    for sld_id in sld_id_lst.findall(f"{{{NS_P}}}sldId"):
        rid = sld_id.get(f"{{{R}}}id")
        target = rid_to_target.get(rid)
        if not target:
            continue
        # Targets are relative to ppt/ (e.g. "slides/slide1.xml")
        ordered.append("ppt/" + target.lstrip("/").replace("../", ""))
    return ordered


def _ph_key(ph_el):
    """
    Stable key for matching a placeholder across slide / layout / master.

    PowerPoint matches by idx when present; title-family placeholders
    (title / ctrTitle) all map to the master's title slot. We normalize to a
    pair (family, idx) where family collapses the title variants.
    """
    if ph_el is None:
        return None
    t = ph_el.get("type", "body")
    idx = ph_el.get("idx")
    family = "title" if t in ("title", "ctrTitle") else t
    return (family, idx)


def _xfrm_bbox(el):
    """Return (x, y, w, h) EMU from an element's a:xfrm, or None."""
    xfrm = el.find(f".//{{{NS_A}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if off is None or ext is None:
        return None
    return (int(off.get("x", 0)), int(off.get("y", 0)),
            int(ext.get("cx", 0)), int(ext.get("cy", 0)))


def _rel_target(rels_part: str, rel_type: str, names: set, zf):
    """Resolve a single relationship target (as a ppt/-relative part name)."""
    if rels_part not in names:
        return None
    rels = etree.fromstring(zf.read(rels_part))
    for rel in rels.findall(f"{{{NS_REL}}}Relationship"):
        if rel.get("Type") == rel_type:
            target = rel.get("Target", "")
            part = "ppt/" + target.lstrip("/").replace("../", "")
            return part if part in names else None
    return None


def _geom_from_part(part: str, names: set, zf):
    """Build {placeholder-key: (x,y,w,h)} for placeholders that carry an xfrm."""
    if part is None or part not in names:
        return {}
    root = etree.fromstring(zf.read(part))
    geom = {}
    for sp in root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        bbox = _xfrm_bbox(sp)
        if ph is not None and bbox is not None:
            geom[_ph_key(ph)] = bbox
    return geom


def _layout_placeholder_geometry(slide_part: str, names: set, zf):
    """
    Build {placeholder-key: (x,y,w,h)} by walking the inheritance chain a slide's
    placeholders resolve against: master first, then layout overrides it. Lets us
    locate placeholders that inherit geometry rather than declaring their own.
    """
    REL_LAYOUT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                  "relationships/slideLayout")
    REL_MASTER = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                  "relationships/slideMaster")

    slide_name = slide_part.rsplit("/", 1)[-1]
    slide_rels = f"ppt/slides/_rels/{slide_name}.rels"
    layout_part = _rel_target(slide_rels, REL_LAYOUT, names, zf)

    geom = {}
    if layout_part:
        layout_name = layout_part.rsplit("/", 1)[-1]
        layout_rels = f"ppt/slideLayouts/_rels/{layout_name}.rels"
        master_part = _rel_target(layout_rels, REL_MASTER, names, zf)
        geom.update(_geom_from_part(master_part, names, zf))   # base
    geom.update(_geom_from_part(layout_part, names, zf))        # layout overrides
    return geom


def _shapes_for_slide(slide_xml: bytes, slide_cx: int, slide_cy: int,
                      layout_geom: dict | None = None):
    """
    Extract simple shape geometry from one slide's XML.

    Returns a list of dicts with name, text, and bbox expressed in slide-relative
    percentages (xPct, yPct, wPct, hPct) so the front end can hit-test clicks.

    Placeholders without an explicit a:xfrm inherit their geometry from the
    layout (looked up via `layout_geom`, keyed by placeholder type+idx).
    """
    layout_geom = layout_geom or {}
    root = etree.fromstring(slide_xml)
    shapes = []
    # sp = autoshape/textbox, pic = picture, graphicFrame = table/chart
    for tag in ("sp", "pic", "graphicFrame", "cxnSp"):
        for el in root.iter(f"{{{NS_P}}}{tag}"):
            bbox = _xfrm_bbox(el)
            if bbox is None:
                ph = el.find(f".//{{{NS_P}}}ph")
                bbox = layout_geom.get(_ph_key(ph)) if ph is not None else None
            if bbox is None:
                continue
            x, y, w, h = bbox

            # Shape name from nvPr cNvPr@name
            name = ""
            cnvpr = el.find(f".//{{{NS_P}}}cNvPr")
            if cnvpr is not None:
                name = cnvpr.get("name", "")

            # Concatenate text runs
            text = "".join(
                (t.text or "") for t in el.iter(f"{{{NS_A}}}t")
            ).replace("\n", " ").strip()

            shapes.append(
                {
                    "name": name,
                    "text": text[:200],
                    "xPct": round(x / slide_cx * 100, 2),
                    "yPct": round(y / slide_cy * 100, 2),
                    "wPct": round(w / slide_cx * 100, 2),
                    "hPct": round(h / slide_cy * 100, 2),
                }
            )
    return shapes


def extract_slide_shapes(pptx_path: Path):
    """Return a list (per slide, in order) of shape-geometry lists."""
    cx, cy = _slide_size_emu(pptx_path)
    slide_parts = _ordered_slide_files(pptx_path)
    out = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = set(z.namelist())
        for part in slide_parts:
            if part not in names:
                out.append([])
                continue
            layout_geom = _layout_placeholder_geometry(part, names, z)
            out.append(_shapes_for_slide(z.read(part), cx, cy, layout_geom))
    return out


# --------------------------------------------------------------------------- #
# Rendering: PPTX -> PDF (LibreOffice) -> PNG per slide (pdftoppm)
# --------------------------------------------------------------------------- #
def render_slide_pngs(pptx_path: Path, out_dir: Path, dpi: int = 96):
    """
    Render every slide of the PPTX to a PNG in out_dir.

    Returns the list of PNG filenames (relative to out_dir), in slide order.
    Requires `libreoffice` and `pdftoppm` on PATH.
    """
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not shutil.which("libreoffice"):
        raise RuntimeError("libreoffice not found on PATH — required to render slides")
    if not shutil.which("pdftoppm"):
        raise RuntimeError("pdftoppm not found on PATH — required to split the PDF")

    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile, \
         tempfile.TemporaryDirectory(prefix="lo-out-") as lo_out:
        subprocess.run(
            [
                "libreoffice", "--headless",
                f"-env:UserInstallation=file://{profile}",
                "--convert-to", "pdf", "--outdir", lo_out, str(pptx_path),
            ],
            check=True, capture_output=True, timeout=300,
        )
        pdf = Path(lo_out) / (pptx_path.stem + ".pdf")
        if not pdf.exists():
            raise RuntimeError(f"LibreOffice did not produce {pdf}")

        # pdftoppm writes slide-1.png, slide-2.png, ...
        subprocess.run(
            ["pdftoppm", "-png", "-r", str(dpi), str(pdf),
             str(out_dir / "slide")],
            check=True, capture_output=True, timeout=300,
        )

    pngs = sorted(
        (p for p in out_dir.glob("slide-*.png")),
        key=lambda p: int(p.stem.split("-")[-1]),
    )
    return [p.name for p in pngs]


# --------------------------------------------------------------------------- #
# Viewer generation
# --------------------------------------------------------------------------- #
def build_review_viewer(pptx_path: Path, out_dir: Path | None = None,
                        dpi: int = 96, verbose: bool = True) -> Path:
    """
    Build a self-contained reviewer-mode viewer for a PPTX.

    Renders each slide to a PNG, extracts shape geometry, and writes:
      <out_dir>/index.html        — the viewer + Reviewer Mode overlay
      <out_dir>/slides.json        — slide metadata (title, png, shapes)
      <out_dir>/slide-N.png        — rendered slides
      <out_dir>/serve-review.py    — one-shot save-server shim (serves this dir AND
                                       persists the reviewer's auto-saves to disk)
      <out_dir>/review-comments.json (written by the save-server on auto-save, or by
                                       the viewer's Export; auto-loaded on startup)

    Returns the path to index.html.
    """
    pptx_path = Path(pptx_path).resolve()
    if out_dir is None:
        out_dir = pptx_path.with_name(pptx_path.stem + "-review")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if verbose:
        print(f"Rendering {pptx_path.name} ...")
    pngs = render_slide_pngs(pptx_path, out_dir, dpi=dpi)
    titles = list_slides(pptx_path)
    shapes = extract_slide_shapes(pptx_path)

    n = len(pngs)
    slides = []
    for i in range(n):
        slides.append(
            {
                "index": i,
                "number": i + 1,
                "title": titles[i] if i < len(titles) else f"Slide {i + 1}",
                "png": pngs[i],
                "shapes": shapes[i] if i < len(shapes) else [],
            }
        )

    meta = {"deck": pptx_path.name, "slides": slides}
    (out_dir / "slides.json").write_text(json.dumps(meta, indent=2))

    html = _VIEWER_HTML.replace("__DECK_NAME__", _escape(pptx_path.name))
    (out_dir / "index.html").write_text(html)

    # Drop a one-shot save-server shim next to the viewer so a reviewer can persist
    # auto-saves to disk without the package installed:  python3 serve-review.py
    shim = out_dir / "serve-review.py"
    shim.write_text(_SERVE_SHIM)
    shim.chmod(0o755)

    if verbose:
        print(f"Review viewer: {out_dir / 'index.html'} ({n} slides)")
        print("Open it in a browser, toggle Review (or press R), "
              "click slides to comment. Auto-saves on 2s idle / click-outside.")
    return out_dir / "index.html"


# A standalone shim dropped into the viewer dir. It prefers the installed package's
# save-server; if pptx-slide-merger isn't importable it falls back to an inline
# copy of the same handler so the folder stays self-contained and portable.
_SERVE_SHIM = r'''#!/usr/bin/env python3
"""Serve this review viewer dir AND persist the reviewer's auto-saves to disk.

    python3 serve-review.py [--port 8000]

PUT/POST /review-comments.json writes the comments straight into this folder, so
Reviewer Mode auto-save (2s idle / click-outside) lands on disk with no Export.
"""
import argparse, json, sys
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
COMMENTS_FILE = "review-comments.json"


class H(SimpleHTTPRequestHandler):
    def _write(self):
        p = self.path.split("?", 1)[0]
        if not p.endswith(COMMENTS_FILE):
            self.send_error(404, "not found"); return
        try:
            n = int(self.headers.get("Content-Length", 0))
        except (TypeError, ValueError):
            n = 0
        if n <= 0 or n > 8 * 1024 * 1024:
            self._json(400, {"ok": False, "error": "bad length"}); return
        body = self.rfile.read(n)
        try:
            json.loads(body)
        except (ValueError, UnicodeDecodeError):
            self._json(400, {"ok": False, "error": "invalid json"}); return
        (HERE / COMMENTS_FILE).write_bytes(body)
        sys.stderr.write(f"[review] saved {COMMENTS_FILE} ({len(body)} bytes)\n")
        self._json(200, {"ok": True})

    def do_PUT(self): self._write()
    def do_POST(self): self._write()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,PUT,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _json(self, status, payload):
        b = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers(); self.wfile.write(b)

    def log_message(self, fmt, *a):
        sys.stderr.write("[review] " + (fmt % a) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--host", default="127.0.0.1")
    args = ap.parse_args()
    httpd = ThreadingHTTPServer((args.host, args.port),
                                partial(H, directory=str(HERE)))
    print(f"Review save-server on http://{args.host}:{httpd.server_address[1]}/  "
          f"(PUT {COMMENTS_FILE} persists into {HERE})")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nstopped")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
'''


def _escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# --------------------------------------------------------------------------- #
# The viewer (self-contained static HTML + JS). Loads slides.json at runtime.
# --------------------------------------------------------------------------- #
_VIEWER_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Review — __DECK_NAME__</title>
<style>
  :root { --accent:#EB1000; }
  * { box-sizing: border-box; }
  body { margin:0; background:#0c0e16; color:#e7ecf5;
         font:14px/1.4 -apple-system,Segoe UI,Roboto,sans-serif; }
  header { display:flex; align-items:center; gap:12px; padding:10px 16px;
           background:#12141f; border-bottom:1px solid #2a2f3e; position:sticky; top:0; z-index:50; }
  header h1 { font-size:14px; margin:0; font-weight:600; }
  header .deck { color:#8b95a7; font:12px/1.3 monospace; }
  .nav { margin-left:auto; display:flex; gap:6px; align-items:center; }
  button { background:#1b2030; color:#cfd6e4; border:1px solid #313748;
           border-radius:6px; padding:5px 10px; cursor:pointer; font:inherit; }
  button:hover { border-color:var(--accent); color:#fff; }
  button.primary { background:var(--accent); border-color:var(--accent); color:#fff; }
  #count { min-width:18px; text-align:center; background:var(--accent); color:#fff;
           border-radius:999px; padding:1px 8px; font-weight:700; }
  #stage { padding:24px; display:flex; justify-content:center; }
  #slideWrap { position:relative; max-width:1100px; width:100%;
               box-shadow:0 12px 40px rgba(0,0,0,.6); }
  #slideImg { width:100%; display:block; }
  body.rv-active #slideWrap { cursor:crosshair; }
  .rv-pin { position:absolute; transform:translate(-50%,-50%); z-index:60;
            width:24px; height:24px; border-radius:50% 50% 50% 2px; background:var(--accent);
            color:#fff; font:700 12px/24px monospace; text-align:center; cursor:pointer;
            border:2px solid #fff; box-shadow:0 2px 8px rgba(0,0,0,.5); }
  .rv-shapebox { position:absolute; border:1px dashed rgba(235,16,0,.6); pointer-events:none;
                 z-index:55; display:none; }
  .rv-pop { position:fixed; z-index:1001; width:320px; background:#12141f;
            border:1px solid var(--accent); border-radius:10px; padding:12px; color:#e7ecf5;
            box-shadow:0 10px 30px rgba(0,0,0,.6); }
  .rv-pop textarea { width:100%; height:74px; background:#0c0e16; color:#e7ecf5;
            border:1px solid #2a2f3e; border-radius:6px; padding:6px; font:inherit; resize:vertical; margin:6px 0; }
  .rv-pop .meta { font:11px/1.4 monospace; color:#8b95a7; margin-bottom:4px; word-break:break-word; }
  .rv-pop .row { display:flex; gap:6px; justify-content:flex-end; margin-top:4px; }
  footer { padding:6px 16px; color:#8b95a7; font-size:12px; border-top:1px solid #2a2f3e; }
  #toast { position:fixed; left:16px; bottom:16px; z-index:2000; background:#1b2030; color:#a3e6d8;
           border:1px solid #2bbfa9; border-radius:6px; padding:6px 10px; font-size:12px; display:none; }
  /* Review toggle: opt-in, collapsed by default. Lights up when on. */
  #rvToggle { display:inline-flex; align-items:center; gap:6px; }
  #rvToggle.rv-on { background:var(--accent); border-color:var(--accent); color:#fff; }
  #rvToggle .rv-count { background:#fff; color:var(--accent); border-radius:999px;
           padding:0 7px; font-weight:700; font-size:11px; }
  #rvToggle.rv-on .rv-count { background:#12141f; color:#fff; }
  #rvToggle .rv-count:empty { display:none; }
  /* Review-only controls (panel actions) hidden until review mode is on. */
  .rv-only { display:none; }
  body.rv-active .rv-only { display:inline-block; }
</style>
</head>
<body>
<header>
  <h1>Slide Review</h1>
  <span class="deck">__DECK_NAME__</span>
  <div class="nav">
    <button id="prev">◀</button>
    <span id="pos" style="font:12px monospace;color:#8b95a7"></span>
    <button id="next">▶</button>
    <button id="rvToggle" title="Toggle Reviewer Mode (R)" aria-pressed="false">🖉 Review<span class="rv-count"></span></button>
    <button id="export" class="rv-only" title="Download review-comments.json">⤓ Export</button>
    <button id="copy" class="rv-only" title="Copy all comments as markdown for Claude">⧉ Copy for Claude</button>
    <button id="clear" class="rv-only" title="Clear all comments">✕ Clear</button>
  </div>
</header>

<div id="stage">
  <div id="slideWrap">
    <img id="slideImg" alt="slide">
    <div class="rv-shapebox" id="shapebox"></div>
  </div>
</div>
<footer><span id="footHint">Toggle <b>Review</b> (or press <b>R</b>) to start commenting. Collapsed = clean viewing.</span></footer>
<div id="toast"></div>

<script>
(function () {
  'use strict';
  const STORAGE_KEY = 'pptx-review:' + location.pathname;
  let meta = { deck: '', slides: [] };
  let comments = [];
  let active = false;
  let cur = 0;
  let pop = null;

  const $ = (id) => document.getElementById(id);
  const wrap = $('slideWrap'), img = $('slideImg'), shapebox = $('shapebox');

  // ---------- persistence ----------
  function load() {
    try { comments = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); } catch { comments = []; }
  }
  function save() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(comments));
    renderBadge(); renderPins();
    // best-effort: persist straight into the viewer folder via the save-server (PUT).
    // On a plain static server (no save-server) this 405/404s harmlessly — localStorage
    // + Export still work as the fallback.
    fetch('review-comments.json', {
      method: 'PUT', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(comments, null, 2)
    }).then(r => { if (r.ok) toast('Saved to folder ✓'); }).catch(() => {});
  }
  async function loadFile() {
    try {
      const r = await fetch('review-comments.json', { cache: 'no-store' });
      if (!r.ok) return;
      const fileComments = await r.json();
      if (Array.isArray(fileComments)) {
        const ids = new Set(comments.map(c => c.id));
        fileComments.forEach(c => { if (!ids.has(c.id)) comments.push(c); });
        save();
      }
    } catch { /* no file — fine */ }
  }

  // ---------- helpers ----------
  function slide() { return meta.slides[cur] || { shapes: [], title: '', number: cur + 1 }; }
  function hitShape(xPct, yPct) {
    // smallest-area shape whose bbox contains the click
    let best = null, bestArea = Infinity;
    (slide().shapes || []).forEach(s => {
      if (xPct >= s.xPct && xPct <= s.xPct + s.wPct &&
          yPct >= s.yPct && yPct <= s.yPct + s.hPct) {
        const area = s.wPct * s.hPct;
        if (area < bestArea) { bestArea = area; best = s; }
      }
    });
    return best;
  }

  // ---------- render slide ----------
  function showSlide(i) {
    cur = Math.max(0, Math.min(meta.slides.length - 1, i));
    const s = slide();
    img.src = s.png;
    $('pos').textContent = (cur + 1) + ' / ' + meta.slides.length;
    closePop();
    renderPins();
  }

  function renderBadge() {
    const c = document.querySelector('#rvToggle .rv-count');
    if (c) c.textContent = comments.length ? comments.length : '';
  }

  function renderPins() {
    wrap.querySelectorAll('.rv-pin').forEach(p => p.remove());
    if (!active) return;   // pins only show in review mode (clean viewing otherwise)
    comments.filter(c => c.slideIndex === cur).forEach((c) => {
      const pin = document.createElement('div');
      pin.className = 'rv-pin';
      pin.textContent = (comments.indexOf(c) + 1);
      pin.style.left = c.xPct + '%';
      pin.style.top = c.yPct + '%';
      pin.onclick = (e) => { e.stopPropagation(); openView(c, e.clientX, e.clientY); };
      wrap.appendChild(pin);
    });
  }

  function setActive(on) {
    active = on;
    document.body.classList.toggle('rv-active', on);
    const t = $('rvToggle');
    t.classList.toggle('rv-on', on);
    t.setAttribute('aria-pressed', on ? 'true' : 'false');
    $('footHint').innerHTML = on
      ? 'Click a slide to comment. Bubble auto-saves on 2s idle or click-outside. <b>R</b> / ✕ Review to exit.'
      : 'Toggle <b>Review</b> (or press <b>R</b>) to start commenting. Collapsed = clean viewing.';
    if (!on) { closePop(); shapebox.style.display = 'none'; }
    renderPins();
  }

  // ---------- popovers ----------
  // commitPending: while a new-comment bubble is open, this holds the function that
  // commits it. A click outside (or 2s idle) calls it; closePop clears it.
  let commitPending = null, idleTimer = null;
  function closePop() {
    if (idleTimer) { clearTimeout(idleTimer); idleTimer = null; }
    commitPending = null;
    if (pop) { pop.remove(); pop = null; }
  }
  function place(el, x, y) {
    el.style.left = Math.min(x, window.innerWidth - 340) + 'px';
    el.style.top = Math.min(y, window.innerHeight - 230) + 'px';
  }
  function openNew(xPct, yPct, target, clientX, clientY) {
    closePop();
    const s = slide();
    pop = document.createElement('div'); pop.className = 'rv-pop';
    const tdesc = target
      ? ('shape: "' + (target.name || '?') + '"' +
         (target.text ? ' — "' + target.text.slice(0, 60) + '"' : ''))
      : '(no shape at this point — region comment)';
    pop.innerHTML =
      '<div class="meta">Slide ' + s.number + ' · "' + esc((s.title || '').slice(0, 40)) + '"<br>' +
      esc(tdesc) + '<br>@ ' + xPct + '% , ' + yPct + '%</div>' +
      '<textarea placeholder="Your comment / change request…"></textarea>' +
      '<div class="row"><button class="cancel">Cancel</button>' +
      '<button class="primary save">Save</button></div>';
    document.body.appendChild(pop); place(pop, clientX, clientY);
    const ta = pop.querySelector('textarea'); ta.focus();
    let done = false;
    function commitComment() {
      if (done) return; done = true;
      const text = ta.value.trim();
      if (text) {
        comments.push({
          id: 'c' + Date.now() + Math.floor(Math.random() * 1e4),
          type: 'comment',
          slideIndex: cur,
          slideNumber: s.number,
          slideTitle: s.title,
          xPct: xPct, yPct: yPct,
          targetShape: target ? {
            name: target.name, text: target.text,
            xPct: target.xPct, yPct: target.yPct, wPct: target.wPct, hPct: target.hPct
          } : null,
          comment: text,
          ts: new Date().toISOString()
        });
        save();
      }
      closePop();
    }
    commitPending = commitComment;
    // auto-save: 2s after the last keystroke
    ta.addEventListener('input', () => {
      if (idleTimer) clearTimeout(idleTimer);
      idleTimer = setTimeout(commitComment, 2000);
    });
    pop.querySelector('.cancel').onclick = () => { done = true; closePop(); }; // discard
    pop.querySelector('.save').onclick = commitComment;
  }
  function openView(c, x, y) {
    closePop();
    pop = document.createElement('div'); pop.className = 'rv-pop';
    const t = c.targetShape;
    pop.innerHTML =
      '<div class="meta">Slide ' + c.slideNumber + ' · ' + esc(c.type) + '<br>' +
      (t ? 'shape: "' + esc(t.name || '?') + '"' : 'region @ ' + c.xPct + '%,' + c.yPct + '%') +
      '</div><div>' + esc(c.comment || '') + '</div>' +
      '<div class="row"><button class="del">Delete</button>' +
      '<button class="primary close">Close</button></div>';
    document.body.appendChild(pop); place(pop, x, y);
    pop.querySelector('.close').onclick = closePop;
    pop.querySelector('.del').onclick = () => {
      comments = comments.filter(z => z.id !== c.id); save(); closePop();
    };
  }

  // ---------- export ----------
  function exportJson() {
    const blob = new Blob([JSON.stringify(comments, null, 2)], { type: 'application/json' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
    a.download = 'review-comments.json'; a.click(); URL.revokeObjectURL(a.href);
    toast('Saved review-comments.json — drop it next to index.html');
  }
  function copyForClaude() {
    const by = {};
    comments.forEach(c => { (by[c.slideIndex] = by[c.slideIndex] || []).push(c); });
    let md = '# Deck review comments (' + comments.length + ') — ' + meta.deck + '\n\n';
    Object.keys(by).sort((a, b) => a - b).forEach(i => {
      const first = by[i][0];
      md += '## Slide ' + first.slideNumber + ' — ' + (first.slideTitle || '') + '\n';
      by[i].forEach(c => {
        const where = c.targetShape
          ? 'shape "' + (c.targetShape.name || '?') + '"' +
            (c.targetShape.text ? ' ("' + c.targetShape.text.slice(0, 50) + '")' : '')
          : 'region @ ' + c.xPct + '%,' + c.yPct + '%';
        md += '- on ' + where + ': ' + c.comment + '\n';
      });
      md += '\n';
    });
    navigator.clipboard.writeText(md).then(() => toast('Copied for Claude ✓'),
      () => toast('Copy failed — see console') );
  }
  function toast(msg) {
    const t = $('toast'); t.textContent = msg; t.style.display = 'block';
    clearTimeout(toast._t); toast._t = setTimeout(() => t.style.display = 'none', 2400);
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>]/g, m =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m])); }

  // ---------- interactions ----------
  wrap.addEventListener('mousemove', (e) => {
    if (!active) return;
    const r = wrap.getBoundingClientRect();
    const xPct = (e.clientX - r.left) / r.width * 100;
    const yPct = (e.clientY - r.top) / r.height * 100;
    const sh = hitShape(xPct, yPct);
    if (sh) {
      shapebox.style.display = 'block';
      shapebox.style.left = sh.xPct + '%';
      shapebox.style.top = sh.yPct + '%';
      shapebox.style.width = sh.wPct + '%';
      shapebox.style.height = sh.hPct + '%';
    } else { shapebox.style.display = 'none'; }
  });
  // Capture-phase document click: while a new-comment bubble is open, a click
  // anywhere outside it auto-saves the bubble AND is consumed, so it never also
  // drops a stray second comment. Otherwise a click on the slide opens a new one.
  document.addEventListener('click', (e) => {
    if (!active) return;
    const t = e.target;
    // ignore clicks on our own chrome / the open bubble / a pin
    if (t.closest('.rv-pop') || t.closest('header') || t.classList.contains('rv-pin')) return;
    if (commitPending) {
      e.preventDefault(); e.stopPropagation();
      commitPending();
      return;
    }
    if (!wrap.contains(t)) return;     // clicks off the slide image do nothing
    const r = wrap.getBoundingClientRect();
    const xPct = +(((e.clientX - r.left) / r.width) * 100).toFixed(2);
    const yPct = +(((e.clientY - r.top) / r.height) * 100).toFixed(2);
    openNew(Math.max(0, Math.min(100, xPct)), Math.max(0, Math.min(100, yPct)),
            hitShape(xPct, yPct), e.clientX, e.clientY);
  }, true);

  $('rvToggle').onclick = () => setActive(!active);
  $('prev').onclick = () => showSlide(cur - 1);
  $('next').onclick = () => showSlide(cur + 1);
  $('export').onclick = exportJson;
  $('copy').onclick = copyForClaude;
  $('clear').onclick = () => { if (confirm('Clear all ' + comments.length + ' comments?')) { comments = []; save(); } };

  // Keyboard shield: typing inside a comment field must never reach the viewer's
  // global key handler (R toggle, arrows). Capture phase + stopPropagation so
  // letters like r/x/p type into the textarea instead of navigating. Do NOT
  // preventDefault — let the character through.
  const isTypingTarget = (t) => t && (t.tagName === 'TEXTAREA' || t.tagName === 'INPUT' ||
      t.isContentEditable === true || (t.closest && t.closest('.rv-pop')));
  ['keydown', 'keypress', 'keyup'].forEach(ev =>
    document.addEventListener(ev, (e) => { if (isTypingTarget(e.target)) e.stopPropagation(); }, true));

  window.addEventListener('keydown', (e) => {
    if (isTypingTarget(e.target)) return;
    if (e.key === 'r' || e.key === 'R') { if (!e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); setActive(!active); } }
    else if (e.key === 'ArrowLeft') showSlide(cur - 1);
    else if (e.key === 'ArrowRight') showSlide(cur + 1);
    else if (e.key === 'Escape') closePop();
  });

  // ---------- init ----------
  (async function init() {
    try {
      const r = await fetch('slides.json', { cache: 'no-store' });
      meta = await r.json();
    } catch (err) {
      document.getElementById('stage').innerHTML =
        '<div style="color:#ff8a80">Could not load slides.json — serve this folder over HTTP (file:// blocks fetch).</div>';
      return;
    }
    load(); renderBadge();
    showSlide(0);
    await loadFile();
  })();
})();
</script>
</body>
</html>
"""
