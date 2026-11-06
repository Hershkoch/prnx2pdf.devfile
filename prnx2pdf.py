#!/usr/bin/env python3
"""
prnx2pdf.py - Convert DevExpress .prnx report documents (e.g. from BASSnet) to PDF.

A .prnx file is a gzip-compressed XML snapshot of a report as it looked in a
DevExpress Print Preview window. This script reads that XML and redraws every
page (labels, tables, lines, images, headers/footers, page numbers) into a PDF.

Requirements:
    pip install reportlab

Usage:
    python prnx2pdf.py report.prnx                  # -> report.pdf next to it
    python prnx2pdf.py report.prnx -o out.pdf
    python prnx2pdf.py *.prnx -d converted/         # batch, into a folder
    python prnx2pdf.py report.prnx --printed "3/25/2026 9:52 AM"

Notes / limitations:
    * Fonts: Microsoft Sans Serif / Arial are substituted with Arial (Windows/macOS)
      or Liberation Sans (Linux). Line breaks may differ very slightly from the
      original, so long text that is split across pages is worth a quick check.
    * "Printed Date & Time" fields are evaluated at print time and are not stored
      in the file; they default to the current time (override with --printed).
    * Wingdings bullets (U+F0A7) are shown as a small square bullet; tabs become spaces.
"""

import argparse
import base64
import datetime
import glob
import gzip
import io
import os
import re
import sys
import xml.etree.ElementTree as ET

from reportlab.lib.colors import Color
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# ----------------------------------------------------------------------------
# Fonts
# ----------------------------------------------------------------------------

FONT_CANDIDATES = {
    # variant: list of possible files (first one found wins)
    "R": ["arial.ttf", "Arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"],
    "B": ["arialbd.ttf", "Arial Bold.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"],
    "I": ["ariali.ttf", "Arial Italic.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf"],
    "BI": ["arialbi.ttf", "Arial Bold Italic.ttf", "LiberationSans-BoldItalic.ttf",
           "DejaVuSans-BoldOblique.ttf"],
}
FONT_DIRS = [
    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
    "/Library/Fonts", "/System/Library/Fonts/Supplemental", os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
]

_fonts_ready = False
FONT_NAMES = {}


def _find_font_file(names):
    for d in FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for name in names:
            p = os.path.join(d, name)
            if os.path.isfile(p):
                return p
            hits = glob.glob(os.path.join(d, "**", name), recursive=True)
            if hits:
                return hits[0]
    return None


def setup_fonts():
    """Register TrueType fonts; fall back to built-in Helvetica if none found."""
    global _fonts_ready
    if _fonts_ready:
        return
    builtin = {"R": "Helvetica", "B": "Helvetica-Bold", "I": "Helvetica-Oblique",
               "BI": "Helvetica-BoldOblique"}
    for var, names in FONT_CANDIDATES.items():
        path = _find_font_file(names)
        if path:
            fname = "PRNX-" + var
            pdfmetrics.registerFont(TTFont(fname, path))
            FONT_NAMES[var] = fname
        else:
            FONT_NAMES[var] = builtin[var]
    if FONT_NAMES["R"] == "Helvetica":
        print("warning: no TrueType sans font found; using Helvetica "
              "(characters like bullets may not render)", file=sys.stderr)
    _fonts_ready = True


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

K = 72 / 300.0  # DevExpress document units (1/300 inch) -> PDF points

NAMED_COLORS = {"Black": (0, 0, 0), "Gray": (.5, .5, .5), "White": (1, 1, 1),
                "Red": (1, 0, 0), "Blue": (0, 0, 1), "Green": (0, .5, 0),
                "Silver": (.75, .75, .75), "DarkGray": (.66, .66, .66),
                "LightGray": (.83, .83, .83)}

# Paper sizes in document units (1/300 inch) for common PaperKind values
PAPER = {"A4": (2480.315, 3507.874), "Letter": (2550, 3300), "Legal": (2550, 4200),
         "A3": (3507.874, 4960.63), "A5": (1748.031, 2480.315)}


def read_prnx(path):
    with open(path, "rb") as f:
        data = f.read()
    if data[:2] == b"\x1f\x8b":
        data = gzip.decompress(data)
    return ET.fromstring(data.decode("utf-8-sig"))


def parse_rect(e):
    return [float(v) for v in e.attrib["Rect"].split(",")]


def parse_color(v):
    if not v or v == "Transparent":
        return None
    if v in NAMED_COLORS:
        return Color(*NAMED_COLORS[v])
    try:
        p = [int(x) for x in v.split(",")]
    except ValueError:
        return Color(0, 0, 0)
    if len(p) == 4:
        a, r, g, b = p
        return None if a == 0 else Color(r / 255, g / 255, b / 255)
    return Color(*[x / 255 for x in p[:3]])


def parse_font(st):
    parts = [p.strip() for p in st.get("Font", "Arial, 9.75pt").split(",")]
    size = float(re.sub(r"[^\d.]", "", parts[1]) or 9.75) if len(parts) > 1 else 9.75
    if len(parts) > 1 and parts[1].endswith("px"):
        size *= 0.75
    style = " ".join(parts[2:])
    bold, italic = "Bold" in style, "Italic" in style
    var = "BI" if bold and italic else "B" if bold else "I" if italic else "R"
    return FONT_NAMES[var], size, "Underline" in style


def clean_text(text):
    return (text.replace("\uf0a7", "\u25aa")   # Wingdings square bullet
                .replace("\uf0b7", "\u2022")   # Symbol bullet
                .replace("\u2011", "-")        # non-breaking hyphen
                .replace("\t", "   "))


def wrap_text(text, fname, fsize, width, nowrap):
    sw = pdfmetrics.stringWidth
    lines = []
    for para in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if nowrap:
            lines.append(para)
            continue
        cur = ""
        for w in re.split(r"(\s+)", para):
            if not w:
                continue
            trial = cur + w
            if sw(trial.rstrip(), fname, fsize) <= width or not cur.strip():
                if not cur.strip() and not w.isspace() and sw(trial, fname, fsize) > width:
                    piece = ""  # a single word wider than the box: break by character
                    for ch in trial:
                        if piece and sw(piece + ch, fname, fsize) > width:
                            lines.append(piece)
                            piece = ch
                        else:
                            piece += ch
                    cur = piece
                else:
                    cur = trial
            else:
                lines.append(cur.rstrip())
                cur = "" if w.isspace() else w
        lines.append(cur.rstrip())
    return lines


# ----------------------------------------------------------------------------
# Renderer
# ----------------------------------------------------------------------------

class PrnxRenderer:
    def __init__(self, root, printed=None):
        self.root = root
        self.opts = root.find("Options")
        self.printed = printed or datetime.datetime.now()

        self.styles = {}
        ss = self.opts.find("SharedStyles")
        for s in (ss if ss is not None else []):
            self.styles[int(s.attrib["Index"])] = s.attrib

        self.images = {}
        si = self.opts.find("SharedImages")
        for i, s in enumerate(si if si is not None else []):
            src = s.attrib.get("ImageSource", "")
            if "," in src:
                try:
                    raw = base64.b64decode(src.split(",", 1)[1])
                    self.images[int(s.attrib.get("Index", i))] = ImageReader(io.BytesIO(raw))
                except Exception:
                    pass

        self.shared = {}
        sb = self.opts.find("SharedBricks")
        for b in (sb if sb is not None else []):
            self.shared[int(b.attrib["Index"])] = b

        self.pages = [p for p in root if re.fullmatch(r"Page\d+", p.tag)]
        self.npages = len(self.pages)

        # paper size
        self.pw, self.ph = PAPER["A4"]
        pd = self.opts.find("PageData")
        if pd is not None and len(pd):
            a = pd[0].attrib
            if "SizeF" in a:  # SizeF is in 1/100 inch
                w, h = [float(v) for v in a["SizeF"].split(",")]
                self.pw, self.ph = w * 3, h * 3
            elif a.get("PaperKind") in PAPER:
                self.pw, self.ph = PAPER[a["PaperKind"]]
            if a.get("Landscape") == "true" and self.pw < self.ph:
                self.pw, self.ph = self.ph, self.pw
        self.PH = self.ph * K
        self.c = None

    # -- style lookup -------------------------------------------------------
    def style_of(self, e):
        st = e.find("Style")
        if st is not None:
            return st.attrib
        if "Style" in e.attrib:
            return self.styles.get(int(e.attrib["Style"]), {})
        return {}

    # -- primitives ---------------------------------------------------------
    def draw_text(self, txt, st, x, y, w, h, clip=None):
        if not txt:
            return
        txt = clean_text(txt)
        c, PH = self.c, self.PH
        fname, fsize, underline = parse_font(st)
        pad = [float(v) for v in st.get("Padding", "0,0,0,0,100").split(",")]
        dpi = pad[4] if len(pad) > 4 else 100
        pl, pr, pt, pb = [v * 300.0 / dpi for v in pad[:4]]
        nowrap = "NoWrap" in st.get("StringFormat", "")
        align = st.get("TextAlignment", "TopLeft")

        inner_w = (w - pl - pr) * K
        lines = wrap_text(txt, fname, fsize, inner_w, nowrap)
        ascent = pdfmetrics.getAscent(fname) / 1000 * fsize
        lh = fsize * 1.149
        total = lh * len(lines)
        top = (y + pt) * K
        inner_h = (h - pt - pb) * K
        if align.startswith("Middle"):
            top += (inner_h - total) / 2
        elif align.startswith("Bottom"):
            top += inner_h - total

        col = parse_color(st.get("ForeColor", "Black")) or Color(0, 0, 0)
        c.setFillColor(col)
        c.setFont(fname, fsize)
        for i, ln in enumerate(lines):
            ly = top + i * lh
            if clip and not (clip[0] * K <= ly + lh / 2 < clip[1] * K):
                continue  # line belongs to the part of a split brick on another page
            base = PH - (ly + ascent)
            tw = pdfmetrics.stringWidth(ln, fname, fsize)
            if align.endswith("Right"):
                lx = (x + w - pr) * K - tw
            elif align.endswith("Center"):
                lx = (x + pl) * K + (inner_w - tw) / 2
            else:
                lx = (x + pl) * K
            c.drawString(lx, base, ln)
            if underline and ln.strip():
                c.setLineWidth(fsize / 14)
                c.setStrokeColor(col)
                c.line(lx, base - fsize * 0.12, lx + tw, base - fsize * 0.12)

    def draw_background(self, st, x, y, w, h):
        col = parse_color(st.get("BackColor", "Transparent"))
        if col:
            self.c.setFillColor(col)
            self.c.rect(x * K, self.PH - (y + h) * K, w * K, h * K, stroke=0, fill=1)

    def draw_borders(self, st, x, y, w, h):
        sides = st.get("Sides", "None")
        col = parse_color(st.get("BorderColor", "Black"))
        if sides == "None" or col is None:
            return
        c, PH = self.c, self.PH
        c.setStrokeColor(col)
        c.setLineWidth(float(st.get("BorderWidthSerializable", "1")) * 0.75)
        x0, x1 = x * K, (x + w) * K
        y0, y1 = PH - y * K, PH - (y + h) * K
        every = sides == "All"
        if every or "Top" in sides:
            c.line(x0, y0, x1, y0)
        if every or "Bottom" in sides:
            c.line(x0, y1, x1, y1)
        if every or "Left" in sides:
            c.line(x0, y0, x0, y1)
        if every or "Right" in sides:
            c.line(x1, y0, x1, y1)

    def page_info_text(self, e, pageno):
        fmt = e.attrib.get("Format", "{0}")
        kind = e.attrib.get("PageInfo")
        d = self.printed
        if kind == "DateTime":
            g = f"{d.month}/{d.day}/{d.year} {d.strftime('%I:%M %p').lstrip('0')}"
            out = fmt.replace("{0:g}", g).replace("{0:d}", f"{d.month}/{d.day}/{d.year}")
            return re.sub(r"\{0(:[^}]*)?\}", g, out)
        if kind == "NumberOfTotal":
            return fmt.replace("{0}", str(pageno)).replace("{1}", str(self.npages))
        if kind in ("Number", "RomLowNumber", "RomHiNumber"):
            return fmt.replace("{0}", str(pageno))
        if kind == "Total":
            return fmt.replace("{0}", str(self.npages))
        return fmt

    # -- bricks -------------------------------------------------------------
    def draw_brick(self, e, ox, oy, pageno, clip=None, pos=None):
        if e.attrib.get("IsVisible") == "false" or "Rect" not in e.attrib:
            return
        btype = e.attrib.get("BrickType")
        rx, ry, w, h = parse_rect(e)
        if pos:
            rx, ry = pos
        x, y = ox + rx, oy + ry
        st = self.style_of(e)
        c, PH = self.c, self.PH

        if btype == "Container":
            # A window onto a shared brick, used when a brick is split across pages.
            target = self.shared.get(int(e.attrib["Brick"]))
            if target is None:
                return
            off = float(e.attrib.get("BrickOffsetY", "0"))
            c.saveState()
            p = c.beginPath()
            p.rect(x * K, PH - (y + h) * K, w * K, h * K)
            c.clipPath(p, stroke=0, fill=0)
            self.draw_brick(target, x, y + off, pageno, clip=(y, y + h), pos=(0, 0))
            c.restoreState()
            return

        if btype == "CompositeBrick":
            inner = e.find("InnerBricks")
            for ch in (inner if inner is not None else []):
                self.draw_brick(ch, x, y, pageno, clip)
            return

        self.draw_background(st, x, y, w, h)

        if btype in ("Label", "Text", "CheckBoxText"):
            self.draw_text(e.attrib.get("Text"), st, x, y, w, h, clip)
        elif btype == "PageInfoText":
            self.draw_text(self.page_info_text(e, pageno), st, x, y, w, h, clip)
        elif btype == "Line":
            lw = float(e.attrib.get("LineWidth", "1"))
            col = parse_color(st.get("ForeColor", "Black")) or Color(0, 0, 0)
            c.setStrokeColor(col)
            c.setLineWidth(lw * 0.75)
            if e.attrib.get("HtmlLineDirection", "Horizontal") == "Horizontal":
                yy = PH - (y + h / 2) * K
                c.line(x * K, yy, (x + w) * K, yy)
            else:
                xx = (x + w / 2) * K
                c.line(xx, PH - y * K, xx, PH - (y + h) * K)
        elif btype == "Image":
            img = self.images.get(int(e.attrib.get("ImageEntry", "0")))
            if img:
                iw, ih = img.getSize()
                scale = min(w * K / iw, h * K / ih)
                if e.attrib.get("SizeMode") == "Squeeze":
                    scale = min(scale, 0.75)  # 96-dpi pixels -> points, shrink only
                dw, dh = iw * scale, ih * scale
                al = e.attrib.get("ImageAlignment", "MiddleLeft")
                dx = x * K
                if al.endswith("Center"):
                    dx += (w * K - dw) / 2
                elif al.endswith("Right"):
                    dx += w * K - dw
                dy = y * K + (h * K - dh) / 2
                c.drawImage(img, dx, PH - dy - dh, dw, dh, mask="auto")

        kids = e.find("Bricks")  # Panel / Table / Row children
        for ch in (kids if kids is not None else []):
            self.draw_brick(ch, x, y, pageno, clip)

        self.draw_borders(st, x, y, w, h)

    # -- output -------------------------------------------------------------
    def save(self, out_path):
        setup_fonts()
        self.c = canvas.Canvas(out_path, pagesize=(self.pw * K, self.ph * K))
        self.c.setTitle(self.opts.attrib.get("Name", "Report"))
        for n, page in enumerate(self.pages, 1):
            inner = page.find("InnerBricks")
            for ch in (inner if inner is not None else []):
                self.draw_brick(ch, 0, 0, n)
            self.c.showPage()
        self.c.save()
        return self.npages


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------

def convert(src, dst, printed=None):
    root = read_prnx(src)
    if root.tag != "PreviewSerializer":
        raise ValueError("not a DevExpress .prnx document")
    return PrnxRenderer(root, printed).save(dst)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Convert DevExpress .prnx reports to PDF.")
    ap.add_argument("inputs", nargs="+", help=".prnx file(s); wildcards allowed")
    ap.add_argument("-o", "--output", help="output PDF path (single input only)")
    ap.add_argument("-d", "--outdir", help="folder for output PDFs")
    ap.add_argument("--printed", help='value for "Printed Date & Time" fields, '
                                      'e.g. "3/25/2026 9:52 AM" (default: now)')
    args = ap.parse_args(argv)

    files = []
    for pat in args.inputs:
        files.extend(sorted(glob.glob(pat)) or [pat])
    if args.output and len(files) > 1:
        ap.error("--output can only be used with a single input; use --outdir")

    printed = None
    if args.printed:
        for fmt in ("%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M", "%m/%d/%Y", "%Y-%m-%d"):
            try:
                printed = datetime.datetime.strptime(args.printed, fmt)
                break
            except ValueError:
                pass
        if printed is None:
            ap.error(f"could not parse --printed value: {args.printed}")

    if args.outdir:
        os.makedirs(args.outdir, exist_ok=True)

    failed = 0
    for src in files:
        if args.output:
            dst = args.output
        else:
            base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
            dst = os.path.join(args.outdir or os.path.dirname(src) or ".", base)
        try:
            n = convert(src, dst, printed)
            print(f"{src} -> {dst} ({n} pages)")
        except Exception as ex:
            failed += 1
            print(f"error: {src}: {ex}", file=sys.stderr)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
