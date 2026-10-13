"""Convert lecture-slide PDFs to editable Word documents, one slide per page.

Native slide text becomes editable Word text (keeping bold/italic/colour/
super- and subscripts). Photos, charts and diagrams are kept as pictures,
rendered from the slide without the template background. Tables that exist
in the PDF only as pictures can be supplied as editable tables through a
JSON file (see --tables).

Usage: python3 pdf2word.py OUT_DIR LECTURE.pdf [...] [--tables tables.json]
"""
import argparse
import collections
import io
import json
import os
import re

import pymupdf
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT_SCALE = 0.6          # slide font sizes are far larger than document text
MIN_PT, MAX_BODY_PT = 9.5, 16
TITLE_MIN_SIZE = 24        # slide font size from which a line counts as a title
FIG_DPI = 200
PAGE_W, PAGE_H = Cm(29.7), Cm(21.0)
MARGIN = Cm(1.6)
CONTENT_W = PAGE_W - 2 * MARGIN
BULLETS = "•–-▪■●○◦➢➤►▶✓"


# ---------------------------------------------------------------- analysis
def template_xrefs(doc):
    """Images repeated on most slides are the slide template (background, logo)."""
    cnt = collections.Counter()
    for page in doc:
        for x in {i["xref"] for i in page.get_image_info(xrefs=True)}:
            cnt[x] += 1
    return {x for x, c in cnt.items() if c >= 0.5 * len(doc)}


def cluster_rects(rects, gap=6):
    rects = [pymupdf.Rect(r) for r in rects]
    merged = True
    while merged:
        merged = False
        out = []
        while rects:
            r = rects.pop()
            i = 0
            while i < len(rects):
                if (r + (-gap, -gap, gap, gap)).intersects(rects[i]):
                    r |= rects.pop(i)
                    merged = True
                else:
                    i += 1
            out.append(r)
        rects = out
    return rects


def figure_regions(page, tmpl, lines=None):
    """Content images and vector-drawing clusters, as (rect, kind, xref)."""
    pr = page.rect
    regs = []
    for info in page.get_image_info(xrefs=True):
        if info["xref"] in tmpl:
            continue
        r = pymupdf.Rect(info["bbox"]) & pr
        if r.is_empty or r.width < 12 or r.height < 12:
            continue
        if r.height < 60 and r.width < 260:
            continue                     # small inline picture: handled as equation
        regs.append([r, "image", info["xref"]])
    solid, thin = [], []
    for d in page.get_drawings():
        r = pymupdf.Rect(d["rect"]) & pr
        if r.is_empty and not (r.width or r.height):
            continue
        (thin if min(r.width, r.height) < 3 and max(r.width, r.height) > 20 else solid).append(r)
    bad = []
    for ln in lines or []:
        for sp in ln["spans"]:
            if any(_bad_char(ch, False) for ch in sp["text"] if ch.strip()):
                bad.append(pymupdf.Rect(sp["bbox"]))
    for r in cluster_rects(solid, gap=20):
        if r.width * max(r.height, 1) < 2500 or min(r.width, r.height) < 8:
            continue                     # glyph-sized: left to equation detection
        if r.height < 70 or (r.height < 110 and any(r.intersects(b) for b in bad)):
            continue                     # a typeset equation, not a diagram
        for t in thin:
            if (r + (-4, -4, 4, 4)).intersects(t):
                r |= t
        # short labels written next to a diagram belong to it
        grew = True
        while grew:
            grew = False
            for ln in lines or []:
                b = ln["bbox"]
                txt = "".join(sp["text"] for sp in ln["spans"]).strip()
                if b in r or txt[:1] in BULLETS:
                    continue
                same_box = any(o["block"] == ln["block"] and o["bbox"] in r for o in lines)
                short = len(txt) <= 14 or b.width <= 110
                if (same_box and (r + (-30, -30, 30, 30)).intersects(b)) or (
                        short and (r + (-18, -18, 18, 18)).intersects(b)):
                    r |= b
                    grew = True
        regs.append([r, "drawing", None])
    # merge overlapping regions so a diagram drawn over a picture stays one figure
    changed = True
    while changed:
        changed = False
        for i in range(len(regs)):
            for j in range(i + 1, len(regs)):
                a, b = regs[i][0], regs[j][0]
                g = 30 if regs[i][1] != "image" and regs[j][1] != "image" else 4
                if not (a + (-g, -g, g, g)).intersects(b):
                    continue
                if True:
                    regs[i][0] = a | b
                    regs[i][1] = "mixed" if regs[i][1] != regs[j][1] else regs[i][1]
                    regs[i][2] = regs[i][2] or regs[j][2]
                    regs.pop(j)
                    changed = True
                    break
            if changed:
                break
    return regs


def _is_math(ch):
    return 0x1D400 <= ord(ch) <= 0x1D7FF


def span_text(span):
    """Rebuild span text from glyph positions; restore spaces the PDF omits."""
    chars = span["chars"]
    out = []
    for i, c in enumerate(chars):
        if i:
            a = chars[i - 1]
            gap = c["bbox"][0] - a["bbox"][2]
            lim = 0.05
            if _is_math(a["c"]) or _is_math(c["c"]):
                lim = 0.15
            elif a["c"].isupper() and c["c"].isupper():
                lim = 0.1
            if a["c"] != " " and c["c"] != " " and gap > lim * span["size"]:
                out.append(" ")
        out.append(c["c"])
    return "".join(out)


def slide_lines(page):
    raw = []
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b["lines"]:
            spans = []
            for s in ln["spans"]:
                s = dict(s)
                s["text"] = span_text(s)
                if s["text"]:
                    spans.append(s)
            if not "".join(s["text"] for s in spans).strip():
                continue
            raw.append({"bbox": pymupdf.Rect(ln["bbox"]), "spans": spans,
                        "dir": ln["dir"], "block": b["number"], "box": pymupdf.Rect(b["bbox"])})
    # justified text arrives as several "lines" on one baseline: join them
    raw.sort(key=lambda l: (l["block"], round(l["bbox"].y1), l["bbox"].x0))
    lines = []
    for ln in raw:
        prev = lines[-1] if lines else None
        if (prev and prev["block"] == ln["block"] and prev["dir"] == ln["dir"] == (1.0, 0.0)
                and abs(prev["bbox"].y1 - ln["bbox"].y1) < 2.5
                and -2 < ln["bbox"].x0 - prev["bbox"].x1 < 60):
            if not prev["spans"][-1]["text"].endswith(" ") and not ln["spans"][0]["text"].startswith(" "):
                sp = dict(ln["spans"][0]); sp["text"] = " " + sp["text"]
                ln["spans"][0] = sp
            prev["spans"] += ln["spans"]
            prev["bbox"] |= ln["bbox"]
            continue
        lines.append(ln)
    for ln in lines:
        ln["size"] = max(s["size"] for s in ln["spans"] if s["text"].strip())
    return lines


# ---------------------------------------------------------------- equations
def _bad_char(ch, first):
    o = ord(ch)
    if 0x0B80 <= o <= 0x0DFF or 0x1D400 <= o <= 0x1D7FF:
        return True                      # maths glyphs with broken unicode mapping
    return 0xF000 <= o <= 0xF0FF and not first   # symbol-font glyph inside a line


def equation_zones(page, lines, tmpl, figs):
    """Areas holding equations (broken maths text, small inline images, glyph
    drawings). Each is kept as one picture together with the text on its row."""
    seeds = []
    for ln in lines:
        first = True
        for s in ln["spans"]:
            for ch in s["text"]:
                if ch.strip():
                    if _bad_char(ch, first):
                        seeds.append(pymupdf.Rect(s["bbox"]))
                        break
                    first = False
    for info in page.get_image_info(xrefs=True):
        r = pymupdf.Rect(info["bbox"])
        if info["xref"] not in tmpl and r.height < 60 and r.width < 260:
            seeds.append(r)
    thin = []
    for d in page.get_drawings():
        r = pymupdf.Rect(d["rect"])
        if r.width * max(r.height, 1) >= 2500:
            continue
        (seeds if min(r.width, r.height) >= 3 else thin).append(r)
    def in_fig(r):
        return any((f + (-3, -3, 3, 3)).contains(r) for f in figs)
    seeds = [r for r in seeds if not in_fig(r)]
    thin = [r for r in thin if not in_fig(r)]
    if not seeds:
        return []
    zones = cluster_rects(seeds, gap=12)
    changed = True
    while changed:
        changed = False
        for i, z in enumerate(zones):
            grown = pymupdf.Rect(z)
            for ln in lines:
                b = ln["bbox"]
                ov = min(b.y1, grown.y1) - max(b.y0, grown.y0)
                xgap = max(b.x0 - grown.x1, grown.x0 - b.x1, 0)
                if (ov > 0.4 * min(b.height, grown.height) and xgap < 70 and b.width < 500
                        and not in_fig(b)):
                    grown |= b
            for r in thin:
                if (grown + (-4, -4, 4, 4)).intersects(r):
                    grown |= r
            if grown != z:
                zones[i] = grown
                changed = True
        merged = cluster_rects(zones, gap=3)
        if len(merged) != len(zones):
            zones = merged
            changed = True
    return zones


# ---------------------------------------------------------------- figures
def render_figure(work_page, rect):
    pad = 2
    clip = (rect + (-pad, -pad, pad, pad)) & work_page.rect
    pix = work_page.get_pixmap(clip=clip, dpi=FIG_DPI, alpha=False)
    return pix.tobytes("png"), clip


# ---------------------------------------------------------------- docx utils
def set_cell_border(table, show):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single" if show else "nil")
        if show:
            el.set(qn("w:sz"), "4")
            el.set(qn("w:color"), "808080")
        borders.append(el)
    tblPr.append(borders)


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hex_fill)
    tcPr.append(sh)


def color_of(span):
    c = span["color"]
    return RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)


def add_runs(par, spans, base_line_size):
    for s in spans:
        txt = s["text"]
        if not txt:
            continue
        run = par.add_run(txt)
        f = run.font
        flags = s["flags"]
        sup = bool(flags & 1) or (s["size"] < 0.75 * base_line_size and
                                   s["origin"][1] < s.get("_line_base", s["origin"][1]) - 1)
        f.bold = bool(flags & 16) or "Bold" in s["font"]
        f.italic = bool(flags & 2) or "Italic" in s["font"] or "Oblique" in s["font"]
        size = s["size"] if not sup else base_line_size
        f.size = Pt(max(MIN_PT, min(round(size * FONT_SCALE * 2) / 2, 30)))
        col = color_of(s)
        if col != RGBColor(0, 0, 0):
            f.color.rgb = col
        if s.get("_script") == "sup":
            f.superscript = True
        elif s.get("_script") == "sub":
            f.subscript = True
        if "Symbol" in s["font"]:
            f.name = "Symbol"


def mark_scripts(line):
    """Tag spans that are raised/lowered relative to the line's main text."""
    spans = line["spans"]
    main = max(spans, key=lambda s: (s["size"], len(s["text"].strip())))
    base_y, base_size = main["origin"][1], main["size"]
    for s in spans:
        if s["size"] < 0.85 * base_size and s["text"].strip():
            dy = s["origin"][1] - base_y
            if dy < -0.15 * base_size:
                s["_script"] = "sup"
            elif dy > 0.08 * base_size:
                s["_script"] = "sub"


# ---------------------------------------------------------------- paragraphs
def build_paragraphs(lines):
    """Group lines into paragraphs: same block, same indent, no bullet start."""
    paras = []
    for ln in sorted(lines, key=lambda l: (l["block"], round(l["bbox"].y0 / 3), l["bbox"].x0)):
        text = "".join(s["text"] for s in ln["spans"]).strip()
        starts_bullet = text[:1] in BULLETS or bool(re.match(r"\(?\d{1,2}[.)]\s", text))
        prev = paras[-1] if paras else None
        if prev is not None:
            pl = prev["lines"][-1]
            same_col = abs(ln["bbox"].x0 - pl["bbox"].x0) < 4 or (
                prev["bullet"] and ln["bbox"].x0 > prev["lines"][0]["bbox"].x0 + 2
                and abs(ln["bbox"].x0 - pl["bbox"].x0) < 4) or (
                prev["bullet"] and len(prev["lines"]) == 1 and abs(ln["bbox"].x0 - prev["text_x"]) < 4)
            gap = ln["bbox"].y0 - pl["bbox"].y1
            if (not starts_bullet and same_col and abs(ln["size"] - pl["size"]) < 1.5
                    and -2 < gap < 0.6 * ln["size"] and ln["block"] == pl["block"]):
                prev["lines"].append(ln)
                continue
        text_x = ln["bbox"].x0
        if starts_bullet and len(ln["spans"]) > 1:
            text_x = ln["spans"][1]["bbox"][0]
        paras.append({"lines": [ln], "bullet": starts_bullet, "text_x": text_x})
    for p in paras:
        p["bbox"] = pymupdf.Rect(p["lines"][0]["bbox"])
        for l in p["lines"]:
            p["bbox"] |= l["bbox"]
        p["size"] = max(l["size"] for l in p["lines"])
    return paras


def write_paragraph(container, para, base_x, title=False):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(2)
    pf.space_before = Pt(0)
    indent = max(0, para["lines"][0]["bbox"].x0 - base_x)
    if not title and indent > 3:
        pf.left_indent = Pt(min(indent * 0.6, 120))
    if para["bullet"]:
        pf.left_indent = Pt((pf.left_indent.pt if pf.left_indent else 0) + 12)
        pf.first_line_indent = Pt(-12)
    centred = [l for l in para["lines"] if l["dir"] != (1.0, 0.0)]
    for i, ln in enumerate(para["lines"]):
        mark_scripts(ln)
        spans = ln["spans"]
        if i > 0:
            prev_txt = "".join(s["text"] for s in para["lines"][i - 1]["spans"])
            if not prev_txt.endswith((" ", "-")) and not spans[0]["text"].startswith(" "):
                p.add_run(" ")
        add_runs(p, spans, ln["size"])
    if title:
        for r in p.runs:
            r.font.size = Pt(max(14, min(20, para["size"] * 0.55)))
            r.font.bold = True
    return p


# ---------------------------------------------------------------- tables
def write_table(container, spec, width):
    rows = spec["rows"]
    ncol = max(len(r) for r in rows)
    t = container.add_table(rows=len(rows), cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_border(t, True)
    hdr = spec.get("header_rows", 1)
    fs = Pt(spec.get("font_pt", 8))
    for i, row in enumerate(rows):
        for j in range(ncol):
            cell = t.cell(i, j)
            txt = row[j] if j < len(row) else ""
            cell.text = ""
            par = cell.paragraphs[0]
            par.paragraph_format.space_after = Pt(0)
            write_rich(par, txt, fs, bold=i < hdr)
            if i < hdr:
                shade(cell, "DCE6F2")
    if spec.get("caption"):
        cap = container.add_paragraph()
        write_rich(cap, spec["caption"], Pt(9), bold=True)
        # caption above the table
        t._tbl.addprevious(cap._p)
    if spec.get("note"):
        n = container.add_paragraph()
        write_rich(n, spec["note"], Pt(8))
    return t


_RICH = re.compile(r"(\^\{[^}]*\}|_\{[^}]*\}|\*\*[^*]+\*\*)")


def write_rich(par, text, size, bold=False):
    """Minimal markup: ^{sup} _{sub} **bold**."""
    for part in _RICH.split(text):
        if not part:
            continue
        r = None
        if part.startswith("^{"):
            r = par.add_run(part[2:-1]); r.font.superscript = True
        elif part.startswith("_{"):
            r = par.add_run(part[2:-1]); r.font.subscript = True
        elif part.startswith("**"):
            r = par.add_run(part[2:-2]); r.font.bold = True
        else:
            r = par.add_run(part)
        r.font.size = size
        if bold:
            r.font.bold = True


# ---------------------------------------------------------------- main
def convert(pdf_path, out_path, table_specs, force_picture=()):
    src = pymupdf.open(pdf_path)
    work = pymupdf.open(pdf_path)       # copy used for rendering figures
    tmpl = template_xrefs(src)
    lecture = os.path.splitext(os.path.basename(pdf_path))[0]

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = PAGE_W, PAGE_H
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, MARGIN)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
    st.font.size = Pt(11)
    hp = sec.header.paragraphs[0]
    hp.text = lecture
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in hp.runs:
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    report = []
    for page in src:
        pno = page.number + 1
        wpage = work[page.number]
        lines = slide_lines(page)
        regs = figure_regions(page, tmpl, lines)
        zones = equation_zones(page, lines, tmpl, [g[0] for g in regs])
        regs = [g for g in regs
                if not any((z + (-2, -2, 2, 2)).contains(g[0]) for z in zones)]
        for z in zones:
            host = next((g for g in regs if (g[0] + (-8, -8, 8, 8)).intersects(z)), None)
            if host:
                host[0] |= z              # equation drawn inside a diagram
            else:
                regs.append([z, "equation", None])
        if pno in force_picture:
            # fallback for slides too intricate to rebuild: keep the slide body as
            # one picture and only the title as text
            body = [ln for ln in lines if ln["size"] < TITLE_MIN_SIZE or ln["bbox"].y0 > 0.3 * page.rect.height]
            u = None
            for r in [g[0] for g in regs] + [ln["bbox"] for ln in body]:
                u = pymupdf.Rect(r) if u is None else u | r
            regs = [[u, "drawing", None]] if u else []
        specs = table_specs.get(str(pno), {})

        # decide which text is editable and which belongs to a figure
        editable, labels = [], []
        for ln in lines:
            c = (ln["bbox"].tl + ln["bbox"].br) / 2
            inside = [r for r in regs if (r[0] + (-3, -3, 3, 3)).contains(c)]
            is_label = False
            for r, kind, _ in inside:
                if kind != "image" or ln["size"] < 16 or ln["dir"] != (1.0, 0.0):
                    is_label = True
            (labels if is_label else editable).append(ln)

        # remove editable text and template images from the render copy
        for ln in editable:
            wpage.add_redact_annot(ln["bbox"])
        wpage.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                               graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
        for x in tmpl:
            try:
                wpage.delete_image(x)
            except Exception:
                pass

        paras = build_paragraphs(editable)
        title = None
        big = [p for p in paras if p["size"] >= TITLE_MIN_SIZE and not p["bullet"]
               and p["bbox"].y0 < 0.3 * page.rect.height]
        if big:
            top = max(p["size"] for p in big)
            tparas = [p for p in big if p["size"] >= top - 1]
            tparas.sort(key=lambda p: p["bbox"].y0)
            title = tparas[0]
            paras.remove(title)

        # page heading
        hd = doc.add_paragraph()
        hd.paragraph_format.space_after = Pt(4)
        tag = hd.add_run(f"Slide {pno}   ")
        tag.font.size = Pt(9)
        tag.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        if title:
            p = write_paragraph(doc, title, title["bbox"].x0, title=True)
            # move title runs onto the heading line
            for r in list(p.runs):
                hd._p.append(r._r)
            p._p.getparent().remove(p._p)
        hd.style = doc.styles["Heading 2"]
        hd.paragraph_format.page_break_before = page.number > 0
        hd.paragraph_format.space_before = Pt(0)

        items = [("text", p, p["bbox"]) for p in paras]
        for r, kind, xref in regs:
            if xref is not None and str(xref) in specs:
                items.append(("table", specs[str(xref)], r))
            elif kind == "equation":
                items.append(("equation", None, r))
            else:
                items.append(("figure", None, r))
        base_x = min((p["bbox"].x0 for p in paras), default=0)

        # two-column layout when the slide content sits in two side-by-side columns
        two_col, split = find_split(items)
        if two_col:
            left = [it for it in items if (it[2].x0 + it[2].x1) / 2 < split]
            right = [it for it in items if (it[2].x0 + it[2].x1) / 2 >= split]
            lw = (split - page.rect.x0) / page.rect.width
            lw = min(max(lw, 0.3), 0.7)
            t = doc.add_table(rows=1, cols=2)
            set_cell_border(t, False)
            t.autofit = False
            widths = [int(CONTENT_W * lw), int(CONTENT_W * (1 - lw))]
            for col, (cell, group) in enumerate(zip(t.rows[0].cells, (left, right))):
                cell.width = widths[col]
                cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
                emit(cell, group, wpage, widths[col], base_x, report, pno)
                if not cell.paragraphs and not cell.tables:
                    cell.add_paragraph()
        else:
            emit(doc, items, wpage, CONTENT_W, base_x, report, pno)

        report.append((pno, len(editable), len(labels), len(regs)))

    doc.save(out_path)
    return report


BODY_H_PT = (PAGE_H - 2 * MARGIN) / 12700 - 40   # room left under the slide heading
FIG_SCALE = 0.85                          # slide points -> document points


def text_height(para):
    return sum(l["size"] for l in para["lines"]) * FONT_SCALE * 1.3 + 3


def find_split(items):
    """An x position dividing the items into a left and a right column, or None."""
    if len(items) < 2:
        return False, None
    best = None
    for cand in sorted({it[2].x1 for it in items}):
        left = [it for it in items if it[2].x1 <= cand + 25]
        right = [it for it in items if it[2].x0 >= cand - 25]
        straddle = len(items) - len(left) - len(right)
        if not left or not right or straddle > 0.15 * len(items):
            continue
        ly = (min(i[2].y0 for i in left), max(i[2].y1 for i in left))
        ry = (min(i[2].y0 for i in right), max(i[2].y1 for i in right))
        if min(ly[1], ry[1]) - max(ly[0], ry[0]) < 40:
            continue                      # columns that do not sit side by side
        if best is None or straddle < best[0]:
            best = (straddle, cand + 2)
    return (True, best[1]) if best else (False, None)


def emit(container, items, wpage, width, base_x, report, pno):
    items = sorted(items, key=lambda it: (round(it[2].y0 / 8), it[2].x0))
    used = sum(text_height(o) for k, o, _ in items if k == "text")
    used += sum(len(o["rows"]) * 11 + 14 for k, o, _ in items if k == "table")
    used += sum(r.height * FIG_SCALE + 4 for k, _, r in items if k == "equation")
    figs = [r for k, _, r in items if k == "figure"]
    # figures that sit side by side share a row; count heights per row
    rows = cluster_rows(figs)
    fig_h = sum(max(r.height for r in row) for row in rows) * FIG_SCALE
    avail = max(BODY_H_PT - used, 120)
    shrink = min(1.0, avail / fig_h) if fig_h else 1.0
    # each row of figures is placed as one unit, in reading order with the text
    units = [(r[0].y0, "row", r) for r in rows]
    units += [(rect.y0, kind, obj if kind != "equation" else rect)
              for kind, obj, rect in items if kind != "figure"]
    units.sort(key=lambda u: u[0])
    width_pt = width / 12700
    for _, kind, obj in units:
        if kind == "text":
            write_paragraph(container, obj, base_x)
        elif kind == "table":
            write_table(container, obj, width)
        elif kind == "equation":
            png, clip = render_figure(wpage, obj)
            w = min(clip.width * FIG_SCALE, width_pt * 0.97)
            p = container.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Pt(max(0, min((obj.x0 - base_x) * 0.6, 120)))
            p.add_run().add_picture(io.BytesIO(png), width=Pt(w))
        else:
            row = sorted(obj, key=lambda r: r.x0)
            pics = [render_figure(wpage, r) for r in row]
            ws = [c.width * FIG_SCALE * shrink for _, c in pics]
            if sum(ws) > width_pt * 0.97:
                k = width_pt * 0.97 / sum(ws)
                ws = [w * k for w in ws]
            if len(row) == 1:
                cells = [container.add_paragraph()]
            else:
                t = container.add_table(rows=1, cols=len(row))
                set_cell_border(t, False)
                t.alignment = WD_TABLE_ALIGNMENT.CENTER
                cells = []
                for cell, w in zip(t.rows[0].cells, ws):
                    cell.width = Pt(w + 6)
                    cells.append(cell.paragraphs[0])
            for p, (png, _), w in zip(cells, pics, ws):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(2)
                p.add_run().add_picture(io.BytesIO(png), width=Pt(w))


def cluster_rows(rects):
    rows = []
    for r in sorted(rects, key=lambda r: (r.y0, r.x0)):
        for row in rows:
            if any(min(r.y1, q.y1) - max(r.y0, q.y0) > 0.5 * min(r.height, q.height) for q in row):
                row.append(r)
                break
        else:
            rows.append([r])
    return rows


def overflowing_slides(docx_path):
    """Slides whose Word page spills onto a second page (checked via LibreOffice)."""
    import subprocess, tempfile, shutil
    d = tempfile.mkdtemp()
    shutil.copy(docx_path, os.path.join(d, "x.docx"))
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "x.docx"],
                   cwd=d, capture_output=True)
    out, cur = set(), 0
    for p in pymupdf.open(os.path.join(d, "x.pdf")):
        m = re.search(r"^Slide (\d+)", p.get_text(), re.M)
        if m:
            cur = int(m.group(1))
        else:
            out.add(cur)
    shutil.rmtree(d)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out_dir")
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("--tables", help="JSON: {lecture: {page: {xref: table}}}")
    a = ap.parse_args()
    specs = json.load(open(a.tables)) if a.tables else {}
    os.makedirs(a.out_dir, exist_ok=True)
    for pdf in a.pdfs:
        name = os.path.splitext(os.path.basename(pdf))[0]
        out = os.path.join(a.out_dir, name + ".docx")
        rep = convert(pdf, out, specs.get(name, {}))
        over = overflowing_slides(out)
        if over:
            rep = convert(pdf, out, specs.get(name, {}), force_picture=over)
            print(name, "kept as one picture:", sorted(over), "still over:", sorted(overflowing_slides(out)))
        print(name, "slides:", len(rep), "editable lines:", sum(r[1] for r in rep),
              "figures:", sum(r[3] for r in rep))


if __name__ == "__main__":
    main()
