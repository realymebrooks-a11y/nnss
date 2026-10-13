"""Build the bilingual (English/Chinese) lecture notes PDF.

Every original slide is placed on the page unchanged (as vector PDF content);
below it comes an English/Chinese table for the slide text, optional notes,
and each lecture ends with worked examples. Content lives in notes/LNN.txt:

  @lecture N               lecture number (matches "Lecture N.pdf")
  @skip 1-7,8              slides left out (title, schedule, section dividers)
  @title EN ||| ZH         lecture title
  @summary text            Chinese overview of the lecture
  @glossary                followed by "EN ||| ZH" term lines
  @slide K                 followed by "EN ||| ZH" lines and "! note" lines
  @example title           followed by "Q: EN ||| ZH", "S: step", "A: answer"

Inline HTML (<b>, <sub>, <sup>) is allowed in any text.

Usage: python3 build_notes.py OUT.pdf notes/L01.txt [notes/L02.txt ...]
"""
import html
import os
import re
import sys

import pymupdf

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_DIR = "/usr/share/fonts/truetype/wqy"
FONT_FILE = "wqy-zenhei.ttc"
PAGE = pymupdf.paper_size("a4")
M = 30                                   # page margin (pt)
TOP, BOTTOM = M + 4, PAGE[1] - 26        # usable band of each page
GAP = 6                                  # space after each block
SLIDE_W = 370                            # width of an embedded slide (pt)

CSS = f"""
@font-face {{ font-family: zh; src: url({FONT_FILE}); }}
* {{ font-family: sans-serif, zh; }}
body {{ font-size: 9.5pt; line-height: 1.35; color: #222; }}
h1 {{ font-size: 18pt; color: #0b6e8a; margin: 0 0 6pt 0; }}
h2 {{ font-size: 15pt; color: #0b6e8a; margin: 10pt 0 4pt 0; }}
h3 {{ font-size: 11.5pt; color: #b3372e; margin: 10pt 0 3pt 0; }}
p {{ margin: 2pt 0; }}
table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
td, th {{ border: 0.5pt solid #9aa; padding: 2pt 4pt; vertical-align: top; }}
th {{ background-color: #dfecef; font-size: 9pt; }}
.note {{ background-color: #fff6dc; border: 0.5pt solid #e0c36a; padding: 3pt 5pt; margin-top: 4pt; }}
.tag {{ color: #b3372e; font-weight: bold; }}
.q {{ background-color: #eef5f7; padding: 3pt 5pt; margin: 2pt 0; }}
.ans {{ background-color: #eaf6e6; padding: 3pt 5pt; margin: 3pt 0 8pt 0; }}
.small {{ font-size: 8.5pt; color: #666; }}
"""


# ---------------------------------------------------------------- parsing
def parse(path):
    lec = {"slides": {}, "glossary": [], "examples": [], "summary": "", "skip": set()}
    section, cur = None, None
    for raw in open(path, encoding="utf-8"):
        line = raw.rstrip("\n")
        if not line.strip():
            continue
        if line.startswith("@lecture"):
            lec["n"] = int(line.split()[1])
        elif line.startswith("@skip"):
            for part in line[5:].replace(" ", "").split(","):
                a, _, b = part.partition("-")
                lec["skip"].update(range(int(a), int(b or a) + 1))
        elif line.startswith("@title"):
            en, zh = pair(line[6:])
            lec["title"] = (en, zh)
        elif line.startswith("@summary"):
            lec["summary"] = line[8:].strip()
        elif line.startswith("@glossary"):
            section = "glossary"
        elif line.startswith("@slide"):
            section = "slide"
            cur = lec["slides"].setdefault(int(line.split()[1]), {"rows": [], "notes": []})
        elif line.startswith("@example"):
            section = "example"
            cur = {"title": line[8:].strip(), "q": [], "steps": [], "answer": []}
            lec["examples"].append(cur)
        elif section == "glossary":
            lec["glossary"].append(pair(line))
        elif section == "slide":
            if line.startswith("!"):
                cur["notes"].append(line[1:].strip())
            else:
                cur["rows"].append(pair(line))
        elif section == "example":
            tag, _, body = line.partition(":")
            body = body.strip()
            if tag == "Q":
                cur["q"].append(pair(body))
            elif tag == "S":
                cur["steps"].append(body)
            elif tag == "A":
                cur["answer"].append(body)
            else:
                raise ValueError(f"{path}: bad example line: {line}")
    return lec


def pair(text):
    if "|||" not in text:
        raise ValueError(f"missing '|||' in: {text}")
    en, zh = text.split("|||", 1)
    return en.strip(), zh.strip()


def cols(*pcts):
    return "<colgroup>" + "".join(f"<col style='width:{p}%'>" for p in pcts) + "</colgroup>"


def esc(text):
    """Escape text but keep the small set of inline tags we allow."""
    t = html.escape(text, quote=False)
    return re.sub(r"&lt;(/?(?:b|i|sub|sup|br))&gt;", r"<\1>", t)


# ---------------------------------------------------------------- html
def slide_html(lec, k, slide):
    rows = "".join(f"<tr><td>{esc(en)}</td><td>{esc(zh)}</td></tr>"
                   for en, zh in slide["rows"])
    out = ""
    if rows:
        out += f"<table>{cols(50, 50)}<tr><th>English 原文</th><th>中文译文</th></tr>{rows}</table>"
    for n in slide["notes"]:
        out += f"<div class='note'><span class='tag'>【讲解】</span>{esc(n)}</div>"
    return out


def lecture_intro_html(lec):
    en, zh = lec["title"]
    out = (f"<h1>Lecture {lec['n']} 第 {lec['n']} 讲</h1>"
           f"<h2 style='margin-top:0'>{esc(en)}　{esc(zh)}</h2>")
    if lec["summary"]:
        out += f"<h3>本讲概要 Overview</h3><p>{esc(lec['summary'])}</p>"
    if lec["glossary"]:
        g = lec["glossary"]
        half = (len(g) + 1) // 2
        left, right = g[:half], g[half:] + [("", "")] * (half - len(g[half:]))
        rows = "".join(f"<tr><td>{esc(a)}</td><td>{esc(b)}</td><td>{esc(c)}</td><td>{esc(d)}</td></tr>"
                       for (a, b), (c, d) in zip(left, right))
        out += (f"<h3>关键术语 Key Terms</h3><table>{cols(26, 24, 26, 24)}"
                f"<tr><th>English</th><th>中文</th><th>English</th><th>中文</th></tr>"
                f"{rows}</table>")
    return out


def examples_html(lec):
    out = f"<h2>例题讲解 Worked Examples（Lecture {lec['n']}）</h2>"
    for i, ex in enumerate(lec["examples"], 1):
        out += f"<h3>例题 {i}　{esc(ex['title'])}</h3>"
        for en, zh in ex["q"]:
            out += f"<div class='q'><p><b>Question:</b> {esc(en)}</p><p><b>题目：</b>{esc(zh)}</p></div>"
        out += "<p class='tag'>解答 Solution</p>"
        for s in ex["steps"]:
            out += f"<p>{esc(s)}</p>"
        for a in ex["answer"]:
            out += f"<div class='ans'><b>答案 Answer：</b>{esc(a)}</div>"
    return out


# ---------------------------------------------------------------- layout
class Builder:
    """Flows content continuously: slides, tables and notes follow each other
    on the page instead of each slide starting a new page."""

    def __init__(self, out_path):
        self.out_path = out_path
        self.tmp = out_path + ".tmp.pdf"
        self.writer = pymupdf.DocumentWriter(self.tmp)
        self.archive = pymupdf.Archive(FONT_DIR)
        self.stamps = []        # (page index, src pdf, slide index, rect, label)
        self.toc = []           # [level, title, page]
        self.headers = {}
        self.header = ""
        self.page = 0
        self.dev = None
        self.y = TOP

    def new_page(self):
        if self.dev is not None:
            self.writer.end_page()
        self.dev = self.writer.begin_page(pymupdf.Rect(0, 0, *PAGE))
        self.page += 1
        self.headers[self.page - 1] = self.header
        self.y = TOP

    def ensure(self, h):
        if self.dev is None or self.y + h > BOTTOM:
            self.new_page()

    def html(self, text):
        story = pymupdf.Story(html=text, user_css=CSS, archive=self.archive)
        more = True
        while more:
            self.ensure(30)
            more, filled = story.place(pymupdf.Rect(M, self.y, PAGE[0] - M, BOTTOM))
            story.draw(self.dev)
            if more:
                self.new_page()
            else:
                self.y = pymupdf.Rect(filled).y1 + GAP

    def slide(self, lec, k, slide, src):
        sr = src[k - 1].rect
        h = SLIDE_W * sr.height / sr.width
        self.ensure(h + 12 + (40 if slide["rows"] or slide["notes"] else 0))
        x0 = (PAGE[0] - SLIDE_W) / 2
        box = pymupdf.Rect(x0, self.y + 10, x0 + SLIDE_W, self.y + 10 + h)
        self.stamps.append((self.page - 1, src, k - 1, box, f"Lecture {lec['n']} · Slide {k}"))
        self.y = box.y1 + 5
        body = slide_html(lec, k, slide)
        if body:
            self.html(body)
        self.y += 4

    def finish(self):
        self.writer.end_page()
        self.writer.close()
        doc = pymupdf.open(self.tmp)
        for p in doc:
            p.wrap_contents()          # isolate the writer's coordinate transform
        font = pymupdf.Font(fontfile=os.path.join(FONT_DIR, FONT_FILE))
        for pno, src, sidx, box, label in self.stamps:
            p = doc[pno]
            p.show_pdf_page(box, src, sidx)
            p.draw_rect(box, color=(0.6, 0.6, 0.6), width=0.5)
            tw = pymupdf.TextWriter(p.rect)
            tw.append((box.x0, box.y0 - 2.5), label, font=font, fontsize=7.5)
            tw.write_text(p, color=(0.04, 0.43, 0.54))
        for i, p in enumerate(doc):
            tw = pymupdf.TextWriter(p.rect)
            hdr = self.headers.get(i, "")
            if hdr:
                tw.append((PAGE[0] - M - font.text_length(hdr, 7.5), M - 6), hdr, font=font, fontsize=7.5)
            num = f"{i + 1}"
            tw.append((PAGE[0] / 2 - font.text_length(num, 8) / 2, PAGE[1] - 14), num,
                      font=font, fontsize=8)
            tw.write_text(p, color=(0.45, 0.45, 0.45))
            if hdr:
                p.draw_line((M, M - 3), (PAGE[0] - M, M - 3), color=(0.75, 0.75, 0.75), width=0.4)
        doc.set_toc(self.toc)
        doc.subset_fonts()
        doc.save(self.out_path, garbage=3, deflate=True)
        os.remove(self.tmp)


def main():
    out, paths = sys.argv[1], sys.argv[2:]
    lectures = sorted((parse(p) for p in paths), key=lambda l: l["n"])
    b = Builder(out)
    for lec in lectures:
        src = pymupdf.open(os.path.join(ROOT, f"Lecture {lec['n']}.pdf"))
        b.header = f"Lecture {lec['n']}  {lec['title'][0]}  {lec['title'][1]}"
        b.new_page()
        b.toc.append([1, f"Lecture {lec['n']}  {lec['title'][1]}", b.page])
        b.html(lecture_intro_html(lec))
        keep = [k for k in range(1, len(src) + 1) if k not in lec["skip"]]
        missing = [k for k in keep if k not in lec["slides"]]
        if missing:
            print(f"Lecture {lec['n']}: no notes for slides {missing}", file=sys.stderr)
        for i, k in enumerate(keep):
            b.slide(lec, k, lec["slides"].get(k, {"rows": [], "notes": []}), src)
            if i == 0 or k % 5 == 0:
                b.toc.append([2, f"Slide {k}", b.page])
        if lec["examples"]:
            b.ensure(120)
            b.toc.append([2, "例题讲解 Worked Examples", b.page])
            b.html(examples_html(lec))
    b.finish()
    print("pages:", b.page)


if __name__ == "__main__":
    main()
