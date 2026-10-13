"""Build the bilingual (English/Chinese) lecture notes PDF.

Every original slide is placed on the page unchanged (as vector PDF content);
below it comes an English/Chinese table for the slide text, optional notes,
and each lecture ends with worked examples. Content lives in notes/LNN.txt:

  @lecture N               lecture number (matches "Lecture N.pdf")
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
M = 40                                   # page margin (pt)
SLIDE_W = PAGE[0] - 2 * M

CSS = f"""
@font-face {{ font-family: zh; src: url({FONT_FILE}); }}
* {{ font-family: sans-serif, zh; }}
body {{ font-size: 9.5pt; line-height: 1.35; color: #222; }}
h1 {{ font-size: 22pt; color: #0b6e8a; margin: 0 0 6pt 0; }}
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
    lec = {"slides": {}, "glossary": [], "examples": [], "summary": ""}
    section, cur = None, None
    for raw in open(path, encoding="utf-8"):
        line = raw.rstrip("\n")
        if not line.strip():
            continue
        if line.startswith("@lecture"):
            lec["n"] = int(line.split()[1])
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
           f"<h2>{esc(en)}<br>{esc(zh)}</h2>")
    if lec["summary"]:
        out += f"<h3>本讲概要 Overview</h3><p>{esc(lec['summary'])}</p>"
    if lec["glossary"]:
        rows = "".join(f"<tr><td>{esc(e)}</td><td>{esc(z)}</td></tr>"
                       for e, z in lec["glossary"])
        out += (f"<h3>关键术语 Key Terms</h3><table>{cols(50, 50)}<tr><th>English</th><th>中文</th></tr>"
                f"{rows}</table>")
    return out


def examples_html(lec):
    out = f"<h1>例题讲解 Worked Examples</h1><p class='small'>Lecture {lec['n']}：{esc(lec['title'][1])}</p>"
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


def cover_html(lectures):
    items = "".join(f"<tr><td>Lecture {l['n']}</td><td>{esc(l['title'][0])}</td><td>{esc(l['title'][1])}</td></tr>"
                    for l in lectures)
    return (f"<h1>Properties of Materials</h1><h2>材料性能 —— 中英文对照讲义</h2>"
            f"<p>Dr. Alkis Tsamis · CTGU/UoL JEP, September 2024</p>"
            f"<p class='small'>说明：每页上方为原课件幻灯片（原样保留、未作修改），下方为逐句英汉对照、"
            f"【讲解】要点；每讲开头附关键术语表，结尾附例题讲解。</p>"
            f"<h3>目录 Contents</h3><table>{cols(16, 46, 38)}<tr><th>Lecture</th><th>English</th><th>中文</th></tr>{items}</table>")


# ---------------------------------------------------------------- layout
class Builder:
    def __init__(self, out_path):
        self.out_path = out_path
        self.tmp = out_path + ".tmp.pdf"
        self.writer = pymupdf.DocumentWriter(self.tmp)
        self.archive = pymupdf.Archive(FONT_DIR)
        self.stamps = []        # (page index, src pdf, slide no, rect)
        self.toc = []           # [level, title, page]
        self.page = 0

    def flow(self, html_text, header=""):
        """Place HTML on new pages, continuing until it is all placed."""
        story = pymupdf.Story(html=html_text, user_css=CSS, archive=self.archive)
        full = pymupdf.Rect(M, M + 14, PAGE[0] - M, PAGE[1] - M)
        more, first = True, True
        while more:
            self.begin(header if first else header + "（续）")
            more, _ = story.place(full)
            story.draw(self.dev)
            self.end()
            first = False

    def begin(self, header):
        self.dev = self.writer.begin_page(pymupdf.Rect(0, 0, *PAGE))
        self.header = header
        self.page += 1

    def end(self):
        self.writer.end_page()
        self.headers = getattr(self, "headers", {})
        self.headers[self.page - 1] = self.header

    def slide_page(self, lec, k, slide, src):
        sr = src[k - 1].rect
        h = SLIDE_W * sr.height / sr.width
        box = pymupdf.Rect(M, M + 14, M + SLIDE_W, M + 14 + h)
        header = f"Lecture {lec['n']} · Slide {k}"
        self.begin(header)
        self.stamps.append((self.page - 1, src, k - 1, box))
        body = slide_html(lec, k, slide)
        rect = pymupdf.Rect(M, box.y1 + 8, PAGE[0] - M, PAGE[1] - M)
        if not body:
            self.end()
            return
        story = pymupdf.Story(html=body, user_css=CSS, archive=self.archive)
        more, _ = story.place(rect)
        story.draw(self.dev)
        self.end()
        while more:                                    # continuation pages
            self.begin(header + "（续）")
            more, _ = story.place(pymupdf.Rect(M, M + 14, PAGE[0] - M, PAGE[1] - M))
            story.draw(self.dev)
            self.end()

    def finish(self):
        self.writer.close()
        doc = pymupdf.open(self.tmp)
        for pno, src, sidx, box in self.stamps:
            p = doc[pno]
            p.show_pdf_page(box, src, sidx)
            p.draw_rect(box, color=(0.6, 0.6, 0.6), width=0.5)
        font = pymupdf.Font(fontfile=os.path.join(FONT_DIR, FONT_FILE))
        for i, p in enumerate(doc):
            tw = pymupdf.TextWriter(p.rect)
            hdr = self.headers.get(i, "")
            if hdr:
                tw.append((M, M + 2), hdr, font=font, fontsize=8)
            num = f"{i + 1}"
            tw.append((PAGE[0] / 2 - font.text_length(num, 8) / 2, PAGE[1] - M / 2), num,
                      font=font, fontsize=8)
            tw.write_text(p, color=(0.45, 0.45, 0.45))
        doc.set_toc(self.toc)
        doc.subset_fonts()
        doc.save(self.out_path, garbage=3, deflate=True)
        os.remove(self.tmp)


def main():
    out, paths = sys.argv[1], sys.argv[2:]
    lectures = sorted((parse(p) for p in paths), key=lambda l: l["n"])
    b = Builder(out)
    b.flow(cover_html(lectures), header="")
    b.toc.append([1, "目录 Contents", 1])
    for lec in lectures:
        src = pymupdf.open(os.path.join(ROOT, f"Lecture {lec['n']}.pdf"))
        b.toc.append([1, f"Lecture {lec['n']}  {lec['title'][1]}", b.page + 1])
        b.flow(lecture_intro_html(lec), header=f"Lecture {lec['n']} · 概要与术语")
        missing = [k for k in range(1, len(src) + 1) if k not in lec["slides"]]
        if missing:
            print(f"Lecture {lec['n']}: no notes for slides {missing}", file=sys.stderr)
        for k in range(1, len(src) + 1):
            if k == 1 or k % 5 == 0:
                b.toc.append([2, f"Slide {k}", b.page + 1])
            b.slide_page(lec, k, lec["slides"].get(k, {"rows": [], "notes": []}), src)
        if lec["examples"]:
            b.toc.append([2, "例题讲解 Worked Examples", b.page + 1])
            b.flow(examples_html(lec), header=f"Lecture {lec['n']} · 例题讲解")
    b.finish()
    print("pages:", b.page)


if __name__ == "__main__":
    main()
