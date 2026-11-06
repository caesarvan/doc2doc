"""Compact Markdown ("pptx2md-mini v1") -> .pptx.

  python md2ppt.py prompts deck.md                 # JSON lines: every picture the agent must create
  python md2ppt.py check   deck.md --images DIR    # which pictures are still missing
  python md2ppt.py build   deck.md out.pptx --images DIR [--draft]
  python md2ppt.py preview out.pptx OUTDIR         # slide PNGs via LibreOffice, for visual checks

Pictures: ```chart``` blocks are redrawn from their data (specs.py, or native PowerPoint charts);
```prompt``` blocks must exist as DIR/<id>.png|jpg|jpeg|webp|svg, made by the agent itself.
--draft puts a grey box where a picture is missing (layout preview only; never ship it).
"""
import argparse
import io
import json
import re
import sys
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

sys.path.insert(0, str(Path(__file__).parent))
import specs  # noqa: E402

CM = 360000
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
ELEM = re.compile(r"^<!-- @(-?[\d.]+),(-?[\d.]+) ([\d.]+)x([\d.]+) (\S+)(.*?) -->$")
FENCE = re.compile(r"^```(prompt|chart) (.*)$")
SPAN = re.compile(r"\[([^\]]*)\]\{([^}]*)\}")
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".webp", ".svg")


def parse_attrs(text):
    out = {}
    for k, v, flag in re.findall(r'(\w+)=("[^"]*"|\S+)|(\w+)', text or ""):
        if flag:
            out[flag] = "1"
        else:
            out[k] = v.strip('"')
    return out


def parse_md(path):
    """-> deck, sections {"master": sec, "layout:<name>": sec}, slides [sec], assets {id: (kind, attrs, body)}
    where sec = {"attrs": {...}, "bg": hex|None, "elems": [...], "notes": [...]}."""
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    deck, sections, slides, assets = {}, {}, [], {}
    sec, elem, in_notes, i = None, None, False, 0
    while i < len(lines):
        line = lines[i]
        m = FENCE.match(line)
        if line.startswith("<!-- deck "):
            deck = parse_attrs(line[len("<!-- deck "):-3])
            w, h = line.split()[2].split("x")
            deck["w"], deck["h"] = float(w), float(h)
        elif line.startswith("## "):
            head, _, comment = line[3:].partition("<!--")
            sec = {"attrs": parse_attrs(comment.replace("-->", "")), "bg": None, "elems": [], "notes": []}
            name = head.strip()
            if name == "母版":
                sections["master"] = sec
            elif name.startswith("版式"):
                sections["layout:" + name[2:].strip()] = sec
            else:
                slides.append(sec)
            elem, in_notes = None, False
        elif m:
            attrs = parse_attrs(m.group(2))
            body = []
            i += 1
            while i < len(lines) and lines[i] != "```":
                body.append(lines[i])
                i += 1
            assets[attrs["id"]] = (m.group(1), attrs, "\n".join(body))
        elif sec is not None and line.startswith("<!-- bg="):
            sec["bg"] = line[len("<!-- bg="):-3].strip()
        elif sec is not None and line.startswith("<!-- notes -->"):
            elem, in_notes = None, True
        elif sec is not None and (em := ELEM.match(line)):
            x, y, w, h = (float(v) for v in em.group(1, 2, 3, 4))
            elem = {"box": (x, y, w, h), "kind": em.group(5), "attrs": parse_attrs(em.group(6)), "text": []}
            sec["elems"].append(elem)
            in_notes = False
        elif in_notes and sec is not None:
            sec["notes"].append(line)
        elif elem is not None and line.strip():
            elem["text"].append(line)
        i += 1
    return deck, sections, slides, assets


def all_elems(sections, slides):
    for sec in list(sections.values()) + slides:
        yield from sec["elems"]


# ---------------------------------------------------------------- pictures

def resolve(ident, assets):
    seen = set()
    while ident in assets and assets[ident][1].get("same") and ident not in seen:
        seen.add(ident)
        ident = assets[ident][1]["same"]
    return ident


def image_file(ident, images_dir):
    if not images_dir:
        return None
    for ext in IMAGE_EXT:
        f = Path(images_dir) / f"{ident}{ext}"
        if f.exists():
            return f
    return None


def svg_to_png(svg_path, width_px):
    import pymupdf as fitz
    doc = fitz.open(str(svg_path))
    page = doc[0]
    k = width_px / max(page.rect.width, 1)
    return page.get_pixmap(matrix=fitz.Matrix(k, k), alpha=True).tobytes("png")


def draft_box(text, aspect):
    w = 900
    h = max(60, round(w / max(aspect, 0.05)))
    img = Image.new("RGB", (w, h), (242, 242, 242))
    d = ImageDraw.Draw(img)
    d.rectangle([2, 2, w - 3, h - 3], outline=(160, 160, 160), width=3)
    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except OSError:
        font = ImageFont.load_default()
    d.text((14, 10), text[:300], fill=(90, 90, 90), font=font)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def picture(ident, assets, images_dir, box, draft, missing):
    """-> (png/jpg bytes or None, 'chart'|'native'|'image')"""
    rid = resolve(ident, assets)
    kind, attrs, body = assets.get(rid, ("prompt", {}, ""))
    if kind == "chart":
        spec = yaml.safe_load(body)
        if spec.get("renderer") == "pptx-chart":
            return spec, "native"
        buf = io.BytesIO()
        specs.render(spec, tuple(spec["px"])).save(buf, "PNG")
        return buf.getvalue(), "chart"
    f = image_file(rid, images_dir) or image_file(ident, images_dir)
    if f is None:
        missing.add(rid)
        return (draft_box(f"[missing {rid}] {body}", box[2] / max(box[3], 0.01)) if draft else None), "image"
    if f.suffix.lower() == ".svg":
        return svg_to_png(f, max(600, round(box[2] / 2.54 * 220))), "image"
    if f.suffix.lower() == ".webp":
        buf = io.BytesIO()
        Image.open(f).save(buf, "PNG")
        return buf.getvalue(), "image"
    return f.read_bytes(), "image"


# ---------------------------------------------------------------- shapes & text

def rgb(hexv):
    return RGBColor.from_string(hexv)


def emu(*vals):
    return [Emu(round(v * CM)) for v in vals]


def apply_line(shape, a):
    if a.get("line") == "none":
        shape.line.fill.background()
    elif a.get("line"):
        shape.line.color.rgb = rgb(a["line"])
    if a.get("lw"):
        shape.line.width = Pt(float(a["lw"]))
    ln = shape.line._get_or_add_ln()
    if a.get("dash"):
        etree.SubElement(ln, qn("a:prstDash")).set("val", a["dash"])
    for key, tag in (("head", "a:headEnd"), ("tail", "a:tailEnd")):
        if a.get(key):
            etree.SubElement(ln, qn(tag)).set("type", a[key])


def set_font(run, st):
    f = run.font
    f.name, f.size = st["f"], Pt(st["sz"])
    f.bold, f.italic, f.underline = st["b"], st["i"], st["u"]
    f.color.rgb = rgb(st["c"])
    if st.get("fe"):
        rpr = run._r.get_or_add_rPr()
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", st["fe"])


def base_style(a, deck):
    return {"f": a.get("f", deck.get("f", "Calibri")), "fe": a.get("fe", deck.get("fe")),
            "sz": float(a.get("sz", deck.get("sz", 18))), "c": a.get("c", "000000"),
            "b": a.get("b") == "1", "i": a.get("i") == "1", "u": a.get("u") == "1"}


def fill_paragraph(p, line, base):
    pos, chunks = 0, []
    for m in SPAN.finditer(line):
        chunks.append((line[pos:m.start()], {}))
        chunks.append((m.group(1), parse_attrs(m.group(2))))
        pos = m.end()
    chunks.append((line[pos:], {}))
    for text, over in chunks:
        if not text:
            continue
        st = dict(base)
        for k, v in over.items():
            st[k] = float(v) if k == "sz" else (v != "0" if k in ("b", "i", "u") else v)
        for n, piece in enumerate(text.split("\\n")):
            if n:
                p._p.add_br()
            if piece:
                r = p.add_run()
                r.text = piece
                set_font(r, st)


ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}


def set_text(tf, lines, a, deck):
    base = base_style(a, deck)
    tf.word_wrap = a.get("wrap") != "0"
    tf.auto_size = None
    if a.get("va"):
        tf.vertical_anchor = {"m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM, "t": MSO_ANCHOR.TOP}[a["va"]]
    if a.get("ins"):
        l, t, r, b = (Emu(round(float(v) * CM)) for v in a["ins"].split(","))
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = l, t, r, b
    for n, line in enumerate(lines):
        p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
        pa = {}
        if (m := re.search(r" \{([^{}]*)\}$", line)) and "=" in m.group(1):
            pa, line = parse_attrs(m.group(1)), line[: m.start()]
        p.alignment = ALIGN[pa.get("al", a.get("al", "l"))]
        lvl = int(pa.get("lvl", 0))
        if lvl:
            p.level = min(lvl, 8)
        ppr = p._p.get_or_add_pPr()
        if pa.get("bu"):
            ppr.set("marL", str(round((lvl + 1) * 0.5 * CM)))
            ppr.set("indent", str(round(-0.5 * CM)))
            if pa["bu"] == "1.":
                etree.SubElement(ppr, qn("a:buAutoNum")).set("type", "arabicPeriod")
            else:
                etree.SubElement(ppr, qn("a:buChar")).set("char", pa["bu"])
        else:
            etree.SubElement(ppr, qn("a:buNone"))
        if line != "\\":
            fill_paragraph(p, line, base)


def add_table(slide, e, deck):
    x, y, w, h = e["box"]
    a = e["attrs"]
    rows = []
    for line in e["text"]:
        if re.match(r"^\|(\s*-+\s*\|)+$", line.replace(":", "")):
            continue
        cells = [c.strip().replace("\\|", "|") for c in re.split(r"(?<!\\)\|", line.strip())[1:-1]]
        rows.append(cells)
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    gf = slide.shapes.add_table(len(rows), ncol, *emu(x, y, w, h))
    tbl = gf.table
    if a.get("cols"):
        for col, v in zip(tbl.columns, a["cols"].split(",")):
            col.width = Emu(round(float(v) * CM))
    if a.get("rows"):
        for row, v in zip(tbl.rows, a["rows"].split(",")):
            row.height = Emu(round(float(v) * CM))
    tblpr = gf._element.find(f".//{A}tblPr")
    if a.get("style"):
        sid = tblpr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
        sid.text = a["style"]
    flags = a.get("flags", "").split(",") if "flags" in a else None
    if flags is not None:
        for k in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
            tblpr.set(k, "1" if k in flags else "0")
    base = base_style(a, deck)
    for i, row in enumerate(rows):
        for j in range(ncol):
            raw = row[j] if j < len(row) else ""
            if raw == "^":
                continue
            ca = {}
            if (m := re.search(r" \{([^{}]*)\}$", raw)):
                ca, raw = parse_attrs(m.group(1)), raw[: m.start()]
            cell = tbl.cell(i, j)
            span, rspan = int(ca.get("span", 1)), int(ca.get("rspan", 1))
            if span > 1 or rspan > 1:
                cell.merge(tbl.cell(min(i + rspan, len(rows)) - 1, min(j + span, ncol) - 1))
            if ca.get("fill"):
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(ca["fill"])
            tf = cell.text_frame
            for n, part in enumerate(raw.strip().split("<br>")):
                p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
                if part.strip():
                    fill_paragraph(p, part.strip(), base)


def add_native_chart(slide, e, spec):
    x, y, w, h = e["box"]
    ctype = getattr(XL_CHART_TYPE, spec.get("type", "COLUMN_CLUSTERED"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    data = CategoryChartData()
    cats = spec.get("categories") or [str(i + 1) for i in range(len(spec["series"][0]["values"]))]
    data.categories = cats
    for s in spec["series"]:
        data.add_series(s.get("name") or "", s["values"])
    gf = slide.shapes.add_chart(ctype, *emu(x, y, w, h), data)
    ch = gf.chart
    ch.has_legend = bool(spec.get("legend", True))
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    if spec.get("title"):
        ch.has_title = True
        ch.chart_title.text_frame.text = spec["title"]


def add_element(slide, e, deck, assets, images_dir, slide_no, draft, missing):
    x, y, w, h = e["box"]
    a, kind = e["attrs"], e["kind"]
    if kind in ("img", "video", "chart"):
        data, src = picture(a["id"], assets, images_dir, e["box"], draft, missing)
        if data is None:
            return
        if src == "native":
            add_native_chart(slide, e, data)
            return
        pic = slide.shapes.add_picture(io.BytesIO(data), *emu(x, y, w, h))
        if src == "chart":  # generated pictures already show exactly the visible part
            if a.get("crop"):
                pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = (float(v) for v in a["crop"].split(","))
            if a.get("rot"):
                pic.rotation = float(a["rot"])
        return
    if kind == "table":
        add_table(slide, e, deck)
        return
    if kind == "line":
        x1, y1, x2, y2 = x, y, x + w, y + h
        if a.get("flipH"):
            x1, x2 = x2, x1
        if a.get("flipV"):
            y1, y2 = y2, y1
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *emu(x1, y1, x2, y2))
        apply_line(c, a)
        return
    if kind in ("text", "num"):
        shape = slide.shapes.add_textbox(*emu(x, y, w, h))
    else:
        try:
            prst = MSO_SHAPE.RECTANGLE if kind == "rect" else MSO_SHAPE.from_xml(kind)
        except Exception:
            print(f"WARN slide {slide_no}: unknown shape '{kind}', drawn as rectangle")
            prst = MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(prst, *emu(x, y, w, h))
        shape.shadow.inherit = False
        if a.get("fill") and a["fill"] != "none":
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(a["fill"])
        else:
            shape.fill.background()
        if "line" not in a:
            shape.line.fill.background()
        apply_line(shape, a)
        if a.get("adj"):
            for n, v in enumerate(a["adj"].split(",")):
                try:
                    shape.adjustments[n] = int(v) / 100000
                except (IndexError, ValueError):
                    pass
    if a.get("rot"):
        shape.rotation = float(a["rot"])
    xfrm = shape._element.find(f".//{A}xfrm")
    for k in ("flipH", "flipV"):
        if a.get(k):
            xfrm.set(k, "1")
    lines = [str(slide_no)] if kind == "num" else e["text"]
    if lines:
        set_text(shape.text_frame, lines, a, deck)


def set_bg(slide, colour):
    if colour:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(colour)


def build(md_path, out, images_dir=None, draft=False):
    deck, sections, slides, assets = parse_md(md_path)
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(deck["w"], deck["h"])
    blank = prs.slide_layouts[6]
    master = sections.get("master", {"elems": [], "bg": None, "attrs": {}})
    missing = set()
    for n, sec in enumerate(slides, 1):
        slide = prs.slides.add_slide(blank)
        layout = sections.get("layout:" + sec["attrs"].get("layout", ""), {"elems": [], "bg": None, "attrs": {}})
        set_bg(slide, sec["bg"] or layout["bg"] or master["bg"])
        hide_slide = sec["attrs"].get("nomaster")
        layer = []
        if not hide_slide and not layout["attrs"].get("nomaster"):
            layer += master["elems"]
        if not hide_slide:
            layer += layout["elems"]
        for e in layer + sec["elems"]:
            add_element(slide, e, deck, assets, images_dir, n, draft, missing)
        if sec["notes"]:
            slide.notes_slide.notes_text_frame.text = "\n".join(sec["notes"]).strip()
        if sec["attrs"].get("hidden"):
            slide._element.set("show", "0")
    if missing and not draft:
        raise SystemExit("缺少图片，未生成 PPT：" + " ".join(sorted(missing)) +
                         f"\n请把它们存为 {images_dir or 'images'}/<id>.png|jpg|svg 后重试")
    prs.save(out)
    print(f"{out}: {len(slides)} slides" + (f" (draft, {len(missing)} missing pictures)" if missing else ""))


def prompt_list(md_path):
    deck, sections, slides, assets = parse_md(md_path)
    boxes = {}
    for e in all_elems(sections, slides):
        if e["kind"] in ("img", "video"):
            boxes.setdefault(resolve(e["attrs"]["id"], assets), e["box"])
    out = []
    for ident, (kind, attrs, body) in assets.items():
        if kind == "prompt" and not attrs.get("same"):
            b = boxes.get(ident, (0, 0, 1, 1))
            out.append({"id": ident, "aspect": round(b[2] / max(b[3], 0.01), 3),
                        "box_cm": [b[2], b[3]], "orig_px": attrs.get("px"), "prompt": " ".join(body.split())})
    return out


def check(md_path, images_dir):
    items = prompt_list(md_path)
    have = [i for i in items if image_file(i["id"], images_dir)]
    print(f"{len(have)}/{len(items)} pictures present in {images_dir}")
    for i in items:
        f = image_file(i["id"], images_dir)
        if not f:
            print(f"  MISSING {i['id']} (aspect {i['aspect']})")
        elif f.suffix.lower() != ".svg":
            im = Image.open(f)
            ratio = im.width / im.height
            if abs(ratio - i["aspect"]) / i["aspect"] > 0.08:
                print(f"  {i['id']}: aspect {ratio:.2f} vs box {i['aspect']:.2f} (will be stretched)")


def preview(src, outdir):
    import shutil
    import subprocess
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if not soffice and Path(cand).exists():
            soffice = cand
    if not soffice:
        raise SystemExit("LibreOffice (soffice) not found")
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)], check=True,
                       capture_output=True, timeout=600)
        import pymupdf as fitz
        pdf = fitz.open(next(Path(tmp).glob("*.pdf")))
        for i, page in enumerate(pdf, 1):
            page.get_pixmap(matrix=fitz.Matrix(1600 / page.rect.width, 1600 / page.rect.width)).save(
                str(out / f"slide{i:02d}.png"))
    print(f"{len(pdf)} slides -> {out}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("prompts"); p.add_argument("md")
    p = sub.add_parser("check"); p.add_argument("md"); p.add_argument("--images", default="images")
    p = sub.add_parser("build"); p.add_argument("md"); p.add_argument("out")
    p.add_argument("--images", default="images"); p.add_argument("--draft", action="store_true")
    p = sub.add_parser("preview"); p.add_argument("src"); p.add_argument("outdir")
    a = ap.parse_args()
    if a.cmd == "prompts":
        for item in prompt_list(a.md):
            print(json.dumps(item, ensure_ascii=False))
    elif a.cmd == "check":
        check(a.md, a.images)
    elif a.cmd == "build":
        build(a.md, a.out, a.images, a.draft)
    else:
        preview(a.src, a.outdir)
